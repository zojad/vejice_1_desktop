/* global Office, Word */
import { popraviPoved } from "../api/apiVejice.js";

/** ─────────────────────────────────────────────────────────
 *  Helpers: znaki & pravila
 *  ───────────────────────────────────────────────────────── */
const QUOTES = new Set(['"', "'", "“", "”", "„", "«", "»"]);
const isDigit = (ch) => ch >= "0" && ch <= "9";
const charAtSafe = (s, i) => (i >= 0 && i < s.length ? s[i] : "");

/** tisočiški ločilo: digit ',' digit  (npr. 1,000) */
function isThousandsSeparator(original, corrected, kind, pos) {
  const s = kind === "delete" ? original : corrected;
  const prev = charAtSafe(s, pos - 1);
  const next = charAtSafe(s, pos + 1);
  return isDigit(prev) && isDigit(next);
}

/** hitri stražar: ali so se spremenile samo vejice */
function onlyCommasChanged(original, corrected) {
  const strip = (x) => x.replace(/,/g, "");
  return strip(original) === strip(corrected);
}

/** ─────────────────────────────────────────────────────────
 *  Minimalni diff, ki vrne SAMO operacije z vejicami
 *  - insert: vejica obstaja v corrected, ne v original
 *  - delete: vejica obstaja v original, ne v corrected
 *  Pozicije so v ustreznem stringu.
 *  ───────────────────────────────────────────────────────── */
function diffCommasOnly(original, corrected) {
  const ops = [];
  let i = 0,
    j = 0;
  while (i < original.length || j < corrected.length) {
    const o = original[i] ?? "";
    const c = corrected[j] ?? "";

    if (o === c) {
      i++;
      j++;
      continue;
    }
    if (c === "," && o !== ",") {
      ops.push({ kind: "insert", pos: j });
      j++;
      continue;
    }
    if (o === "," && c !== ",") {
      ops.push({ kind: "delete", pos: i });
      i++;
      continue;
    }
    if (o) i++;
    if (c) j++;
  }
  return ops;
}

/** Filtriraj operacije:
 *  - izloči tisočiške ločila
 *  - dovoljuj “vejica brez presledka” samo pred narekovajem (premi govor)
 *    (če API vrne ",ki", bomo dodali presledek naknadno)
 */
function filterCommaOps(original, corrected, ops) {
  return ops.filter((op) => {
    if (isThousandsSeparator(original, corrected, op.kind, op.pos)) return false;

    if (op.kind === "insert") {
      const next = charAtSafe(corrected, op.pos + 1);
      const noSpaceAfter = next && !/\s/.test(next);
      if (noSpaceAfter && !QUOTES.has(next)) {
        // dovolimo (in bomo dodali presledek posebej)
        return true;
      }
    }
    return true;
  });
}

/** ─────────────────────────────────────────────────────────
 *  Anchor-based mikro urejanje (ohrani formatiranje)
 *  ───────────────────────────────────────────────────────── */
function makeAnchor(text, idx, span = 10) {
  const left = text.slice(Math.max(0, idx - span), idx);
  const right = text.slice(idx, Math.min(text.length, idx + span));
  return { left, right };
}

// Vstavi vejico na podlagi sidra (urejanje samo znaka)
async function insertCommaAt(context, paragraph, original, corrected, atCorrectedPos) {
  const { left, right } = makeAnchor(corrected, atCorrectedPos);
  const pr = paragraph.getRange();

  if (left.length > 0) {
    const m = pr.search(left, { matchCase: false, matchWholeWord: false });
    m.load("items");
    await context.sync();
    if (!m.items.length) return;

    const after = m.items[0].getRange("After");
    after.insertText(",", Word.InsertLocation.before);
  } else {
    // Če je na začetku odstavka, uporabi desno sidro
    if (!right) return; // nič za sidrati
    const m = pr.search(right, { matchCase: false, matchWholeWord: false });
    m.load("items");
    await context.sync();
    if (!m.items.length) return;

    const before = m.items[0].getRange("Before");
    before.insertText(",", Word.InsertLocation.before);
  }
}

// Po potrebi dodaj presledek po vejici (razen pred narekovaji)
async function ensureSpaceAfterComma(context, paragraph, corrected, atCorrectedPos) {
  const next = charAtSafe(corrected, atCorrectedPos + 1);
  if (!next || /\s/.test(next) || QUOTES.has(next)) return;

  const { left, right } = makeAnchor(corrected, atCorrectedPos + 1);
  const pr = paragraph.getRange();

  if (left.length > 0) {
    const m = pr.search(left, { matchCase: false, matchWholeWord: false });
    m.load("items");
    await context.sync();
    if (!m.items.length) return;

    const beforeRight = m.items[0].getRange("Before");
    beforeRight.insertText(" ", Word.InsertLocation.before);
  } else if (right.length > 0) {
    const m = pr.search(right, { matchCase: false, matchWholeWord: false });
    m.load("items");
    await context.sync();
    if (!m.items.length) return;

    const before = m.items[0].getRange("Before");
    before.insertText(" ", Word.InsertLocation.before);
  }
}

// Briši samo znak vejice
async function deleteCommaAt(context, paragraph, original, atOriginalPos) {
  const { left, right } = makeAnchor(original, atOriginalPos);
  const pr = paragraph.getRange();

  if (left.length > 0) {
    const lm = pr.search(left, { matchCase: false, matchWholeWord: false });
    lm.load("items");
    await context.sync();
    if (!lm.items.length) return;

    const afterLeft = lm.items[0].getRange("After");
    const comma = afterLeft.search(",", { matchCase: false, matchWholeWord: false });
    comma.load("items");
    await context.sync();
    if (!comma.items.length) return;

    comma.items[0].insertText("", Word.InsertLocation.replace);
  } else {
    // Začetek odstavka: poišči desno sidro, nato najbližjo vejico pred njim
    if (!right) return;
    const rm = pr.search(right, { matchCase: false, matchWholeWord: false });
    rm.load("items");
    await context.sync();
    if (!rm.items.length) return;

    const beforeRight = rm.items[0].getRange("Before");
    const comma = beforeRight.search(",", { matchCase: false, matchWholeWord: false });
    comma.load("items");
    await context.sync();
    if (!comma.items.length) return;

    comma.items[0].insertText("", Word.InsertLocation.replace);
  }
}

/** ─────────────────────────────────────────────────────────
 *  MAIN: Preveri vejice – celoten dokument, po odstavkih
 *  ───────────────────────────────────────────────────────── */
export async function checkDocumentText(event) {
  try {
    await Word.run(async (context) => {
      // Ustvari/ponovno uporabi “cono” na celotnem telesu
      const bodyRange = context.document.body.getRange();
      const boxes = context.document.contentControls;
      boxes.load("items/tag");
      await context.sync();

      let bodyBox = boxes.items.find((c) => (c.tag || "").startsWith("vejice-body-"));
      if (!bodyBox) {
        bodyBox = bodyRange.insertContentControl();
        bodyBox.tag = `vejice-body-${Date.now()}`;
        bodyBox.title = "Vejice – celotno telo";
        bodyBox.appearance = "Hidden";
        await context.sync();
      }

      // Naloži trenutni state Track Changes in ga varno obnovi na koncu
      const doc = context.document;
      doc.load("trackRevisions");
      await context.sync();
      const prevTrack = doc.trackRevisions;
      doc.trackRevisions = true;

      // Odstavki telesa dokumenta
      const paras = context.document.body.paragraphs;
      paras.load("items/text");
      await context.sync();

      for (const p of paras.items) {
        const original = p.text || "";
        if (!original.trim()) continue;

        const corrected = await popraviPoved(original);

        // Če API ni spremenil samo vejic, preskoči (varnost)
        if (!onlyCommasChanged(original, corrected)) continue;

        const ops = filterCommaOps(original, corrected, diffCommasOnly(original, corrected));
        if (!ops.length) continue;

        for (const op of ops) {
          if (op.kind === "insert") {
            await insertCommaAt(context, p, original, corrected, op.pos);
            await ensureSpaceAfterComma(context, p, corrected, op.pos);
          } else {
            await deleteCommaAt(context, p, original, op.pos);
          }
        }

        // majhen sync po odstavku, da je UI odziven
        await context.sync();
      }

      // Povrni stanje Track Changes
      doc.trackRevisions = prevTrack;
      await context.sync();
    });
  } catch (e) {
    console.error("checkDocumentText error:", e);
  } finally {
    if (event && event.completed) event.completed();
  }
}
