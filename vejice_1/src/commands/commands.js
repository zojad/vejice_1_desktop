/* global Office, Word */
import { checkDocumentText as runCheckVejice } from "../logic/preveriVejice.js";

Office.onReady(() => {
  // Office.js is ready
});

/** ─────────────────────────────────────────────────────────
 * Helpers
 * ───────────────────────────────────────────────────────── */
const isCommaish = (t) => {
  const s = (t || "").trim();
  return s === "," || s === "‚" || s === "，"; // extend if you need more variants
};

async function getLatestBodyZone(context) {
  const boxes = context.document.contentControls;
  boxes.load("items/tag");
  await context.sync();

  const candidates = boxes.items
    .filter((c) => (c.tag || "").startsWith("vejice-body-"))
    .sort((a, b) => (a.tag > b.tag ? -1 : 1)); // newest first
  return candidates[0] || null;
}

/** Accept only comma revisions inside the body zone */
async function acceptCommasInZone(event) {
  try {
    await Word.run(async (context) => {
      const box = await getLatestBodyZone(context);
      if (!box) {
        // nothing to do
        return;
      }
      const boxRange = box.getRange();

      const revisions = context.document.body.revisions;
      revisions.load("items/range,text,type");
      await context.sync();

      for (const rev of revisions.items) {
        // Scope to the zone only
        const relation = rev.range.compareLocationWith(boxRange);
        const inside =
          relation === Word.LocationRelation.inside ||
          relation === Word.LocationRelation.equal;

        if (!inside) continue;

        // Accept only comma-like edits
        rev.range.load("text");
        await context.sync();
        if (isCommaish(rev.range.text)) {
          rev.accept();
        }
      }
      await context.sync();

      // Optional: remove the zone after applying
      // box.delete(true); await context.sync();
    });
  } catch (e) {
    console.error("acceptAllChanges error:", e);
  } finally {
    event.completed();
  }
}

/** Reject only comma revisions inside the body zone */
async function rejectCommasInZone(event) {
  try {
    await Word.run(async (context) => {
      const box = await getLatestBodyZone(context);
      if (!box) {
        return;
      }
      const boxRange = box.getRange();

      const revisions = context.document.body.revisions;
      revisions.load("items/range,text,type");
      await context.sync();

      for (const rev of revisions.items) {
        const relation = rev.range.compareLocationWith(boxRange);
        const inside =
          relation === Word.LocationRelation.inside ||
          relation === Word.LocationRelation.equal;

        if (!inside) continue;

        rev.range.load("text");
        await context.sync();
        if (isCommaish(rev.range.text)) {
          rev.reject();
        }
      }
      await context.sync();

      // Optional: remove the zone after applying
      // box.delete(true); await context.sync();
    });
  } catch (e) {
    console.error("rejectAllChanges error:", e);
  } finally {
    event.completed();
  }
}

/** ─────────────────────────────────────────────────────────
 * Ribbon command entry points (match manifest FunctionName)
 * ───────────────────────────────────────────────────────── */
async function checkDocumentText(event) {
  await runCheckVejice(event); // runs the whole-doc, paragraph-by-paragraph logic
}

async function acceptAllChanges(event) {
  await acceptCommasInZone(event);
}

async function rejectAllChanges(event) {
  await rejectCommasInZone(event);
}

/** Register the functions with Office (names must match manifest FunctionName) */
Office.actions.associate("checkDocumentText", checkDocumentText);
Office.actions.associate("acceptAllChanges", acceptAllChanges);
Office.actions.associate("rejectAllChanges", rejectAllChanges);
