"use strict";
self["webpackHotUpdateoffice_addin_taskpane_js"]("commands",{

/***/ "./src/logic/preveriVejice.js":
/*!************************************!*\
  !*** ./src/logic/preveriVejice.js ***!
  \************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   applyAllSuggestionsOnline: function() { return /* binding */ applyAllSuggestionsOnline; },
/* harmony export */   checkDocumentText: function() { return /* binding */ checkDocumentText; },
/* harmony export */   getPendingSuggestionsOnline: function() { return /* binding */ getPendingSuggestionsOnline; },
/* harmony export */   rejectAllSuggestionsOnline: function() { return /* binding */ rejectAllSuggestionsOnline; }
/* harmony export */ });
/* harmony import */ var _api_apiVejice_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../api/apiVejice.js */ "./src/api/apiVejice.js");
/* harmony import */ var _utils_host_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/host.js */ "./src/utils/host.js");
var _Office;
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
/* global Word, window, process, performance, console, Office, URL */



/** ─────────────────────────────────────────────────────────
 *  DEBUG helpers (flip DEBUG=false to silence logs)
 *  ───────────────────────────────────────────────────────── */
var envIsProd = function envIsProd() {
  var _process$env;
  return typeof process !== "undefined" && ((_process$env = process.env) === null || _process$env === void 0 ? void 0 : "development") === "production" || typeof window !== "undefined" && window.__VEJICE_ENV__ === "production";
};
var DEBUG_OVERRIDE = typeof window !== "undefined" && typeof window.__VEJICE_DEBUG__ === "boolean" ? window.__VEJICE_DEBUG__ : undefined;
var DEBUG = typeof DEBUG_OVERRIDE === "boolean" ? DEBUG_OVERRIDE : !envIsProd();
var log = function log() {
  var _console;
  for (var _len = arguments.length, a = new Array(_len), _key = 0; _key < _len; _key++) {
    a[_key] = arguments[_key];
  }
  return DEBUG && (_console = console).log.apply(_console, ["[Vejice CHECK]"].concat(a));
};
var warn = function warn() {
  var _console2;
  for (var _len2 = arguments.length, a = new Array(_len2), _key2 = 0; _key2 < _len2; _key2++) {
    a[_key2] = arguments[_key2];
  }
  return DEBUG && (_console2 = console).warn.apply(_console2, ["[Vejice CHECK]"].concat(a));
};
var errL = function errL() {
  var _console3;
  for (var _len3 = arguments.length, a = new Array(_len3), _key3 = 0; _key3 < _len3; _key3++) {
    a[_key3] = arguments[_key3];
  }
  return (_console3 = console).error.apply(_console3, ["[Vejice CHECK]"].concat(a));
};
var tnow = function tnow() {
  var _performance$now, _performance, _performance$now2;
  return (_performance$now = (_performance = performance) === null || _performance === void 0 || (_performance$now2 = _performance.now) === null || _performance$now2 === void 0 ? void 0 : _performance$now2.call(_performance)) !== null && _performance$now !== void 0 ? _performance$now : Date.now();
};
var SNIP = function SNIP(s) {
  var n = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 80;
  return typeof s === "string" ? s.slice(0, n) : s;
};
var MAX_AUTOFIX_PASSES = typeof Office !== "undefined" && ((_Office = Office) === null || _Office === void 0 || (_Office = _Office.context) === null || _Office === void 0 ? void 0 : _Office.platform) === "PC" ? 3 : 2;
var HIGHLIGHT_INSERT = "#FFF9C4"; // light yellow
var HIGHLIGHT_DELETE = "#FFCDD2"; // light red

var pendingSuggestionsOnline = [];
var MAX_PARAGRAPH_CHARS = 1000;
var LONG_PARAGRAPH_MESSAGE = "Odstavek je predolg za preverjanje. Razdelite ga na več odstavkov in poskusite znova.";
function resetPendingSuggestionsOnline() {
  pendingSuggestionsOnline.length = 0;
}
function addPendingSuggestionOnline(suggestion) {
  pendingSuggestionsOnline.push(suggestion);
}
function getPendingSuggestionsOnline() {
  var debugSnapshot = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : false;
  if (!debugSnapshot) return pendingSuggestionsOnline;
  return pendingSuggestionsOnline.map(function (sug) {
    return {
      id: sug === null || sug === void 0 ? void 0 : sug.id,
      kind: sug === null || sug === void 0 ? void 0 : sug.kind,
      paragraphIndex: sug === null || sug === void 0 ? void 0 : sug.paragraphIndex,
      metadata: sug === null || sug === void 0 ? void 0 : sug.metadata,
      originalPos: sug === null || sug === void 0 ? void 0 : sug.originalPos,
      leftWord: sug === null || sug === void 0 ? void 0 : sug.leftWord,
      leftSnippet: sug === null || sug === void 0 ? void 0 : sug.leftSnippet,
      rightSnippet: sug === null || sug === void 0 ? void 0 : sug.rightSnippet
    };
  });
}
if (typeof window !== "undefined") {
  window.__VEJICE_DEBUG_STATE__ = window.__VEJICE_DEBUG_STATE__ || {};
  window.__VEJICE_DEBUG_STATE__.getPendingSuggestionsOnline = getPendingSuggestionsOnline;
  window.__VEJICE_DEBUG_STATE__.getParagraphAnchorsOnline = function () {
    return paragraphTokenAnchorsOnline;
  };
  window.getPendingSuggestionsOnline = getPendingSuggestionsOnline;
  window.getPendingSuggestionsSnapshot = function () {
    return getPendingSuggestionsOnline(true);
  };
}
var paragraphsTouchedOnline = new Set();
function resetParagraphsTouchedOnline() {
  paragraphsTouchedOnline.clear();
}
function markParagraphTouched(paragraphIndex) {
  if (typeof paragraphIndex === "number" && paragraphIndex >= 0) {
    paragraphsTouchedOnline.add(paragraphIndex);
  }
}
var toastDialog = null;
function showToastNotification(message) {
  var _Office$context;
  if (!message) return;
  if (typeof Office === "undefined" || !((_Office$context = Office.context) !== null && _Office$context !== void 0 && (_Office$context = _Office$context.ui) !== null && _Office$context !== void 0 && _Office$context.displayDialogAsync)) {
    warn("Toast notification unavailable", message);
    return;
  }
  var origin = typeof window !== "undefined" && window.location && window.location.origin || null;
  if (!origin) {
    warn("Toast notification: origin unavailable");
    return;
  }
  var toastUrl = new URL("toast.html", origin);
  toastUrl.searchParams.set("message", message);
  Office.context.ui.displayDialogAsync(toastUrl.toString(), {
    height: 20,
    width: 30,
    displayInIframe: true
  }, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      warn("Toast notification failed", asyncResult.error);
      return;
    }
    if (toastDialog) {
      try {
        toastDialog.close();
      } catch (err) {
        warn("Toast notification: failed to close previous dialog", err);
      }
    }
    toastDialog = asyncResult.value;
    var closeDialog = function closeDialog() {
      if (!toastDialog) return;
      try {
        toastDialog.close();
      } catch (err) {
        warn("Toast notification: failed to close dialog", err);
      } finally {
        toastDialog = null;
      }
    };
    toastDialog.addEventHandler(Office.EventType.DialogMessageReceived, closeDialog);
    toastDialog.addEventHandler(Office.EventType.DialogEventReceived, closeDialog);
  });
}
function notifyParagraphTooLong(paragraphIndex, length) {
  var label = paragraphIndex + 1;
  var msg = "Odstavek ".concat(label, ": ").concat(LONG_PARAGRAPH_MESSAGE, " (").concat(length, " znakov).");
  warn("Paragraph too long – skipped", {
    paragraphIndex: paragraphIndex,
    length: length
  });
  showToastNotification(msg);
}
var paragraphTokenAnchorsOnline = [];
function resetParagraphTokenAnchorsOnline() {
  paragraphTokenAnchorsOnline.length = 0;
}
function setParagraphTokenAnchorsOnline(paragraphIndex, anchors) {
  paragraphTokenAnchorsOnline[paragraphIndex] = anchors;
}
function getParagraphTokenAnchorsOnline(paragraphIndex) {
  return paragraphTokenAnchorsOnline[paragraphIndex];
}
function createParagraphTokenAnchors(_ref) {
  var paragraphIndex = _ref.paragraphIndex,
    _ref$originalText = _ref.originalText,
    originalText = _ref$originalText === void 0 ? "" : _ref$originalText,
    _ref$correctedText = _ref.correctedText,
    correctedText = _ref$correctedText === void 0 ? "" : _ref$correctedText,
    _ref$sourceTokens = _ref.sourceTokens,
    sourceTokens = _ref$sourceTokens === void 0 ? [] : _ref$sourceTokens,
    _ref$targetTokens = _ref.targetTokens,
    targetTokens = _ref$targetTokens === void 0 ? [] : _ref$targetTokens,
    _ref$documentOffset = _ref.documentOffset,
    documentOffset = _ref$documentOffset === void 0 ? 0 : _ref$documentOffset;
  var safeOriginal = typeof originalText === "string" ? originalText : "";
  var safeCorrected = typeof correctedText === "string" ? correctedText : "";
  var normalizedSource = normalizeTokenList(sourceTokens, "s");
  var normalizedTarget = normalizeTokenList(targetTokens, "t");
  var entry = {
    paragraphIndex: paragraphIndex,
    documentOffset: documentOffset,
    originalText: safeOriginal,
    correctedText: safeCorrected,
    sourceTokens: normalizedSource,
    targetTokens: normalizedTarget,
    sourceAnchors: mapTokensToParagraphText(paragraphIndex, safeOriginal, normalizedSource, documentOffset),
    targetAnchors: mapTokensToParagraphText(paragraphIndex, safeCorrected, normalizedTarget, documentOffset)
  };
  setParagraphTokenAnchorsOnline(paragraphIndex, entry);
  return entry;
}
function normalizeTokenList(tokens, prefix) {
  if (!Array.isArray(tokens)) return [];
  var normalized = [];
  var idCounts = Object.create(null);
  for (var i = 0; i < tokens.length; i++) {
    var _idCounts$baseId;
    var token = normalizeToken(tokens[i], prefix, i);
    if (!token) continue;
    var baseId = typeof token.id === "string" && token.id.length ? token.id : "".concat(prefix).concat(i + 1);
    var count = (_idCounts$baseId = idCounts[baseId]) !== null && _idCounts$baseId !== void 0 ? _idCounts$baseId : 0;
    idCounts[baseId] = count + 1;
    token.id = count === 0 ? baseId : "".concat(baseId, "#").concat(count + 1);
    token.originalId = baseId;
    token.occurrence = count;
    normalized.push(token);
  }
  return normalized;
}
function normalizeToken(rawToken, prefix, index) {
  if (rawToken === null || typeof rawToken === "undefined") return null;
  if (typeof rawToken === "string") {
    return {
      id: "".concat(prefix).concat(index + 1),
      text: rawToken,
      raw: rawToken
    };
  }
  if (_typeof(rawToken) === "object") {
    var _ref2, _ref3, _ref4, _ref5, _rawToken$token_id, _ref6, _ref7, _ref8, _ref9, _rawToken$token, _ref0, _ref1, _ref10, _ref11, _rawToken$whitespace, _ref12, _ref13, _ref14, _rawToken$leading_ws;
    var idCandidate = (_ref2 = (_ref3 = (_ref4 = (_ref5 = (_rawToken$token_id = rawToken.token_id) !== null && _rawToken$token_id !== void 0 ? _rawToken$token_id : rawToken.tokenId) !== null && _ref5 !== void 0 ? _ref5 : rawToken.id) !== null && _ref4 !== void 0 ? _ref4 : rawToken.ID) !== null && _ref3 !== void 0 ? _ref3 : rawToken.name) !== null && _ref2 !== void 0 ? _ref2 : rawToken.key;
    var textCandidate = (_ref6 = (_ref7 = (_ref8 = (_ref9 = (_rawToken$token = rawToken.token) !== null && _rawToken$token !== void 0 ? _rawToken$token : rawToken.text) !== null && _ref9 !== void 0 ? _ref9 : rawToken.form) !== null && _ref8 !== void 0 ? _ref8 : rawToken.value) !== null && _ref7 !== void 0 ? _ref7 : rawToken.surface) !== null && _ref6 !== void 0 ? _ref6 : rawToken.word;
    var trailing = (_ref0 = (_ref1 = (_ref10 = (_ref11 = (_rawToken$whitespace = rawToken.whitespace) !== null && _rawToken$whitespace !== void 0 ? _rawToken$whitespace : rawToken.trailing_ws) !== null && _ref11 !== void 0 ? _ref11 : rawToken.trailingWhitespace) !== null && _ref10 !== void 0 ? _ref10 : rawToken.after) !== null && _ref1 !== void 0 ? _ref1 : rawToken.space) !== null && _ref0 !== void 0 ? _ref0 : "";
    var leading = (_ref12 = (_ref13 = (_ref14 = (_rawToken$leading_ws = rawToken.leading_ws) !== null && _rawToken$leading_ws !== void 0 ? _rawToken$leading_ws : rawToken.leadingWhitespace) !== null && _ref14 !== void 0 ? _ref14 : rawToken.before) !== null && _ref13 !== void 0 ? _ref13 : rawToken.prefix) !== null && _ref12 !== void 0 ? _ref12 : "";
    return {
      id: typeof idCandidate === "string" ? idCandidate : "".concat(prefix).concat(index + 1),
      text: typeof textCandidate === "string" ? textCandidate : "",
      trailingWhitespace: typeof trailing === "string" ? trailing : "",
      leadingWhitespace: typeof leading === "string" ? leading : "",
      raw: rawToken
    };
  }
  return null;
}
function mapTokensToParagraphText(paragraphIndex, paragraphText, tokens) {
  var documentOffset = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : 0;
  var byId = Object.create(null);
  var ordered = [];
  if (!Array.isArray(tokens) || !tokens.length) {
    return {
      byId: byId,
      ordered: ordered
    };
  }
  var safeParagraph = typeof paragraphText === "string" ? paragraphText : "";
  var textOccurrences = Object.create(null);
  var trimmedOccurrences = Object.create(null);
  var cursor = 0;
  for (var i = 0; i < tokens.length; i++) {
    var _token$text, _token$id, _textOccurrences$text, _trimmedOccurrences$t;
    var token = tokens[i];
    var tokenText = (_token$text = token === null || token === void 0 ? void 0 : token.text) !== null && _token$text !== void 0 ? _token$text : "";
    var tokenId = (_token$id = token === null || token === void 0 ? void 0 : token.id) !== null && _token$id !== void 0 ? _token$id : "tok".concat(i + 1);
    var tokenLength = tokenText.length;
    var charStart = resolveTokenPosition(safeParagraph, tokenText, cursor);
    var charEnd = charStart >= 0 ? charStart + tokenLength : -1;
    if (charStart >= 0) {
      cursor = charEnd;
    } else if (tokenText) {
      warn("Token mapping failed", {
        paragraphIndex: paragraphIndex,
        tokenId: tokenId,
        tokenText: tokenText,
        cursor: cursor
      });
    }
    var textKey = tokenText || "";
    var trimmedKey = textKey.trim();
    var occurrence = (_textOccurrences$text = textOccurrences[textKey]) !== null && _textOccurrences$text !== void 0 ? _textOccurrences$text : 0;
    textOccurrences[textKey] = occurrence + 1;
    var trimmedOccurrence = trimmedKey && trimmedKey !== textKey ? (_trimmedOccurrences$t = trimmedOccurrences[trimmedKey]) !== null && _trimmedOccurrences$t !== void 0 ? _trimmedOccurrences$t : 0 : occurrence;
    if (trimmedKey && trimmedKey !== textKey) {
      trimmedOccurrences[trimmedKey] = trimmedOccurrence + 1;
    }
    var anchor = {
      paragraphIndex: paragraphIndex,
      tokenId: tokenId,
      tokenIndex: i,
      tokenText: tokenText,
      length: tokenLength,
      textOccurrence: occurrence,
      trimmedTextOccurrence: trimmedKey ? trimmedOccurrence : occurrence,
      charStart: charStart,
      charEnd: charEnd,
      documentCharStart: charStart >= 0 ? documentOffset + charStart : -1,
      documentCharEnd: charEnd >= 0 ? documentOffset + charEnd : -1,
      matched: charStart >= 0
    };
    byId[tokenId] = anchor;
    ordered.push(anchor);
  }
  return {
    byId: byId,
    ordered: ordered
  };
}
function resolveTokenPosition(text, tokenText, fromIndex) {
  if (!tokenText || typeof text !== "string") return -1;
  var textLength = text.length;
  if (!textLength) return -1;
  var searchStart = fromIndex;
  if (searchStart < 0) searchStart = 0;
  if (searchStart > textLength) searchStart = textLength;
  var idx = text.indexOf(tokenText, searchStart);
  if (idx !== -1) return idx;
  var trimmed = tokenText.trim();
  if (trimmed && trimmed !== tokenText) {
    idx = text.indexOf(trimmed, searchStart);
    if (idx !== -1) return idx;
  }
  if (searchStart > 0) {
    var retryStart = Math.max(0, searchStart - tokenText.length - 1);
    idx = text.indexOf(tokenText, retryStart);
    if (idx !== -1) return idx;
  }
  return text.indexOf(tokenText);
}
function selectTokenAnchors(entry, type) {
  if (!entry) return null;
  return type === "target" ? entry.targetAnchors : entry.sourceAnchors;
}
function snapshotAnchor(anchor) {
  if (!anchor) return undefined;
  return {
    tokenId: anchor.tokenId,
    tokenIndex: anchor.tokenIndex,
    tokenText: anchor.tokenText,
    textOccurrence: anchor.textOccurrence,
    trimmedTextOccurrence: anchor.trimmedTextOccurrence,
    charStart: anchor.charStart,
    charEnd: anchor.charEnd,
    documentCharStart: anchor.documentCharStart,
    documentCharEnd: anchor.documentCharEnd,
    length: anchor.length,
    matched: anchor.matched
  };
}
function findAnchorsNearChar(entry, type, charIndex) {
  var _collection$ordered;
  var collection = selectTokenAnchors(entry, type);
  if (!(collection !== null && collection !== void 0 && (_collection$ordered = collection.ordered) !== null && _collection$ordered !== void 0 && _collection$ordered.length) || typeof charIndex !== "number" || charIndex < 0) {
    return {
      before: null,
      at: null,
      after: null
    };
  }
  var before = null;
  for (var i = 0; i < collection.ordered.length; i++) {
    var anchor = collection.ordered[i];
    if (!anchor || anchor.charStart < 0) continue;
    if (charIndex >= anchor.charStart && charIndex <= anchor.charEnd) {
      return {
        before: before !== null && before !== void 0 ? before : anchor,
        at: anchor,
        after: findNextAnchorWithPosition(collection.ordered, i + 1)
      };
    }
    if (anchor.charStart > charIndex) {
      return {
        before: before,
        at: null,
        after: anchor
      };
    }
    before = anchor;
  }
  return {
    before: before,
    at: null,
    after: null
  };
}
function findNextAnchorWithPosition(list, startIndex) {
  if (!Array.isArray(list)) return null;
  for (var i = startIndex; i < list.length; i++) {
    var anchor = list[i];
    if (anchor && anchor.charStart >= 0) return anchor;
  }
  return null;
}
function countSnippetOccurrencesBefore(text, snippet, limit) {
  if (!snippet) return 0;
  var safeText = typeof text === "string" ? text : "";
  var hop = Math.max(1, snippet.length);
  var count = 0;
  var idx = safeText.indexOf(snippet);
  while (idx !== -1 && idx < limit) {
    count++;
    idx = safeText.indexOf(snippet, idx + hop);
  }
  return count;
}
function getRangeForCharacterSpan(_x, _x2, _x3, _x4, _x5) {
  return _getRangeForCharacterSpan.apply(this, arguments);
}
function _getRangeForCharacterSpan() {
  _getRangeForCharacterSpan = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(context, paragraph, paragraphText, charStart, charEnd) {
    var reason,
      fallbackSnippet,
      text,
      safeStart,
      computedEnd,
      safeEnd,
      snippet,
      matches,
      occurrence,
      idx,
      _args = arguments,
      _t;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.p = _context.n) {
        case 0:
          reason = _args.length > 5 && _args[5] !== undefined ? _args[5] : "span";
          fallbackSnippet = _args.length > 6 ? _args[6] : undefined;
          if (!(!paragraph || typeof paragraph.getRange !== "function")) {
            _context.n = 1;
            break;
          }
          return _context.a(2, null);
        case 1:
          if (!(!Number.isFinite(charStart) || charStart < 0)) {
            _context.n = 2;
            break;
          }
          return _context.a(2, null);
        case 2:
          text = typeof paragraphText === "string" ? paragraphText : paragraph.text || "";
          if (text) {
            _context.n = 3;
            break;
          }
          return _context.a(2, null);
        case 3:
          safeStart = Math.max(0, Math.min(Math.floor(charStart), text.length ? text.length - 1 : 0));
          computedEnd = Math.max(safeStart + 1, Math.floor(charEnd !== null && charEnd !== void 0 ? charEnd : safeStart + 1));
          safeEnd = Math.min(computedEnd, text.length);
          snippet = text.slice(safeStart, safeEnd);
          if (!snippet && typeof fallbackSnippet === "string" && fallbackSnippet.length) {
            snippet = fallbackSnippet;
          }
          if (snippet) {
            _context.n = 4;
            break;
          }
          return _context.a(2, null);
        case 4:
          _context.p = 4;
          matches = paragraph.getRange().search(snippet, {
            matchCase: true,
            matchWholeWord: false,
            ignoreSpace: false,
            ignorePunct: false
          });
          matches.load("items");
          _context.n = 5;
          return context.sync();
        case 5:
          if (matches.items.length) {
            _context.n = 6;
            break;
          }
          warn("getRangeForCharacterSpan(".concat(reason, "): snippet not found"), {
            snippet: snippet,
            safeStart: safeStart
          });
          return _context.a(2, null);
        case 6:
          occurrence = countSnippetOccurrencesBefore(text, snippet, safeStart);
          idx = Math.min(occurrence, matches.items.length - 1);
          return _context.a(2, matches.items[idx]);
        case 7:
          _context.p = 7;
          _t = _context.v;
          warn("getRangeForCharacterSpan(".concat(reason, ") failed"), _t);
          return _context.a(2, null);
      }
    }, _callee, null, [[4, 7]]);
  }));
  return _getRangeForCharacterSpan.apply(this, arguments);
}
function buildDeleteSuggestionMetadata(entry, charIndex) {
  var _entry$documentOffset, _entry$originalText, _entry$paragraphIndex;
  var sourceAround = findAnchorsNearChar(entry, "source", charIndex);
  var documentOffset = (_entry$documentOffset = entry === null || entry === void 0 ? void 0 : entry.documentOffset) !== null && _entry$documentOffset !== void 0 ? _entry$documentOffset : 0;
  var charStart = Math.max(0, charIndex);
  var charEnd = charStart + 1;
  var paragraphText = (_entry$originalText = entry === null || entry === void 0 ? void 0 : entry.originalText) !== null && _entry$originalText !== void 0 ? _entry$originalText : "";
  var highlightText = paragraphText.slice(charStart, charEnd) || ",";
  return {
    kind: "delete",
    paragraphIndex: (_entry$paragraphIndex = entry === null || entry === void 0 ? void 0 : entry.paragraphIndex) !== null && _entry$paragraphIndex !== void 0 ? _entry$paragraphIndex : -1,
    charStart: charStart,
    charEnd: charEnd,
    documentCharStart: documentOffset + charStart,
    documentCharEnd: documentOffset + charEnd,
    sourceTokenBefore: snapshotAnchor(sourceAround.before),
    sourceTokenAt: snapshotAnchor(sourceAround.at),
    sourceTokenAfter: snapshotAnchor(sourceAround.after),
    highlightText: highlightText
  };
}
function buildInsertSuggestionMetadata(entry, _ref15) {
  var _entry$documentOffset2, _ref16, _ref17, _ref18, _ref19, _sourceAround$after, _highlightAnchor$char, _highlightAnchor$char2, _entry$originalText2, _entry$paragraphIndex2;
  var originalCharIndex = _ref15.originalCharIndex,
    targetCharIndex = _ref15.targetCharIndex;
  var srcIndex = typeof originalCharIndex === "number" ? originalCharIndex : -1;
  var targetIndex = typeof targetCharIndex === "number" ? targetCharIndex : srcIndex;
  var sourceAround = findAnchorsNearChar(entry, "source", srcIndex);
  var targetAround = findAnchorsNearChar(entry, "target", targetIndex);
  var documentOffset = (_entry$documentOffset2 = entry === null || entry === void 0 ? void 0 : entry.documentOffset) !== null && _entry$documentOffset2 !== void 0 ? _entry$documentOffset2 : 0;
  var highlightAnchor = (_ref16 = (_ref17 = (_ref18 = (_ref19 = (_sourceAround$after = sourceAround.after) !== null && _sourceAround$after !== void 0 ? _sourceAround$after : sourceAround.at) !== null && _ref19 !== void 0 ? _ref19 : sourceAround.before) !== null && _ref18 !== void 0 ? _ref18 : targetAround.before) !== null && _ref17 !== void 0 ? _ref17 : targetAround.at) !== null && _ref16 !== void 0 ? _ref16 : targetAround.after;
  var highlightCharStart = (_highlightAnchor$char = highlightAnchor === null || highlightAnchor === void 0 ? void 0 : highlightAnchor.charStart) !== null && _highlightAnchor$char !== void 0 ? _highlightAnchor$char : srcIndex;
  var highlightCharEnd = (_highlightAnchor$char2 = highlightAnchor === null || highlightAnchor === void 0 ? void 0 : highlightAnchor.charEnd) !== null && _highlightAnchor$char2 !== void 0 ? _highlightAnchor$char2 : srcIndex;
  var paragraphText = (_entry$originalText2 = entry === null || entry === void 0 ? void 0 : entry.originalText) !== null && _entry$originalText2 !== void 0 ? _entry$originalText2 : "";
  var highlightText = "";
  if (highlightCharStart >= 0 && highlightCharEnd > highlightCharStart) {
    highlightText = paragraphText.slice(highlightCharStart, highlightCharEnd);
  }
  if (!highlightText && highlightCharStart >= 0) {
    highlightText = paragraphText.slice(highlightCharStart, highlightCharStart + 1);
  }
  return {
    kind: "insert",
    paragraphIndex: (_entry$paragraphIndex2 = entry === null || entry === void 0 ? void 0 : entry.paragraphIndex) !== null && _entry$paragraphIndex2 !== void 0 ? _entry$paragraphIndex2 : -1,
    targetCharStart: targetIndex,
    targetCharEnd: targetIndex >= 0 ? targetIndex + 1 : targetIndex,
    targetDocumentCharStart: targetIndex >= 0 ? documentOffset + targetIndex : targetIndex,
    targetDocumentCharEnd: targetIndex >= 0 ? documentOffset + targetIndex + 1 : targetIndex,
    highlightCharStart: highlightCharStart,
    highlightCharEnd: highlightCharEnd,
    highlightText: highlightText,
    sourceTokenBefore: snapshotAnchor(sourceAround.before),
    sourceTokenAt: snapshotAnchor(sourceAround.at),
    sourceTokenAfter: snapshotAnchor(sourceAround.after),
    targetTokenBefore: snapshotAnchor(targetAround.before),
    targetTokenAt: snapshotAnchor(targetAround.at),
    targetTokenAfter: snapshotAnchor(targetAround.after),
    highlightAnchorTarget: snapshotAnchor(highlightAnchor)
  };
}

/** ─────────────────────────────────────────────────────────
 *  Helpers: znaki & pravila
 *  ───────────────────────────────────────────────────────── */
var QUOTES = new Set(['"', "'", "“", "”", "„", "«", "»"]);
var isDigit = function isDigit(ch) {
  return ch >= "0" && ch <= "9";
};
var charAtSafe = function charAtSafe(s, i) {
  return i >= 0 && i < s.length ? s[i] : "";
};

/** Številčni vejici (decimalna ali tisočiška) */
function isNumericComma(original, corrected, kind, pos) {
  var s = kind === "delete" ? original : corrected;
  var prev = charAtSafe(s, pos - 1);
  var next = charAtSafe(s, pos + 1);
  return isDigit(prev) && isDigit(next);
}

/** Guard: ali so se spremenile samo vejice */
function onlyCommasChanged(original, corrected) {
  var strip = function strip(x) {
    return x.replace(/,/g, "");
  };
  return strip(original) === strip(corrected);
}

/** Minimalni diff: samo operacije z vejicami */
function diffCommasOnly(original, corrected) {
  var ops = [];
  var i = 0,
    j = 0;
  while (i < original.length || j < corrected.length) {
    var _original$i, _corrected$j;
    var o = (_original$i = original[i]) !== null && _original$i !== void 0 ? _original$i : "";
    var c = (_corrected$j = corrected[j]) !== null && _corrected$j !== void 0 ? _corrected$j : "";
    if (o === c) {
      i++;
      j++;
      continue;
    }
    if (c === "," && o !== ",") {
      ops.push({
        kind: "insert",
        pos: j,
        originalPos: i,
        correctedPos: j
      });
      j++;
      continue;
    }
    if (o === "," && c !== ",") {
      ops.push({
        kind: "delete",
        pos: i,
        originalPos: i,
        correctedPos: j
      });
      i++;
      continue;
    }
    if (o) i++;
    if (c) j++;
  }
  return ops;
}

/** Filtriraj operacije (izloči številčne vejice, dodaj presledek kasneje, če treba) */
function filterCommaOps(original, corrected, ops) {
  return ops.filter(function (op) {
    if (isNumericComma(original, corrected, op.kind, op.pos)) return false;
    if (op.kind === "insert") {
      var next = charAtSafe(corrected, op.pos + 1);
      var noSpaceAfter = next && !/\s/.test(next);
      if (noSpaceAfter && !QUOTES.has(next)) {
        // dovolimo; presledek dodamo naknadno
        return true;
      }
    }
    return true;
  });
}

/** Anchor-based mikro urejanje (ohrani formatiranje) */
function makeAnchor(text, idx) {
  var span = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : 16;
  var left = text.slice(Math.max(0, idx - span), idx);
  var right = text.slice(idx, Math.min(text.length, idx + span));
  return {
    left: left,
    right: right
  };
}
function findInsertIndex() {
  var original = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : "";
  var left = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var right = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  if (left && right) {
    var candidates = [];
    var pos = original.indexOf(left);
    while (pos >= 0) {
      var idx = pos + left.length;
      var tail = original.slice(idx, idx + right.length);
      if (!right || tail === right.slice(0, tail.length)) candidates.push(idx);
      pos = original.indexOf(left, pos + 1);
    }
    if (candidates.length) return candidates[0];
  }
  if (left) {
    var _idx = original.lastIndexOf(left);
    if (_idx >= 0) return _idx + left.length;
  }
  if (right) {
    var _idx2 = original.indexOf(right);
    if (_idx2 >= 0) return _idx2;
  }
  return -1;
}
function commaAlreadyPresent() {
  var original = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : "";
  var left = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var right = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var insertIdx = findInsertIndex(original, left, right);
  if (insertIdx < 0) return false;
  var prev = original[insertIdx - 1];
  var at = original[insertIdx];
  return prev === "," || at === ",";
}
function resolveOpPositions(opOrPos) {
  if (opOrPos && _typeof(opOrPos) === "object") {
    var correctedPos = typeof opOrPos.correctedPos === "number" ? opOrPos.correctedPos : typeof opOrPos.pos === "number" ? opOrPos.pos : -1;
    var originalPos = typeof opOrPos.originalPos === "number" ? opOrPos.originalPos : correctedPos;
    return {
      correctedPos: correctedPos,
      originalPos: originalPos
    };
  }
  var pos = typeof opOrPos === "number" ? opOrPos : -1;
  return {
    correctedPos: pos,
    originalPos: pos
  };
}
function removeDoubleCommas(_x6, _x7) {
  return _removeDoubleCommas.apply(this, arguments);
} // Vstavi vejico na podlagi sidra
function _removeDoubleCommas() {
  _removeDoubleCommas = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(context, paragraph) {
    var passes, matches;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          passes = 0;
        case 1:
          if (!(passes < 3)) {
            _context2.n = 5;
            break;
          }
          matches = paragraph.search(",,", {
            matchCase: false,
            matchWholeWord: false
          });
          matches.load("items");
          _context2.n = 2;
          return context.sync();
        case 2:
          if (matches.items.length) {
            _context2.n = 3;
            break;
          }
          return _context2.a(3, 5);
        case 3:
          matches.items.forEach(function (r) {
            return r.insertText(",", Word.InsertLocation.replace);
          });
          _context2.n = 4;
          return context.sync();
        case 4:
          passes++;
          _context2.n = 1;
          break;
        case 5:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _removeDoubleCommas.apply(this, arguments);
}
function insertCommaAt(_x8, _x9, _x0, _x1, _x10) {
  return _insertCommaAt.apply(this, arguments);
} // Po potrebi dodaj presledek po vejici (razen pred narekovaji ali števkami)
function _insertCommaAt() {
  _insertCommaAt = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(context, paragraph, original, corrected, opOrPos) {
    var _resolveOpPositions, correctedPos, originalPos, pos, prevChar, atChar, _makeAnchor, left, right, pr, m, targetIdx, ordinal, after, rightClean, snippet, _m, before;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          _resolveOpPositions = resolveOpPositions(opOrPos), correctedPos = _resolveOpPositions.correctedPos, originalPos = _resolveOpPositions.originalPos;
          if (!(!Number.isFinite(correctedPos) || correctedPos < 0)) {
            _context3.n = 1;
            break;
          }
          return _context3.a(2);
        case 1:
          // Direct position guard: if a comma is already at/adjacent to the target, skip.
          pos = Math.max(0, correctedPos);
          prevChar = pos > 0 ? original[pos - 1] : "";
          atChar = pos < original.length ? original[pos] : "";
          if (!(prevChar === "," || atChar === ",")) {
            _context3.n = 2;
            break;
          }
          warn("insert: comma already at/near target (pos check); skipping");
          return _context3.a(2);
        case 2:
          _makeAnchor = makeAnchor(corrected, correctedPos), left = _makeAnchor.left, right = _makeAnchor.right;
          if (!commaAlreadyPresent(original, left, right)) {
            _context3.n = 3;
            break;
          }
          warn("insert: comma already present at target; skipping");
          return _context3.a(2);
        case 3:
          pr = paragraph.getRange();
          if (!(left.length > 0)) {
            _context3.n = 6;
            break;
          }
          m = pr.search(left, {
            matchCase: false,
            matchWholeWord: false
          });
          m.load("items");
          _context3.n = 4;
          return context.sync();
        case 4:
          if (m.items.length) {
            _context3.n = 5;
            break;
          }
          warn("insert: left anchor not found");
          return _context3.a(2);
        case 5:
          targetIdx = m.items.length - 1;
          if (typeof originalPos === "number" && originalPos >= 0) {
            ordinal = countSnippetOccurrencesBefore(original, left, originalPos);
            if (ordinal > 0) {
              targetIdx = Math.min(ordinal - 1, m.items.length - 1);
            }
          }
          after = m.items[targetIdx].getRange("After");
          after.insertText(",", Word.InsertLocation.before);
          _context3.n = 11;
          break;
        case 6:
          if (right) {
            _context3.n = 7;
            break;
          }
          warn("insert: no right anchor at paragraph start");
          return _context3.a(2);
        case 7:
          rightClean = typeof right === "string" ? right.replace(/^,+/, "").replace(/,/g, "").trim() : "";
          if (rightClean) {
            _context3.n = 8;
            break;
          }
          warn("insert: right anchor unavailable");
          return _context3.a(2);
        case 8:
          snippet = rightClean.slice(0, 12);
          _m = pr.search(snippet, {
            matchCase: false,
            matchWholeWord: false
          });
          _m.load("items");
          _context3.n = 9;
          return context.sync();
        case 9:
          if (_m.items.length) {
            _context3.n = 10;
            break;
          }
          warn("insert: right anchor not found");
          return _context3.a(2);
        case 10:
          before = _m.items[0].getRange("Before");
          before.insertText(",", Word.InsertLocation.before);
        case 11:
          return _context3.a(2);
      }
    }, _callee3);
  }));
  return _insertCommaAt.apply(this, arguments);
}
function ensureSpaceAfterComma(_x11, _x12, _x13, _x14, _x15) {
  return _ensureSpaceAfterComma.apply(this, arguments);
} // Briši samo znak vejice
function _ensureSpaceAfterComma() {
  _ensureSpaceAfterComma = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(context, paragraph, original, corrected, opOrPos) {
    var _resolveOpPositions2, correctedPos, originalPos, next, _makeAnchor2, left, right, pr, m, targetIdx, ordinal, beforeRight, rightClean, snippet, _m2, before;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          _resolveOpPositions2 = resolveOpPositions(opOrPos), correctedPos = _resolveOpPositions2.correctedPos, originalPos = _resolveOpPositions2.originalPos;
          if (!(!Number.isFinite(correctedPos) || correctedPos < 0)) {
            _context4.n = 1;
            break;
          }
          return _context4.a(2);
        case 1:
          next = charAtSafe(corrected, correctedPos + 1);
          if (!(!next || /\s/.test(next) || QUOTES.has(next) || isDigit(next))) {
            _context4.n = 2;
            break;
          }
          return _context4.a(2);
        case 2:
          _makeAnchor2 = makeAnchor(corrected, correctedPos + 1), left = _makeAnchor2.left, right = _makeAnchor2.right;
          pr = paragraph.getRange();
          if (!(left.length > 0)) {
            _context4.n = 5;
            break;
          }
          m = pr.search(left, {
            matchCase: false,
            matchWholeWord: false
          });
          m.load("items");
          _context4.n = 3;
          return context.sync();
        case 3:
          if (m.items.length) {
            _context4.n = 4;
            break;
          }
          warn("space-after: left anchor not found");
          return _context4.a(2);
        case 4:
          targetIdx = m.items.length - 1;
          if (typeof originalPos === "number" && originalPos >= 0) {
            ordinal = countSnippetOccurrencesBefore(original, left, originalPos);
            if (ordinal > 0) {
              targetIdx = Math.min(ordinal - 1, m.items.length - 1);
            }
          }
          beforeRight = m.items[targetIdx].getRange("Before");
          beforeRight.insertText(" ", Word.InsertLocation.before);
          _context4.n = 9;
          break;
        case 5:
          if (!(right.length > 0)) {
            _context4.n = 9;
            break;
          }
          rightClean = typeof right === "string" ? right.replace(/^,+/, "").replace(/,/g, "").trim() : "";
          if (rightClean) {
            _context4.n = 6;
            break;
          }
          warn("space-after: right anchor unavailable");
          return _context4.a(2);
        case 6:
          snippet = rightClean.slice(0, 12);
          _m2 = pr.search(snippet, {
            matchCase: false,
            matchWholeWord: false
          });
          _m2.load("items");
          _context4.n = 7;
          return context.sync();
        case 7:
          if (_m2.items.length) {
            _context4.n = 8;
            break;
          }
          warn("space-after: right anchor not found");
          return _context4.a(2);
        case 8:
          before = _m2.items[0].getRange("Before");
          before.insertText(" ", Word.InsertLocation.before);
        case 9:
          return _context4.a(2);
      }
    }, _callee4);
  }));
  return _ensureSpaceAfterComma.apply(this, arguments);
}
function deleteCommaAt(_x16, _x17, _x18, _x19) {
  return _deleteCommaAt.apply(this, arguments);
}
function _deleteCommaAt() {
  _deleteCommaAt = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(context, paragraph, original, atOriginalPos) {
    var pr, ordinal, i, matches, idx;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.n) {
        case 0:
          pr = paragraph.getRange();
          ordinal = 0;
          for (i = 0; i <= atOriginalPos && i < original.length; i++) {
            if (original[i] === ",") ordinal++;
          }
          if (!(ordinal === 0)) {
            _context5.n = 1;
            break;
          }
          warn("delete: no comma found in original at pos", atOriginalPos);
          return _context5.a(2);
        case 1:
          matches = pr.search(",", {
            matchCase: false,
            matchWholeWord: false
          });
          matches.load("items");
          _context5.n = 2;
          return context.sync();
        case 2:
          idx = ordinal - 1;
          if (!(idx >= matches.items.length)) {
            _context5.n = 3;
            break;
          }
          warn("delete: comma ordinal out of range", ordinal, "/", matches.items.length);
          return _context5.a(2);
        case 3:
          matches.items[idx].insertText("", Word.InsertLocation.replace);
        case 4:
          return _context5.a(2);
      }
    }, _callee5);
  }));
  return _deleteCommaAt.apply(this, arguments);
}
function createSuggestionId(kind, paragraphIndex, pos) {
  return "".concat(kind, "-").concat(paragraphIndex, "-").concat(pos, "-").concat(pendingSuggestionsOnline.length);
}
function highlightSuggestionOnline(_x20, _x21, _x22, _x23, _x24, _x25, _x26) {
  return _highlightSuggestionOnline.apply(this, arguments);
}
function _highlightSuggestionOnline() {
  _highlightSuggestionOnline = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(context, paragraph, original, corrected, op, paragraphIndex, paragraphAnchors) {
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.n) {
        case 0:
          if (!(op.kind === "delete")) {
            _context6.n = 1;
            break;
          }
          return _context6.a(2, highlightDeleteSuggestion(context, paragraph, original, op, paragraphIndex, paragraphAnchors));
        case 1:
          return _context6.a(2, highlightInsertSuggestion(context, paragraph, corrected, op, paragraphIndex, paragraphAnchors));
      }
    }, _callee6);
  }));
  return _highlightSuggestionOnline.apply(this, arguments);
}
function countCommasUpTo(text, pos) {
  var count = 0;
  for (var i = 0; i <= pos && i < text.length; i++) {
    if (text[i] === ",") count++;
  }
  return count;
}
function highlightDeleteSuggestion(_x27, _x28, _x29, _x30, _x31, _x32) {
  return _highlightDeleteSuggestion.apply(this, arguments);
}
function _highlightDeleteSuggestion() {
  _highlightDeleteSuggestion = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7(context, paragraph, original, op, paragraphIndex, anchorsEntry) {
    var _op$originalPos;
    var metadata, targetRange;
    return _regenerator().w(function (_context7) {
      while (1) switch (_context7.n) {
        case 0:
          metadata = buildDeleteSuggestionMetadata(anchorsEntry, (_op$originalPos = op.originalPos) !== null && _op$originalPos !== void 0 ? _op$originalPos : op.pos);
          _context7.n = 1;
          return getRangeForCharacterSpan(context, paragraph, original, metadata.charStart, metadata.charEnd, "highlight-delete", metadata.highlightText);
        case 1:
          targetRange = _context7.v;
          if (targetRange) {
            _context7.n = 3;
            break;
          }
          _context7.n = 2;
          return findCommaRangeByOrdinal(context, paragraph, original, op);
        case 2:
          targetRange = _context7.v;
          if (targetRange) {
            _context7.n = 3;
            break;
          }
          return _context7.a(2, false);
        case 3:
          targetRange.font.highlightColor = HIGHLIGHT_DELETE;
          context.trackedObjects.add(targetRange);
          addPendingSuggestionOnline({
            id: createSuggestionId("del", paragraphIndex, op.pos),
            kind: "delete",
            paragraphIndex: paragraphIndex,
            originalPos: op.pos,
            highlightRange: targetRange,
            metadata: metadata
          });
          markParagraphTouched(paragraphIndex);
          return _context7.a(2, true);
      }
    }, _callee7);
  }));
  return _highlightDeleteSuggestion.apply(this, arguments);
}
function highlightInsertSuggestion(_x33, _x34, _x35, _x36, _x37, _x38) {
  return _highlightInsertSuggestion.apply(this, arguments);
}
function _highlightInsertSuggestion() {
  _highlightInsertSuggestion = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8(context, paragraph, corrected, op, paragraphIndex, anchorsEntry) {
    var _op$originalPos2, _op$correctedPos;
    var metadata, anchor, rawLeft, rawRight, leftSnippetStored, rightSnippetStored, lastWord, leftContext, searchOpts, range, wordSearch, leftSearch, rightSnippet, rightSearch, _anchorsEntry$origina;
    return _regenerator().w(function (_context8) {
      while (1) switch (_context8.n) {
        case 0:
          metadata = buildInsertSuggestionMetadata(anchorsEntry, {
            originalCharIndex: (_op$originalPos2 = op.originalPos) !== null && _op$originalPos2 !== void 0 ? _op$originalPos2 : op.pos,
            targetCharIndex: (_op$correctedPos = op.correctedPos) !== null && _op$correctedPos !== void 0 ? _op$correctedPos : op.pos
          });
          anchor = makeAnchor(corrected, op.pos);
          rawLeft = anchor.left || "";
          rawRight = anchor.right || corrected.slice(op.pos, op.pos + 24);
          leftSnippetStored = rawLeft.slice(-40);
          rightSnippetStored = rawRight.slice(0, 40);
          lastWord = extractLastWord(rawLeft);
          leftContext = rawLeft.slice(-20).replace(/[\r\n]+/g, " ");
          searchOpts = {
            matchCase: false,
            matchWholeWord: false
          };
          range = null;
          if (!(!range && lastWord)) {
            _context8.n = 2;
            break;
          }
          wordSearch = paragraph.getRange().search(lastWord, {
            matchCase: false,
            matchWholeWord: true
          });
          wordSearch.load("items");
          _context8.n = 1;
          return context.sync();
        case 1:
          if (wordSearch.items.length) {
            range = wordSearch.items[wordSearch.items.length - 1];
          }
        case 2:
          if (!(!range && leftContext.trim())) {
            _context8.n = 4;
            break;
          }
          leftSearch = paragraph.getRange().search(leftContext.trim(), searchOpts);
          leftSearch.load("items");
          _context8.n = 3;
          return context.sync();
        case 3:
          if (leftSearch.items.length) {
            range = leftSearch.items[leftSearch.items.length - 1];
          }
        case 4:
          if (range) {
            _context8.n = 6;
            break;
          }
          rightSnippet = rightSnippetStored.replace(/,/g, "").trim();
          rightSnippet = rightSnippet.slice(0, 8);
          if (!rightSnippet) {
            _context8.n = 6;
            break;
          }
          rightSearch = paragraph.getRange().search(rightSnippet, searchOpts);
          rightSearch.load("items");
          _context8.n = 5;
          return context.sync();
        case 5:
          if (rightSearch.items.length) {
            range = rightSearch.items[0];
          }
        case 6:
          if (range) {
            _context8.n = 8;
            break;
          }
          warn("highlight insert: could not locate snippet");
          _context8.n = 7;
          return getRangeForCharacterSpan(context, paragraph, (_anchorsEntry$origina = anchorsEntry === null || anchorsEntry === void 0 ? void 0 : anchorsEntry.originalText) !== null && _anchorsEntry$origina !== void 0 ? _anchorsEntry$origina : corrected, metadata.highlightCharStart, metadata.highlightCharEnd, "highlight-insert", metadata.highlightText);
        case 7:
          range = _context8.v;
          if (range) {
            _context8.n = 8;
            break;
          }
          return _context8.a(2, false);
        case 8:
          try {
            range = range.getRange("Content");
          } catch (err) {
            warn("highlight insert: failed to focus range", err);
          }
          range.font.highlightColor = HIGHLIGHT_INSERT;
          context.trackedObjects.add(range);
          addPendingSuggestionOnline({
            id: createSuggestionId("ins", paragraphIndex, op.pos),
            kind: "insert",
            paragraphIndex: paragraphIndex,
            leftWord: lastWord,
            leftSnippet: leftSnippetStored,
            rightSnippet: rightSnippetStored,
            highlightRange: range,
            metadata: metadata
          });
          markParagraphTouched(paragraphIndex);
          return _context8.a(2, true);
      }
    }, _callee8);
  }));
  return _highlightInsertSuggestion.apply(this, arguments);
}
function findCommaRangeByOrdinal(_x39, _x40, _x41, _x42) {
  return _findCommaRangeByOrdinal.apply(this, arguments);
}
function _findCommaRangeByOrdinal() {
  _findCommaRangeByOrdinal = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9(context, paragraph, original, op) {
    var ordinal, commaSearch;
    return _regenerator().w(function (_context9) {
      while (1) switch (_context9.n) {
        case 0:
          ordinal = countCommasUpTo(original, op.pos);
          if (!(ordinal <= 0)) {
            _context9.n = 1;
            break;
          }
          warn("highlight delete: no comma ordinal", op);
          return _context9.a(2, null);
        case 1:
          commaSearch = paragraph.getRange().search(",", {
            matchCase: false,
            matchWholeWord: false
          });
          commaSearch.load("items");
          _context9.n = 2;
          return context.sync();
        case 2:
          if (!(!commaSearch.items.length || ordinal > commaSearch.items.length)) {
            _context9.n = 3;
            break;
          }
          warn("highlight delete: comma search out of range");
          return _context9.a(2, null);
        case 3:
          return _context9.a(2, commaSearch.items[ordinal - 1]);
      }
    }, _callee9);
  }));
  return _findCommaRangeByOrdinal.apply(this, arguments);
}
function extractLastWord(text) {
  var match = text.match(/((?:[0-9A-Za-z\xAA\xB5\xBA\xC0-\xD6\xD8-\xF6\xF8-\u02C1\u02C6-\u02D1\u02E0-\u02E4\u02EC\u02EE\u0370-\u0374\u0376\u0377\u037A-\u037D\u037F\u0386\u0388-\u038A\u038C\u038E-\u03A1\u03A3-\u03F5\u03F7-\u0481\u048A-\u052F\u0531-\u0556\u0559\u0560-\u0588\u05D0-\u05EA\u05EF-\u05F2\u0620-\u064A\u066E\u066F\u0671-\u06D3\u06D5\u06E5\u06E6\u06EE\u06EF\u06FA-\u06FC\u06FF\u0710\u0712-\u072F\u074D-\u07A5\u07B1\u07CA-\u07EA\u07F4\u07F5\u07FA\u0800-\u0815\u081A\u0824\u0828\u0840-\u0858\u0860-\u086A\u0870-\u0887\u0889-\u088F\u08A0-\u08C9\u0904-\u0939\u093D\u0950\u0958-\u0961\u0971-\u0980\u0985-\u098C\u098F\u0990\u0993-\u09A8\u09AA-\u09B0\u09B2\u09B6-\u09B9\u09BD\u09CE\u09DC\u09DD\u09DF-\u09E1\u09F0\u09F1\u09FC\u0A05-\u0A0A\u0A0F\u0A10\u0A13-\u0A28\u0A2A-\u0A30\u0A32\u0A33\u0A35\u0A36\u0A38\u0A39\u0A59-\u0A5C\u0A5E\u0A72-\u0A74\u0A85-\u0A8D\u0A8F-\u0A91\u0A93-\u0AA8\u0AAA-\u0AB0\u0AB2\u0AB3\u0AB5-\u0AB9\u0ABD\u0AD0\u0AE0\u0AE1\u0AF9\u0B05-\u0B0C\u0B0F\u0B10\u0B13-\u0B28\u0B2A-\u0B30\u0B32\u0B33\u0B35-\u0B39\u0B3D\u0B5C\u0B5D\u0B5F-\u0B61\u0B71\u0B83\u0B85-\u0B8A\u0B8E-\u0B90\u0B92-\u0B95\u0B99\u0B9A\u0B9C\u0B9E\u0B9F\u0BA3\u0BA4\u0BA8-\u0BAA\u0BAE-\u0BB9\u0BD0\u0C05-\u0C0C\u0C0E-\u0C10\u0C12-\u0C28\u0C2A-\u0C39\u0C3D\u0C58-\u0C5A\u0C5C\u0C5D\u0C60\u0C61\u0C80\u0C85-\u0C8C\u0C8E-\u0C90\u0C92-\u0CA8\u0CAA-\u0CB3\u0CB5-\u0CB9\u0CBD\u0CDC-\u0CDE\u0CE0\u0CE1\u0CF1\u0CF2\u0D04-\u0D0C\u0D0E-\u0D10\u0D12-\u0D3A\u0D3D\u0D4E\u0D54-\u0D56\u0D5F-\u0D61\u0D7A-\u0D7F\u0D85-\u0D96\u0D9A-\u0DB1\u0DB3-\u0DBB\u0DBD\u0DC0-\u0DC6\u0E01-\u0E30\u0E32\u0E33\u0E40-\u0E46\u0E81\u0E82\u0E84\u0E86-\u0E8A\u0E8C-\u0EA3\u0EA5\u0EA7-\u0EB0\u0EB2\u0EB3\u0EBD\u0EC0-\u0EC4\u0EC6\u0EDC-\u0EDF\u0F00\u0F40-\u0F47\u0F49-\u0F6C\u0F88-\u0F8C\u1000-\u102A\u103F\u1050-\u1055\u105A-\u105D\u1061\u1065\u1066\u106E-\u1070\u1075-\u1081\u108E\u10A0-\u10C5\u10C7\u10CD\u10D0-\u10FA\u10FC-\u1248\u124A-\u124D\u1250-\u1256\u1258\u125A-\u125D\u1260-\u1288\u128A-\u128D\u1290-\u12B0\u12B2-\u12B5\u12B8-\u12BE\u12C0\u12C2-\u12C5\u12C8-\u12D6\u12D8-\u1310\u1312-\u1315\u1318-\u135A\u1380-\u138F\u13A0-\u13F5\u13F8-\u13FD\u1401-\u166C\u166F-\u167F\u1681-\u169A\u16A0-\u16EA\u16F1-\u16F8\u1700-\u1711\u171F-\u1731\u1740-\u1751\u1760-\u176C\u176E-\u1770\u1780-\u17B3\u17D7\u17DC\u1820-\u1878\u1880-\u1884\u1887-\u18A8\u18AA\u18B0-\u18F5\u1900-\u191E\u1950-\u196D\u1970-\u1974\u1980-\u19AB\u19B0-\u19C9\u1A00-\u1A16\u1A20-\u1A54\u1AA7\u1B05-\u1B33\u1B45-\u1B4C\u1B83-\u1BA0\u1BAE\u1BAF\u1BBA-\u1BE5\u1C00-\u1C23\u1C4D-\u1C4F\u1C5A-\u1C7D\u1C80-\u1C8A\u1C90-\u1CBA\u1CBD-\u1CBF\u1CE9-\u1CEC\u1CEE-\u1CF3\u1CF5\u1CF6\u1CFA\u1D00-\u1DBF\u1E00-\u1F15\u1F18-\u1F1D\u1F20-\u1F45\u1F48-\u1F4D\u1F50-\u1F57\u1F59\u1F5B\u1F5D\u1F5F-\u1F7D\u1F80-\u1FB4\u1FB6-\u1FBC\u1FBE\u1FC2-\u1FC4\u1FC6-\u1FCC\u1FD0-\u1FD3\u1FD6-\u1FDB\u1FE0-\u1FEC\u1FF2-\u1FF4\u1FF6-\u1FFC\u2071\u207F\u2090-\u209C\u2102\u2107\u210A-\u2113\u2115\u2119-\u211D\u2124\u2126\u2128\u212A-\u212D\u212F-\u2139\u213C-\u213F\u2145-\u2149\u214E\u2183\u2184\u2C00-\u2CE4\u2CEB-\u2CEE\u2CF2\u2CF3\u2D00-\u2D25\u2D27\u2D2D\u2D30-\u2D67\u2D6F\u2D80-\u2D96\u2DA0-\u2DA6\u2DA8-\u2DAE\u2DB0-\u2DB6\u2DB8-\u2DBE\u2DC0-\u2DC6\u2DC8-\u2DCE\u2DD0-\u2DD6\u2DD8-\u2DDE\u2E2F\u3005\u3006\u3031-\u3035\u303B\u303C\u3041-\u3096\u309D-\u309F\u30A1-\u30FA\u30FC-\u30FF\u3105-\u312F\u3131-\u318E\u31A0-\u31BF\u31F0-\u31FF\u3400-\u4DBF\u4E00-\uA48C\uA4D0-\uA4FD\uA500-\uA60C\uA610-\uA61F\uA62A\uA62B\uA640-\uA66E\uA67F-\uA69D\uA6A0-\uA6E5\uA717-\uA71F\uA722-\uA788\uA78B-\uA7DC\uA7F1-\uA801\uA803-\uA805\uA807-\uA80A\uA80C-\uA822\uA840-\uA873\uA882-\uA8B3\uA8F2-\uA8F7\uA8FB\uA8FD\uA8FE\uA90A-\uA925\uA930-\uA946\uA960-\uA97C\uA984-\uA9B2\uA9CF\uA9E0-\uA9E4\uA9E6-\uA9EF\uA9FA-\uA9FE\uAA00-\uAA28\uAA40-\uAA42\uAA44-\uAA4B\uAA60-\uAA76\uAA7A\uAA7E-\uAAAF\uAAB1\uAAB5\uAAB6\uAAB9-\uAABD\uAAC0\uAAC2\uAADB-\uAADD\uAAE0-\uAAEA\uAAF2-\uAAF4\uAB01-\uAB06\uAB09-\uAB0E\uAB11-\uAB16\uAB20-\uAB26\uAB28-\uAB2E\uAB30-\uAB5A\uAB5C-\uAB69\uAB70-\uABE2\uAC00-\uD7A3\uD7B0-\uD7C6\uD7CB-\uD7FB\uF900-\uFA6D\uFA70-\uFAD9\uFB00-\uFB06\uFB13-\uFB17\uFB1D\uFB1F-\uFB28\uFB2A-\uFB36\uFB38-\uFB3C\uFB3E\uFB40\uFB41\uFB43\uFB44\uFB46-\uFBB1\uFBD3-\uFD3D\uFD50-\uFD8F\uFD92-\uFDC7\uFDF0-\uFDFB\uFE70-\uFE74\uFE76-\uFEFC\uFF21-\uFF3A\uFF41-\uFF5A\uFF66-\uFFBE\uFFC2-\uFFC7\uFFCA-\uFFCF\uFFD2-\uFFD7\uFFDA-\uFFDC]|\uD800[\uDC00-\uDC0B\uDC0D-\uDC26\uDC28-\uDC3A\uDC3C\uDC3D\uDC3F-\uDC4D\uDC50-\uDC5D\uDC80-\uDCFA\uDE80-\uDE9C\uDEA0-\uDED0\uDF00-\uDF1F\uDF2D-\uDF40\uDF42-\uDF49\uDF50-\uDF75\uDF80-\uDF9D\uDFA0-\uDFC3\uDFC8-\uDFCF]|\uD801[\uDC00-\uDC9D\uDCB0-\uDCD3\uDCD8-\uDCFB\uDD00-\uDD27\uDD30-\uDD63\uDD70-\uDD7A\uDD7C-\uDD8A\uDD8C-\uDD92\uDD94\uDD95\uDD97-\uDDA1\uDDA3-\uDDB1\uDDB3-\uDDB9\uDDBB\uDDBC\uDDC0-\uDDF3\uDE00-\uDF36\uDF40-\uDF55\uDF60-\uDF67\uDF80-\uDF85\uDF87-\uDFB0\uDFB2-\uDFBA]|\uD802[\uDC00-\uDC05\uDC08\uDC0A-\uDC35\uDC37\uDC38\uDC3C\uDC3F-\uDC55\uDC60-\uDC76\uDC80-\uDC9E\uDCE0-\uDCF2\uDCF4\uDCF5\uDD00-\uDD15\uDD20-\uDD39\uDD40-\uDD59\uDD80-\uDDB7\uDDBE\uDDBF\uDE00\uDE10-\uDE13\uDE15-\uDE17\uDE19-\uDE35\uDE60-\uDE7C\uDE80-\uDE9C\uDEC0-\uDEC7\uDEC9-\uDEE4\uDF00-\uDF35\uDF40-\uDF55\uDF60-\uDF72\uDF80-\uDF91]|\uD803[\uDC00-\uDC48\uDC80-\uDCB2\uDCC0-\uDCF2\uDD00-\uDD23\uDD4A-\uDD65\uDD6F-\uDD85\uDE80-\uDEA9\uDEB0\uDEB1\uDEC2-\uDEC7\uDF00-\uDF1C\uDF27\uDF30-\uDF45\uDF70-\uDF81\uDFB0-\uDFC4\uDFE0-\uDFF6]|\uD804[\uDC03-\uDC37\uDC71\uDC72\uDC75\uDC83-\uDCAF\uDCD0-\uDCE8\uDD03-\uDD26\uDD44\uDD47\uDD50-\uDD72\uDD76\uDD83-\uDDB2\uDDC1-\uDDC4\uDDDA\uDDDC\uDE00-\uDE11\uDE13-\uDE2B\uDE3F\uDE40\uDE80-\uDE86\uDE88\uDE8A-\uDE8D\uDE8F-\uDE9D\uDE9F-\uDEA8\uDEB0-\uDEDE\uDF05-\uDF0C\uDF0F\uDF10\uDF13-\uDF28\uDF2A-\uDF30\uDF32\uDF33\uDF35-\uDF39\uDF3D\uDF50\uDF5D-\uDF61\uDF80-\uDF89\uDF8B\uDF8E\uDF90-\uDFB5\uDFB7\uDFD1\uDFD3]|\uD805[\uDC00-\uDC34\uDC47-\uDC4A\uDC5F-\uDC61\uDC80-\uDCAF\uDCC4\uDCC5\uDCC7\uDD80-\uDDAE\uDDD8-\uDDDB\uDE00-\uDE2F\uDE44\uDE80-\uDEAA\uDEB8\uDF00-\uDF1A\uDF40-\uDF46]|\uD806[\uDC00-\uDC2B\uDCA0-\uDCDF\uDCFF-\uDD06\uDD09\uDD0C-\uDD13\uDD15\uDD16\uDD18-\uDD2F\uDD3F\uDD41\uDDA0-\uDDA7\uDDAA-\uDDD0\uDDE1\uDDE3\uDE00\uDE0B-\uDE32\uDE3A\uDE50\uDE5C-\uDE89\uDE9D\uDEB0-\uDEF8\uDFC0-\uDFE0]|\uD807[\uDC00-\uDC08\uDC0A-\uDC2E\uDC40\uDC72-\uDC8F\uDD00-\uDD06\uDD08\uDD09\uDD0B-\uDD30\uDD46\uDD60-\uDD65\uDD67\uDD68\uDD6A-\uDD89\uDD98\uDDB0-\uDDDB\uDEE0-\uDEF2\uDF02\uDF04-\uDF10\uDF12-\uDF33\uDFB0]|\uD808[\uDC00-\uDF99]|\uD809[\uDC80-\uDD43]|\uD80B[\uDF90-\uDFF0]|[\uD80C\uD80E\uD80F\uD81C-\uD822\uD840-\uD868\uD86A-\uD86D\uD86F-\uD872\uD874-\uD879\uD880-\uD883\uD885-\uD88C][\uDC00-\uDFFF]|\uD80D[\uDC00-\uDC2F\uDC41-\uDC46\uDC60-\uDFFF]|\uD810[\uDC00-\uDFFA]|\uD811[\uDC00-\uDE46]|\uD818[\uDD00-\uDD1D]|\uD81A[\uDC00-\uDE38\uDE40-\uDE5E\uDE70-\uDEBE\uDED0-\uDEED\uDF00-\uDF2F\uDF40-\uDF43\uDF63-\uDF77\uDF7D-\uDF8F]|\uD81B[\uDD40-\uDD6C\uDE40-\uDE7F\uDEA0-\uDEB8\uDEBB-\uDED3\uDF00-\uDF4A\uDF50\uDF93-\uDF9F\uDFE0\uDFE1\uDFE3\uDFF2\uDFF3]|\uD823[\uDC00-\uDCD5\uDCFF-\uDD1E\uDD80-\uDDF2]|\uD82B[\uDFF0-\uDFF3\uDFF5-\uDFFB\uDFFD\uDFFE]|\uD82C[\uDC00-\uDD22\uDD32\uDD50-\uDD52\uDD55\uDD64-\uDD67\uDD70-\uDEFB]|\uD82F[\uDC00-\uDC6A\uDC70-\uDC7C\uDC80-\uDC88\uDC90-\uDC99]|\uD835[\uDC00-\uDC54\uDC56-\uDC9C\uDC9E\uDC9F\uDCA2\uDCA5\uDCA6\uDCA9-\uDCAC\uDCAE-\uDCB9\uDCBB\uDCBD-\uDCC3\uDCC5-\uDD05\uDD07-\uDD0A\uDD0D-\uDD14\uDD16-\uDD1C\uDD1E-\uDD39\uDD3B-\uDD3E\uDD40-\uDD44\uDD46\uDD4A-\uDD50\uDD52-\uDEA5\uDEA8-\uDEC0\uDEC2-\uDEDA\uDEDC-\uDEFA\uDEFC-\uDF14\uDF16-\uDF34\uDF36-\uDF4E\uDF50-\uDF6E\uDF70-\uDF88\uDF8A-\uDFA8\uDFAA-\uDFC2\uDFC4-\uDFCB]|\uD837[\uDF00-\uDF1E\uDF25-\uDF2A]|\uD838[\uDC30-\uDC6D\uDD00-\uDD2C\uDD37-\uDD3D\uDD4E\uDE90-\uDEAD\uDEC0-\uDEEB]|\uD839[\uDCD0-\uDCEB\uDDD0-\uDDED\uDDF0\uDEC0-\uDEDE\uDEE0-\uDEE2\uDEE4\uDEE5\uDEE7-\uDEED\uDEF0-\uDEF4\uDEFE\uDEFF\uDFE0-\uDFE6\uDFE8-\uDFEB\uDFED\uDFEE\uDFF0-\uDFFE]|\uD83A[\uDC00-\uDCC4\uDD00-\uDD43\uDD4B]|\uD83B[\uDE00-\uDE03\uDE05-\uDE1F\uDE21\uDE22\uDE24\uDE27\uDE29-\uDE32\uDE34-\uDE37\uDE39\uDE3B\uDE42\uDE47\uDE49\uDE4B\uDE4D-\uDE4F\uDE51\uDE52\uDE54\uDE57\uDE59\uDE5B\uDE5D\uDE5F\uDE61\uDE62\uDE64\uDE67-\uDE6A\uDE6C-\uDE72\uDE74-\uDE77\uDE79-\uDE7C\uDE7E\uDE80-\uDE89\uDE8B-\uDE9B\uDEA1-\uDEA3\uDEA5-\uDEA9\uDEAB-\uDEBB]|\uD869[\uDC00-\uDEDF\uDF00-\uDFFF]|\uD86E[\uDC00-\uDC1D\uDC20-\uDFFF]|\uD873[\uDC00-\uDEAD\uDEB0-\uDFFF]|\uD87A[\uDC00-\uDFE0\uDFF0-\uDFFF]|\uD87B[\uDC00-\uDE5D]|\uD87E[\uDC00-\uDE1D]|\uD884[\uDC00-\uDF4A\uDF50-\uDFFF]|\uD88D[\uDC00-\uDC79])+)(?:[\0-\/:-@\[-`\{-\xA9\xAB-\xB4\xB6-\xB9\xBB-\xBF\xD7\xF7\u02C2-\u02C5\u02D2-\u02DF\u02E5-\u02EB\u02ED\u02EF-\u036F\u0375\u0378\u0379\u037E\u0380-\u0385\u0387\u038B\u038D\u03A2\u03F6\u0482-\u0489\u0530\u0557\u0558\u055A-\u055F\u0589-\u05CF\u05EB-\u05EE\u05F3-\u061F\u064B-\u066D\u0670\u06D4\u06D6-\u06E4\u06E7-\u06ED\u06F0-\u06F9\u06FD\u06FE\u0700-\u070F\u0711\u0730-\u074C\u07A6-\u07B0\u07B2-\u07C9\u07EB-\u07F3\u07F6-\u07F9\u07FB-\u07FF\u0816-\u0819\u081B-\u0823\u0825-\u0827\u0829-\u083F\u0859-\u085F\u086B-\u086F\u0888\u0890-\u089F\u08CA-\u0903\u093A-\u093C\u093E-\u094F\u0951-\u0957\u0962-\u0970\u0981-\u0984\u098D\u098E\u0991\u0992\u09A9\u09B1\u09B3-\u09B5\u09BA-\u09BC\u09BE-\u09CD\u09CF-\u09DB\u09DE\u09E2-\u09EF\u09F2-\u09FB\u09FD-\u0A04\u0A0B-\u0A0E\u0A11\u0A12\u0A29\u0A31\u0A34\u0A37\u0A3A-\u0A58\u0A5D\u0A5F-\u0A71\u0A75-\u0A84\u0A8E\u0A92\u0AA9\u0AB1\u0AB4\u0ABA-\u0ABC\u0ABE-\u0ACF\u0AD1-\u0ADF\u0AE2-\u0AF8\u0AFA-\u0B04\u0B0D\u0B0E\u0B11\u0B12\u0B29\u0B31\u0B34\u0B3A-\u0B3C\u0B3E-\u0B5B\u0B5E\u0B62-\u0B70\u0B72-\u0B82\u0B84\u0B8B-\u0B8D\u0B91\u0B96-\u0B98\u0B9B\u0B9D\u0BA0-\u0BA2\u0BA5-\u0BA7\u0BAB-\u0BAD\u0BBA-\u0BCF\u0BD1-\u0C04\u0C0D\u0C11\u0C29\u0C3A-\u0C3C\u0C3E-\u0C57\u0C5B\u0C5E\u0C5F\u0C62-\u0C7F\u0C81-\u0C84\u0C8D\u0C91\u0CA9\u0CB4\u0CBA-\u0CBC\u0CBE-\u0CDB\u0CDF\u0CE2-\u0CF0\u0CF3-\u0D03\u0D0D\u0D11\u0D3B\u0D3C\u0D3E-\u0D4D\u0D4F-\u0D53\u0D57-\u0D5E\u0D62-\u0D79\u0D80-\u0D84\u0D97-\u0D99\u0DB2\u0DBC\u0DBE\u0DBF\u0DC7-\u0E00\u0E31\u0E34-\u0E3F\u0E47-\u0E80\u0E83\u0E85\u0E8B\u0EA4\u0EA6\u0EB1\u0EB4-\u0EBC\u0EBE\u0EBF\u0EC5\u0EC7-\u0EDB\u0EE0-\u0EFF\u0F01-\u0F3F\u0F48\u0F6D-\u0F87\u0F8D-\u0FFF\u102B-\u103E\u1040-\u104F\u1056-\u1059\u105E-\u1060\u1062-\u1064\u1067-\u106D\u1071-\u1074\u1082-\u108D\u108F-\u109F\u10C6\u10C8-\u10CC\u10CE\u10CF\u10FB\u1249\u124E\u124F\u1257\u1259\u125E\u125F\u1289\u128E\u128F\u12B1\u12B6\u12B7\u12BF\u12C1\u12C6\u12C7\u12D7\u1311\u1316\u1317\u135B-\u137F\u1390-\u139F\u13F6\u13F7\u13FE-\u1400\u166D\u166E\u1680\u169B-\u169F\u16EB-\u16F0\u16F9-\u16FF\u1712-\u171E\u1732-\u173F\u1752-\u175F\u176D\u1771-\u177F\u17B4-\u17D6\u17D8-\u17DB\u17DD-\u181F\u1879-\u187F\u1885\u1886\u18A9\u18AB-\u18AF\u18F6-\u18FF\u191F-\u194F\u196E\u196F\u1975-\u197F\u19AC-\u19AF\u19CA-\u19FF\u1A17-\u1A1F\u1A55-\u1AA6\u1AA8-\u1B04\u1B34-\u1B44\u1B4D-\u1B82\u1BA1-\u1BAD\u1BB0-\u1BB9\u1BE6-\u1BFF\u1C24-\u1C4C\u1C50-\u1C59\u1C7E\u1C7F\u1C8B-\u1C8F\u1CBB\u1CBC\u1CC0-\u1CE8\u1CED\u1CF4\u1CF7-\u1CF9\u1CFB-\u1CFF\u1DC0-\u1DFF\u1F16\u1F17\u1F1E\u1F1F\u1F46\u1F47\u1F4E\u1F4F\u1F58\u1F5A\u1F5C\u1F5E\u1F7E\u1F7F\u1FB5\u1FBD\u1FBF-\u1FC1\u1FC5\u1FCD-\u1FCF\u1FD4\u1FD5\u1FDC-\u1FDF\u1FED-\u1FF1\u1FF5\u1FFD-\u2070\u2072-\u207E\u2080-\u208F\u209D-\u2101\u2103-\u2106\u2108\u2109\u2114\u2116-\u2118\u211E-\u2123\u2125\u2127\u2129\u212E\u213A\u213B\u2140-\u2144\u214A-\u214D\u214F-\u2182\u2185-\u2BFF\u2CE5-\u2CEA\u2CEF-\u2CF1\u2CF4-\u2CFF\u2D26\u2D28-\u2D2C\u2D2E\u2D2F\u2D68-\u2D6E\u2D70-\u2D7F\u2D97-\u2D9F\u2DA7\u2DAF\u2DB7\u2DBF\u2DC7\u2DCF\u2DD7\u2DDF-\u2E2E\u2E30-\u3004\u3007-\u3030\u3036-\u303A\u303D-\u3040\u3097-\u309C\u30A0\u30FB\u3100-\u3104\u3130\u318F-\u319F\u31C0-\u31EF\u3200-\u33FF\u4DC0-\u4DFF\uA48D-\uA4CF\uA4FE\uA4FF\uA60D-\uA60F\uA620-\uA629\uA62C-\uA63F\uA66F-\uA67E\uA69E\uA69F\uA6E6-\uA716\uA720\uA721\uA789\uA78A\uA7DD-\uA7F0\uA802\uA806\uA80B\uA823-\uA83F\uA874-\uA881\uA8B4-\uA8F1\uA8F8-\uA8FA\uA8FC\uA8FF-\uA909\uA926-\uA92F\uA947-\uA95F\uA97D-\uA983\uA9B3-\uA9CE\uA9D0-\uA9DF\uA9E5\uA9F0-\uA9F9\uA9FF\uAA29-\uAA3F\uAA43\uAA4C-\uAA5F\uAA77-\uAA79\uAA7B-\uAA7D\uAAB0\uAAB2-\uAAB4\uAAB7\uAAB8\uAABE\uAABF\uAAC1\uAAC3-\uAADA\uAADE\uAADF\uAAEB-\uAAF1\uAAF5-\uAB00\uAB07\uAB08\uAB0F\uAB10\uAB17-\uAB1F\uAB27\uAB2F\uAB5B\uAB6A-\uAB6F\uABE3-\uABFF\uD7A4-\uD7AF\uD7C7-\uD7CA\uD7FC-\uD7FF\uE000-\uF8FF\uFA6E\uFA6F\uFADA-\uFAFF\uFB07-\uFB12\uFB18-\uFB1C\uFB1E\uFB29\uFB37\uFB3D\uFB3F\uFB42\uFB45\uFBB2-\uFBD2\uFD3E-\uFD4F\uFD90\uFD91\uFDC8-\uFDEF\uFDFC-\uFE6F\uFE75\uFEFD-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65\uFFBF-\uFFC1\uFFC8\uFFC9\uFFD0\uFFD1\uFFD8\uFFD9\uFFDD-\uFFFF]|\uD800[\uDC0C\uDC27\uDC3B\uDC3E\uDC4E\uDC4F\uDC5E-\uDC7F\uDCFB-\uDE7F\uDE9D-\uDE9F\uDED1-\uDEFF\uDF20-\uDF2C\uDF41\uDF4A-\uDF4F\uDF76-\uDF7F\uDF9E\uDF9F\uDFC4-\uDFC7\uDFD0-\uDFFF]|\uD801[\uDC9E-\uDCAF\uDCD4-\uDCD7\uDCFC-\uDCFF\uDD28-\uDD2F\uDD64-\uDD6F\uDD7B\uDD8B\uDD93\uDD96\uDDA2\uDDB2\uDDBA\uDDBD-\uDDBF\uDDF4-\uDDFF\uDF37-\uDF3F\uDF56-\uDF5F\uDF68-\uDF7F\uDF86\uDFB1\uDFBB-\uDFFF]|\uD802[\uDC06\uDC07\uDC09\uDC36\uDC39-\uDC3B\uDC3D\uDC3E\uDC56-\uDC5F\uDC77-\uDC7F\uDC9F-\uDCDF\uDCF3\uDCF6-\uDCFF\uDD16-\uDD1F\uDD3A-\uDD3F\uDD5A-\uDD7F\uDDB8-\uDDBD\uDDC0-\uDDFF\uDE01-\uDE0F\uDE14\uDE18\uDE36-\uDE5F\uDE7D-\uDE7F\uDE9D-\uDEBF\uDEC8\uDEE5-\uDEFF\uDF36-\uDF3F\uDF56-\uDF5F\uDF73-\uDF7F\uDF92-\uDFFF]|\uD803[\uDC49-\uDC7F\uDCB3-\uDCBF\uDCF3-\uDCFF\uDD24-\uDD49\uDD66-\uDD6E\uDD86-\uDE7F\uDEAA-\uDEAF\uDEB2-\uDEC1\uDEC8-\uDEFF\uDF1D-\uDF26\uDF28-\uDF2F\uDF46-\uDF6F\uDF82-\uDFAF\uDFC5-\uDFDF\uDFF7-\uDFFF]|\uD804[\uDC00-\uDC02\uDC38-\uDC70\uDC73\uDC74\uDC76-\uDC82\uDCB0-\uDCCF\uDCE9-\uDD02\uDD27-\uDD43\uDD45\uDD46\uDD48-\uDD4F\uDD73-\uDD75\uDD77-\uDD82\uDDB3-\uDDC0\uDDC5-\uDDD9\uDDDB\uDDDD-\uDDFF\uDE12\uDE2C-\uDE3E\uDE41-\uDE7F\uDE87\uDE89\uDE8E\uDE9E\uDEA9-\uDEAF\uDEDF-\uDF04\uDF0D\uDF0E\uDF11\uDF12\uDF29\uDF31\uDF34\uDF3A-\uDF3C\uDF3E-\uDF4F\uDF51-\uDF5C\uDF62-\uDF7F\uDF8A\uDF8C\uDF8D\uDF8F\uDFB6\uDFB8-\uDFD0\uDFD2\uDFD4-\uDFFF]|\uD805[\uDC35-\uDC46\uDC4B-\uDC5E\uDC62-\uDC7F\uDCB0-\uDCC3\uDCC6\uDCC8-\uDD7F\uDDAF-\uDDD7\uDDDC-\uDDFF\uDE30-\uDE43\uDE45-\uDE7F\uDEAB-\uDEB7\uDEB9-\uDEFF\uDF1B-\uDF3F\uDF47-\uDFFF]|\uD806[\uDC2C-\uDC9F\uDCE0-\uDCFE\uDD07\uDD08\uDD0A\uDD0B\uDD14\uDD17\uDD30-\uDD3E\uDD40\uDD42-\uDD9F\uDDA8\uDDA9\uDDD1-\uDDE0\uDDE2\uDDE4-\uDDFF\uDE01-\uDE0A\uDE33-\uDE39\uDE3B-\uDE4F\uDE51-\uDE5B\uDE8A-\uDE9C\uDE9E-\uDEAF\uDEF9-\uDFBF\uDFE1-\uDFFF]|\uD807[\uDC09\uDC2F-\uDC3F\uDC41-\uDC71\uDC90-\uDCFF\uDD07\uDD0A\uDD31-\uDD45\uDD47-\uDD5F\uDD66\uDD69\uDD8A-\uDD97\uDD99-\uDDAF\uDDDC-\uDEDF\uDEF3-\uDF01\uDF03\uDF11\uDF34-\uDFAF\uDFB1-\uDFFF]|\uD808[\uDF9A-\uDFFF]|\uD809[\uDC00-\uDC7F\uDD44-\uDFFF]|[\uD80A\uD812-\uD817\uD819\uD824-\uD82A\uD82D\uD82E\uD830-\uD834\uD836\uD83C-\uD83F\uD87C\uD87D\uD87F\uD88E-\uDBFF][\uDC00-\uDFFF]|\uD80B[\uDC00-\uDF8F\uDFF1-\uDFFF]|\uD80D[\uDC30-\uDC40\uDC47-\uDC5F]|\uD810[\uDFFB-\uDFFF]|\uD811[\uDE47-\uDFFF]|\uD818[\uDC00-\uDCFF\uDD1E-\uDFFF]|\uD81A[\uDE39-\uDE3F\uDE5F-\uDE6F\uDEBF-\uDECF\uDEEE-\uDEFF\uDF30-\uDF3F\uDF44-\uDF62\uDF78-\uDF7C\uDF90-\uDFFF]|\uD81B[\uDC00-\uDD3F\uDD6D-\uDE3F\uDE80-\uDE9F\uDEB9\uDEBA\uDED4-\uDEFF\uDF4B-\uDF4F\uDF51-\uDF92\uDFA0-\uDFDF\uDFE2\uDFE4-\uDFF1\uDFF4-\uDFFF]|\uD823[\uDCD6-\uDCFE\uDD1F-\uDD7F\uDDF3-\uDFFF]|\uD82B[\uDC00-\uDFEF\uDFF4\uDFFC\uDFFF]|\uD82C[\uDD23-\uDD31\uDD33-\uDD4F\uDD53\uDD54\uDD56-\uDD63\uDD68-\uDD6F\uDEFC-\uDFFF]|\uD82F[\uDC6B-\uDC6F\uDC7D-\uDC7F\uDC89-\uDC8F\uDC9A-\uDFFF]|\uD835[\uDC55\uDC9D\uDCA0\uDCA1\uDCA3\uDCA4\uDCA7\uDCA8\uDCAD\uDCBA\uDCBC\uDCC4\uDD06\uDD0B\uDD0C\uDD15\uDD1D\uDD3A\uDD3F\uDD45\uDD47-\uDD49\uDD51\uDEA6\uDEA7\uDEC1\uDEDB\uDEFB\uDF15\uDF35\uDF4F\uDF6F\uDF89\uDFA9\uDFC3\uDFCC-\uDFFF]|\uD837[\uDC00-\uDEFF\uDF1F-\uDF24\uDF2B-\uDFFF]|\uD838[\uDC00-\uDC2F\uDC6E-\uDCFF\uDD2D-\uDD36\uDD3E-\uDD4D\uDD4F-\uDE8F\uDEAE-\uDEBF\uDEEC-\uDFFF]|\uD839[\uDC00-\uDCCF\uDCEC-\uDDCF\uDDEE\uDDEF\uDDF1-\uDEBF\uDEDF\uDEE3\uDEE6\uDEEE\uDEEF\uDEF5-\uDEFD\uDF00-\uDFDF\uDFE7\uDFEC\uDFEF\uDFFF]|\uD83A[\uDCC5-\uDCFF\uDD44-\uDD4A\uDD4C-\uDFFF]|\uD83B[\uDC00-\uDDFF\uDE04\uDE20\uDE23\uDE25\uDE26\uDE28\uDE33\uDE38\uDE3A\uDE3C-\uDE41\uDE43-\uDE46\uDE48\uDE4A\uDE4C\uDE50\uDE53\uDE55\uDE56\uDE58\uDE5A\uDE5C\uDE5E\uDE60\uDE63\uDE65\uDE66\uDE6B\uDE73\uDE78\uDE7D\uDE7F\uDE8A\uDE9C-\uDEA0\uDEA4\uDEAA\uDEBC-\uDFFF]|\uD869[\uDEE0-\uDEFF]|\uD86E[\uDC1E\uDC1F]|\uD873[\uDEAE\uDEAF]|\uD87A[\uDFE1-\uDFEF]|\uD87B[\uDE5E-\uDFFF]|\uD87E[\uDE1E-\uDFFF]|\uD884[\uDF4B-\uDF4F]|\uD88D[\uDC7A-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF])*$/);
  return match ? match[1] : "";
}
function tryApplyDeleteUsingMetadata(_x43, _x44, _x45) {
  return _tryApplyDeleteUsingMetadata.apply(this, arguments);
}
function _tryApplyDeleteUsingMetadata() {
  _tryApplyDeleteUsingMetadata = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0(context, paragraph, suggestion) {
    var _meta$sourceTokenAt, _meta$sourceTokenAfte, _meta$sourceTokenBefo, _entry$originalText3;
    var meta, commaAnchor, tokenRange, text, commaIndex, newText, commaSearch, entry, range;
    return _regenerator().w(function (_context0) {
      while (1) switch (_context0.n) {
        case 0:
          meta = suggestion === null || suggestion === void 0 ? void 0 : suggestion.metadata;
          if (meta) {
            _context0.n = 1;
            break;
          }
          return _context0.a(2, false);
        case 1:
          commaAnchor = ((_meta$sourceTokenAt = meta.sourceTokenAt) === null || _meta$sourceTokenAt === void 0 || (_meta$sourceTokenAt = _meta$sourceTokenAt.tokenText) === null || _meta$sourceTokenAt === void 0 ? void 0 : _meta$sourceTokenAt.includes(",")) && meta.sourceTokenAt || ((_meta$sourceTokenAfte = meta.sourceTokenAfter) === null || _meta$sourceTokenAfte === void 0 || (_meta$sourceTokenAfte = _meta$sourceTokenAfte.tokenText) === null || _meta$sourceTokenAfte === void 0 ? void 0 : _meta$sourceTokenAfte.includes(",")) && meta.sourceTokenAfter || ((_meta$sourceTokenBefo = meta.sourceTokenBefore) === null || _meta$sourceTokenBefo === void 0 || (_meta$sourceTokenBefo = _meta$sourceTokenBefo.tokenText) === null || _meta$sourceTokenBefo === void 0 ? void 0 : _meta$sourceTokenBefo.includes(",")) && meta.sourceTokenBefore;
          if (!commaAnchor) {
            _context0.n = 6;
            break;
          }
          _context0.n = 2;
          return findTokenRangeForAnchor(context, paragraph, commaAnchor);
        case 2:
          tokenRange = _context0.v;
          if (!tokenRange) {
            _context0.n = 6;
            break;
          }
          tokenRange.load("text");
          _context0.n = 3;
          return context.sync();
        case 3:
          text = tokenRange.text || "";
          commaIndex = text.indexOf(",");
          if (!(commaIndex >= 0)) {
            _context0.n = 4;
            break;
          }
          newText = text.slice(0, commaIndex) + text.slice(commaIndex + 1);
          tokenRange.insertText(newText, Word.InsertLocation.replace);
          return _context0.a(2, true);
        case 4:
          commaSearch = tokenRange.search(",", {
            matchCase: false,
            matchWholeWord: false
          });
          commaSearch.load("items");
          _context0.n = 5;
          return context.sync();
        case 5:
          if (!commaSearch.items.length) {
            _context0.n = 6;
            break;
          }
          commaSearch.items[0].insertText("", Word.InsertLocation.replace);
          return _context0.a(2, true);
        case 6:
          if (!(!Number.isFinite(meta.charStart) || meta.charStart < 0)) {
            _context0.n = 7;
            break;
          }
          return _context0.a(2, false);
        case 7:
          entry = getParagraphTokenAnchorsOnline(suggestion.paragraphIndex);
          _context0.n = 8;
          return getRangeForCharacterSpan(context, paragraph, (_entry$originalText3 = entry === null || entry === void 0 ? void 0 : entry.originalText) !== null && _entry$originalText3 !== void 0 ? _entry$originalText3 : paragraph.text, meta.charStart, meta.charEnd, "apply-delete", meta.highlightText);
        case 8:
          range = _context0.v;
          if (range) {
            _context0.n = 9;
            break;
          }
          return _context0.a(2, false);
        case 9:
          range.insertText("", Word.InsertLocation.replace);
          return _context0.a(2, true);
      }
    }, _callee0);
  }));
  return _tryApplyDeleteUsingMetadata.apply(this, arguments);
}
function tryApplyDeleteUsingHighlight(_x46) {
  return _tryApplyDeleteUsingHighlight.apply(this, arguments);
}
function _tryApplyDeleteUsingHighlight() {
  _tryApplyDeleteUsingHighlight = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee1(suggestion) {
    var _t2;
    return _regenerator().w(function (_context1) {
      while (1) switch (_context1.p = _context1.n) {
        case 0:
          if (suggestion !== null && suggestion !== void 0 && suggestion.highlightRange) {
            _context1.n = 1;
            break;
          }
          return _context1.a(2, false);
        case 1:
          _context1.p = 1;
          suggestion.highlightRange.insertText("", Word.InsertLocation.replace);
          return _context1.a(2, true);
        case 2:
          _context1.p = 2;
          _t2 = _context1.v;
          warn("apply delete: highlight range removal failed", _t2);
          return _context1.a(2, false);
      }
    }, _callee1, null, [[1, 2]]);
  }));
  return _tryApplyDeleteUsingHighlight.apply(this, arguments);
}
function applyDeleteSuggestionLegacy(_x47, _x48, _x49) {
  return _applyDeleteSuggestionLegacy.apply(this, arguments);
}
function _applyDeleteSuggestionLegacy() {
  _applyDeleteSuggestionLegacy = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee10(context, paragraph, suggestion) {
    var ordinal, commaSearch, idx;
    return _regenerator().w(function (_context10) {
      while (1) switch (_context10.n) {
        case 0:
          ordinal = countCommasUpTo(paragraph.text || "", suggestion.originalPos);
          if (!(ordinal <= 0)) {
            _context10.n = 1;
            break;
          }
          warn("apply delete: no ordinal");
          return _context10.a(2);
        case 1:
          commaSearch = paragraph.getRange().search(",", {
            matchCase: false,
            matchWholeWord: false
          });
          commaSearch.load("items");
          _context10.n = 2;
          return context.sync();
        case 2:
          idx = ordinal - 1;
          if (!(!commaSearch.items.length || idx >= commaSearch.items.length)) {
            _context10.n = 3;
            break;
          }
          warn("apply delete: ordinal out of range");
          return _context10.a(2);
        case 3:
          commaSearch.items[idx].insertText("", Word.InsertLocation.replace);
        case 4:
          return _context10.a(2);
      }
    }, _callee10);
  }));
  return _applyDeleteSuggestionLegacy.apply(this, arguments);
}
function applyDeleteSuggestion(_x50, _x51, _x52) {
  return _applyDeleteSuggestion.apply(this, arguments);
}
function _applyDeleteSuggestion() {
  _applyDeleteSuggestion = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee11(context, paragraph, suggestion) {
    return _regenerator().w(function (_context11) {
      while (1) switch (_context11.n) {
        case 0:
          _context11.n = 1;
          return tryApplyDeleteUsingMetadata(context, paragraph, suggestion);
        case 1:
          if (!_context11.v) {
            _context11.n = 2;
            break;
          }
          return _context11.a(2);
        case 2:
          _context11.n = 3;
          return tryApplyDeleteUsingHighlight(suggestion);
        case 3:
          if (!_context11.v) {
            _context11.n = 4;
            break;
          }
          return _context11.a(2);
        case 4:
          _context11.n = 5;
          return applyDeleteSuggestionLegacy(context, paragraph, suggestion);
        case 5:
          return _context11.a(2);
      }
    }, _callee11);
  }));
  return _applyDeleteSuggestion.apply(this, arguments);
}
function findTokenRangeForAnchor(_x53, _x54, _x55) {
  return _findTokenRangeForAnchor.apply(this, arguments);
}
function _findTokenRangeForAnchor() {
  _findTokenRangeForAnchor = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee13(context, paragraph, anchorSnapshot) {
    var fallbackOrdinal, tryFind, range, trimmed;
    return _regenerator().w(function (_context13) {
      while (1) switch (_context13.n) {
        case 0:
          if (anchorSnapshot !== null && anchorSnapshot !== void 0 && anchorSnapshot.tokenText) {
            _context13.n = 1;
            break;
          }
          return _context13.a(2, null);
        case 1:
          fallbackOrdinal = typeof anchorSnapshot.textOccurrence === "number" ? anchorSnapshot.textOccurrence : typeof anchorSnapshot.tokenIndex === "number" ? anchorSnapshot.tokenIndex : 0;
          tryFind = /*#__PURE__*/function () {
            var _ref20 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee12(text, ordinalHint) {
              var matches, ordinal, targetIndex;
              return _regenerator().w(function (_context12) {
                while (1) switch (_context12.n) {
                  case 0:
                    if (text) {
                      _context12.n = 1;
                      break;
                    }
                    return _context12.a(2, null);
                  case 1:
                    matches = paragraph.getRange().search(text, {
                      matchCase: false,
                      matchWholeWord: false
                    });
                    matches.load("items");
                    _context12.n = 2;
                    return context.sync();
                  case 2:
                    if (matches.items.length) {
                      _context12.n = 3;
                      break;
                    }
                    return _context12.a(2, null);
                  case 3:
                    ordinal = typeof ordinalHint === "number" ? ordinalHint : typeof anchorSnapshot.tokenIndex === "number" ? anchorSnapshot.tokenIndex : fallbackOrdinal;
                    targetIndex = Math.max(0, Math.min(ordinal, matches.items.length - 1));
                    return _context12.a(2, matches.items[targetIndex]);
                }
              }, _callee12);
            }));
            return function tryFind(_x79, _x80) {
              return _ref20.apply(this, arguments);
            };
          }();
          _context13.n = 2;
          return tryFind(anchorSnapshot.tokenText, anchorSnapshot.textOccurrence);
        case 2:
          range = _context13.v;
          if (!range) {
            _context13.n = 3;
            break;
          }
          return _context13.a(2, range);
        case 3:
          trimmed = anchorSnapshot.tokenText.trim();
          if (!(trimmed && trimmed !== anchorSnapshot.tokenText)) {
            _context13.n = 5;
            break;
          }
          _context13.n = 4;
          return tryFind(trimmed, anchorSnapshot.trimmedTextOccurrence);
        case 4:
          range = _context13.v;
          if (!range) {
            _context13.n = 5;
            break;
          }
          return _context13.a(2, range);
        case 5:
          return _context13.a(2, null);
      }
    }, _callee13);
  }));
  return _findTokenRangeForAnchor.apply(this, arguments);
}
function selectInsertAnchor(meta) {
  if (!meta) return null;
  var candidates = [meta.sourceTokenAfter ? {
    anchor: meta.sourceTokenAfter,
    location: Word.InsertLocation.before
  } : null, meta.sourceTokenAt ? {
    anchor: meta.sourceTokenAt,
    location: Word.InsertLocation.after
  } : null, meta.sourceTokenBefore ? {
    anchor: meta.sourceTokenBefore,
    location: Word.InsertLocation.after
  } : null, meta.targetTokenBefore ? {
    anchor: meta.targetTokenBefore,
    location: Word.InsertLocation.before
  } : null, meta.targetTokenAt ? {
    anchor: meta.targetTokenAt,
    location: Word.InsertLocation.after
  } : null].filter(Boolean);
  var _iterator = _createForOfIteratorHelper(candidates),
    _step;
  try {
    for (_iterator.s(); !(_step = _iterator.n()).done;) {
      var _candidate$anchor;
      var candidate = _step.value;
      if (candidate !== null && candidate !== void 0 && (_candidate$anchor = candidate.anchor) !== null && _candidate$anchor !== void 0 && _candidate$anchor.matched && Number.isFinite(candidate.anchor.charStart) && candidate.anchor.charStart >= 0) {
        return candidate;
      }
    }
  } catch (err) {
    _iterator.e(err);
  } finally {
    _iterator.f();
  }
  return null;
}
function tryApplyInsertUsingMetadata(_x56, _x57, _x58) {
  return _tryApplyInsertUsingMetadata.apply(this, arguments);
}
function _tryApplyInsertUsingMetadata() {
  _tryApplyInsertUsingMetadata = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee14(context, paragraph, suggestion) {
    var _entry$originalText4;
    var meta, anchorInfo, entry, range, _t3;
    return _regenerator().w(function (_context14) {
      while (1) switch (_context14.p = _context14.n) {
        case 0:
          meta = suggestion === null || suggestion === void 0 ? void 0 : suggestion.metadata;
          if (meta) {
            _context14.n = 1;
            break;
          }
          return _context14.a(2, false);
        case 1:
          anchorInfo = selectInsertAnchor(meta);
          if (anchorInfo) {
            _context14.n = 2;
            break;
          }
          return _context14.a(2, false);
        case 2:
          entry = getParagraphTokenAnchorsOnline(suggestion.paragraphIndex);
          _context14.n = 3;
          return getRangeForCharacterSpan(context, paragraph, (_entry$originalText4 = entry === null || entry === void 0 ? void 0 : entry.originalText) !== null && _entry$originalText4 !== void 0 ? _entry$originalText4 : paragraph.text, anchorInfo.anchor.charStart, anchorInfo.anchor.charEnd, "apply-insert-anchor", anchorInfo.anchor.tokenText || meta.highlightText);
        case 3:
          range = _context14.v;
          if (range) {
            _context14.n = 4;
            break;
          }
          return _context14.a(2, false);
        case 4:
          _context14.p = 4;
          if (anchorInfo.location === Word.InsertLocation.before) {
            range.insertText(",", Word.InsertLocation.before);
          } else {
            range.getRange("After").insertText(",", Word.InsertLocation.before);
          }
          _context14.n = 6;
          break;
        case 5:
          _context14.p = 5;
          _t3 = _context14.v;
          warn("apply insert metadata: failed to insert via anchor", _t3);
          return _context14.a(2, false);
        case 6:
          return _context14.a(2, true);
      }
    }, _callee14, null, [[4, 5]]);
  }));
  return _tryApplyInsertUsingMetadata.apply(this, arguments);
}
function applyInsertSuggestionLegacy(_x59, _x60, _x61) {
  return _applyInsertSuggestionLegacy.apply(this, arguments);
}
function _applyInsertSuggestionLegacy() {
  _applyInsertSuggestionLegacy = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee15(context, paragraph, suggestion) {
    var range, after;
    return _regenerator().w(function (_context15) {
      while (1) switch (_context15.n) {
        case 0:
          _context15.n = 1;
          return findRangeForInsert(context, paragraph, suggestion);
        case 1:
          range = _context15.v;
          if (range) {
            _context15.n = 2;
            break;
          }
          warn("apply insert: unable to locate range");
          return _context15.a(2);
        case 2:
          after = range.getRange("After");
          after.insertText(",", Word.InsertLocation.before);
        case 3:
          return _context15.a(2);
      }
    }, _callee15);
  }));
  return _applyInsertSuggestionLegacy.apply(this, arguments);
}
function applyInsertSuggestion(_x62, _x63, _x64) {
  return _applyInsertSuggestion.apply(this, arguments);
}
function _applyInsertSuggestion() {
  _applyInsertSuggestion = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee16(context, paragraph, suggestion) {
    return _regenerator().w(function (_context16) {
      while (1) switch (_context16.n) {
        case 0:
          _context16.n = 1;
          return tryApplyInsertUsingMetadata(context, paragraph, suggestion);
        case 1:
          if (!_context16.v) {
            _context16.n = 2;
            break;
          }
          return _context16.a(2);
        case 2:
          _context16.n = 3;
          return applyInsertSuggestionLegacy(context, paragraph, suggestion);
        case 3:
          return _context16.a(2);
      }
    }, _callee16);
  }));
  return _applyInsertSuggestion.apply(this, arguments);
}
function normalizeCommaSpacingInParagraph(_x65, _x66) {
  return _normalizeCommaSpacingInParagraph.apply(this, arguments);
}
function _normalizeCommaSpacingInParagraph() {
  _normalizeCommaSpacingInParagraph = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee17(context, paragraph) {
    var text, idx, _text, toTrim, nextChar, afterRange;
    return _regenerator().w(function (_context17) {
      while (1) switch (_context17.n) {
        case 0:
          paragraph.load("text");
          _context17.n = 1;
          return context.sync();
        case 1:
          text = paragraph.text || "";
          if (text.includes(",")) {
            _context17.n = 2;
            break;
          }
          return _context17.a(2);
        case 2:
          idx = text.length - 1;
        case 3:
          if (!(idx >= 0)) {
            _context17.n = 10;
            break;
          }
          if (!(text[idx] !== ",")) {
            _context17.n = 4;
            break;
          }
          return _context17.a(3, 9);
        case 4:
          if (!(idx > 0 && /\s/.test(text[idx - 1]))) {
            _context17.n = 6;
            break;
          }
          _context17.n = 5;
          return getRangeForCharacterSpan(context, paragraph, text, idx - 1, idx, "trim-space-before-comma", " ");
        case 5:
          toTrim = _context17.v;
          if (toTrim) {
            toTrim.insertText("", Word.InsertLocation.replace);
          }
        case 6:
          nextChar = (_text = text[idx + 1]) !== null && _text !== void 0 ? _text : "";
          if (nextChar) {
            _context17.n = 7;
            break;
          }
          return _context17.a(3, 9);
        case 7:
          if (!(!/\s/.test(nextChar) && !QUOTES.has(nextChar) && !isDigit(nextChar))) {
            _context17.n = 9;
            break;
          }
          _context17.n = 8;
          return getRangeForCharacterSpan(context, paragraph, text, idx + 1, idx + 2, "space-after-comma", nextChar);
        case 8:
          afterRange = _context17.v;
          if (afterRange) {
            afterRange.insertText(" ", Word.InsertLocation.before);
          }
        case 9:
          idx--;
          _context17.n = 3;
          break;
        case 10:
          return _context17.a(2);
      }
    }, _callee17);
  }));
  return _normalizeCommaSpacingInParagraph.apply(this, arguments);
}
function cleanupCommaSpacingForParagraphs(_x67, _x68, _x69) {
  return _cleanupCommaSpacingForParagraphs.apply(this, arguments);
}
function _cleanupCommaSpacingForParagraphs() {
  _cleanupCommaSpacingForParagraphs = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee18(context, paragraphs, indexes) {
    var _iterator2, _step2, idx, paragraph, _t4, _t5;
    return _regenerator().w(function (_context18) {
      while (1) switch (_context18.p = _context18.n) {
        case 0:
          if (indexes !== null && indexes !== void 0 && indexes.size) {
            _context18.n = 1;
            break;
          }
          return _context18.a(2);
        case 1:
          _iterator2 = _createForOfIteratorHelper(indexes);
          _context18.p = 2;
          _iterator2.s();
        case 3:
          if ((_step2 = _iterator2.n()).done) {
            _context18.n = 8;
            break;
          }
          idx = _step2.value;
          paragraph = paragraphs.items[idx];
          if (paragraph) {
            _context18.n = 4;
            break;
          }
          return _context18.a(3, 7);
        case 4:
          _context18.p = 4;
          _context18.n = 5;
          return normalizeCommaSpacingInParagraph(context, paragraph);
        case 5:
          _context18.n = 7;
          break;
        case 6:
          _context18.p = 6;
          _t4 = _context18.v;
          warn("Failed to normalize comma spacing", _t4);
        case 7:
          _context18.n = 3;
          break;
        case 8:
          _context18.n = 10;
          break;
        case 9:
          _context18.p = 9;
          _t5 = _context18.v;
          _iterator2.e(_t5);
        case 10:
          _context18.p = 10;
          _iterator2.f();
          return _context18.f(10);
        case 11:
          return _context18.a(2);
      }
    }, _callee18, null, [[4, 6], [2, 9, 10, 11]]);
  }));
  return _cleanupCommaSpacingForParagraphs.apply(this, arguments);
}
function findRangeForInsert(_x70, _x71, _x72) {
  return _findRangeForInsert.apply(this, arguments);
}
function _findRangeForInsert() {
  _findRangeForInsert = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee19(context, paragraph, suggestion) {
    var searchOpts, range, wordSearch, leftFrag, leftSearch, rightFrag, rightSearch;
    return _regenerator().w(function (_context19) {
      while (1) switch (_context19.n) {
        case 0:
          searchOpts = {
            matchCase: false,
            matchWholeWord: false
          };
          range = null;
          if (!suggestion.leftWord) {
            _context19.n = 2;
            break;
          }
          wordSearch = paragraph.getRange().search(suggestion.leftWord, {
            matchCase: false,
            matchWholeWord: true
          });
          wordSearch.load("items");
          _context19.n = 1;
          return context.sync();
        case 1:
          if (wordSearch.items.length) {
            range = wordSearch.items[wordSearch.items.length - 1];
          }
        case 2:
          leftFrag = (suggestion.leftSnippet || "").slice(-20).replace(/[\r\n]+/g, " ");
          if (!(!range && leftFrag.trim())) {
            _context19.n = 4;
            break;
          }
          leftSearch = paragraph.getRange().search(leftFrag.trim(), searchOpts);
          leftSearch.load("items");
          _context19.n = 3;
          return context.sync();
        case 3:
          if (leftSearch.items.length) {
            range = leftSearch.items[leftSearch.items.length - 1];
          }
        case 4:
          if (range) {
            _context19.n = 6;
            break;
          }
          rightFrag = (suggestion.rightSnippet || "").replace(/,/g, "").trim();
          rightFrag = rightFrag.slice(0, 8);
          if (!rightFrag) {
            _context19.n = 6;
            break;
          }
          rightSearch = paragraph.getRange().search(rightFrag, searchOpts);
          rightSearch.load("items");
          _context19.n = 5;
          return context.sync();
        case 5:
          if (rightSearch.items.length) {
            range = rightSearch.items[0];
          }
        case 6:
          return _context19.a(2, range);
      }
    }, _callee19);
  }));
  return _findRangeForInsert.apply(this, arguments);
}
function clearHighlightForSuggestion(_x73, _x74, _x75) {
  return _clearHighlightForSuggestion.apply(this, arguments);
}
function _clearHighlightForSuggestion() {
  _clearHighlightForSuggestion = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee20(context, paragraph, suggestion) {
    var _ref21, _entry$originalText5, _meta$highlightAnchor;
    var meta, entry, paragraphText, charStart, charEnd, range;
    return _regenerator().w(function (_context20) {
      while (1) switch (_context20.n) {
        case 0:
          if (suggestion) {
            _context20.n = 1;
            break;
          }
          return _context20.a(2);
        case 1:
          if (!suggestion.highlightRange) {
            _context20.n = 2;
            break;
          }
          try {
            suggestion.highlightRange.font.highlightColor = null;
            context.trackedObjects.remove(suggestion.highlightRange);
          } catch (err) {
            warn("clearHighlightForSuggestion: failed via highlightRange", err);
          } finally {
            suggestion.highlightRange = null;
          }
          return _context20.a(2);
        case 2:
          meta = suggestion.metadata;
          if (meta) {
            _context20.n = 3;
            break;
          }
          return _context20.a(2);
        case 3:
          entry = paragraphTokenAnchorsOnline[suggestion.paragraphIndex];
          paragraphText = (_ref21 = (_entry$originalText5 = entry === null || entry === void 0 ? void 0 : entry.originalText) !== null && _entry$originalText5 !== void 0 ? _entry$originalText5 : paragraph === null || paragraph === void 0 ? void 0 : paragraph.text) !== null && _ref21 !== void 0 ? _ref21 : "";
          charStart = typeof meta.highlightCharStart === "number" ? meta.highlightCharStart : meta.charStart;
          charEnd = typeof meta.highlightCharEnd === "number" ? meta.highlightCharEnd : meta.charEnd;
          if (!(!paragraph || !paragraphText || !Number.isFinite(charStart))) {
            _context20.n = 4;
            break;
          }
          return _context20.a(2);
        case 4:
          _context20.n = 5;
          return getRangeForCharacterSpan(context, paragraph, paragraphText, charStart, charEnd, "clear-highlight", meta.highlightText || ((_meta$highlightAnchor = meta.highlightAnchorTarget) === null || _meta$highlightAnchor === void 0 ? void 0 : _meta$highlightAnchor.tokenText));
        case 5:
          range = _context20.v;
          if (range) {
            range.font.highlightColor = null;
          }
        case 6:
          return _context20.a(2);
      }
    }, _callee20);
  }));
  return _clearHighlightForSuggestion.apply(this, arguments);
}
function clearOnlineSuggestionMarkers(_x76, _x77, _x78) {
  return _clearOnlineSuggestionMarkers.apply(this, arguments);
}
function _clearOnlineSuggestionMarkers() {
  _clearOnlineSuggestionMarkers = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee21(context, suggestionsOverride, paragraphs) {
    var source, clearHighlight, _iterator3, _step3, _item$suggestion, _item$paragraph, _paragraphs$items, item, suggestion, paragraph, _t6;
    return _regenerator().w(function (_context21) {
      while (1) switch (_context21.p = _context21.n) {
        case 0:
          source = Array.isArray(suggestionsOverride) && suggestionsOverride.length ? suggestionsOverride : pendingSuggestionsOnline;
          clearHighlight = function clearHighlight(sug) {
            if (!(sug !== null && sug !== void 0 && sug.highlightRange)) return;
            try {
              sug.highlightRange.font.highlightColor = null;
              context.trackedObjects.remove(sug.highlightRange);
            } catch (err) {
              warn("Failed to clear highlight", err);
            } finally {
              sug.highlightRange = null;
            }
          };
          if (source.length) {
            _context21.n = 2;
            break;
          }
          context.document.body.font.highlightColor = null;
          _context21.n = 1;
          return context.sync();
        case 1:
          return _context21.a(2);
        case 2:
          _iterator3 = _createForOfIteratorHelper(source);
          _context21.p = 3;
          _iterator3.s();
        case 4:
          if ((_step3 = _iterator3.n()).done) {
            _context21.n = 9;
            break;
          }
          item = _step3.value;
          suggestion = (_item$suggestion = item === null || item === void 0 ? void 0 : item.suggestion) !== null && _item$suggestion !== void 0 ? _item$suggestion : item;
          if (suggestion) {
            _context21.n = 5;
            break;
          }
          return _context21.a(3, 8);
        case 5:
          paragraph = (_item$paragraph = item === null || item === void 0 ? void 0 : item.paragraph) !== null && _item$paragraph !== void 0 ? _item$paragraph : paragraphs === null || paragraphs === void 0 || (_paragraphs$items = paragraphs.items) === null || _paragraphs$items === void 0 ? void 0 : _paragraphs$items[suggestion.paragraphIndex];
          if (!paragraph) {
            _context21.n = 7;
            break;
          }
          _context21.n = 6;
          return clearHighlightForSuggestion(context, paragraph, suggestion);
        case 6:
          _context21.n = 8;
          break;
        case 7:
          clearHighlight(suggestion);
        case 8:
          _context21.n = 4;
          break;
        case 9:
          _context21.n = 11;
          break;
        case 10:
          _context21.p = 10;
          _t6 = _context21.v;
          _iterator3.e(_t6);
        case 11:
          _context21.p = 11;
          _iterator3.f();
          return _context21.f(11);
        case 12:
          _context21.n = 13;
          return context.sync();
        case 13:
          if (!suggestionsOverride) {
            resetPendingSuggestionsOnline();
          }
        case 14:
          return _context21.a(2);
      }
    }, _callee21, null, [[3, 10, 11, 12]]);
  }));
  return _clearOnlineSuggestionMarkers.apply(this, arguments);
}
function applyAllSuggestionsOnline() {
  return _applyAllSuggestionsOnline.apply(this, arguments);
}
function _applyAllSuggestionsOnline() {
  _applyAllSuggestionsOnline = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee23() {
    return _regenerator().w(function (_context23) {
      while (1) switch (_context23.n) {
        case 0:
          if (pendingSuggestionsOnline.length) {
            _context23.n = 1;
            break;
          }
          return _context23.a(2);
        case 1:
          _context23.n = 2;
          return Word.run(/*#__PURE__*/function () {
            var _ref22 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee22(context) {
              var paras, touchedIndexes, processedSuggestions, _iterator4, _step4, sug, p, _t7, _t8;
              return _regenerator().w(function (_context22) {
                while (1) switch (_context22.p = _context22.n) {
                  case 0:
                    paras = context.document.body.paragraphs;
                    paras.load("items/text");
                    _context22.n = 1;
                    return context.sync();
                  case 1:
                    touchedIndexes = new Set(paragraphsTouchedOnline);
                    processedSuggestions = [];
                    _iterator4 = _createForOfIteratorHelper(pendingSuggestionsOnline);
                    _context22.p = 2;
                    _iterator4.s();
                  case 3:
                    if ((_step4 = _iterator4.n()).done) {
                      _context22.n = 11;
                      break;
                    }
                    sug = _step4.value;
                    p = paras.items[sug.paragraphIndex];
                    if (p) {
                      _context22.n = 4;
                      break;
                    }
                    return _context22.a(3, 10);
                  case 4:
                    _context22.p = 4;
                    if (!(sug.kind === "delete")) {
                      _context22.n = 6;
                      break;
                    }
                    _context22.n = 5;
                    return applyDeleteSuggestion(context, p, sug);
                  case 5:
                    _context22.n = 7;
                    break;
                  case 6:
                    _context22.n = 7;
                    return applyInsertSuggestion(context, p, sug);
                  case 7:
                    p.load("text");
                    // Keep paragraph.text up-to-date for subsequent metadata lookups.
                    // eslint-disable-next-line office-addins/no-context-sync-in-loop
                    _context22.n = 8;
                    return context.sync();
                  case 8:
                    processedSuggestions.push({
                      suggestion: sug,
                      paragraph: p
                    });
                    _context22.n = 10;
                    break;
                  case 9:
                    _context22.p = 9;
                    _t7 = _context22.v;
                    warn("applyAllSuggestionsOnline: failed to apply suggestion", _t7);
                  case 10:
                    _context22.n = 3;
                    break;
                  case 11:
                    _context22.n = 13;
                    break;
                  case 12:
                    _context22.p = 12;
                    _t8 = _context22.v;
                    _iterator4.e(_t8);
                  case 13:
                    _context22.p = 13;
                    _iterator4.f();
                    return _context22.f(13);
                  case 14:
                    _context22.n = 15;
                    return context.sync();
                  case 15:
                    _context22.n = 16;
                    return cleanupCommaSpacingForParagraphs(context, paras, touchedIndexes);
                  case 16:
                    resetParagraphsTouchedOnline();
                    _context22.n = 17;
                    return clearOnlineSuggestionMarkers(context, processedSuggestions);
                  case 17:
                    resetParagraphTokenAnchorsOnline();
                    resetPendingSuggestionsOnline();
                    context.document.body.font.highlightColor = null;
                    _context22.n = 18;
                    return context.sync();
                  case 18:
                    return _context22.a(2);
                }
              }, _callee22, null, [[4, 9], [2, 12, 13, 14]]);
            }));
            return function (_x81) {
              return _ref22.apply(this, arguments);
            };
          }());
        case 2:
          return _context23.a(2);
      }
    }, _callee23);
  }));
  return _applyAllSuggestionsOnline.apply(this, arguments);
}
function rejectAllSuggestionsOnline() {
  return _rejectAllSuggestionsOnline.apply(this, arguments);
}
/** ─────────────────────────────────────────────────────────
 *  MAIN: Preveri vejice – celoten dokument, po odstavkih
 *  ───────────────────────────────────────────────────────── */
function _rejectAllSuggestionsOnline() {
  _rejectAllSuggestionsOnline = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee25() {
    return _regenerator().w(function (_context25) {
      while (1) switch (_context25.n) {
        case 0:
          _context25.n = 1;
          return Word.run(/*#__PURE__*/function () {
            var _ref23 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee24(context) {
              var paras;
              return _regenerator().w(function (_context24) {
                while (1) switch (_context24.n) {
                  case 0:
                    paras = context.document.body.paragraphs;
                    paras.load("items/text");
                    _context24.n = 1;
                    return context.sync();
                  case 1:
                    _context24.n = 2;
                    return clearOnlineSuggestionMarkers(context, null, paras);
                  case 2:
                    context.document.body.font.highlightColor = null;
                    _context24.n = 3;
                    return context.sync();
                  case 3:
                    return _context24.a(2);
                }
              }, _callee24);
            }));
            return function (_x82) {
              return _ref23.apply(this, arguments);
            };
          }());
        case 1:
          return _context25.a(2);
      }
    }, _callee25);
  }));
  return _rejectAllSuggestionsOnline.apply(this, arguments);
}
function checkDocumentText() {
  return _checkDocumentText.apply(this, arguments);
}
function _checkDocumentText() {
  _checkDocumentText = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee26() {
    return _regenerator().w(function (_context26) {
      while (1) switch (_context26.n) {
        case 0:
          if (!(0,_utils_host_js__WEBPACK_IMPORTED_MODULE_1__.isWordOnline)()) {
            _context26.n = 1;
            break;
          }
          return _context26.a(2, checkDocumentTextOnline());
        case 1:
          return _context26.a(2, checkDocumentTextDesktop());
      }
    }, _callee26);
  }));
  return _checkDocumentText.apply(this, arguments);
}
function checkDocumentTextDesktop() {
  return _checkDocumentTextDesktop.apply(this, arguments);
}
function _checkDocumentTextDesktop() {
  _checkDocumentTextDesktop = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee28() {
    var totalInserted, totalDeleted, paragraphsProcessed, apiErrors, _t1;
    return _regenerator().w(function (_context28) {
      while (1) switch (_context28.p = _context28.n) {
        case 0:
          log("START checkDocumentText()");
          totalInserted = 0;
          totalDeleted = 0;
          paragraphsProcessed = 0;
          apiErrors = 0;
          _context28.p = 1;
          _context28.n = 2;
          return Word.run(/*#__PURE__*/function () {
            var _ref24 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee27(context) {
              var doc, trackToggleSupported, prevTrack, paras, idx, p, sourceText, trimmed, pStart, passText, pass, corrected, opsAll, ops, _iterator5, _step5, op, updated, _t9, _t0;
              return _regenerator().w(function (_context27) {
                while (1) switch (_context27.p = _context27.n) {
                  case 0:
                    // naloži in začasno vključi sledenje spremembam
                    doc = context.document;
                    trackToggleSupported = false;
                    prevTrack = false;
                    log("TrackRevisions toggle skipped; please enable Track Changes in Word if desired.");
                    _context27.p = 1;
                    // pridobi odstavke
                    paras = context.document.body.paragraphs;
                    paras.load("items/text");
                    _context27.n = 2;
                    return context.sync();
                  case 2:
                    log("Paragraphs found:", paras.items.length);
                    idx = 0;
                  case 3:
                    if (!(idx < paras.items.length)) {
                      _context27.n = 29;
                      break;
                    }
                    p = paras.items[idx];
                    sourceText = p.text || "";
                    trimmed = sourceText.trim();
                    if (trimmed) {
                      _context27.n = 4;
                      break;
                    }
                    return _context27.a(3, 28);
                  case 4:
                    if (!(trimmed.length > MAX_PARAGRAPH_CHARS)) {
                      _context27.n = 5;
                      break;
                    }
                    notifyParagraphTooLong(idx, trimmed.length);
                    return _context27.a(3, 28);
                  case 5:
                    pStart = tnow();
                    paragraphsProcessed++;
                    log("P".concat(idx, ": len=").concat(sourceText.length, " | \"").concat(SNIP(trimmed), "\""));
                    passText = sourceText;
                    pass = 0;
                  case 6:
                    if (!(pass < MAX_AUTOFIX_PASSES)) {
                      _context27.n = 27;
                      break;
                    }
                    corrected = void 0;
                    _context27.p = 7;
                    _context27.n = 8;
                    return (0,_api_apiVejice_js__WEBPACK_IMPORTED_MODULE_0__.popraviPoved)(passText);
                  case 8:
                    corrected = _context27.v;
                    _context27.n = 10;
                    break;
                  case 9:
                    _context27.p = 9;
                    _t9 = _context27.v;
                    apiErrors++;
                    warn("P".concat(idx, " pass ").concat(pass, ": API call failed -> stop paragraph"), _t9);
                    return _context27.a(3, 27);
                  case 10:
                    log("P".concat(idx, " pass ").concat(pass, ": corrected -> \"").concat(SNIP(corrected), "\""));
                    opsAll = diffCommasOnly(passText, corrected);
                    ops = filterCommaOps(passText, corrected, opsAll);
                    log("P".concat(idx, " pass ").concat(pass, ": ops candidate=").concat(opsAll.length, ", after filter=").concat(ops.length));
                    if (ops.length) {
                      _context27.n = 11;
                      break;
                    }
                    if (pass === 0) log("P".concat(idx, ": no applicable comma ops"));
                    return _context27.a(3, 27);
                  case 11:
                    _iterator5 = _createForOfIteratorHelper(ops);
                    _context27.p = 12;
                    _iterator5.s();
                  case 13:
                    if ((_step5 = _iterator5.n()).done) {
                      _context27.n = 19;
                      break;
                    }
                    op = _step5.value;
                    if (!(op.kind === "insert")) {
                      _context27.n = 16;
                      break;
                    }
                    _context27.n = 14;
                    return insertCommaAt(context, p, passText, corrected, op);
                  case 14:
                    _context27.n = 15;
                    return ensureSpaceAfterComma(context, p, passText, corrected, op);
                  case 15:
                    totalInserted++;
                    _context27.n = 18;
                    break;
                  case 16:
                    _context27.n = 17;
                    return deleteCommaAt(context, p, passText, op.pos);
                  case 17:
                    totalDeleted++;
                  case 18:
                    _context27.n = 13;
                    break;
                  case 19:
                    _context27.n = 21;
                    break;
                  case 20:
                    _context27.p = 20;
                    _t0 = _context27.v;
                    _iterator5.e(_t0);
                  case 21:
                    _context27.p = 21;
                    _iterator5.f();
                    return _context27.f(21);
                  case 22:
                    _context27.n = 23;
                    return context.sync();
                  case 23:
                    p.load("text");
                    // eslint-disable-next-line office-addins/no-context-sync-in-loop
                    _context27.n = 24;
                    return context.sync();
                  case 24:
                    updated = p.text || "";
                    if (!(!updated || updated === passText)) {
                      _context27.n = 25;
                      break;
                    }
                    return _context27.a(3, 27);
                  case 25:
                    passText = updated;
                  case 26:
                    pass++;
                    _context27.n = 6;
                    break;
                  case 27:
                    log("P".concat(idx, ": applied (ins=").concat(totalInserted, ", del=").concat(totalDeleted, ") | ").concat(Math.round(tnow() - pStart), " ms"));
                  case 28:
                    idx++;
                    _context27.n = 3;
                    break;
                  case 29:
                    _context27.p = 29;
                    // Leave trackRevisions as-is; no restore to avoid host errors
                    void prevTrack;
                    return _context27.f(29);
                  case 30:
                    return _context27.a(2);
                }
              }, _callee27, null, [[12, 20, 21, 22], [7, 9], [1,, 29, 30]]);
            }));
            return function (_x83) {
              return _ref24.apply(this, arguments);
            };
          }());
        case 2:
          log("DONE checkDocumentText() | paragraphs:", paragraphsProcessed, "| inserted:", totalInserted, "| deleted:", totalDeleted, "| apiErrors:", apiErrors);
          _context28.n = 4;
          break;
        case 3:
          _context28.p = 3;
          _t1 = _context28.v;
          errL("ERROR in checkDocumentText:", _t1);
        case 4:
          return _context28.a(2);
      }
    }, _callee28, null, [[1, 3]]);
  }));
  return _checkDocumentTextDesktop.apply(this, arguments);
}
function checkDocumentTextOnline() {
  return _checkDocumentTextOnline.apply(this, arguments);
}
function _checkDocumentTextOnline() {
  _checkDocumentTextOnline = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee30() {
    var paragraphsProcessed, suggestions, apiErrors, _t12;
    return _regenerator().w(function (_context30) {
      while (1) switch (_context30.p = _context30.n) {
        case 0:
          log("START checkDocumentTextOnline()");
          paragraphsProcessed = 0;
          suggestions = 0;
          apiErrors = 0;
          _context30.p = 1;
          _context30.n = 2;
          return Word.run(/*#__PURE__*/function () {
            var _ref25 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee29(context) {
              var paras, documentCharOffset, idx, p, original, trimmed, paragraphDocOffset, paragraphAnchors, detail, corrected, ops, _iterator6, _step6, op, marked, _t10, _t11;
              return _regenerator().w(function (_context29) {
                while (1) switch (_context29.p = _context29.n) {
                  case 0:
                    paras = context.document.body.paragraphs;
                    paras.load("items/text");
                    _context29.n = 1;
                    return context.sync();
                  case 1:
                    _context29.n = 2;
                    return clearOnlineSuggestionMarkers(context, null, paras);
                  case 2:
                    resetPendingSuggestionsOnline();
                    resetParagraphsTouchedOnline();
                    resetParagraphTokenAnchorsOnline();
                    documentCharOffset = 0;
                    idx = 0;
                  case 3:
                    if (!(idx < paras.items.length)) {
                      _context29.n = 20;
                      break;
                    }
                    p = paras.items[idx];
                    original = p.text || "";
                    trimmed = original.trim();
                    paragraphDocOffset = documentCharOffset;
                    documentCharOffset += original.length + 1;
                    paragraphAnchors = null;
                    if (trimmed) {
                      _context29.n = 4;
                      break;
                    }
                    paragraphAnchors = createParagraphTokenAnchors({
                      paragraphIndex: idx,
                      originalText: original,
                      correctedText: original,
                      sourceTokens: [],
                      targetTokens: [],
                      documentOffset: paragraphDocOffset
                    });
                    return _context29.a(3, 19);
                  case 4:
                    paragraphsProcessed++;
                    log("P".concat(idx, " ONLINE: len=").concat(original.length, " | \"").concat(SNIP(trimmed), "\""));
                    if (!(trimmed.length > MAX_PARAGRAPH_CHARS)) {
                      _context29.n = 5;
                      break;
                    }
                    notifyParagraphTooLong(idx, trimmed.length);
                    return _context29.a(3, 19);
                  case 5:
                    detail = void 0;
                    _context29.p = 6;
                    _context29.n = 7;
                    return (0,_api_apiVejice_js__WEBPACK_IMPORTED_MODULE_0__.popraviPovedDetailed)(original);
                  case 7:
                    detail = _context29.v;
                    _context29.n = 9;
                    break;
                  case 8:
                    _context29.p = 8;
                    _t10 = _context29.v;
                    apiErrors++;
                    warn("P".concat(idx, ": API call failed -> skip paragraph"), _t10);
                    paragraphAnchors = createParagraphTokenAnchors({
                      paragraphIndex: idx,
                      originalText: original,
                      correctedText: original,
                      sourceTokens: [],
                      targetTokens: [],
                      documentOffset: paragraphDocOffset
                    });
                    return _context29.a(3, 19);
                  case 9:
                    corrected = detail.correctedText;
                    paragraphAnchors = createParagraphTokenAnchors({
                      paragraphIndex: idx,
                      originalText: original,
                      correctedText: corrected,
                      sourceTokens: detail.sourceTokens,
                      targetTokens: detail.targetTokens,
                      documentOffset: paragraphDocOffset
                    });
                    if (onlyCommasChanged(original, corrected)) {
                      _context29.n = 10;
                      break;
                    }
                    log("P".concat(idx, ": API changed more than commas -> SKIP"));
                    return _context29.a(3, 19);
                  case 10:
                    ops = filterCommaOps(original, corrected, diffCommasOnly(original, corrected));
                    if (ops.length) {
                      _context29.n = 11;
                      break;
                    }
                    return _context29.a(3, 19);
                  case 11:
                    _iterator6 = _createForOfIteratorHelper(ops);
                    _context29.p = 12;
                    _iterator6.s();
                  case 13:
                    if ((_step6 = _iterator6.n()).done) {
                      _context29.n = 16;
                      break;
                    }
                    op = _step6.value;
                    _context29.n = 14;
                    return highlightSuggestionOnline(context, p, original, corrected, op, idx, paragraphAnchors);
                  case 14:
                    marked = _context29.v;
                    if (marked) suggestions++;
                  case 15:
                    _context29.n = 13;
                    break;
                  case 16:
                    _context29.n = 18;
                    break;
                  case 17:
                    _context29.p = 17;
                    _t11 = _context29.v;
                    _iterator6.e(_t11);
                  case 18:
                    _context29.p = 18;
                    _iterator6.f();
                    return _context29.f(18);
                  case 19:
                    idx++;
                    _context29.n = 3;
                    break;
                  case 20:
                    _context29.n = 21;
                    return context.sync();
                  case 21:
                    return _context29.a(2);
                }
              }, _callee29, null, [[12, 17, 18, 19], [6, 8]]);
            }));
            return function (_x84) {
              return _ref25.apply(this, arguments);
            };
          }());
        case 2:
          log("DONE checkDocumentTextOnline() | paragraphs:", paragraphsProcessed, "| suggestions:", suggestions, "| apiErrors:", apiErrors);
          _context30.n = 4;
          break;
        case 3:
          _context30.p = 3;
          _t12 = _context30.v;
          errL("ERROR in checkDocumentTextOnline:", _t12);
        case 4:
          return _context30.a(2);
      }
    }, _callee30, null, [[1, 3]]);
  }));
  return _checkDocumentTextOnline.apply(this, arguments);
}

/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ !function() {
/******/ 	__webpack_require__.h = function() { return "cb10e04a28ff7f80df25"; }
/******/ }();
/******/ 
/******/ }
);
//# sourceMappingURL=commands.16f2c544c124940be931.hot-update.js.map