const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

global.EN2Zotero = {};

function loadAddonScript(relativePath) {
  const fullPath = path.join(__dirname, "..", relativePath);
  const code = fs.readFileSync(fullPath, "utf8");
  vm.runInNewContext(code, global, { filename: fullPath });
}

loadAddonScript("content/matcher.js");

function testMatchAllKeepsEndNoteProvenance() {
  EN2Zotero.Matcher._doiIndex = new Map();
  EN2Zotero.Matcher._titleYearIndex = new Map();
  EN2Zotero.Matcher._titleOnlyIndex = new Map();
  EN2Zotero.Matcher._itemCache = new Map();

  const results = EN2Zotero.Matcher.matchAll([{
    sourceId: "word/document.xml:0",
    citationItems: [{
      paperpileItemId: "39",
      endnoteRecNum: "39",
      endnoteDbId: "DB-1",
      cslData: { type: "legal_case", title: "Example Case" },
      suppressYear: true
    }]
  }], "doi_title_year");

  assert.strictEqual(results.length, 1);
  assert.strictEqual(results[0].itemMatches[0].endnoteRecNum, "39");
  assert.strictEqual(results[0].itemMatches[0].endnoteDbId, "DB-1");
  assert.strictEqual(results[0].itemMatches[0].suppressYear, true);
}

function testMapsEndNoteSpecificCSLTypes() {
  assert.strictEqual(EN2Zotero.Matcher._mapCSLTypeToZotero("article-newspaper"), "newspaperArticle");
  assert.strictEqual(EN2Zotero.Matcher._mapCSLTypeToZotero("article-magazine"), "magazineArticle");
  assert.strictEqual(EN2Zotero.Matcher._mapCSLTypeToZotero("legislation"), "statute");
  assert.strictEqual(EN2Zotero.Matcher._mapCSLTypeToZotero("legal_case"), "case");
  assert.strictEqual(EN2Zotero.Matcher._mapCSLTypeToZotero("motion_picture"), "film");
}

testMatchAllKeepsEndNoteProvenance();
testMapsEndNoteSpecificCSLTypes();
console.log("EndNote matcher tests passed");
