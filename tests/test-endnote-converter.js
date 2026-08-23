const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const JSZip = require("jszip");
const {
  citeXml,
  documentXml,
  endnoteInstr,
  endnoteRecordXml,
  fieldBlock,
  fieldBlockWithFldData,
  makeDocx,
  paragraph,
  runText
} = require("./helpers/make-docx");

global.JSZip = JSZip;
global.EN2Zotero = {};

function loadAddonScript(relativePath) {
  const fullPath = path.join(__dirname, "..", relativePath);
  const code = fs.readFileSync(fullPath, "utf8");
  vm.runInNewContext(code, global, { filename: fullPath });
}

loadAddonScript("content/converter.js");

EN2Zotero.Matcher = {
  getCSLJSON() {
    return {
      type: "article-journal",
      title: "Blood biomarkers for traumatic brain injury",
      author: [{ family: "Smith", given: "John A." }],
      issued: { "date-parts": [[2020]] },
      DOI: "10.1016/S1474-4422(20)30123-X"
    };
  },
  buildURI() {
    return "http://zotero.org/users/local/test/items/ABC123";
  }
};

async function readDocumentXml(docxData) {
  const zip = await JSZip.loadAsync(docxData);
  return zip.file("word/document.xml").async("string");
}

function fakeMatchResult(sourceId = "word/document.xml:0") {
  return {
    citation: {
      fieldIndex: 0,
      sourceId,
      endnoteId: sourceId,
      paperpileId: sourceId,
      formattedText: "(Smith et al., 2020)",
      isInFootnote: false,
      isInEndnote: false
    },
    itemMatches: [{
      paperpileItemId: "39",
      endnoteRecNum: "39",
      cslData: { DOI: "10.1016/S1474-4422(20)30123-X" },
      locator: "45-50",
      locatorType: "page",
      prefix: "see ",
      suffix: ", ch. 2",
      suppressAuthor: true,
      match: { id: 101, key: "ABC123" },
      matchType: "doi",
      confidence: 100
    }]
  };
}

async function testReplacesEndNotePairWithZoteroField() {
  const cite = endnoteInstr(
    "EN.CITE",
    citeXml({ recNum: "39", attrs: 'ExcludeAuth="1" PageNum="45-50"' })
  );
  const data = endnoteInstr(
    "EN.CITE.DATA",
    citeXml({ recNum: "39", record: endnoteRecordXml() })
  );
  const docxData = await makeDocx({
    document: documentXml(paragraph(
      runText("Before ") +
      fieldBlock(cite, "(Smith et al., 2020)") +
      fieldBlock(data, "") +
      runText(" after.")
    ))
  });

  const converted = await EN2Zotero.Converter.convert(docxData, [fakeMatchResult()], {
    fieldMode: "fieldcodes"
  });
  const xml = await readDocumentXml(converted);

  assert.ok(xml.includes("ADDIN ZOTERO_ITEM CSL_CITATION"));
  assert.ok(xml.includes("&quot;locator&quot;:&quot;45-50&quot;"));
  assert.ok(xml.includes("&quot;suppress-author&quot;:true"));
  assert.ok(!xml.includes("ADDIN EN.CITE "));
  assert.ok(!xml.includes("ADDIN EN.CITE.DATA"));
  assert.ok(xml.includes("(Smith et al., 2020)"));
  assert.ok(xml.includes("Before "));
  assert.ok(xml.includes(" after."));
}

async function testReplacesNestedFldDataEndNoteField() {
  const metadata = `<EndNote>${citeXml({ recNum: "39", record: endnoteRecordXml() })}</EndNote>`;
  const docxData = await makeDocx({
    document: documentXml(paragraph(
      runText("Before ") +
      fieldBlockWithFldData("ADDIN EN.CITE", metadata, "(Smith et al., 2020)", {
        nestedDataXml: metadata
      }) +
      runText(" after.")
    ))
  });

  const converted = await EN2Zotero.Converter.convert(docxData, [fakeMatchResult()], {
    fieldMode: "fieldcodes"
  });
  const xml = await readDocumentXml(converted);

  assert.ok(xml.includes("ADDIN ZOTERO_ITEM CSL_CITATION"));
  assert.ok(!xml.includes("ADDIN EN.CITE "));
  assert.ok(!xml.includes("ADDIN EN.CITE.DATA"));
  assert.ok(!xml.includes("<w:fldData"));
  assert.ok(xml.includes("(Smith et al., 2020)"));
  assert.ok(xml.includes("Before "));
  assert.ok(xml.includes(" after."));
}

async function testIgnoresExistingZoteroFieldBeforeNestedFldDataEndNoteField() {
  const metadata = `<EndNote>${citeXml({ recNum: "39", record: endnoteRecordXml() })}</EndNote>`;
  const existingZoteroField = fieldBlock(
    'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"existing","citationItems":[]}',
    "(Existing Zotero)"
  );
  const docxData = await makeDocx({
    document: documentXml(paragraph(
      existingZoteroField +
      runText(" before ") +
      fieldBlockWithFldData("ADDIN EN.CITE", metadata, "(Smith et al., 2020)", {
        nestedDataXml: metadata
      })
    ))
  });

  const converted = await EN2Zotero.Converter.convert(docxData, [fakeMatchResult("word/document.xml:0")], {
    fieldMode: "fieldcodes"
  });
  const xml = await readDocumentXml(converted);

  assert.ok(xml.includes("(Existing Zotero)"));
  assert.ok(xml.includes("(Smith et al., 2020)"));
  assert.strictEqual((xml.match(/ADDIN ZOTERO_ITEM CSL_CITATION/g) || []).length, 2);
  assert.ok(!xml.includes("ADDIN EN.CITE "));
  assert.ok(!xml.includes("ADDIN EN.CITE.DATA"));
}

async function testReplacesEndNoteBibliographyWithZoteroBibliography() {
  const docxData = await makeDocx({
    document: documentXml(paragraph(fieldBlock("ADDIN EN.REFLIST", "References")))
  });

  const converted = await EN2Zotero.Converter.convert(docxData, [], {
    fieldMode: "fieldcodes"
  });
  const xml = await readDocumentXml(converted);

  assert.ok(xml.includes("ADDIN ZOTERO_BIBL"));
  assert.ok(xml.includes("CSL_BIBLIOGRAPHY"));
  assert.ok(!xml.includes("ADDIN EN.REFLIST"));
  assert.ok(xml.includes("References"));
}

async function testBookmarksUseEndNoteSourceIds() {
  const cite = endnoteInstr("EN.CITE", citeXml({ recNum: "39" }));
  const data = endnoteInstr("EN.CITE.DATA", citeXml({ recNum: "39", record: endnoteRecordXml() }));
  const docxData = await makeDocx({
    document: documentXml(paragraph(fieldBlock(cite, "(Smith et al., 2020)") + fieldBlock(data, "")))
  });

  const converted = await EN2Zotero.Converter.convert(docxData, [fakeMatchResult()], {
    fieldMode: "bookmarks"
  });
  const zip = await JSZip.loadAsync(converted);
  const document = await zip.file("word/document.xml").async("string");
  const settings = await zip.file("word/settings.xml").async("string");

  assert.ok(document.includes("w:bookmarkStart"));
  assert.ok(!document.includes("ADDIN EN.CITE "));
  assert.ok(!document.includes("ADDIN EN.CITE.DATA"));
  assert.ok(settings.includes("ZOTERO_PREF_1"));
  assert.ok(settings.includes("ADDIN ZOTERO_ITEM CSL_CITATION"));
}

async function testReferenceMarksAreDisabledForMVP() {
  const docxData = await makeDocx({
    document: documentXml(paragraph(runText("No citations")))
  });

  await assert.rejects(
    () => EN2Zotero.Converter.convert(docxData, [], { fieldMode: "referencemarks" }),
    /Reference Marks/
  );
}

async function main() {
  await testReplacesEndNotePairWithZoteroField();
  await testReplacesNestedFldDataEndNoteField();
  await testIgnoresExistingZoteroFieldBeforeNestedFldDataEndNoteField();
  await testReplacesEndNoteBibliographyWithZoteroBibliography();
  await testBookmarksUseEndNoteSourceIds();
  await testReferenceMarksAreDisabledForMVP();
  console.log("EndNote converter tests passed");
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
