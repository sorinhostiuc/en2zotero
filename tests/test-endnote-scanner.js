const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");
const {
  citeXml,
  documentXml,
  endnoteInstr,
  endnoteRecordXml,
  fieldBlock,
  fieldBlockWithFldData,
  footnotesXml,
  makeDocx,
  paragraph,
  runText
} = require("./helpers/make-docx");

global.DOMParser = DOMParser;
global.JSZip = JSZip;
global.EN2Zotero = {};

function loadAddonScript(relativePath) {
  const fullPath = path.join(__dirname, "..", relativePath);
  const code = fs.readFileSync(fullPath, "utf8");
  vm.runInNewContext(code, global, { filename: fullPath });
}

loadAddonScript("content/scanner.js");

async function scanDocument(documentBody, extras = {}) {
  const docxData = await makeDocx({
    document: documentXml(documentBody),
    footnotes: extras.footnotes,
    endnotes: extras.endnotes
  });
  return EN2Zotero.Scanner.scan(docxData);
}

async function testParsesEndNoteTravelingLibraryPair() {
  const cite = endnoteInstr(
    "EN.CITE",
    citeXml({
      recNum: "39",
      attrs: 'ExcludeAuth="1" PrefixText="see " SuffixText=", ch. 2" PageNum="45-50"'
    })
  );
  const data = endnoteInstr(
    "EN.CITE.DATA",
    citeXml({ recNum: "39", record: endnoteRecordXml() })
  );
  const body = paragraph(
    runText("Before ") +
    fieldBlock(cite, "(Smith et al., 2020)") +
    fieldBlock(data, "") +
    runText(" after.")
  );

  const results = await scanDocument(body);

  assert.strictEqual(results.errors.length, 0, JSON.stringify(results.errors));
  assert.strictEqual(results.citations.length, 1);
  const citation = results.citations[0];
  assert.strictEqual(citation.sourceId, "word/document.xml:0");
  assert.strictEqual(citation.hasDataField, true);
  assert.strictEqual(citation.formattedText, "(Smith et al., 2020)");
  assert.strictEqual(citation.citationItems.length, 1);

  const item = citation.citationItems[0];
  assert.strictEqual(item.endnoteRecNum, "39");
  assert.strictEqual(item.endnoteDbId, "DB-1");
  assert.strictEqual(item.cslData.DOI, "10.1016/S1474-4422(20)30123-X");
  assert.strictEqual(item.cslData.PMID, "12345678");
  assert.strictEqual(item.cslData.type, "article-journal");
  assert.strictEqual(item.cslData.title, "Blood biomarkers for traumatic brain injury");
  assert.strictEqual(item.cslData["container-title"], "Lancet Neurology");
  assert.strictEqual(item.cslData.journalAbbreviation, "Lancet Neurol");
  assert.strictEqual(item.cslData.author[0].family, "Smith");
  assert.strictEqual(item.cslData.author[0].given, "John A.");
  assert.strictEqual(item.locator, "45-50");
  assert.strictEqual(item.locatorType, "page");
  assert.strictEqual(item.prefix, "see ");
  assert.strictEqual(item.suffix, ", ch. 2");
  assert.strictEqual(item.suppressAuthor, true);
}

async function testParsesEndNoteFldDataWithNestedDataField() {
  const metadata = `<EndNote>${
    citeXml({ recNum: "65", record: endnoteRecordXml({ recNum: "65", title: "Patient-centred medical ethics", year: "2018" }) })
  }${
    citeXml({ recNum: "72", author: "Walter", year: "2014", record: endnoteRecordXml({ recNum: "72", title: "Ethics and family medicine", year: "2014" }) })
  }</EndNote>`;
  const body = paragraph(
    fieldBlockWithFldData("ADDIN EN.CITE", metadata, "(Tunzi and Ventres, 2018, Walter and Ross, 2014)", {
      nestedDataXml: metadata
    })
  );

  const results = await scanDocument(body);

  assert.strictEqual(results.errors.length, 0, JSON.stringify(results.errors));
  assert.strictEqual(results.citations.length, 1);
  assert.strictEqual(results.citations[0].sourceId, "word/document.xml:0");
  assert.strictEqual(results.citations[0].hasDataField, true);
  assert.strictEqual(results.citations[0].citationItems.length, 2);
  assert.strictEqual(results.citations[0].citationItems[0].endnoteRecNum, "65");
  assert.strictEqual(results.citations[0].citationItems[0].cslData.title, "Patient-centred medical ethics");
  assert.strictEqual(results.citations[0].citationItems[1].endnoteRecNum, "72");
  assert.strictEqual(results.citations[0].citationItems[1].cslData.title, "Ethics and family medicine");
}

async function testParsesHexEncodedEndNoteDataField() {
  const metadata = `<EndNote>${citeXml({
    recNum: "88",
    record: endnoteRecordXml({
      recNum: "88",
      title: "Hex encoded EndNote traveling library",
      year: "2024",
      doi: "10.1000/hex-endnote"
    })
  })}</EndNote>`;
  const hexMetadata = Buffer.from(metadata, "utf8").toString("hex").toUpperCase();
  const body = paragraph(
    fieldBlock("ADDIN EN.CITE", "(Smith et al., 2024)") +
    fieldBlock("ADDIN EN.CITE.DATA " + hexMetadata, "")
  );

  const results = await scanDocument(body);

  assert.strictEqual(results.errors.length, 0, JSON.stringify(results.errors));
  assert.strictEqual(results.citations.length, 1);
  assert.strictEqual(results.citations[0].hasDataField, true);
  assert.strictEqual(results.citations[0].citationItems.length, 1);
  assert.strictEqual(results.citations[0].citationItems[0].endnoteRecNum, "88");
  assert.strictEqual(results.citations[0].citationItems[0].cslData.title, "Hex encoded EndNote traveling library");
  assert.strictEqual(results.citations[0].citationItems[0].cslData.DOI, "10.1000/hex-endnote");
}

async function testParsesClusterAndFragmentedInstrText() {
  const cite = endnoteInstr(
    "EN.CITE",
    citeXml({ recNum: "39" }) + citeXml({ recNum: "42", author: "Jones", year: "2019" })
  );
  const data = endnoteInstr(
    "EN.CITE.DATA",
    citeXml({ recNum: "39", record: endnoteRecordXml({ recNum: "39" }) }) +
    citeXml({
      recNum: "42",
      author: "Jones",
      year: "2019",
      record: endnoteRecordXml({
        recNum: "42",
        title: "Neurotrauma outcomes in emergency medicine",
        year: "2019",
        doi: "10.1000/example-42"
      })
    })
  );
  const body = paragraph(
    fieldBlock(cite, "(Smith et al., 2020; Jones et al., 2019)") +
    fieldBlock(data, "", { chunkSize: 80 })
  );

  const results = await scanDocument(body);

  assert.strictEqual(results.errors.length, 0, JSON.stringify(results.errors));
  assert.strictEqual(results.citations.length, 1);
  assert.strictEqual(results.citations[0].citationItems.length, 2);
  assert.strictEqual(results.citations[0].citationItems[0].endnoteRecNum, "39");
  assert.strictEqual(results.citations[0].citationItems[1].endnoteRecNum, "42");
  assert.strictEqual(results.citations[0].citationItems[1].cslData.title, "Neurotrauma outcomes in emergency medicine");
}

async function testScansFootnotes() {
  const cite = endnoteInstr("EN.CITE", citeXml({ recNum: "39" }));
  const data = endnoteInstr("EN.CITE.DATA", citeXml({ recNum: "39", record: endnoteRecordXml() }));
  const footnote = footnotesXml(paragraph(fieldBlock(cite, "Smith, Blood biomarkers") + fieldBlock(data, "")));

  const results = await scanDocument(paragraph(runText("Body text")), { footnotes: footnote });

  assert.strictEqual(results.errors.length, 0, JSON.stringify(results.errors));
  assert.strictEqual(results.citations.length, 1);
  assert.strictEqual(results.citations[0].isInFootnote, true);
  assert.strictEqual(results.citations[0].footnoteId, "2");
  assert.strictEqual(results.citations[0].sourceId, "word/footnotes.xml:0");
}

async function testFallsBackWhenDataFieldIsMissing() {
  const cite = endnoteInstr(
    "EN.CITE",
    citeXml({ recNum: "39", record: endnoteRecordXml({ recNum: "39" }) })
  );
  const body = paragraph(fieldBlock(cite, "(Smith et al., 2020)"));

  const results = await scanDocument(body);

  assert.strictEqual(results.errors.length, 0, JSON.stringify(results.errors));
  assert.strictEqual(results.citations.length, 1);
  assert.strictEqual(results.citations[0].hasDataField, false);
  assert.strictEqual(results.citations[0].citationItems[0].cslData.title, "Blood biomarkers for traumatic brain injury");
}

async function testDetectsEndNoteBibliography() {
  const body = paragraph(fieldBlock("ADDIN EN.REFLIST", "References"));

  const results = await scanDocument(body);

  assert.ok(results.bibliography);
  assert.strictEqual(results.bibliography.sourceId, "word/document.xml:bibliography:0");
  assert.strictEqual(results.bibliography.formattedText, "References");
}

function testSanitizesBareAmpersandsInEndNoteXml() {
  const xml = "<EndNote><Cite><record><titles><title>A & B &amp; C &#xD; D</title></titles></record></Cite></EndNote>";
  const sanitized = EN2Zotero.Scanner._sanitizeEndNoteXml(xml);

  assert.ok(sanitized.includes("A &amp; B &amp; C &#xD; D"));
}

async function main() {
  await testParsesEndNoteTravelingLibraryPair();
  await testParsesEndNoteFldDataWithNestedDataField();
  await testParsesHexEncodedEndNoteDataField();
  await testParsesClusterAndFragmentedInstrText();
  await testScansFootnotes();
  await testFallsBackWhenDataFieldIsMissing();
  await testDetectsEndNoteBibliography();
  testSanitizesBareAmpersandsInEndNoteXml();
  console.log("EndNote scanner tests passed");
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
