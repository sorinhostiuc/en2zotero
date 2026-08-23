const JSZip = require("jszip");

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runText(text) {
  return `<w:r><w:t>${escapeXml(text)}</w:t></w:r>`;
}

function instrRuns(instrText, chunkSize) {
  const chunks = [];
  if (chunkSize && instrText.length > chunkSize) {
    for (let i = 0; i < instrText.length; i += chunkSize) {
      chunks.push(instrText.slice(i, i + chunkSize));
    }
  } else {
    chunks.push(instrText);
  }
  return chunks
    .map(chunk => `<w:r><w:instrText xml:space="preserve">${escapeXml(chunk)}</w:instrText></w:r>`)
    .join("");
}

function fieldBlock(instrText, displayText, options = {}) {
  const display = displayText ? runText(displayText) : "";
  return [
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    instrRuns(instrText, options.chunkSize || 0),
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    display,
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  ].join("");
}

function fldDataValue(xml) {
  const raw = Buffer.from(xml, "utf8").toString("base64");
  return raw.replace(/(.{76})/g, "$1\r\n");
}

function fieldBlockWithFldData(instrText, fldDataXml, displayText, options = {}) {
  const display = displayText ? runText(displayText) : "";
  const nestedData = options.nestedDataXml
    ? [
        `<w:r><w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">${fldDataValue(options.nestedDataXml)}</w:fldData></w:fldChar></w:r>`,
        '<w:r><w:instrText xml:space="preserve"> ADDIN EN.CITE.DATA </w:instrText></w:r>',
        '<w:r></w:r>',
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
      ].join("")
    : "";
  return [
    `<w:r><w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">${fldDataValue(fldDataXml)}</w:fldData></w:fldChar></w:r>`,
    instrRuns(instrText, options.chunkSize || 0),
    nestedData,
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    display,
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  ].join("");
}

function paragraph(innerXml) {
  return `<w:p>${innerXml}</w:p>`;
}

function documentXml(bodyInner) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyInner}<w:sectPr/></w:body>
</w:document>`;
}

function footnotesXml(innerXml) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"/>
  <w:footnote w:type="continuationSeparator" w:id="0"/>
  <w:footnote w:id="2">${innerXml}</w:footnote>
</w:footnotes>`;
}

async function makeDocx({ document, footnotes, endnotes }) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.folder("word").file("document.xml", document);
  zip.folder("word").file("settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`);
  if (footnotes) zip.folder("word").file("footnotes.xml", footnotes);
  if (endnotes) zip.folder("word").file("endnotes.xml", endnotes);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

function endnoteRecordXml(overrides = {}) {
  const recNum = overrides.recNum || "39";
  const refType = overrides.refType || "Journal Article";
  const title = overrides.title || "Blood biomarkers for traumatic brain injury";
  const container = overrides.container || "Lancet Neurology";
  const year = overrides.year || "2020";
  const doi = overrides.doi || "10.1016/S1474-4422(20)30123-X";
  return `<record>
    <rec-number>${recNum}</rec-number>
    <foreign-keys><key app="EN" db-id="${overrides.dbId || "DB-1"}">${recNum}</key></foreign-keys>
    <ref-type name="${refType}">17</ref-type>
    <contributors>
      <authors>
        <author>Smith, John A.</author>
        <author>Johnson, Kate</author>
      </authors>
    </contributors>
    <titles>
      <title>${title}</title>
      <secondary-title>${container}</secondary-title>
    </titles>
    <periodical><full-title>${container}</full-title><abbr-1>Lancet Neurol</abbr-1></periodical>
    <pages>234-245</pages>
    <volume>19</volume>
    <number>3</number>
    <dates><year>${year}</year></dates>
    <isbn>1474-4422</isbn>
    <electronic-resource-num>${doi}</electronic-resource-num>
    <accession-num>12345678</accession-num>
    <urls><related-urls><url>https://doi.org/${doi}</url></related-urls></urls>
  </record>`;
}

function citeXml({ recNum = "39", attrs = "", record = "", author = "Smith", year = "2020" } = {}) {
  return `<Cite${attrs ? " " + attrs : ""}>
    <Author>${author}</Author>
    <Year>${year}</Year>
    <RecNum>${recNum}</RecNum>
    <DisplayText>(${author} et al., ${year})</DisplayText>
    ${record}
  </Cite>`;
}

function endnoteInstr(prefix, cites) {
  return `ADDIN ${prefix} <EndNote>${cites}</EndNote>`;
}

module.exports = {
  documentXml,
  endnoteInstr,
  endnoteRecordXml,
  fieldBlock,
  fieldBlockWithFldData,
  footnotesXml,
  makeDocx,
  paragraph,
  runText,
  citeXml
};
