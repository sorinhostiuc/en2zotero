/* Scanner Module - extracts EndNote CWYW citations from .docx OOXML */

if (typeof EN2Zotero === "undefined") var EN2Zotero = {};

EN2Zotero.Scanner = {
  FIELD_PREFIX_CITE: "ADDIN EN.CITE",
  FIELD_PREFIX_DATA: "ADDIN EN.CITE.DATA",
  FIELD_PREFIX_REFLIST: "ADDIN EN.REFLIST",
  W_NS: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",

  async scan(docxData) {
    if (docxData && docxData.byteLength !== undefined && !(docxData instanceof Uint8Array)) {
      docxData = new Uint8Array(docxData);
    }
    const JSZipRef = typeof JSZip !== "undefined" ? JSZip : (await import("./lib/jszip.min.js")).default;
    const zip = await JSZipRef.loadAsync(docxData);

    const results = { citations: [], bibliography: null, errors: [] };
    const xmlFiles = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];

    for (const xmlPath of xmlFiles) {
      const file = zip.file(xmlPath);
      if (!file) continue;
      const text = await file.async("string");
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(text, "application/xml");
      const context = {
        xmlPath,
        isFootnote: xmlPath === "word/footnotes.xml",
        isEndnote: xmlPath === "word/endnotes.xml"
      };
      this._extractFields(xmlDoc, results, context);
    }

    return results;
  },

  _extractFields(xmlDoc, results, context) {
    const fields = this._collectFields(xmlDoc);
    let localEndNoteIndex = 0;
    let bibliographyIndex = 0;

    for (let i = 0; i < fields.length; i++) {
      const field = fields[i];
      const type = this._classifyEndNoteField(field.instrText);

      if (type === "reflist") {
        results.bibliography = {
          sourceId: context.xmlPath + ":bibliography:" + bibliographyIndex++,
          rawInstrText: field.instrText,
          formattedText: field.displayText,
          fieldBeginIndex: field.beginIndex,
          fieldEndIndex: field.endIndex,
          isInFootnote: context.isFootnote,
          isInEndnote: context.isEndnote,
          footnoteId: field.footnoteId || null
        };
        continue;
      }

      if (type !== "cite") continue;

      const nextField = fields[i + 1] || null;
      const nextType = nextField ? this._classifyEndNoteField(nextField.instrText) : "other";
      const dataField = nextType === "data" ? nextField : null;
      const sourceId = context.xmlPath + ":" + localEndNoteIndex++;

      this._processEndNoteCitation(field, dataField, results, context, sourceId);
      if (dataField) i++;
    }
  },

  _collectFields(xmlDoc) {
    const fields = [];
    const allRuns = xmlDoc.getElementsByTagNameNS(this.W_NS, "r");
    let i = 0;

    while (i < allRuns.length) {
      const run = allRuns[i];
      const fldChar = run.getElementsByTagNameNS(this.W_NS, "fldChar")[0];

      if (fldChar && fldChar.getAttribute("w:fldCharType") === "begin") {
        const field = this._collectField(allRuns, i);
        if (field) {
          fields.push(field);
          i = field.endIndex + 1;
          continue;
        }
      }
      i++;
    }

    return fields;
  },

  _collectField(allRuns, beginIndex) {
    const instrTextParts = [];
    const displayTextParts = [];
    const fldDataParts = [];
    const nestedFields = [];
    let currentNestedField = null;
    let phase = "instr";
    let endIndex = beginIndex;
    let nestLevel = 1;
    let footnoteId = null;

    this._appendFldData(allRuns[beginIndex], fldDataParts);

    for (let i = beginIndex + 1; i < allRuns.length; i++) {
      const run = allRuns[i];
      const fldChar = run.getElementsByTagNameNS(this.W_NS, "fldChar")[0];

      if (fldChar) {
        const charType = fldChar.getAttribute("w:fldCharType");
        if (charType === "begin") {
          if (nestLevel === 1) {
            currentNestedField = { instrTextParts: [], fldDataParts: [] };
            this._appendFldData(run, currentNestedField.fldDataParts);
          }
          nestLevel++;
        } else if (charType === "separate") {
          if (nestLevel === 1) phase = "display";
        } else if (charType === "end") {
          if (nestLevel === 2 && currentNestedField) {
            nestedFields.push({
              instrText: currentNestedField.instrTextParts.join("").trim(),
              fldDataText: currentNestedField.fldDataParts.join("").trim()
            });
            currentNestedField = null;
          }
          nestLevel--;
          if (nestLevel === 0) {
            endIndex = i;
            break;
          }
        }
        continue;
      }

      if (phase === "instr") {
        const instrEl = run.getElementsByTagNameNS(this.W_NS, "instrText")[0];
        if (instrEl) {
          if (nestLevel === 1) {
            instrTextParts.push(instrEl.textContent);
          } else if (nestLevel === 2 && currentNestedField) {
            currentNestedField.instrTextParts.push(instrEl.textContent);
          }
        }
      } else if (phase === "display") {
        const tEl = run.getElementsByTagNameNS(this.W_NS, "t")[0];
        if (tEl && nestLevel === 1) displayTextParts.push(tEl.textContent);
      }
    }

    let fnParent = allRuns[beginIndex];
    while (fnParent && fnParent.parentNode) {
      fnParent = fnParent.parentNode;
      if (fnParent.localName === "footnote") {
        footnoteId = fnParent.getAttribute("w:id");
        break;
      }
      if (fnParent.localName === "endnote") {
        footnoteId = fnParent.getAttribute("w:id");
        break;
      }
    }

    const instrText = instrTextParts.join("").trim();
    if (!instrText) return null;

    return {
      instrText,
      fldDataText: fldDataParts.join("").trim(),
      nestedFields,
      displayText: displayTextParts.join(""),
      beginIndex,
      endIndex,
      footnoteId
    };
  },

  _appendFldData(run, target) {
    if (!run) return;
    const fldData = run.getElementsByTagNameNS(this.W_NS, "fldData")[0];
    if (fldData && fldData.textContent) target.push(fldData.textContent);
  },

  _classifyEndNoteField(instrText) {
    if (/\bADDIN\s+EN\.REFLIST\b/.test(instrText)) return "reflist";
    if (/\bADDIN\s+EN\.CITE\.DATA\b/.test(instrText)) return "data";
    if (/\bADDIN\s+EN\.CITE\b/.test(instrText)) return "cite";
    return "other";
  },

  _processEndNoteCitation(citeField, dataField, results, context, sourceId) {
    const nestedDataField = this._findNestedDataField(citeField);
    const effectiveDataField = dataField || nestedDataField;
    const citeXml = this._extractEndNoteXmlFromField(citeField);
    const dataXml = effectiveDataField ? this._extractEndNoteXmlFromField(effectiveDataField) : "";
    const metadataXml = dataXml || citeXml;

    if (!metadataXml) {
      results.errors.push({
        type: "citation_no_xml",
        message: "No EndNote XML found in field code",
        displayText: citeField.displayText
      });
      return;
    }

    const optionDoc = citeXml ? this._parseEndNoteDocument(citeXml) : null;
    const metadataDoc = this._parseEndNoteDocument(metadataXml);
    if (!metadataDoc) {
      results.errors.push({
        type: "citation_xml_parse_error",
        message: "Could not parse EndNote XML",
        displayText: citeField.displayText
      });
      return;
    }

    const optionCites = optionDoc ? this._getElements(optionDoc, "Cite") : [];
    const metadataCites = this._getElements(metadataDoc, "Cite");
    const citeElements = metadataCites.length ? metadataCites : optionCites;
    const citationItems = [];

    for (let i = 0; i < citeElements.length; i++) {
      const metadataCite = citeElements[i];
      const metadataRecNum = this._extractText(metadataCite, "RecNum");
      const optionCite = this._findOptionCite(optionCites, metadataRecNum, i) || metadataCite;
      const recNum = this._extractText(optionCite, "RecNum")
        || metadataRecNum
        || this._extractText(metadataCite, "rec-number")
        || "";
      const recordEl = this._firstElement(metadataCite, "record") || this._firstElement(optionCite, "record");
      const cslData = recordEl
        ? this._endnoteRecordToCSL(recordEl)
        : this._minimalCiteToCSL(optionCite || metadataCite);
      const pageNum = this._getAttr(optionCite, "PageNum");

      citationItems.push({
        paperpileItemId: recNum,
        endnoteRecNum: recNum,
        endnoteDbId: recordEl ? this._extractEndNoteDbId(recordEl) : "",
        cslData,
        locator: pageNum || null,
        locatorType: pageNum ? "page" : null,
        prefix: this._getAttr(optionCite, "PrefixText") || null,
        suffix: this._getAttr(optionCite, "SuffixText") || null,
        suppressAuthor: this._getAttr(optionCite, "ExcludeAuth") === "1",
        suppressYear: this._getAttr(optionCite, "ExcludeYear") === "1"
      });
    }

    if (citationItems.length === 0) {
      results.errors.push({
        type: "citation_no_items",
        message: "No Cite blocks found",
        displayText: citeField.displayText
      });
      return;
    }

    results.citations.push({
      fieldIndex: results.citations.length,
      sourceId,
      endnoteId: sourceId,
      paperpileId: sourceId,
      rawInstrText: citeField.instrText,
      rawDataInstrText: effectiveDataField ? effectiveDataField.instrText : "",
      citationItems,
      formattedText: citeField.displayText,
      fieldBeginIndex: citeField.beginIndex,
      fieldEndIndex: dataField ? dataField.endIndex : citeField.endIndex,
      isInFootnote: context.isFootnote,
      isInEndnote: context.isEndnote,
      footnoteId: citeField.footnoteId || null,
      hasDataField: !!effectiveDataField
    });
  },

  _findNestedDataField(field) {
    const nestedFields = field && field.nestedFields ? field.nestedFields : [];
    for (let i = 0; i < nestedFields.length; i++) {
      if (this._classifyEndNoteField(nestedFields[i].instrText) === "data") return nestedFields[i];
    }
    return null;
  },

  _findOptionCite(optionCites, recNum, index) {
    if (recNum) {
      for (let i = 0; i < optionCites.length; i++) {
        if (this._extractText(optionCites[i], "RecNum") === recNum) return optionCites[i];
      }
    }
    return optionCites[index] || null;
  },

  _extractEndNoteXml(instrText) {
    const decoded = this._decodeXmlEntities(instrText || "");
    const directXml = this._sliceEndNoteXml(decoded);
    if (directXml) return directXml;

    const hexDecoded = this._decodeHexEndNotePayload(decoded);
    if (hexDecoded) return this._sliceEndNoteXml(hexDecoded);

    return "";
  },

  _sliceEndNoteXml(decodedText) {
    const decoded = String(decodedText || "");
    const start = decoded.indexOf("<EndNote");
    if (start === -1) return "";
    const close = "</EndNote>";
    const end = decoded.indexOf(close, start);
    if (end === -1) return "";
    return decoded.substring(start, end + close.length);
  },

  _extractEndNoteXmlFromField(field) {
    if (!field) return "";
    const instrXml = this._extractEndNoteXml(field.instrText || "");
    if (instrXml) return instrXml;
    const decodedFldData = this._decodeFldData(field.fldDataText || "");
    return this._extractEndNoteXml(decodedFldData);
  },

  _decodeFldData(fldDataText) {
    const raw = String(fldDataText || "").trim();
    if (!raw) return "";
    if (raw.indexOf("<") !== -1 || raw.indexOf("&lt;") !== -1) {
      return this._decodeXmlEntities(raw).replace(/\0/g, "");
    }

    const hexDecoded = this._decodeHexEndNotePayload(raw);
    if (hexDecoded) return hexDecoded;

    const clean = raw.replace(/\s+/g, "");
    if (!clean) return "";

    try {
      if (typeof Buffer !== "undefined") {
        return Buffer.from(clean, "base64").toString("utf8").replace(/\0/g, "");
      }

      if (typeof atob === "undefined") return "";

      const binary = atob(clean);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

      if (typeof TextDecoder !== "undefined") {
        return new TextDecoder("utf-8").decode(bytes).replace(/\0/g, "");
      }

      let escaped = "";
      for (let i = 0; i < binary.length; i++) {
        const hex = binary.charCodeAt(i).toString(16).padStart(2, "0");
        escaped += "%" + hex;
      }
      return decodeURIComponent(escaped).replace(/\0/g, "");
    } catch (e) {
      return "";
    }
  },

  _decodeHexEndNotePayload(text) {
    const hexRuns = String(text || "").match(/[0-9A-Fa-f](?:\s*[0-9A-Fa-f]){39,}/g) || [];
    for (let i = 0; i < hexRuns.length; i++) {
      const clean = this._trimHexToEndNotePayload(hexRuns[i].replace(/\s+/g, ""));
      if (clean.length % 2 !== 0) continue;

      const bytes = new Uint8Array(clean.length / 2);
      let valid = true;
      for (let j = 0; j < clean.length; j += 2) {
        const byte = parseInt(clean.substring(j, j + 2), 16);
        if (isNaN(byte)) {
          valid = false;
          break;
        }
        bytes[j / 2] = byte;
      }
      if (!valid) continue;

      const decodedCandidates = this._decodeByteStringCandidates(bytes);
      for (let j = 0; j < decodedCandidates.length; j++) {
        const decoded = this._decodeXmlEntities(decodedCandidates[j]).replace(/\0/g, "");
        if (decoded.indexOf("<EndNote") !== -1 && decoded.indexOf("</EndNote>") !== -1) {
          return decoded;
        }
      }
    }
    return "";
  },

  _trimHexToEndNotePayload(hexText) {
    const clean = String(hexText || "");
    const markers = [
      /EFBBBF3C456E644E6F7465/i,
      /3C456E644E6F7465/i,
      /FFFE3C0045006E0064004E006F00740065/i,
      /3C0045006E0064004E006F00740065/i,
      /FEFF003C0045006E0064004E006F00740065/i,
      /003C0045006E0064004E006F00740065/i
    ];

    let bestIndex = -1;
    for (let i = 0; i < markers.length; i++) {
      const match = markers[i].exec(clean);
      if (match && (bestIndex === -1 || match.index < bestIndex)) bestIndex = match.index;
    }

    return bestIndex === -1 ? clean : clean.substring(bestIndex);
  },

  _decodeByteStringCandidates(bytes) {
    const candidates = [];
    const add = value => {
      if (value && candidates.indexOf(value) === -1) candidates.push(value);
    };

    add(this._decodeBytes(bytes, "utf-8"));
    add(this._decodeBytes(bytes, "utf-16le"));
    add(this._decodeBytes(bytes, "utf-16be"));

    return candidates;
  },

  _decodeBytes(bytes, encoding) {
    try {
      if (typeof TextDecoder !== "undefined") {
        return new TextDecoder(encoding).decode(bytes);
      }
    } catch (e) { /* try other decoders */ }

    try {
      if (typeof Buffer !== "undefined") {
        if (encoding === "utf-16le") return Buffer.from(bytes).toString("utf16le");
        if (encoding === "utf-8") return Buffer.from(bytes).toString("utf8");
      }
    } catch (e) { /* use manual fallback */ }

    if (encoding === "utf-16le" || encoding === "utf-16be") {
      let text = "";
      for (let i = 0; i + 1 < bytes.length; i += 2) {
        const code = encoding === "utf-16le"
          ? bytes[i] | (bytes[i + 1] << 8)
          : (bytes[i] << 8) | bytes[i + 1];
        text += String.fromCharCode(code);
      }
      return text;
    }

    let binary = "";
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    try {
      return decodeURIComponent(escape(binary));
    } catch (e) {
      return binary;
    }
  },

  _decodeXmlEntities(text) {
    return String(text)
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  },

  _parseEndNoteDocument(xmlString) {
    try {
      const doc = new DOMParser().parseFromString(this._sanitizeEndNoteXml(xmlString), "application/xml");
      const parserErrors = doc.getElementsByTagName("parsererror");
      if (parserErrors && parserErrors.length) return null;
      return doc;
    } catch (e) {
      return null;
    }
  },

  _sanitizeEndNoteXml(xmlString) {
    return String(xmlString || "").replace(
      /&(?!#\d+;|#x[0-9a-fA-F]+;|[A-Za-z][A-Za-z0-9._:-]*;)/g,
      "&amp;"
    );
  },

  _endnoteRecordToCSL(recordEl) {
    const csl = {};
    const refTypeEl = this._firstElement(recordEl, "ref-type");
    const refTypeName = refTypeEl ? this._getAttr(refTypeEl, "name") : "Generic";
    csl.type = this._mapRefType(refTypeName);

    const title = this._extractText(recordEl, "title");
    if (title) csl.title = title;

    const secondaryTitle = this._extractText(recordEl, "secondary-title");
    const fullTitle = this._extractText(recordEl, "full-title");
    if (secondaryTitle) csl["container-title"] = secondaryTitle;
    else if (fullTitle) csl["container-title"] = fullTitle;

    const abbr = this._extractText(recordEl, "abbr-1");
    if (abbr) csl.journalAbbreviation = abbr;

    const tertiaryTitle = this._extractText(recordEl, "tertiary-title");
    if (tertiaryTitle) csl["collection-title"] = tertiaryTitle;

    const shortTitle = this._extractText(recordEl, "short-title");
    if (shortTitle) csl["title-short"] = shortTitle;

    this._copyTextField(recordEl, csl, "volume", "volume");
    this._copyTextField(recordEl, csl, "number", "issue");
    this._copyTextField(recordEl, csl, "pages", "page");
    this._copyTextField(recordEl, csl, "edition", "edition");
    this._copyTextField(recordEl, csl, "publisher", "publisher");
    this._copyTextField(recordEl, csl, "pub-location", "publisher-place");
    this._copyTextField(recordEl, csl, "language", "language");

    const doi = this._extractText(recordEl, "electronic-resource-num");
    if (doi) csl.DOI = doi;

    const accession = this._extractText(recordEl, "accession-num");
    if (accession && /^\d+$/.test(accession.trim())) csl.PMID = accession.trim();

    const url = this._extractText(recordEl, "url");
    if (url) csl.URL = url;

    const isbn = this._extractText(recordEl, "isbn");
    if (isbn) {
      const trimmed = isbn.trim();
      if (/^\d{4}-\d{3}[\dX]$/i.test(trimmed)) csl.ISSN = trimmed;
      else csl.ISBN = trimmed;
    }

    const issued = this._extractDateParts(recordEl);
    if (issued) csl.issued = issued;

    const authors = this._parseEndNoteAuthors(recordEl, "authors");
    if (authors.length) csl.author = authors;

    const editors = this._parseEndNoteAuthors(recordEl, "secondary-authors");
    if (editors.length) csl.editor = editors;

    const translators = this._parseEndNoteAuthors(recordEl, "tertiary-authors");
    if (translators.length) csl.translator = translators;

    return csl;
  },

  _minimalCiteToCSL(citeEl) {
    const csl = { type: "article" };
    const author = this._extractText(citeEl, "Author");
    if (author) csl.author = [{ family: author, given: "" }];
    const year = this._extractText(citeEl, "Year");
    if (year) {
      const parsed = parseInt(year, 10);
      csl.issued = { "date-parts": [[isNaN(parsed) ? year : parsed]] };
    }
    return csl;
  },

  _mapRefType(refTypeName) {
    const typeMap = {
      "Journal Article": "article-journal",
      "Electronic Article": "article-journal",
      "Book": "book",
      "Edited Book": "book",
      "Book Section": "chapter",
      "Conference Proceedings": "paper-conference",
      "Conference Paper": "paper-conference",
      "Report": "report",
      "Thesis": "thesis",
      "Web Page": "webpage",
      "Patent": "patent",
      "Newspaper Article": "article-newspaper",
      "Magazine Article": "article-magazine",
      "Legal Rule or Regulation": "legislation",
      "Case": "legal_case",
      "Film or Broadcast": "motion_picture",
      "Generic": "article"
    };
    return typeMap[refTypeName] || "article";
  },

  _parseEndNoteAuthors(recordEl, containerTag) {
    const container = this._firstElement(recordEl, containerTag);
    if (!container) return [];

    const result = [];
    const authors = this._getElements(container, "author");
    for (let i = 0; i < authors.length; i++) {
      const name = (authors[i].textContent || "").trim();
      if (!name) continue;

      const commaIdx = name.indexOf(",");
      if (commaIdx !== -1) {
        result.push({
          family: name.substring(0, commaIdx).trim(),
          given: name.substring(commaIdx + 1).trim()
        });
      } else if (/\s/.test(name)) {
        result.push({ literal: name });
      } else {
        result.push({ family: name, given: "" });
      }
    }
    return result;
  },

  _extractDateParts(recordEl) {
    const dateText = this._extractText(recordEl, "date");
    const yearText = this._extractText(recordEl, "year");
    const source = dateText || yearText;
    if (!source) return null;

    const parts = String(source).match(/\d+/g);
    if (!parts || !parts.length) return null;

    const dateParts = parts.slice(0, 3).map(part => parseInt(part, 10));
    return { "date-parts": [dateParts] };
  },

  _extractEndNoteDbId(recordEl) {
    const keys = this._getElements(recordEl, "key");
    for (let i = 0; i < keys.length; i++) {
      const dbId = this._getAttr(keys[i], "db-id");
      if (dbId) return dbId;
    }
    return "";
  },

  _copyTextField(recordEl, csl, endnoteTag, cslKey) {
    const value = this._extractText(recordEl, endnoteTag);
    if (value) csl[cslKey] = value;
  },

  _extractText(parentEl, tagName) {
    const el = this._firstElement(parentEl, tagName);
    return el ? (el.textContent || "").trim() : "";
  },

  _firstElement(parentEl, tagName) {
    if (!parentEl) return null;
    const els = parentEl.getElementsByTagName(tagName);
    return els && els.length ? els[0] : null;
  },

  _getElements(parentEl, tagName) {
    if (!parentEl) return [];
    const els = parentEl.getElementsByTagName(tagName);
    const result = [];
    for (let i = 0; i < els.length; i++) result.push(els[i]);
    return result;
  },

  _getAttr(el, name) {
    if (!el || !el.getAttribute) return "";
    return el.getAttribute(name) || "";
  }
};
