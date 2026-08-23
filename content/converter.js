/* Converter Module — rewrites EndNote field codes to Zotero format in .docx */

if (typeof EN2Zotero === "undefined") var EN2Zotero = {};

EN2Zotero.Converter = {
  ZOTERO_SCHEMA: "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",

  async convert(docxData, matchResults, options = {}) {
    if (docxData && docxData.byteLength !== undefined && !(docxData instanceof Uint8Array)) {
      docxData = new Uint8Array(docxData);
    }
    const JSZipRef = typeof JSZip !== "undefined" ? JSZip : (await import("./lib/jszip.min.js")).default;
    const zip = await JSZipRef.loadAsync(docxData);

    const xmlFiles = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];
    const allComments = [];
    let nextCommentId = 1;
    // ODT reference marks need a dedicated EndNote/Windows port of odf-writer.js.
    if (options.fieldMode === "referencemarks") {
      throw new Error("Reference Marks (ODF) output is not available in EN2Zotero 1.0.2");
    }

    const useBookmarks = options.fieldMode === "bookmarks";
    const allBookmarkData = {};

    for (const xmlPath of xmlFiles) {
      const file = zip.file(xmlPath);
      if (!file) continue;

      let xmlText = await file.async("string");

      if (useBookmarks) {
        const result = this._rewriteFieldsAsBookmarks(xmlText, matchResults, options, xmlPath);
        xmlText = result.xmlText;
        Object.assign(allBookmarkData, result.bookmarkData);
      } else {
        xmlText = this._rewriteFields(xmlText, matchResults, options, xmlPath);
      }

      if (options.addComments) {
        const result = this._addSkippedComments(xmlText, matchResults, nextCommentId, xmlPath);
        xmlText = result.xmlText;
        allComments.push(...result.comments);
        nextCommentId += result.comments.length;
      }

      zip.file(xmlPath, xmlText);
    }

    if (allComments.length > 0) {
      zip.file("word/comments.xml", this._buildCommentsXml(allComments));
      await this._ensureCommentsRelationship(zip);
    }

    if (useBookmarks && Object.keys(allBookmarkData).length > 0) {
      await this._updateSettingsWithDocVars(zip, allBookmarkData);
    }

    return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  },

  _rewriteFields(xmlText, matchResults, options, xmlPath) {
    const matchMap = new Map();
    for (const mr of matchResults) {
      const id = mr.citation.sourceId || mr.citation.endnoteId || mr.citation.paperpileId;
      if (id) matchMap.set(id, mr);
    }

    // Use block-based approach (same as bookmarks mode) to correctly handle
    // instrText split across multiple <w:r> elements — Word splits instrText
    // at ~255 chars, and EndNote data often exceeds this.
    const blocks = this._findEndNoteFieldBlocks(xmlText, xmlPath);

    // Process in reverse to preserve string positions
    for (let i = blocks.length - 1; i >= 0; i--) {
      const block = blocks[i];

      if (block.isBibliography) {
        const zoteroBibl = { uncited: [], omitted: [], custom: [] };
        const escaped = this._escapeXml(JSON.stringify(zoteroBibl));
        const cleanDisplay = this._stripTrackChangeWrappers(block.displayContent);
        const replacement = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
          + '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL ' + escaped + ' CSL_BIBLIOGRAPHY </w:instrText></w:r>'
          + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
          + cleanDisplay
          + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
        xmlText = this._spliceBalanced(xmlText, block.startPos, block.endPos, replacement);
        continue;
      }

      if (!block.isCitation) continue;

      const matchResult = matchMap.get(block.sourceId);
      if (!matchResult) continue;

      const allSkipped = matchResult.itemMatches.every(m =>
        m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
      );
      if (allSkipped) continue;

      const zoteroCitation = this._buildZoteroCitation(matchResult);
      if (!zoteroCitation) continue;

      const zoteroJSON = JSON.stringify(zoteroCitation);
      const escaped = this._escapeXml(zoteroJSON);
      const cleanDisplay = this._stripTrackChangeWrappers(block.displayContent);

      // Replace the ENTIRE field code block (begin→end) with a clean structure.
      // This ensures no residual EndNote instrText elements remain that would
      // corrupt Zotero's JSON parsing when it concatenates all instrText nodes.
      const replacement = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        + '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION ' + escaped + ' </w:instrText></w:r>'
        + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        + cleanDisplay
        + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
      xmlText = this._spliceBalanced(xmlText, block.startPos, block.endPos, replacement);
    }

    return xmlText;
  },

  // Strip block-level <w:ins>/<w:del> wrappers (open/close pairs) from a fragment.
  // Self-closing forms (e.g. <w:ins ... /> inside <w:rPr>) are run-property metadata
  // and must be kept intact.
  _stripTrackChangeWrappers(xml) {
    return xml
      .replace(/<w:(ins|del)\b([^>]*)>/g, function(match, elem, attrs) {
        return attrs.trimEnd().endsWith('/') ? match : '';
      })
      .replace(/<\/w:(ins|del)>/g, '');
  },

  // Replace xmlText[startPos..endPos] with replacement, preserving balance of any
  // <w:ins>/<w:del> track-change wrappers that crossed the cut boundary. Without
  // this, EndNote fields split across multiple <w:ins> blocks (typical when the
  // citation was inserted via Track Changes) produce invalid OOXML — Word refuses
  // to open the file with "cannot open the file" errors.
  //
  // The cut may straddle several track-change wrappers in two ways:
  //   1. Closes-in-cut without matching opens → prefix has unclosed wrappers.
  //      Normally we'd prepend the missing closes, but if the prefix's last open
  //      is IMMEDIATELY adjacent to our cut start (no real content between them),
  //      prepending a close would leave an empty <w:ins ATTRS></w:ins> pair —
  //      which Word reports as "found unreadable content" even though XML is
  //      well-formed. So we prefer to REMOVE the empty trailing open from prefix.
  //   2. Opens-in-cut without matching closes → suffix has orphan closes. We
  //      DELETE them from suffix (rather than appending empty <w:ins ATTRS> opens
  //      after the replacement, for the same reason).
  _spliceBalanced(xmlText, startPos, endPos, replacement) {
    const cut = xmlText.substring(startPos, endPos);
    const tagRegex = /<(\/?)w:(ins|del)\b([^>]*)>/g;
    const stacks = { ins: [], del: [] };
    const unmatchedCloses = { ins: 0, del: 0 };
    let m;
    while ((m = tagRegex.exec(cut)) !== null) {
      const isClose = m[1] === '/';
      const elem = m[2];
      const attrs = m[3];
      if (isClose) {
        if (stacks[elem].length > 0) stacks[elem].pop();
        else unmatchedCloses[elem]++;
      } else {
        if (attrs.trimEnd().endsWith('/')) continue;
        stacks[elem].push(attrs);
      }
    }

    let prefix = xmlText.substring(0, startPos);
    let suffix = xmlText.substring(endPos);

    // Suffix-side: for each unmatched open in cut, remove its orphan close from suffix.
    function removeNextMatchingClose(suffixText, elemName) {
      const re = new RegExp('<(\\/?)w:' + elemName + '\\b([^>]*)>', 'g');
      const innerStack = [];
      let mm;
      while ((mm = re.exec(suffixText)) !== null) {
        const isClose = mm[1] === '/';
        const attrs = mm[2];
        if (isClose) {
          if (innerStack.length === 0) {
            return suffixText.substring(0, mm.index)
                 + suffixText.substring(mm.index + mm[0].length);
          }
          innerStack.pop();
        } else {
          if (attrs.trimEnd().endsWith('/')) continue;
          innerStack.push(true);
        }
      }
      return null;
    }

    for (let k = 0; k < stacks.ins.length; k++) {
      const newSuffix = removeNextMatchingClose(suffix, 'ins');
      if (newSuffix !== null) suffix = newSuffix;
    }
    for (let k = 0; k < stacks.del.length; k++) {
      const newSuffix = removeNextMatchingClose(suffix, 'del');
      if (newSuffix !== null) suffix = newSuffix;
    }

    // Prefix-side: try to remove each trailing empty open from prefix (avoids
    // empty <w:ins></w:ins> wrappers that trigger Word's "found unreadable
    // content" warning). Only fall back to prepending a close if the trailing
    // open has real content.
    function tryRemoveTrailingEmptyOpen(prefixText, elemName) {
      const re = new RegExp('<w:' + elemName + '\\b([^>]*)>', 'g');
      let last = null;
      let mm;
      while ((mm = re.exec(prefixText)) !== null) {
        const attrs = mm[1];
        if (attrs.trimEnd().endsWith('/')) continue;
        last = mm;
      }
      if (!last) return null;
      const afterOpen = last.index + last[0].length;
      const tail = prefixText.substring(afterOpen);
      if (tail.length === 0 || /^\s*$/.test(tail)) {
        return prefixText.substring(0, last.index) + tail;
      }
      return null;
    }

    let prependCloses = { ins: 0, del: 0 };
    for (let k = 0; k < unmatchedCloses.ins; k++) {
      const newPrefix = tryRemoveTrailingEmptyOpen(prefix, 'ins');
      if (newPrefix !== null) prefix = newPrefix;
      else prependCloses.ins++;
    }
    for (let k = 0; k < unmatchedCloses.del; k++) {
      const newPrefix = tryRemoveTrailingEmptyOpen(prefix, 'del');
      if (newPrefix !== null) prefix = newPrefix;
      else prependCloses.del++;
    }

    const prepend = '</w:ins>'.repeat(prependCloses.ins)
                  + '</w:del>'.repeat(prependCloses.del);

    return prefix + prepend + replacement + suffix;
  },

  _buildZoteroCitation(matchResult) {
    const items = matchResult.itemMatches
      .filter(m => m.match !== null)
      .map(m => {
        const zItem = m.match;
        const cslJSON = EN2Zotero.Matcher.getCSLJSON(zItem);
        const uri = EN2Zotero.Matcher.buildURI(zItem);

        const citItem = {
          id: zItem.id,
          uris: [uri],
          uri: [uri],
          itemData: cslJSON
        };

        if (m.locator) citItem.locator = m.locator;
        if (m.locatorType) citItem.label = m.locatorType;
        if (m.prefix) citItem.prefix = m.prefix;
        if (m.suffix) citItem.suffix = m.suffix;
        if (m.suppressAuthor) citItem["suppress-author"] = true;

        if (typeof Zotero !== "undefined") {
          const authorCount = (cslJSON.author || []).length;
          Zotero.debug("EN2Zotero: citItem id=" + zItem.id + " authors=" + authorCount + " suppress-author=" + (!!m.suppressAuthor) + " title=" + (cslJSON.title || "").substring(0, 40));
        }

        return citItem;
      });

    if (items.length === 0) return null;

    const noteIndex = matchResult.citation.isInFootnote || matchResult.citation.isInEndnote
      ? (matchResult.citation.fieldIndex + 1) : 0;

    return {
      citationID: this._generateCitationID(),
      properties: {
        formattedCitation: "",
        plainCitation: "",
        noteIndex: noteIndex
      },
      citationItems: items,
      schema: this.ZOTERO_SCHEMA
    };
  },

  _generateCitationID() {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "";
    for (let i = 0; i < 8; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  },

  _escapeXml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  },

  _unescapeXml(str) {
    return str
      .replace(/&quot;/g, '"')
      .replace(/&gt;/g, '>')
      .replace(/&lt;/g, '<')
      .replace(/&amp;/g, '&');
  },

  _addSkippedComments(xmlText, matchResults, startId, xmlPath) {
    const comments = [];
    const skippedCitations = [];

    for (const mr of matchResults) {
      const allSkipped = mr.itemMatches.every(m =>
        m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
      );
      if (!allSkipped) continue;

      const sourceId = mr.citation.sourceId || mr.citation.endnoteId || mr.citation.paperpileId;
      if (!sourceId) continue;

      const title = mr.itemMatches[0]?.cslData?.title || "Unknown reference";
      const authors = (mr.itemMatches[0]?.cslData?.author || [])
        .map(a => a.family || a.given || "").filter(Boolean).join(", ");

      skippedCitations.push({ sourceId, title, authors });
    }

    // Process in reverse order to preserve string positions
    const positions = [];
    const blockMap = new Map();
    for (const block of this._findEndNoteFieldBlocks(xmlText, xmlPath)) {
      if (block.isCitation) blockMap.set(block.sourceId, block);
    }
    for (const sc of skippedCitations) {
      const block = blockMap.get(sc.sourceId);
      if (block) positions.push({ pos: block.startPos, ...sc });
    }
    positions.sort((a, b) => b.pos - a.pos);

    for (const item of positions) {
      const commentId = startId + comments.length;
      const commentText = "EN2Zotero: This citation was not converted to Zotero format.\n"
        + (item.title ? "Title: " + item.title + "\n" : "")
        + (item.authors ? "Authors: " + item.authors : "");

      // Find the paragraph containing this position
      let pStart = xmlText.lastIndexOf("<w:p ", item.pos);
      const pStart2 = xmlText.lastIndexOf("<w:p>", item.pos);
      if (pStart2 > pStart) pStart = pStart2;
      const pEnd = xmlText.indexOf("</w:p>", item.pos);

      if (pStart === -1 || pEnd === -1) continue;

      const pOpenEnd = xmlText.indexOf(">", pStart) + 1;
      const startMarker = '<w:commentRangeStart w:id="' + commentId + '"/>';
      const endMarker = '<w:commentRangeEnd w:id="' + commentId + '"/>'
        + '<w:r><w:commentReference w:id="' + commentId + '"/></w:r>';

      // Insert end marker before </w:p> (do this first since it's after pOpenEnd)
      xmlText = xmlText.substring(0, pEnd) + endMarker + xmlText.substring(pEnd);
      // Insert start marker after <w:p...>
      xmlText = xmlText.substring(0, pOpenEnd) + startMarker + xmlText.substring(pOpenEnd);

      comments.push({ id: commentId, text: commentText });
    }

    return { xmlText, comments };
  },

  _buildCommentsXml(comments) {
    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:comments xmlns:w="' + W_NS + '"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    for (const c of comments) {
      const lines = c.text.split("\n");
      xml += '<w:comment w:id="' + c.id + '" w:author="EN2Zotero" w:date="' + now + '">';
      for (const line of lines) {
        xml += '<w:p><w:r><w:t>' + this._escapeXml(line) + '</w:t></w:r></w:p>';
      }
      xml += '</w:comment>';
    }

    xml += '</w:comments>';
    return xml;
  },

  async _ensureCommentsRelationship(zip) {
    // Update [Content_Types].xml
    const ctFile = zip.file("[Content_Types].xml");
    if (ctFile) {
      let ct = await ctFile.async("string");
      if (ct.indexOf("comments.xml") === -1) {
        ct = ct.replace("</Types>",
          '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
        zip.file("[Content_Types].xml", ct);
      }
    }

    // Update word/_rels/document.xml.rels
    const relsFile = zip.file("word/_rels/document.xml.rels");
    if (relsFile) {
      let rels = await relsFile.async("string");
      if (rels.indexOf("comments.xml") === -1) {
        // Find max rId
        const idMatches = [...rels.matchAll(/Id="rId(\d+)"/g)];
        const maxId = idMatches.reduce((max, m) => Math.max(max, parseInt(m[1])), 0);
        const newId = "rId" + (maxId + 1);
        rels = rels.replace("</Relationships>",
          '<Relationship Id="' + newId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>');
        zip.file("word/_rels/document.xml.rels", rels);
      }
    }
  },

  /* ---- Bookmark mode conversion ---- */

  _rewriteFieldsAsBookmarks(xmlText, matchResults, options, xmlPath) {
    const matchMap = new Map();
    for (const mr of matchResults) {
      const id = mr.citation.sourceId || mr.citation.endnoteId || mr.citation.paperpileId;
      if (id) matchMap.set(id, mr);
    }

    const bookmarkData = {};
    let nextBmId = this._getMaxBookmarkId(xmlText) + 1;
    const blocks = this._findEndNoteFieldBlocks(xmlText, xmlPath);

    // Process in reverse to preserve string positions
    for (let i = blocks.length - 1; i >= 0; i--) {
      const block = blocks[i];

      if (block.isBibliography) {
        const bmName = this._generateBookmarkName();
        const bmId = nextBmId++;
        const zoteroBibl = { uncited: [], omitted: [], custom: [] };
        const val = " ADDIN ZOTERO_BIBL " + JSON.stringify(zoteroBibl) + " CSL_BIBLIOGRAPHY ";
        bookmarkData[bmName] = val;

        const cleanDisplay = this._stripTrackChangeWrappers(block.displayContent);
        const replacement = '<w:bookmarkStart w:id="' + bmId + '" w:name="' + bmName + '"/>'
          + cleanDisplay
          + '<w:bookmarkEnd w:id="' + bmId + '"/>';
        xmlText = this._spliceBalanced(xmlText, block.startPos, block.endPos, replacement);
        continue;
      }

      if (!block.isCitation) continue;

      const matchResult = matchMap.get(block.sourceId);
      if (!matchResult) continue;

      const allSkipped = matchResult.itemMatches.every(m =>
        m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
      );
      if (allSkipped) continue;

      const zoteroCitation = this._buildZoteroCitation(matchResult);
      if (!zoteroCitation) continue;

      const bmName = this._generateBookmarkName();
      const bmId = nextBmId++;
      const val = " ADDIN ZOTERO_ITEM CSL_CITATION " + JSON.stringify(zoteroCitation) + " ";
      bookmarkData[bmName] = val;

      const cleanDisplay = this._stripTrackChangeWrappers(block.displayContent);
      const replacement = '<w:bookmarkStart w:id="' + bmId + '" w:name="' + bmName + '"/>'
        + cleanDisplay
        + '<w:bookmarkEnd w:id="' + bmId + '"/>';
      xmlText = this._spliceBalanced(xmlText, block.startPos, block.endPos, replacement);
    }

    return { xmlText, bookmarkData };
  },

  // Find the start of <w:r> or <w:r ...> containing the given position
  // Must NOT match <w:rPr>, <w:rsidR>, etc.
  _findRunStart(xmlText, pos) {
    let searchFrom = pos;
    while (searchFrom >= 0) {
      const idx = xmlText.lastIndexOf("<w:r", searchFrom);
      if (idx === -1) return -1;
      // Check the char after "<w:r" — must be '>', ' ', or '/' (self-closing)
      const nextChar = xmlText[idx + 4];
      if (nextChar === '>' || nextChar === ' ' || nextChar === '/') {
        return idx;
      }
      // It's <w:rPr> or similar — keep searching before this position
      searchFrom = idx - 1;
    }
    return -1;
  },

  _findEndNoteFieldBlocks(xmlText, xmlPath) {
    const rawBlocks = [];

    // Collect all fldChar markers with positions
    const fldCharRegex = /<w:fldChar\b(?=[^>]*w:fldCharType="(?:begin|separate|end)")(?=[^>]*\/>)[^>]*\/>|<w:fldChar\b(?=[^>]*w:fldCharType="(?:begin|separate|end)")[^>]*>[\s\S]*?<\/w:fldChar>/g;
    const markers = [];
    let m;
    while ((m = fldCharRegex.exec(xmlText)) !== null) {
      const charType = this._getFldCharType(m[0]);
      if (charType) markers.push({ type: charType, pos: m.index });
    }

    let i = 0;
    while (i < markers.length) {
      if (markers[i].type !== "begin") { i++; continue; }

      // Find matching separate and end, handling nesting
      let nestLevel = 1;
      let sepIdx = -1;
      let j = i + 1;
      while (j < markers.length && nestLevel > 0) {
        if (markers[j].type === "begin") {
          nestLevel++;
        } else if (markers[j].type === "separate" && nestLevel === 1) {
          sepIdx = j;
        } else if (markers[j].type === "end") {
          nestLevel--;
        }
        if (nestLevel > 0) j++;
      }

      if (nestLevel !== 0 || sepIdx === -1) { i++; continue; }

      // i=begin, sepIdx=separate, j=end
      // Find enclosing <w:r> or <w:r ...> boundaries (NOT <w:rPr>)
      const beginRunStart = this._findRunStart(xmlText, markers[i].pos);
      const beginRunEnd = xmlText.indexOf("</w:r>", markers[i].pos) + 6;

      const sepRunStart = this._findRunStart(xmlText, markers[sepIdx].pos);
      const sepRunEnd = xmlText.indexOf("</w:r>", markers[sepIdx].pos) + 6;

      const endRunStart = this._findRunStart(xmlText, markers[j].pos);
      const endRunEnd = xmlText.indexOf("</w:r>", markers[j].pos) + 6;

      const instrSection = xmlText.substring(beginRunEnd, sepRunStart);
      const instrInfo = this._extractTopLevelInstrText(instrSection);

      // Display content = runs between separate and end (the visible text)
      const displayContent = xmlText.substring(sepRunEnd, endRunStart);

      rawBlocks.push({
        startPos: beginRunStart,
        endPos: endRunEnd,
        instrText: instrInfo.instrText,
        nestedDataInstrText: instrInfo.nestedDataInstrText,
        hasNestedDataField: instrInfo.hasNestedDataField,
        displayContent: displayContent,
        fieldType: this._classifyEndNoteInstr(instrInfo.instrText)
      });

      i = j + 1;
    }

    const blocks = [];
    let localEndNoteIndex = 0;
    let bibliographyIndex = 0;
    const pathPrefix = xmlPath || "word/document.xml";

    for (let k = 0; k < rawBlocks.length; k++) {
      const block = rawBlocks[k];

      if (block.fieldType === "reflist") {
        blocks.push({
          sourceId: pathPrefix + ":bibliography:" + bibliographyIndex++,
          startPos: block.startPos,
          endPos: block.endPos,
          instrText: block.instrText,
          dataInstrText: "",
          displayContent: block.displayContent,
          isCitation: false,
          isBibliography: true,
          hasDataField: false
        });
        continue;
      }

      if (block.fieldType !== "cite") continue;

      const next = rawBlocks[k + 1] || null;
      const dataBlock = next && next.fieldType === "data" ? next : null;
      blocks.push({
        sourceId: pathPrefix + ":" + localEndNoteIndex++,
        startPos: block.startPos,
        endPos: dataBlock ? dataBlock.endPos : block.endPos,
        instrText: block.instrText,
        dataInstrText: dataBlock ? dataBlock.instrText : block.nestedDataInstrText,
        displayContent: block.displayContent,
        isCitation: true,
        isBibliography: false,
        hasDataField: !!dataBlock || block.hasNestedDataField
      });

      if (dataBlock) k++;
    }

    return blocks;
  },

  _extractTopLevelInstrText(instrSection) {
    let instrText = "";
    let nestedDataInstrText = "";
    let hasNestedDataField = false;
    let nestLevel = 1;
    const tokenRegex = /<w:fldChar\b(?=[^>]*w:fldCharType="(?:begin|separate|end)")(?=[^>]*\/>)[^>]*\/>|<w:fldChar\b(?=[^>]*w:fldCharType="(?:begin|separate|end)")[^>]*>[\s\S]*?<\/w:fldChar>|<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g;
    let m;

    while ((m = tokenRegex.exec(instrSection)) !== null) {
      if (m[0].indexOf("<w:fldChar") === 0) {
        const charType = this._getFldCharType(m[0]);
        if (charType === "begin") {
          nestLevel++;
        } else if (charType === "end") {
          nestLevel = Math.max(1, nestLevel - 1);
        }
        continue;
      }

      const text = m[1] || "";
      if (nestLevel === 1) {
        instrText += text;
      } else if (nestLevel === 2) {
        nestedDataInstrText += text;
        if (this._classifyEndNoteInstr(text) === "data") hasNestedDataField = true;
      }
    }

    return { instrText, nestedDataInstrText, hasNestedDataField };
  },

  _getFldCharType(fldCharXml) {
    const match = String(fldCharXml || "").match(/w:fldCharType="(begin|separate|end)"/);
    return match ? match[1] : "";
  },

  _classifyEndNoteInstr(instrText) {
    if (/\bADDIN\s+EN\.REFLIST\b/.test(instrText)) return "reflist";
    if (/\bADDIN\s+EN\.CITE\.DATA\b/.test(instrText)) return "data";
    if (/\bADDIN\s+EN\.CITE\b/.test(instrText)) return "cite";
    return "other";
  },

  _generateBookmarkName() {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "ZOTERO_BREF_";
    for (let i = 0; i < 10; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  },

  _getMaxBookmarkId(xmlText) {
    const idRegex = /w:id="(\d+)"/g;
    let max = 0;
    let m;
    while ((m = idRegex.exec(xmlText)) !== null) {
      const val = parseInt(m[1], 10);
      if (val > max) max = val;
    }
    return max;
  },

  async _updateSettingsWithDocVars(zip, bookmarkData) {
    const settingsFile = zip.file("word/settings.xml");
    if (!settingsFile) return;

    let settings = await settingsFile.async("string");

    // Build docVars XML
    let docVarsContent = "";

    // Add ZOTERO_PREF to tell Zotero this is a bookmark-mode document
    const sessionId = this._generateCitationID();
    const prefXml = '<data data-version="3" zotero-version="7.0.0">'
      + '<session id="' + sessionId + '"/>'
      + '<prefs>'
      + '<pref name="fieldType" value="Bookmark"/>'
      + '</prefs>'
      + '</data>';
    docVarsContent += '<w:docVar w:name="ZOTERO_PREF_1" w:val="' + this._escapeXml(prefXml) + '"/>';

    // Add a docVar for each bookmark
    for (const [name, val] of Object.entries(bookmarkData)) {
      docVarsContent += '<w:docVar w:name="' + this._escapeXml(name) + '" w:val="' + this._escapeXml(val) + '"/>';
    }

    // Insert into settings.xml
    if (settings.indexOf("<w:docVars>") !== -1) {
      // Append to existing docVars
      settings = settings.replace("</w:docVars>", docVarsContent + "</w:docVars>");
    } else if (/<w:settings\b[^>]*\/>/.test(settings)) {
      settings = settings.replace(/<w:settings\b([^>]*)\/>/, "<w:settings$1><w:docVars>" + docVarsContent + "</w:docVars></w:settings>");
    } else {
      // Add before </w:settings>
      settings = settings.replace("</w:settings>", "<w:docVars>" + docVarsContent + "</w:docVars></w:settings>");
    }

    zip.file("word/settings.xml", settings);
  },

  async createBackup(filePath) {
    const backupPath = filePath + ".bak";
    await IOUtils.copy(filePath, backupPath);
    Zotero.debug("EN2Zotero: Backup created at " + backupPath);
    return backupPath;
  },

  async saveFile(filePath, data) {
    await IOUtils.write(filePath, data);
    Zotero.debug("EN2Zotero: Saved converted file to " + filePath);
  }
};
