import { describe, expect, it } from "vitest";
import zlib from "node:zlib";
import { extractUploadedFileContext } from "../server/file-context.mjs";
import { extractPowerPointBrandProfile } from "../server/m365.mjs";

function crc32(buffer) {
  let crc = -1;
  for (const byte of buffer) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ -1) >>> 0;
}

function zipStore(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const [name, contentText] of Object.entries(entries)) {
    const nameBuffer = Buffer.from(name);
    const content = Buffer.from(contentText);
    const compressed = zlib.deflateRawSync(content);
    const crc = crc32(content);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(content.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    localParts.push(local, nameBuffer, compressed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(content.length, 24);
    central.writeUInt16LE(nameBuffer.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuffer);
    offset += local.length + nameBuffer.length + compressed.length;
  }
  const centralOffset = offset;
  const central = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(Object.keys(entries).length, 8);
  end.writeUInt16LE(Object.keys(entries).length, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(centralOffset, 16);
  return Buffer.concat([...localParts, central, end]);
}

describe("uploaded file context extraction", () => {
  it("extracts PowerPoint upload text", () => {
    const pptx = zipStore({
      "[Content_Types].xml": "<Types/>",
      "ppt/presentation.xml": "<p:presentation/>",
      "ppt/slides/slide1.xml": '<p:sld><p:cSld><p:spTree><a:p><a:r><a:t>Quarterly roadmap</a:t></a:r></a:p></p:spTree></p:cSld></p:sld>',
    });
    const result = extractUploadedFileContext({ name: "roadmap.pptx", type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", base64: pptx.toString("base64") });
    expect(result.strategy).toBe("office-open-xml");
    expect(result.extractedText).toContain("Quarterly roadmap");
  });


  it("extracts richer PowerPoint template context including notes and media", () => {
    const pptx = zipStore({
      "[Content_Types].xml": "<Types/>",
      "ppt/presentation.xml": "<p:presentation/>",
      "ppt/slides/slide1.xml": '<p:sld><p:cSld><p:spTree><p:sp/><p:pic/><a:p><a:r><a:t>Executive dolphin strategy</a:t></a:r></a:p></p:spTree></p:cSld></p:sld>',
      "ppt/slides/_rels/slide1.xml.rels": '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/dolphin.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
      "ppt/notesSlides/notesSlide1.xml": '<p:notes><a:p><a:r><a:t>Mention conservation risk.</a:t></a:r></a:p></p:notes>',
      "ppt/theme/theme1.xml": '<a:theme name="Ocean"/>',
      "ppt/slideLayouts/slideLayout1.xml": '<p:sldLayout type="chart"><p:cSld name="Chart Slide"/><p:sp><p:nvSpPr><p:nvPr><p:ph type="chart"/></p:nvPr></p:nvSpPr></p:sp></p:sldLayout>',
      "ppt/slideMasters/slideMaster1.xml": '<p:sldMaster/>',
      "ppt/media/dolphin.png": "fake-image-bytes",
    });
    const result = extractUploadedFileContext({ name: "dolphins.pptx", type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", base64: pptx.toString("base64") });
    expect(result.extractedText).toContain("Themes: theme1.xml");
    expect(result.extractedText).toContain("Slide layouts: 1");
    expect(result.extractedText).toContain("Layout 0: Chart Slide (chart) placeholders: chart");
    expect(result.extractedText).toContain("Embedded media: dolphin.png");
    expect(result.extractedText).toContain("Executive dolphin strategy");
    expect(result.extractedText).toContain("Speaker notes");
    expect(result.extractedText).toContain("Mention conservation risk");
    expect(result.extractedText).toContain("Relationships: image:../media/dolphin.png");
  });

  it("extracts Word styles and comments for review/template grounding", () => {
    const docx = zipStore({
      "[Content_Types].xml": "<Types/>",
      "word/document.xml": '<w:document><w:body><w:p><w:r><w:t>Board memo body</w:t></w:r></w:p></w:body></w:document>',
      "word/styles.xml": '<w:styles><w:style w:styleId="Title"/><w:style w:styleId="Heading1"/></w:styles>',
      "word/comments.xml": '<w:comments><w:comment><w:p><w:r><w:t>Tighten this claim.</w:t></w:r></w:p></w:comment></w:comments>',
      "word/header1.xml": '<w:hdr><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:hdr>',
      "word/media/logo.png": "fake-logo",
    });
    const result = extractUploadedFileContext({ name: "memo.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", base64: docx.toString("base64") });
    expect(result.extractedText).toContain("Document body");
    expect(result.extractedText).toContain("Board memo body");
    expect(result.extractedText).toContain("Styles: Title, Heading1");
    expect(result.extractedText).toContain("Comments");
    expect(result.extractedText).toContain("Tighten this claim");
    expect(result.extractedText).toContain("Embedded media: logo.png");
  });

  it("extracts Excel workbook structure, tables, filters, and chart hints", () => {
    const xlsx = zipStore({
      "[Content_Types].xml": "<Types/>",
      "xl/workbook.xml": '<workbook><sheets><sheet name="Revenue" sheetId="1"/></sheets></workbook>',
      "xl/worksheets/sheet1.xml": '<worksheet><sheetData><row><c/><c/></row><row><c/><c/></row></sheetData><autoFilter ref="A1:B2"/><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>',
      "xl/worksheets/_rels/sheet1.xml.rels": '<Relationships><Relationship Id="rHyperlink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/revenue-source" TargetMode="External"/></Relationships>',
      "xl/tables/table1.xml": '<table name="RevenueTable" ref="A1:B2"/>',
      "xl/sharedStrings.xml": '<sst><si><t>Region</t></si><si><t>Revenue</t></si></sst>',
      "xl/comments/comment1.xml": '<comments><authors><author>Reviewer</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>Validate revenue assumption.</t></r></text></comment></commentList></comments>',
      "xl/charts/chart1.xml": '<c:chartSpace/>',
    });
    const result = extractUploadedFileContext({ name: "revenue.xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", base64: xlsx.toString("base64") });
    expect(result.extractedText).toContain("Workbook structure");
    expect(result.extractedText).toContain("Worksheet 1: 2 rows, 4 cells, 1 merged ranges, autofilter A1:B2");
    expect(result.extractedText).toContain("Table: RevenueTable (A1:B2)");
    expect(result.extractedText).toContain("Hyperlinks: https://example.test/revenue-source");
    expect(result.extractedText).toContain("Charts: chart1.xml");
    expect(result.extractedText).toContain("Comments (comment1.xml)");
    expect(result.extractedText).toContain("Validate revenue assumption");
    expect(result.extractedText).toContain("Region");
  });

  it("extracts a structured PowerPoint brand profile without exposing media bytes", () => {
    const pptx = zipStore({
      "[Content_Types].xml": "<Types/>",
      "ppt/presentation.xml": "<p:presentation/>",
      "ppt/theme/theme1.xml": '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="112233"/></a:dk1><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="CC5500"/></a:accent2><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
      "ppt/slideLayouts/slideLayout1.xml": '<p:sldLayout type="title"><p:cSld name="Executive"><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr></p:sp></p:sldLayout>',
      "ppt/media/company-logo.png": "private-logo-bytes",
    });
    const profile = extractPowerPointBrandProfile(pptx);
    expect(profile.colors).toMatchObject({ dk1: "#112233", accent1: "#336699", accent2: "#CC5500" });
    expect(profile.fonts).toEqual({ major: "Aptos Display", minor: "Aptos" });
    expect(profile.layouts[0]).toMatchObject({ name: "Executive", type: "title", placeholders: ["title"] });
    expect(profile.media[0]).toMatchObject({ name: "company-logo.png", candidate: true });
    expect(profile.media[0]).not.toHaveProperty("base64");
    const extracted = extractUploadedFileContext({ name: "brand.pptx", type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", base64: pptx.toString("base64") });
    expect(extracted.brandProfile.colors.accent1).toBe("#336699");
  });

  it("returns raw fallback for unknown binary uploads", () => {
    const result = extractUploadedFileContext({ name: "template.unknown", type: "application/octet-stream", base64: Buffer.from([0, 1, 2, 3, 255]).toString("base64") });
    expect(result.strategy).toBe("raw-fallback");
    expect(result.extractedText).toContain("Base64 sample");
  });

  it("rejects malformed base64 and Office MIME/signature mismatches", () => {
    expect(() => extractUploadedFileContext({ name: "notes.txt", type: "text/plain", base64: "not base64!" })).toThrow(/base64/i);
    expect(() => extractUploadedFileContext({ name: "report.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", base64: Buffer.from("plain text").toString("base64") })).toThrow(/signature/i);
  });

  it("rejects Office filename and MIME-type mismatches before parsing", () => {
    expect(() => extractUploadedFileContext({
      name: "report.pptx",
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      base64: Buffer.from("not-a-package").toString("base64"),
    })).toThrow(/name and MIME type disagree/i);
  });
  it("rejects a ZIP that is mislabeled as an Office package", () => {
    const mislabeled = zipStore({ "notes.txt": "not an Office package" });
    expect(() => extractUploadedFileContext({ name: "notes.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", base64: mislabeled.toString("base64") })).toThrow(/missing/i);
  });
});
