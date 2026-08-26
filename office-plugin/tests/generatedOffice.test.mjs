import { describe, expect, it } from "vitest";
import zlib from "node:zlib";
import { createGeneratedDocx, createGeneratedPptx, createGeneratedXlsx } from "../server/generated-office.mjs";
import { validateGeneratedOfficePackage } from "../server/artifact-qa.mjs";
import { extractOfficeText } from "../server/m365.mjs";

const MINIMAL_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const fakeImageBase64 = MINIMAL_PNG_BASE64;
function zipEntries(buffer) {
  const entries = new Map();
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65558); i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  const total = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  for (let i = 0; i < total; i += 1) {
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const method = buffer.readUInt16LE(offset + 10);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8");
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    entries.set(name, method === 8 ? zlib.inflateRawSync(compressed) : compressed);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

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
  for (const [name, raw] of Object.entries(entries)) {
    const nameBuffer = Buffer.from(name);
    const content = Buffer.isBuffer(raw) ? raw : Buffer.from(String(raw), "utf8");
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

describe("generated Office artifacts", () => {
  it("rejects embedded image bytes whose signature disagrees with the declared MIME type", () => {
    expect(() => createGeneratedPptx({ slides: [{ title: "Unsafe image", images: [{ name: "unsafe.png", type: "image/png", base64: Buffer.from("not-an-image").toString("base64") }] }] })).toThrow(/does not match its declared type/i);
  });

  it("validates required package parts for generated Office files", () => {
    expect(validateGeneratedOfficePackage(createGeneratedPptx({ slides: [{ title: "QA" }] }), "pptx").entryCount).toBeGreaterThan(4);
    expect(validateGeneratedOfficePackage(createGeneratedDocx({ title: "QA", paragraphs: [{ text: "Body" }] }), "docx").entryCount).toBeGreaterThan(3);
    expect(validateGeneratedOfficePackage(createGeneratedXlsx({ sheets: [{ name: "QA", rows: [["A"]] }] }), "xlsx").entryCount).toBeGreaterThan(3);
  });

  it("rejects unsafe template archive paths before processing", () => {
    const source = createGeneratedPptx({ slides: [{ title: "Template" }] });
    const entries = zipEntries(source);
    entries.set("../unsafe.xml", Buffer.from("unsafe"));
    expect(() => createGeneratedPptx({ templateBase64: zipStore(Object.fromEntries(entries)).toString("base64"), slides: [{ title: "Rejected" }] })).toThrow(/unsafe archive path/i);
  });
  it("creates a PPTX with slide text, speaker notes, and embedded image parts", () => {
    const pptx = createGeneratedPptx({
      fileName: "dolphins.pptx",
      slides: [{
        title: "Dolphin conservation",
        body: "Pods, calves, and echolocation need visual support.",
        backgroundColor: "#E0F2FE",
        notes: "Mention NOAA source in narration.",
        links: [{ text: "NOAA dolphin source", url: "https://example.test/noaa-dolphins", left: 48, top: 470, width: 260, height: 30 }],
        shapes: [{ shapeType: "roundedRectangle", text: "Protect pods", left: 48, top: 310, width: 180, height: 54, fillColor: "#DBEAFE", lineColor: "#2563EB" }],
        tables: [{ name: "Conservation actions", values: [["Action", "Impact"], ["Reduce bycatch", "Protect calves"]], left: 48, top: 380, width: 420, height: 90 }],
        charts: [{ title: "Sightings by month", chartType: "bar", values: [["Jan", 12], ["Feb", 18], ["Mar", 15]], left: 300, top: 180, width: 330, height: 180 }],
        images: [{ name: "dolphin.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), altText: "Dolphin" }],
      }],
    });
    expect(Buffer.isBuffer(pptx)).toBe(true);
    expect(pptx.length).toBeGreaterThan(1000);
    const text = extractOfficeText(pptx, "dolphins.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    expect(text).toContain("Dolphin conservation");
    expect(text).toContain("Speaker notes");
    expect(text).toContain("Mention NOAA source");
    expect(text).toContain("Protect pods");
    expect(text).toContain("Reduce bycatch");
    expect(text).toContain("chart:../charts/chart1.xml");
    expect(text).toContain("Embedded media: image1.png");
    expect(text).toContain("image:../media/image1.png");
    expect(text).toContain("hyperlink:https://example.test/noaa-dolphins");
    const entries = zipEntries(pptx);
    expect(entries.has("ppt/media/image1.png")).toBe(true);
    expect(entries.has("ppt/charts/chart1.xml")).toBe(true);
    expect(entries.has("ppt/charts/_rels/chart1.xml.rels")).toBe(true);
    expect(entries.has("ppt/embeddings/chartData1.xlsx")).toBe(true);
    expect(entries.has("ppt/notesSlides/notesSlide1.xml")).toBe(true);
    const slideXml = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slideXml).toContain("<p:bg>");
    expect(slideXml).toContain("<p:graphicFrame>");
    expect(slideXml).toContain("rChart1");
    expect(slideXml).toContain("Protect pods");
    expect(slideXml).toContain("NOAA dolphin source");
    expect(slideXml).toContain("hlinkClick");
    const slideRels = entries.get("ppt/slides/_rels/slide1.xml.rels").toString("utf8");
    expect(slideRels).toContain("https://example.test/noaa-dolphins");
    expect(slideRels).toContain("TargetMode=\"External\"");
    const chartXml = entries.get("ppt/charts/chart1.xml").toString("utf8");
    expect(chartXml).toContain("Sightings by month");
    expect(chartXml).toContain('c:externalData r:id="rWorkbook"');
    expect(chartXml).toContain("<c:strRef><c:f>&apos;Chart Data&apos;!$A$2:$A$4</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$B$2:$B$4</c:f>");
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("</c:barChart>");
    expect(chartXml.indexOf("<c:barChart>")).toBeLessThan(chartXml.indexOf("<c:ser>"));
    const chartRels = entries.get("ppt/charts/_rels/chart1.xml.rels").toString("utf8");
    expect(chartRels).toContain("relationships/package");
    expect(chartRels).toContain("../embeddings/chartData1.xlsx");
    const contentTypesXml = entries.get("[Content_Types].xml").toString("utf8");
    expect(contentTypesXml).toContain('/ppt/embeddings/chartData1.xlsx');
    const embeddedWorkbookText = extractOfficeText(entries.get("ppt/embeddings/chartData1.xlsx"), "chartData1.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    expect(embeddedWorkbookText).toContain("Workbook structure");
    expect(embeddedWorkbookText).toContain("Jan");
    const embeddedWorkbookEntries = zipEntries(entries.get("ppt/embeddings/chartData1.xlsx"));
    const embeddedSheetXml = embeddedWorkbookEntries.get("xl/worksheets/sheet1.xml").toString("utf8");
    expect(embeddedSheetXml).toContain("<v>18</v>");
  });

  it("applies extracted brand defaults to fallback PowerPoint themes and generated objects", () => {
    const pptx = createGeneratedPptx({
      brandProfile: { colors: { dk1: "#102A43", lt1: "#F0F4F8", accent1: "#D64545", accent2: "#2F855A", accent3: "#805AD5" }, fonts: { major: "Brand Display", minor: "Brand Sans" } },
      slides: [{ title: "Branded", footer: "Internal", shapes: [{ text: "Profile default" }], tables: [{ values: [["Metric", "Value"], ["A", 1]] }], charts: [{ title: "Trend", values: [["A", 1], ["B", 2]] }] }],
    });
    const entries = zipEntries(pptx);
    const theme = entries.get("ppt/theme/theme1.xml").toString("utf8");
    expect(theme).toContain('val="102A43"');
    expect(theme).toContain('val="D64545"');
    expect(theme).toContain('typeface="Brand Display"');
    const slide = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slide).toContain('val="D64545"');
    expect(slide).toContain('val="2F855A"');
    expect(slide).toContain('typeface="Brand Sans"');
    const chart = entries.get("ppt/charts/chart1.xml").toString("utf8");
    expect(chart).toContain('val="D64545"');
  });

  it("keeps explicit PowerPoint styles ahead of extracted brand defaults", () => {
    const pptx = createGeneratedPptx({
      brandProfile: { colors: { accent1: "#D64545", accent2: "#2F855A" }, fonts: { minor: "Brand Sans" } },
      slides: [{ shapes: [{ text: "Explicit", fillColor: "#112233", lineColor: "#445566", fontFamily: "Explicit Font" }], tables: [{ headerFillColor: "#778899", values: [["H"], ["V"]] }], charts: [{ colors: ["#ABCDEF"], values: [["A", 1], ["B", 2]] }] }],
    });
    const entries = zipEntries(pptx);
    const slide = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slide).toContain('val="112233"');
    expect(slide).toContain('val="445566"');
    expect(slide).toContain('typeface="Explicit Font"');
    expect(slide).toContain('val="778899"');
    expect(entries.get("ppt/charts/chart1.xml").toString("utf8")).toContain('val="ABCDEF"');
  });

  it("creates generated PPTX images with fit and fill crop behavior", () => {
    const pptx = createGeneratedPptx({
      fileName: "image-fit-fill.pptx",
      slides: [{
        title: "Image sizing",
        images: [
          { name: "wide.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), left: 100, top: 100, width: 200, height: 200, pixelWidth: 400, pixelHeight: 200, fit: "fit" },
          { name: "tall.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), left: 350, top: 100, width: 200, height: 200, pixelWidth: 100, pixelHeight: 200, fit: "fill" },
          { name: "crop.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), left: 100, top: 340, width: 200, height: 120, crop: { left: 0.1, right: 5, top: 0.2, bottom: 10 } },
        ],
      }],
    });
    const entries = zipEntries(pptx);
    const slideXml = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slideXml).toContain('name="wide.png"');
    expect(slideXml).toContain('<a:off x="1270000" y="1905000"/><a:ext cx="2540000" cy="1270000"/>');
    expect(slideXml).toContain('name="tall.png"');
    expect(slideXml).toContain('<a:srcRect t="25000" b="25000"/>');
    expect(slideXml).toContain('name="crop.png"');
    expect(slideXml).toContain('<a:srcRect l="10000" r="5000" t="20000" b="10000"/>');
  });

  it("creates generated PPTX tables with professional styling controls", () => {
    const pptx = createGeneratedPptx({
      fileName: "styled-table.pptx",
      slides: [{
        title: "Styled table",
        tables: [{
          name: "Executive KPI Table",
          values: [["Metric", "Status"], ["Revenue", { text: "On track", fillColor: "DCFCE7", textColor: "166534", bold: true }], ["Margin", "Watch"]],
          left: 72,
          top: 160,
          width: 480,
          height: 150,
          headerFillColor: "1F4E79",
          headerTextColor: "FFFFFF",
          bandFillColor: "EFF6FF",
          borderColor: "94A3B8",
          borderWidth: 1.25,
          align: "right",
          headerAlign: "center",
          fontSize: 1100,
          headerFontSize: 1300,
          columnWidths: [280, 200],
          rowHeights: [34, 58, 58],
        }],
      }],
    });
    const entries = zipEntries(pptx);
    const slideXml = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slideXml).toContain("Executive KPI Table");
    expect(slideXml).toContain('<a:tblPr firstRow="1" bandRow="1"/>');
    expect(slideXml).toContain('<a:gridCol w="3556000"/><a:gridCol w="2540000"/>');
    expect(slideXml).toContain('<a:tr h="431800">');
    expect(slideXml).toContain('<a:srgbClr val="1F4E79"/>');
    expect(slideXml).toContain('<a:srgbClr val="EFF6FF"/>');
    expect(slideXml).toContain('<a:srgbClr val="DCFCE7"/>');
    expect(slideXml).toContain('<a:srgbClr val="166534"/>');
    expect(slideXml).toContain('<a:pPr algn="ctr"/>');
    expect(slideXml).toContain('<a:pPr algn="r"/>');
    expect(slideXml).toContain('sz="1300" b="1"');
    expect(slideXml).toContain('<a:ln w="15875">');
  });

  it("creates generated PPTX slide footers, dates, labels, and slide numbers", () => {
    const pptx = createGeneratedPptx({
      fileName: "deck-chrome.pptx",
      footer: "CTRL Board Pack",
      dateText: "July 2026",
      confidentialityLabel: "Confidential",
      showSlideNumber: true,
      slideNumberFormat: "{n}/{total}",
      slides: [{ title: "One" }, { title: "Two", footer: "Appendix", showDate: false, confidentialityLabel: "Internal" }],
    });
    const entries = zipEntries(pptx);
    const slide1Xml = entries.get("ppt/slides/slide1.xml").toString("utf8");
    const slide2Xml = entries.get("ppt/slides/slide2.xml").toString("utf8");
    expect(slide1Xml).toContain("CTRL Board Pack");
    expect(slide1Xml).toContain("July 2026");
    expect(slide1Xml).toContain("Confidential");
    expect(slide1Xml).toContain("1/2");
    expect(slide1Xml).toContain('name="Footer"');
    expect(slide1Xml).toContain('name="Slide Number"');
    expect(slide2Xml).toContain("Appendix");
    expect(slide2Xml).not.toContain("July 2026");
    expect(slide2Xml).toContain("Internal");
    expect(slide2Xml).toContain("2/2");
  });

  it("preserves a source PPTX template package while replacing generated slides", () => {
    const templateBuffer = createGeneratedPptx({
      slides: [{ title: "Template starter", body: "Original content", images: [{ name: "brand-logo.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64") }] }],
    });

    const pptx = createGeneratedPptx({
      templateBase64: templateBuffer.toString("base64"),
      fileName: "template-output.pptx",
      slides: [{ title: "New executive story", body: "Generated into the preserved shell.", notes: "Use the brand layout." }],
    });
    const entries = zipEntries(pptx);
    expect(entries.has("ppt/theme/theme1.xml")).toBe(true);
    expect(entries.has("ppt/slideMasters/slideMaster1.xml")).toBe(true);
    expect(entries.has("ppt/slideLayouts/slideLayout1.xml")).toBe(true);
    expect(entries.has("ppt/media/image1.png")).toBe(true);
    expect(entries.get("ppt/slides/slide1.xml").toString("utf8")).toContain("New executive story");
    expect(entries.get("ppt/slides/slide1.xml").toString("utf8")).not.toContain("Original content");
    expect(entries.get("ppt/slides/_rels/slide1.xml.rels").toString("utf8")).toContain("slideLayout");
    expect(entries.get("ppt/notesSlides/notesSlide1.xml").toString("utf8")).toContain("Use the brand layout");
  });

  it("adds generated PPTX slide chrome when preserving a source template", () => {
    const templateBuffer = createGeneratedPptx({ slides: [{ title: "Template starter" }] });
    const pptx = createGeneratedPptx({
      templateBase64: templateBuffer.toString("base64"),
      footer: "Template Footer",
      dateText: "Version 4",
      showSlideNumber: true,
      slides: [{ title: "Templated output" }],
    });
    const entries = zipEntries(pptx);
    const slideXml = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slideXml).toContain("Templated output");
    expect(slideXml).toContain("Template Footer");
    expect(slideXml).toContain("Version 4");
    expect(slideXml).toContain("1");
  });

  it("fills title and body placeholders from a source PowerPoint layout", () => {
    const layoutXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Title and Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="731520"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old title</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1188720"/><a:ext cx="8229600" cy="4572000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
    const template = zipStore({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
      "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
      "ppt/presentation.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="12192000" cy="6858000" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
      "ppt/_rels/presentation.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/slideMasters/slideMaster1.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rLayout1"/></p:sldLayoutIdLst></p:sldMaster>',
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>',
      "ppt/slideLayouts/slideLayout1.xml": layoutXml,
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/theme/theme1.xml": '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Brand"/>',
    });

    const pptx = createGeneratedPptx({
      templateBase64: template.toString("base64"),
      slides: [{ title: "Filled board title", body: "Filled board body" }],
    });
    const slideXml = zipEntries(pptx).get("ppt/slides/slide1.xml").toString("utf8");
    expect(slideXml).toContain('name="Title Placeholder"');
    expect(slideXml).toContain('name="Content Placeholder"');
    expect(slideXml).toContain("Filled board title");
    expect(slideXml).toContain("Filled board body");
    expect(slideXml).not.toContain("Old title");
    expect(slideXml).not.toContain("Old body");
  });

  it("uses image, chart, and table placeholder geometry from a source PowerPoint layout", () => {
    const layoutXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Visual"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="731520"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old title</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Picture Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1270000" y="1524000"/><a:ext cx="2540000" cy="1270000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Chart Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="chart" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="4445000" y="1524000"/><a:ext cx="3175000" cy="1905000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Table Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="tbl" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1270000" y="4064000"/><a:ext cx="6350000" cy="1270000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
    const template = zipStore({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
      "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
      "ppt/presentation.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="12192000" cy="6858000" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
      "ppt/_rels/presentation.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/slideMasters/slideMaster1.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rLayout1"/></p:sldLayoutIdLst></p:sldMaster>',
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>',
      "ppt/slideLayouts/slideLayout1.xml": layoutXml,
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/theme/theme1.xml": '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Brand"/>',
    });

    const pptx = createGeneratedPptx({
      templateBase64: template.toString("base64"),
      slides: [{
        title: "Visual report",
        images: [{ name: "logo.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64") }],
        charts: [{ title: "Trend", chartType: "bar", values: [["A", 1], ["B", 2]] }],
        tables: [{ name: "Data", values: [["Metric", "Value"], ["A", 1]] }],
      }],
    });
    const entries = zipEntries(pptx);
    const slideXml = entries.get("ppt/slides/slide1.xml").toString("utf8");
    expect(slideXml).toContain('name="Picture Placeholder"');
    expect(slideXml).toContain('name="Chart Placeholder"');
    expect(slideXml).toContain('name="Table Placeholder"');
    expect(slideXml).toContain('<a:off x="1270000" y="1524000"/>');
    expect(slideXml).toContain('<a:ext cx="2540000" cy="1270000"/>');
    expect(slideXml).toContain('<a:off x="4445000" y="1524000"/>');
    expect(slideXml).toContain('<a:ext cx="3175000" cy="1905000"/>');
    expect(slideXml).toContain('<a:off x="1270000" y="4064000"/>');
    expect(slideXml).toContain('<a:ext cx="6350000" cy="1270000"/>');
    expect(slideXml).toContain("Visual report");
  });

  it("selects different source PowerPoint layouts per generated slide", () => {
    const titleLayout = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title"><p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="914400"/><a:ext cx="8229600" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Old title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>';
    const chartLayout = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="chart"><p:cSld name="Chart Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="731520"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Chart Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="chart" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1524000" y="1778000"/><a:ext cx="7620000" cy="3302000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>';
    const template = zipStore({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
      "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
      "ppt/presentation.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="12192000" cy="6858000" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
      "ppt/_rels/presentation.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/slideMasters/slideMaster1.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rLayout1"/><p:sldLayoutId id="2" r:id="rLayout2"/></p:sldLayoutIdLst></p:sldMaster>',
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rLayout2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>',
      "ppt/slideLayouts/slideLayout1.xml": titleLayout,
      "ppt/slideLayouts/slideLayout2.xml": chartLayout,
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/slideLayouts/_rels/slideLayout2.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>',
      "ppt/theme/theme1.xml": '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Brand"/>',
    });

    const pptx = createGeneratedPptx({
      templateBase64: template.toString("base64"),
      slides: [
        { title: "Opening", layoutName: "Title Slide" },
        { title: "Revenue trend", layoutType: "chart", charts: [{ title: "Revenue", chartType: "bar", values: [["Q1", 4], ["Q2", 7]] }] },
      ],
    });
    const entries = zipEntries(pptx);
    const slide1Rels = entries.get("ppt/slides/_rels/slide1.xml.rels").toString("utf8");
    const slide2Rels = entries.get("ppt/slides/_rels/slide2.xml.rels").toString("utf8");
    const slide2Xml = entries.get("ppt/slides/slide2.xml").toString("utf8");
    expect(slide1Rels).toContain("../slideLayouts/slideLayout1.xml");
    expect(slide2Rels).toContain("../slideLayouts/slideLayout2.xml");
    expect(slide2Xml).toContain('<a:off x="1524000" y="1778000"/>');
    expect(slide2Xml).toContain('<a:ext cx="7620000" cy="3302000"/>');
    expect(entries.has("ppt/charts/_rels/chart1.xml.rels")).toBe(true);
    expect(entries.has("ppt/embeddings/chartData1.xlsx")).toBe(true);
    expect(entries.get("ppt/charts/chart1.xml").toString("utf8")).toContain('c:externalData r:id="rWorkbook"');
    expect(entries.get("ppt/charts/_rels/chart1.xml.rels").toString("utf8")).toContain("../embeddings/chartData1.xlsx");
  });

  it("creates richer editable PowerPoint chart parts with labels, axes, colors, and additional chart types", () => {
    const pptx = createGeneratedPptx({
      slides: [{
        title: "Chart polish",
        charts: [{
          title: "Market share",
          chartType: "doughnut",
          values: [["A", 45], ["B", 30], ["C", 25]],
          dataLabels: true,
          colors: ["#2563EB"],
          legendPosition: "bottom",
          valueFormat: "0%",
          holeSize: 65,
        }],
      }, {
        title: "Horizontal trend",
        charts: [{
          title: "Revenue",
          chartType: "bar",
          barDirection: "horizontal",
          values: [["Q1", 4], ["Q2", 7]],
          categoryAxisTitle: "Quarter",
          valueAxisTitle: "Revenue ($M)",
          valueFormat: "$#,##0",
          showGridLines: false,
        }],
      }],
    });
    const entries = zipEntries(pptx);
    const doughnutXml = entries.get("ppt/charts/chart1.xml").toString("utf8");
    const barXml = entries.get("ppt/charts/chart2.xml").toString("utf8");
    expect(doughnutXml).toContain("<c:doughnutChart>");
    expect(doughnutXml).toContain('<c:holeSize val="65"/>');
    expect(doughnutXml).toContain("<c:dLbls>");
    expect(doughnutXml).toContain('<a:srgbClr val="2563EB"/>');
    expect(doughnutXml).toContain('<c:formatCode>0%</c:formatCode>');
    expect(barXml).toContain('<c:barDir val="bar"/>');
    expect(barXml).toContain("Quarter");
    expect(barXml).toContain("Revenue ($M)");
    expect(barXml).toContain('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>');
    expect(barXml).not.toContain("<c:majorGridlines/>");
  });

  it("creates multi-series editable PowerPoint chart parts", () => {
    const pptx = createGeneratedPptx({
      slides: [{
        title: "Regional revenue",
        charts: [{
          title: "Revenue by region",
          chartType: "line",
          categories: ["Q1", "Q2", "Q3"],
          series: [
            { name: "North", values: [10, 15, 18] },
            { name: "South", values: [8, 12, 16] },
          ],
          colors: ["#2563EB", "#7C3AED"],
          categoryAxisTitle: "Quarter",
          valueAxisTitle: "Revenue",
          valueFormat: "$#,##0",
          dataLabels: true,
          legendPosition: "bottom",
        }],
      }],
    });
    const entries = zipEntries(pptx);
    const chartXml = entries.get("ppt/charts/chart1.xml").toString("utf8");
    expect(chartXml).toContain("<c:lineChart>");
    expect(chartXml.match(/<c:ser>/g)).toHaveLength(2);
    expect(chartXml).toContain("North");
    expect(chartXml).toContain("South");
    expect(chartXml).toContain("<c:strRef><c:f>&apos;Chart Data&apos;!$A$2:$A$4</c:f>");
    expect(chartXml).toContain("<c:strRef><c:f>&apos;Chart Data&apos;!$B$1</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$B$2:$B$4</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$C$2:$C$4</c:f>");
    expect(chartXml).toContain('<a:srgbClr val="2563EB"/>');
    expect(chartXml).toContain('<a:srgbClr val="7C3AED"/>');
    expect(chartXml.match(/<c:ptCount val="3"\/>/g).length).toBeGreaterThanOrEqual(3);
    expect(chartXml).toContain("Quarter");
    expect(chartXml).toContain("Revenue");
    expect(chartXml).toContain('<c:formatCode>$#,##0</c:formatCode>');
    expect(chartXml).toContain('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>');
    expect(chartXml).toContain('<c:legendPos val="b"/>');
    expect(chartXml).toContain("<c:dLbls>");
  });

  it("creates editable PowerPoint scatter chart parts with numeric x/y series", () => {
    const pptx = createGeneratedPptx({
      slides: [{
        title: "Efficiency analysis",
        charts: [{
          title: "Cost vs output",
          chartType: "scatter",
          xAxisTitle: "Output units",
          yAxisTitle: "Cost per unit",
          scatterStyle: "marker",
          xValueFormat: "0.0",
          yValueFormat: "$0.00",
          colors: ["#7C3AED", "#0284C7"],
          series: [
            { name: "Plant A", points: [[10, 4.2], [20, 3.8], [30, 3.5]] },
            { name: "Plant B", points: [[12, 4.6], [22, 4.0], [32, 3.7]] },
          ],
        }],
      }],
    });
    const entries = zipEntries(pptx);
    const chartXml = entries.get("ppt/charts/chart1.xml").toString("utf8");
    expect(chartXml).toContain("<c:scatterChart>");
    expect(chartXml).toContain("</c:scatterChart>");
    expect(chartXml).toContain('<c:scatterStyle val="marker"/>');
    expect(chartXml).toContain("<c:xVal>");
    expect(chartXml).toContain("<c:yVal>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$A$3:$A$5</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$B$3:$B$5</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$C$3:$C$5</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$D$3:$D$5</c:f>");
    expect(chartXml).toContain("Plant A");
    expect(chartXml).toContain("Plant B");
    expect(chartXml).toContain("Output units");
    expect(chartXml).toContain("Cost per unit");
    expect(chartXml).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"/>');
    expect(chartXml).toContain('<c:numFmt formatCode="$0.00" sourceLinked="0"/>');
    expect(chartXml).toContain('<a:srgbClr val="7C3AED"/>');
    expect(chartXml).not.toContain("screenshot");
  });

  it("creates editable PowerPoint combo chart parts with column and line series", () => {
    const pptx = createGeneratedPptx({
      slides: [{
        title: "Revenue and margin",
        charts: [{
          title: "Revenue with margin rate",
          chartType: "combo",
          categories: ["Q1", "Q2", "Q3"],
          series: [
            { name: "Revenue", values: [12, 18, 22], chartType: "bar" },
            { name: "Margin", values: [0.18, 0.22, 0.25], chartType: "line" },
          ],
          categoryAxisTitle: "Quarter",
          valueAxisTitle: "Revenue / margin",
          valueFormat: "0.0",
          colors: ["#2563EB", "#F97316"],
          legendPosition: "bottom",
        }],
      }],
    });
    const entries = zipEntries(pptx);
    const chartXml = entries.get("ppt/charts/chart1.xml").toString("utf8");
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("</c:barChart>");
    expect(chartXml).toContain("<c:lineChart>");
    expect(chartXml).toContain("</c:lineChart>");
    expect(chartXml.indexOf("<c:barChart>")).toBeLessThan(chartXml.indexOf("<c:lineChart>"));
    expect(chartXml).toContain("Revenue");
    expect(chartXml).toContain("Margin");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$B$2:$B$4</c:f>");
    expect(chartXml).toContain("<c:numRef><c:f>&apos;Chart Data&apos;!$C$2:$C$4</c:f>");
    expect(chartXml).toContain("Quarter");
    expect(chartXml).toContain("Revenue / margin");
    expect(chartXml).toContain('<c:legendPos val="b"/>');
    expect(chartXml.match(/<c:axId val="123456"\/>/g).length).toBeGreaterThanOrEqual(2);
    expect(chartXml).not.toContain("placeholder");
  });
  it("creates a DOCX with headings, body text, real lists, tables, and embedded images", () => {
    const docx = createGeneratedDocx({
      title: "Dolphin Research Brief",
      tableOfContents: true,
      page: { orientation: "landscape", margins: { top: 720, right: 720, bottom: 720, left: 720 } },
      header: "Marine Research Team", footer: "Confidential", links: [{ text: "Research source", url: "https://example.test/research" }], comments: [{ text: "Verify source before publication.", author: "Reviewer" }], footnotes: [{ text: "NOAA dolphin conservation source note." }], endnotes: [{ text: "Endnote source bibliography entry." }], numberedList: ["Review evidence", { text: "Publish appendix", level: 1 }], sections: [{ heading: "Findings", body: "Dolphins use echolocation.", endnote: "Section endnote detail.", links: [{ text: "Echolocation reference", url: "https://example.test/echolocation" }], footnote: "Echolocation summary source.", bullets: ["Pods coordinate", { text: "Calves learn socially", level: 1 }], numberedList: ["Validate sightings", "Summarize findings"], table: [["Topic", "Note"], ["Conservation", "Reduce bycatch"]], images: [{ name: "dolphin.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), altText: "Dolphin photo" }] }],
    });
    const text = extractOfficeText(docx, "brief.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    expect(text).toContain("Dolphin Research Brief");
    expect(text).toContain("Dolphins use echolocation");
    expect(text).toContain("Pods coordinate");
    expect(text).toContain("Conservation");
    expect(text).toContain("Styles:");
    expect(text).toContain("Headers:");
    expect(text).toContain("Marine Research Team");
    expect(text).toContain("Footers:");
    expect(text).toContain("Confidential");
    expect(text).toContain("Comments:");
    expect(text).toContain("Verify source before publication");
    expect(text).toContain("Embedded media: image1.png");
    const entries = zipEntries(docx);
    expect(entries.has("word/media/image1.png")).toBe(true);
    expect(entries.has("word/footnotes.xml")).toBe(true);
    expect(entries.has("word/endnotes.xml")).toBe(true);
    expect(entries.has("word/numbering.xml")).toBe(true);
    expect(entries.has("word/_rels/document.xml.rels")).toBe(true);
    const docRels = entries.get("word/_rels/document.xml.rels").toString("utf8");
    expect(docRels).toContain("https://example.test/research");
    expect(docRels).toContain("https://example.test/echolocation");
    expect(docRels).toContain("TargetMode=\"External\"");
    expect(docRels).toContain("relationships/numbering");
    const documentXml = entries.get("word/document.xml").toString("utf8");
    expect(documentXml).toContain("TOC");
    expect(documentXml).toContain('w:orient="landscape"');
    expect(documentXml).toContain("w:footnoteReference");
    expect(documentXml).toContain("w:endnoteReference");
    expect(documentXml).toContain("<w:numPr>");
    expect(documentXml).toContain('<w:numId w:val="1"/>');
    expect(documentXml).toContain('<w:numId w:val="2"/>');
    expect(documentXml).toContain('<w:ilvl w:val="1"/>');
    expect(documentXml).toContain("Research source");
    expect(documentXml).toContain("Echolocation reference");
    expect(documentXml).toContain("w:hyperlink");
    expect(documentXml).not.toContain(">- Pods coordinate<");
    const numberingXml = entries.get("word/numbering.xml").toString("utf8");
    expect(numberingXml).toContain('<w:numFmt w:val="bullet"/>');
    expect(numberingXml).toContain('<w:numFmt w:val="decimal"/>');
    expect(numberingXml).toContain('<w:lvlText w:val="%1."/>');
    const footnotesXml = entries.get("word/footnotes.xml").toString("utf8");
    expect(footnotesXml).toContain("NOAA dolphin conservation source note");
    expect(footnotesXml).toContain("Echolocation summary source");
    const endnotesXml = entries.get("word/endnotes.xml").toString("utf8");
    expect(endnotesXml).toContain("Endnote source bibliography entry");
    expect(endnotesXml).toContain("Section endnote detail");
  });

  it("preserves a DOCX template package shell while replacing the generated body", () => {
    const source = createGeneratedDocx({ title: "Source template", header: "Template Header", footer: "Template Footer" });
    const sourceEntries = Object.fromEntries([...zipEntries(source)].map(([name, value]) => [name, value]));
    sourceEntries["word/theme/theme1.xml"] = "<w:theme xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">brand-theme</w:theme>";
    sourceEntries["customXml/item1.xml"] = "<brand>CTRL template metadata</brand>";
    const generated = createGeneratedDocx({
      title: "Generated report",
      templateBase64: zipStore(sourceEntries).toString("base64"),
      sections: [{ heading: "Findings", body: "Generated content replaces the template body." }],
    });
    const entries = zipEntries(generated);
    expect(entries.get("word/styles.xml").toString("utf8")).toBe(sourceEntries["word/styles.xml"].toString("utf8"));
    expect(entries.get("word/theme/theme1.xml").toString("utf8")).toContain("brand-theme");
    expect(entries.get("customXml/item1.xml").toString("utf8")).toContain("CTRL template metadata");
    expect(entries.get("word/document.xml").toString("utf8")).toContain("Generated content replaces the template body.");
    expect(entries.get("word/document.xml").toString("utf8")).not.toContain("Source template");
    expect(entries.get("word/_rels/document.xml.rels").toString("utf8")).toContain("header");
    expect(entries.get("word/_rels/document.xml.rels").toString("utf8")).toContain("footer");
    expect(validateGeneratedOfficePackage(generated, "docx").entryCount).toBeGreaterThan(entries.size - 1);
  });

  it("rejects invalid DOCX template input", () => {
    expect(() => createGeneratedDocx({ title: "Invalid", templateBase64: Buffer.from("not-a-docx").toString("base64") })).toThrow(/valid, bounded DOCX/);
  });

  it("creates a DOCX with tracked-change revision markup", () => {
    const docx = createGeneratedDocx({
      title: "Contract Review",
      revisions: [
        { type: "delete", text: "Supplier may terminate at any time.", author: "Legal" },
        { type: "insert", text: "Either party may terminate with 30 days written notice.", author: "Legal" },
      ],
      sections: [{
        heading: "Commercial terms",
        body: "Review language below.",
        revisions: [{ type: "insert", text: "Payment terms are net 30 days.", author: "Finance" }],
      }],
    });
    const entries = zipEntries(docx);
    expect(entries.has("word/settings.xml")).toBe(true);
    const documentXml = entries.get("word/document.xml").toString("utf8");
    const settingsXml = entries.get("word/settings.xml").toString("utf8");
    const relsXml = entries.get("word/_rels/document.xml.rels").toString("utf8");
    expect(documentXml).toContain("<w:del ");
    expect(documentXml).toContain("<w:ins ");
    expect(documentXml).toContain("<w:delText");
    expect(documentXml).toContain("w:author=\"Legal\"");
    expect(documentXml).toContain("Supplier may terminate at any time.");
    expect(documentXml).toContain("Either party may terminate with 30 days written notice.");
    expect(documentXml).toContain("Payment terms are net 30 days.");
    expect(settingsXml).toContain("<w:trackRevisions/>");
    expect(relsXml).toContain("relationships/settings");
  });

  it("creates a DOCX with real columns and section-specific page layout", () => {
    const docx = createGeneratedDocx({
      title: "Market Brief",
      page: { orientation: "portrait", margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      sections: [
        {
          heading: "Two-column insights",
          body: "Demand signals are rising across the segment.",
          columns: { count: 2, space: 540, separator: true },
          page: { margins: { top: 720, right: 900, bottom: 720, left: 900 } },
        },
        {
          heading: "Landscape appendix",
          body: "Detailed tables use a wider page.",
          page: { orientation: "landscape", margins: { top: 720, right: 720, bottom: 720, left: 720 } },
          columns: { count: 1 },
        },
      ],
    });
    const entries = zipEntries(docx);
    const documentXml = entries.get("word/document.xml").toString("utf8");
    expect(documentXml.match(/<w:sectPr>/g).length).toBe(2);
    expect(documentXml).toContain('<w:cols w:num="2" w:space="540" w:sep="1"/>');
    expect(documentXml).toContain('<w:pgMar w:top="720" w:right="900" w:bottom="720" w:left="900"/>');
    expect(documentXml).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>');
    expect(documentXml).not.toContain("Click Layout");
  });

  it("creates a DOCX with real caption fields, bookmarks, and cross-references", () => {
    const docx = createGeneratedDocx({
      title: "Evidence Report",
      images: [{ name: "chart.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), caption: { label: "Figure", text: "Quarterly performance", bookmark: "fig_quarterly_performance" } }],
      crossReferences: [{ bookmark: "fig_quarterly_performance", text: "See", fallback: "Figure 1" }],
      sections: [{ heading: "Appendix", body: "The source table is referenced below.", table: [["Metric", "Value"], ["Growth", "12%"]], tableCaption: { label: "Table", text: "Source metrics", bookmark: "tbl_source_metrics" }, crossReferences: [{ bookmark: "tbl_source_metrics", text: "Refer to", fallback: "Table 1" }] }],
    });
    const entries = zipEntries(docx);
    const documentXml = entries.get("word/document.xml").toString("utf8");
    expect(documentXml).toContain('<w:pStyle w:val="Caption"/>');
    expect(documentXml).toContain('w:name="fig_quarterly_performance"');
    expect(documentXml).toContain('w:name="tbl_source_metrics"');
    expect(documentXml).toContain('w:instr="SEQ Figure \\* ARABIC"');
    expect(documentXml).toContain('w:instr="SEQ Table \\* ARABIC"');
    expect(documentXml).toContain('w:instr="REF fig_quarterly_performance \\h"');
    expect(documentXml).toContain('w:instr="REF tbl_source_metrics \\h"');
    expect(documentXml).toContain("Quarterly performance");
    expect(documentXml).toContain("Source metrics");
    expect(documentXml).not.toContain("insert caption manually");
  });

  it("creates an XLSX with workbook structure, typed rows, formulas, and embedded images", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "dolphin-data.xlsx",
      sheets: [{ name: "Data", rows: [["Month", "Sightings", "Status"], ["Jan", 12, "OK"], ["Feb", 18, "Review"], ["Total", "=SUM(B2:B3)", ""]], tables: [{ name: "SightingsTable", ref: "A1:C4", columns: ["Month", "Sightings", "Status"] }], validations: [{ address: "C2:C4", values: ["OK", "Review"] }], conditionalFormats: [{ address: "B2:B4", operator: "greaterThan", value: 15, fillColor: "#FFF2CC" }], charts: [{ title: "Monthly sightings", chartType: "line", values: [["Jan", 12], ["Feb", 18], ["Mar", 15]], cell: "E12" }], images: [{ name: "dolphin.png", type: "image/png", base64: Buffer.from(MINIMAL_PNG_BASE64, "base64").toString("base64"), cell: "E2", altText: "Dolphin chart companion" }] }],
    });
    const text = extractOfficeText(xlsx, "dolphin-data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    expect(text).toContain("Workbook structure");
    expect(text).toContain("Worksheet 1: 4 rows");
    expect(text).toContain("Month");
    expect(text).toContain("Sightings");
    expect(text).toContain("Table: SightingsTable (A1:C4)");
    expect(text).toContain("Charts: chart1.xml");
    expect(text).toContain("Embedded media: image1.png");
    const entries = zipEntries(xlsx);
    expect(entries.has("docProps/core.xml")).toBe(true);
    expect(entries.has("docProps/app.xml")).toBe(true);
    expect(entries.has("xl/styles.xml")).toBe(true);
    expect(entries.has("xl/tables/table1.xml")).toBe(true);
    expect(entries.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);
    expect(entries.has("xl/media/image1.png")).toBe(true);
    expect(entries.has("xl/charts/chart1.xml")).toBe(true);
    expect(entries.has("xl/drawings/drawing1.xml")).toBe(true);
    expect(entries.has("xl/drawings/_rels/drawing1.xml.rels")).toBe(true);
    const drawingXml = entries.get("xl/drawings/drawing1.xml").toString("utf8");
    expect(drawingXml).toContain("rChart1");
    const chartXml = entries.get("xl/charts/chart1.xml").toString("utf8");
    expect(chartXml).toContain("Monthly sightings");
    expect(chartXml).toContain("<c:lineChart>");
    expect(chartXml).not.toContain('c:externalData r:id="rWorkbook"');
    expect(chartXml).not.toContain("&apos;Chart Data&apos;!");
  });

  it("creates XLSX document properties for professional workbook handoff", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "executive-dashboard.xlsx",
      properties: {
        title: "Executive Dashboard",
        subject: "Quarterly performance",
        creator: "Finance Ops",
        keywords: ["forecast", "board", "revenue"],
        description: "Board-ready workbook generated from CTRL.",
        category: "Executive Reporting",
        company: "Expedient",
        manager: "CEO Office",
      },
      sheets: [{ name: "Summary", rows: [["Metric", "Value"], ["Revenue", 22]] }],
    });
    const entries = zipEntries(xlsx);
    const relsXml = entries.get("_rels/.rels").toString("utf8");
    const contentTypesXml = entries.get("[Content_Types].xml").toString("utf8");
    const coreXml = entries.get("docProps/core.xml").toString("utf8");
    const appXml = entries.get("docProps/app.xml").toString("utf8");
    expect(relsXml).toContain("metadata/core-properties");
    expect(relsXml).toContain("extended-properties");
    expect(contentTypesXml).toContain('/docProps/core.xml');
    expect(contentTypesXml).toContain('/docProps/app.xml');
    expect(coreXml).toContain("<dc:title>Executive Dashboard</dc:title>");
    expect(coreXml).toContain("<dc:subject>Quarterly performance</dc:subject>");
    expect(coreXml).toContain("<dc:creator>Finance Ops</dc:creator>");
    expect(coreXml).toContain("<cp:keywords>forecast, board, revenue</cp:keywords>");
    expect(coreXml).toContain("Board-ready workbook generated from CTRL.");
    expect(coreXml).toContain("<cp:category>Executive Reporting</cp:category>");
    expect(appXml).toContain("<Company>Expedient</Company>");
    expect(appXml).toContain("<Manager>CEO Office</Manager>");
    expect(appXml).toContain("<vt:lpstr>Summary</vt:lpstr>");
  });

  it("creates XLSX worksheet view and page layout settings", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "print-ready.xlsx",
      sheets: [{
        name: "Print Ready",
        rows: [["Region", "Revenue"], ["North", 12], ["South", 18]],
        freeze: { rows: 1, columns: 1 },
        zoomScale: 125,
        orientation: "landscape",
        paperSize: "letter",
        margins: { top: 0.5, right: 0.4, bottom: 0.5, left: 0.4, header: 0.2, footer: 0.2 },
        printArea: "$A$1:$B$3",
        repeatRows: "1:1",
        repeatColumns: "A:A",
        fitToPagesWide: 1,
        fitToPagesTall: 1,
        showGridlines: false,
        showHeadings: true,
        centerHorizontally: true,
        centerVertically: false,
        blackAndWhite: true,
        draftMode: true,
      }],
    });
    const entries = zipEntries(xlsx);
    const sheetXml = entries.get("xl/worksheets/sheet1.xml").toString("utf8");
    const workbookXml = entries.get("xl/workbook.xml").toString("utf8");
    expect(sheetXml).toContain('<sheetView workbookViewId="0" showGridLines="0" showRowColHeaders="1" zoomScale="125">');
    expect(sheetXml).toContain('<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>');
    expect(sheetXml).toContain('<printOptions gridLines="0" headings="1" horizontalCentered="1" verticalCentered="0"/>');
    expect(sheetXml).toContain('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>');
    expect(sheetXml).toContain('<pageSetup orientation="landscape" paperSize="1" fitToWidth="1" fitToHeight="1" blackAndWhite="1" draft="1"/>');
    expect(workbookXml).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Print Ready&apos;!$A$1:$B$3</definedName>');
    expect(workbookXml).toContain('<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;Print Ready&apos;!$1:$1,&apos;Print Ready&apos;!$A:$A</definedName>');
  });

  it("creates XLSX worksheet protection for handoff workbooks", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "protected-model.xlsx",
      sheets: [{
        name: "Model",
        rows: [["Input", "Value"], ["Revenue", 22]],
        protection: { protect: true, password: "review", allowFormatCells: true, allowSort: true, allowAutoFilter: true },
      }],
    });
    const entries = zipEntries(xlsx);
    const sheetXml = entries.get("xl/worksheets/sheet1.xml").toString("utf8");
    expect(sheetXml).toContain("<sheetProtection ");
    expect(sheetXml).toContain('sheet="1"');
    expect(sheetXml).toContain('objects="1"');
    expect(sheetXml).toContain('scenarios="1"');
    expect(sheetXml).toContain('formatCells="0"');
    expect(sheetXml).toContain('sort="0"');
    expect(sheetXml).toContain('autoFilter="0"');
    expect(sheetXml).toMatch(/password="[0-9A-F]{4}"/);
    expect(sheetXml.indexOf("<sheetProtection")).toBeGreaterThan(sheetXml.indexOf("</sheetData>"));
    expect(sheetXml.indexOf("<sheetProtection")).toBeLessThan(sheetXml.indexOf("<pageMargins"));
  });

  it("creates XLSX named ranges for formulas and automation", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "named-ranges.xlsx",
      namedRanges: [{ name: "Revenue_Table", sheetName: "Summary", address: "$A$1:$B$3" }],
      sheets: [{
        name: "Summary",
        rows: [["Region", "Revenue"], ["North", 12], ["South", 18]],
        namedRanges: [
          { name: "Inputs", address: "$B$2:$B$3" },
          { name: "Bad Name 123", address: "$A$1:$A$3", scope: "workbook" },
        ],
      }],
    });
    const entries = zipEntries(xlsx);
    const workbookXml = entries.get("xl/workbook.xml").toString("utf8");
    expect(workbookXml).toContain("<definedNames>");
    expect(workbookXml).toContain('<definedName name="Revenue_Table">&apos;Summary&apos;!$A$1:$B$3</definedName>');
    expect(workbookXml).toContain('<definedName name="Inputs" localSheetId="0">&apos;Summary&apos;!$B$2:$B$3</definedName>');
    expect(workbookXml).toContain('<definedName name="Bad_Name_123">&apos;Summary&apos;!$A$1:$A$3</definedName>');
  });

  it("creates XLSX cell comments for review and audit handoff", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "commented-model.xlsx",
      sheets: [{
        name: "Assumptions",
        rows: [["Metric", "Value"], ["Growth", 0.12]],
        comments: [
          { address: "B2", text: "Validate growth assumption before board review.", author: "Finance" },
          { cell: "A1", text: "Source: FY planning model", author: "CTRL", visible: true },
        ],
      }],
    });
    const entries = zipEntries(xlsx);
    expect(entries.has("xl/comments/comment1.xml")).toBe(true);
    expect(entries.has("xl/drawings/vmlDrawing1.vml")).toBe(true);
    const contentTypesXml = entries.get("[Content_Types].xml").toString("utf8");
    const sheetXml = entries.get("xl/worksheets/sheet1.xml").toString("utf8");
    const sheetRels = entries.get("xl/worksheets/_rels/sheet1.xml.rels").toString("utf8");
    const commentsXml = entries.get("xl/comments/comment1.xml").toString("utf8");
    const vmlXml = entries.get("xl/drawings/vmlDrawing1.vml").toString("utf8");
    expect(contentTypesXml).toContain('/xl/comments/comment1.xml');
    expect(contentTypesXml).toContain('Extension="vml"');
    expect(sheetXml).toContain('<legacyDrawing r:id="rVmlDrawing1"/>');
    expect(sheetRels).toContain("../comments/comment1.xml");
    expect(sheetRels).toContain("../drawings/vmlDrawing1.vml");
    expect(commentsXml).toContain("<author>Finance</author>");
    expect(commentsXml).toContain("<author>CTRL</author>");
    expect(commentsXml).toContain('<comment ref="B2" authorId="0">');
    expect(commentsXml).toContain("Validate growth assumption before board review.");
    expect(commentsXml).toContain('<comment ref="A1" authorId="1">');
    expect(vmlXml).toContain('<x:Row>1</x:Row><x:Column>1</x:Column>');
    expect(vmlXml).toContain('visibility:visible');
  });

  it("creates XLSX clickable hyperlinks for source and drill-through references", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "linked-model.xlsx",
      sheets: [{
        name: "Sources",
        rows: [["Source", "URL"], ["NOAA", "Open source"]],
        links: [
          { address: "B2", url: "https://example.test/source", text: "NOAA source", tooltip: "Open NOAA source" },
          { cell: "A2", target: "https://example.test/detail", display: "Details" },
        ],
      }],
    });
    const entries = zipEntries(xlsx);
    const sheetXml = entries.get("xl/worksheets/sheet1.xml").toString("utf8");
    const sheetRels = entries.get("xl/worksheets/_rels/sheet1.xml.rels").toString("utf8");
    expect(sheetXml).toContain("<hyperlinks>");
    expect(sheetXml).toContain('<hyperlink ref="B2" r:id="rHyperlink1" display="NOAA source" tooltip="Open NOAA source"/>');
    expect(sheetXml).toContain('<hyperlink ref="A2" r:id="rHyperlink2" display="Details" tooltip="Details"/>');
    expect(sheetRels).toContain("relationships/hyperlink");
    expect(sheetRels).toContain('Target="https://example.test/source"');
    expect(sheetRels).toContain('Target="https://example.test/detail"');
    expect(sheetRels).toContain('TargetMode="External"');
  });

  it("creates XLSX column widths and merged cell layout", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "layout-polish.xlsx",
      sheets: [{
        name: "Executive Summary",
        rows: [["Executive Summary", "", ""], ["Metric", "Value", "Note"], ["Revenue", 22, "Draft"]],
        columns: [24, { column: "B", width: 14 }, { min: "C", max: "D", width: 30, hidden: true }],
        merges: ["A1:C1", { ref: "A4:D4" }],
      }],
    });
    const entries = zipEntries(xlsx);
    const sheetXml = entries.get("xl/worksheets/sheet1.xml").toString("utf8");
    expect(sheetXml).toContain('<cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="4" width="30" customWidth="1" hidden="1"/></cols>');
    expect(sheetXml).toContain('<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A4:D4"/></mergeCells>');
    expect(sheetXml.indexOf("<cols>")).toBeLessThan(sheetXml.indexOf("<sheetData>"));
    expect(sheetXml.indexOf("<mergeCells")).toBeGreaterThan(sheetXml.indexOf("</sheetData>"));
  });

  it("creates XLSX row heights and plain range autofilter", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "filterable-report.xlsx",
      sheets: [{
        name: "Report",
        rows: [["Region", "Revenue"], ["North", 12], ["South", 18]],
        rowHeights: [28, { row: 2, height: 20 }, { start: 3, end: 3, height: 18, hidden: true }],
        autoFilter: "A1:B3",
      }],
    });
    const entries = zipEntries(xlsx);
    const sheetXml = entries.get("xl/worksheets/sheet1.xml").toString("utf8");
    expect(sheetXml).toContain('<row r="1" ht="28" customHeight="1">');
    expect(sheetXml).toContain('<row r="2" ht="20" customHeight="1">');
    expect(sheetXml).toContain('<row r="3" ht="18" customHeight="1" hidden="1">');
    expect(sheetXml).toContain('<autoFilter ref="A1:B3"/>');
    expect(sheetXml.indexOf("<autoFilter")).toBeGreaterThan(sheetXml.indexOf("</sheetData>"));
  });

  it("creates richer self-contained XLSX chart parts for generated workbook fallbacks", () => {
    const xlsx = createGeneratedXlsx({
      fileName: "chart-fallbacks.xlsx",
      sheets: [{
        name: "Charts",
        rows: [["Quarter", "Revenue", "Margin"], ["Q1", 12, 0.18], ["Q2", 18, 0.22], ["Q3", 22, 0.25]],
        charts: [
          { title: "Area trend", chartType: "area", categories: ["Q1", "Q2", "Q3"], series: [{ name: "Revenue", values: [12, 18, 22] }], cell: "E2" },
          { title: "Mix", chartType: "doughnut", values: [["Core", 0.72], ["New", 0.28]], holeSize: 60, dataLabels: true, cell: "E18" },
          { title: "Cost curve", chartType: "scatter", series: [{ name: "Plant A", points: [[10, 4.2], [20, 3.8]] }], cell: "M2" },
          { title: "Revenue plus margin", chartType: "combo", categories: ["Q1", "Q2", "Q3"], series: [{ name: "Revenue", values: [12, 18, 22], chartType: "bar" }, { name: "Margin", values: [0.18, 0.22, 0.25], chartType: "line" }], cell: "M18" },
        ],
      }],
    });
    const entries = zipEntries(xlsx);
    expect(entries.has("xl/charts/chart1.xml")).toBe(true);
    expect(entries.has("xl/charts/chart4.xml")).toBe(true);
    expect(entries.get("xl/charts/chart1.xml").toString("utf8")).toContain("<c:areaChart>");
    expect(entries.get("xl/charts/chart2.xml").toString("utf8")).toContain("<c:doughnutChart>");
    expect(entries.get("xl/charts/chart2.xml").toString("utf8")).toContain('<c:holeSize val="60"/>');
    expect(entries.get("xl/charts/chart3.xml").toString("utf8")).toContain("<c:scatterChart>");
    const comboXml = entries.get("xl/charts/chart4.xml").toString("utf8");
    expect(comboXml).toContain("<c:barChart>");
    expect(comboXml).toContain("<c:lineChart>");
    expect(comboXml).not.toContain('c:externalData r:id="rWorkbook"');
    const drawingRels = entries.get("xl/drawings/_rels/drawing1.xml.rels").toString("utf8");
    expect(drawingRels).toContain("../charts/chart1.xml");
    expect(drawingRels).toContain("../charts/chart4.xml");
  });

});
