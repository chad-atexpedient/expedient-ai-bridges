import { describe, expect, it } from "vitest";
import { executeToolCall } from "../src/lib/tools";
import { registerUploadedAsset } from "../src/lib/uploadRegistry";

describe("generated Office tool asset resolution", () => {
  it("creates prepared CTRL demo showcase artifacts for all Office surfaces", async () => {
    const previousFetch = globalThis.fetch;
    const seen: Array<{ path: string; body: any }> = [];
    try {
      globalThis.fetch = async (input, init) => {
        const path = String(input);
        const body = JSON.parse(String(init?.body));
        seen.push({ path, body });
        return new Response(JSON.stringify({ ok: true, downloadUrl: `/api/generated/files/demo-${seen.length}` }), { status: 200 });
      };

      const result = await executeToolCall("excel", {
        id: "demo-all-1",
        name: "ctrl_create_demo_showcase",
        arguments: { surface: "all", audience: "executive", theme: "Quarterly leadership demo" },
      });

      expect(result.ok).toBe(true);
      expect(seen.map((item) => item.path)).toEqual(["/api/generated/xlsx", "/api/generated/docx", "/api/generated/pptx"]);
      expect(seen[0].body.sheets[0].charts[0].chartType).toBe("combo");
      expect(seen[1].body.tableOfContents).toBe(true);
      expect(seen[2].body.showSlideNumber).toBe(true);
      expect(result.content).toContain("Quarterly leadership demo");
      expect(result.content).toContain("markdownLink");
      expect(result.content).toContain("[Download the PowerPoint artifact](/api/generated/files/demo-3)");
      expect(result.content).toContain("demo-3");
    } finally {
      globalThis.fetch = previousFetch;
    }
  });

  it("keeps downloadable demo artifacts available through artifact mode", async () => {
    const previousFetch = globalThis.fetch;
    const seenPaths: string[] = [];
    try {
      globalThis.fetch = async (input, init) => {
        seenPaths.push(String(input));
        const body = JSON.parse(String(init?.body));
        expect(body.fileName).toContain("PowerPoint");
        return new Response(JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/ppt-demo" }), { status: 200 });
      };

      const result = await executeToolCall("powerpoint", {
        id: "demo-current-1",
        name: "ctrl_create_demo_showcase",
        arguments: { surface: "current", mode: "artifact" },
      });

      expect(result.ok).toBe(true);
      expect(result.content).toContain("[Download the PowerPoint artifact](/api/generated/files/ppt-demo)");
      expect(seenPaths).toEqual(["/api/generated/pptx"]);
    } finally {
      globalThis.fetch = previousFetch;
    }
  });

  it("does not silently generate an artifact when live demo mode is explicitly requested without Office APIs", async () => {
    const result = await executeToolCall("powerpoint", {
      id: "demo-live-1",
      name: "ctrl_create_demo_showcase",
      arguments: { surface: "current", mode: "live" },
    });

    expect(result.ok).toBe(false);
    expect(result.content).toContain("Could not create the live powerpoint demo");
    expect(result.content).toContain("not available in powerpoint");
  });

  it("resolves uploaded image names before creating generated PowerPoint files", async () => {
    registerUploadedAsset({
      id: "asset-1",
      name: "dolphin.png",
      type: "image/png",
      size: 10,
      dataUrl: "data:image/png;base64,ZmFrZQ==",
      base64: "ZmFrZQ==",
    });

    const previousFetch = globalThis.fetch;
    try {
      globalThis.fetch = async (input, init) => {
        expect(String(input)).toBe("/api/generated/pptx");
        const body = JSON.parse(String(init?.body));
        expect(body.slides[0].images[0]).toMatchObject({
          assetName: "dolphin.png",
          name: "dolphin.png",
          type: "image/png",
          base64: "ZmFrZQ==",
        });
        return new Response(JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/test.pptx" }), { status: 200 });
      };

      const result = await executeToolCall("powerpoint", {
        id: "gen-1",
        name: "powerpoint_generate_deck_file",
        arguments: { slides: [{ title: "Dolphins", images: [{ assetName: "dolphin.png" }] }] },
      });
      expect(result.ok).toBe(true);
      expect(result.content).toContain("downloadUrl");
      expect(result.content).toContain("[Download the PowerPoint artifact](/api/generated/files/test.pptx)");
    } finally {
      globalThis.fetch = previousFetch;
    }
  });

  it("resolves public image URLs before creating generated PowerPoint files", async () => {
    const previousFetch = globalThis.fetch;
    try {
      globalThis.fetch = async (input, init) => {
        if (String(input).startsWith("/api/image-asset")) {
          return new Response(JSON.stringify({ name: "web-dolphin.jpg", type: "image/jpeg", size: 10, dataUrl: "data:image/jpeg;base64,d2Vi", base64: "d2Vi" }), { status: 200 });
        }
        expect(String(input)).toBe("/api/generated/pptx");
        const body = JSON.parse(String(init?.body));
        expect(body.slides[0].images[0]).toMatchObject({
          imageUrl: "https://example.test/dolphin.jpg",
          name: "web-dolphin.jpg",
          type: "image/jpeg",
          base64: "d2Vi",
        });
        return new Response(JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/test.pptx" }), { status: 200 });
      };

      const result = await executeToolCall("powerpoint", {
        id: "gen-2",
        name: "powerpoint_generate_deck_file",
        arguments: { slides: [{ title: "Dolphins", images: [{ imageUrl: "https://example.test/dolphin.jpg" }] }] },
      });
      expect(result.ok).toBe(true);
    } finally {
      globalThis.fetch = previousFetch;
    }
  });

  it("resolves uploaded PowerPoint template names before creating generated PowerPoint files", async () => {
    registerUploadedAsset({
      id: "template-1",
      name: "board-template.pptx",
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      size: 20,
      dataUrl: "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,cHB0eA==",
      base64: "cHB0eA==",
      kind: "office-template",
      brandProfile: { source: "theme1.xml", colors: { accent1: "#336699" }, fonts: { major: "Aptos Display", minor: "Aptos" }, layouts: [], media: [], guidance: ["Use the extracted theme colors for accents and backgrounds."] },
    });

    const previousFetch = globalThis.fetch;
    try {
      globalThis.fetch = async (input, init) => {
        expect(String(input)).toBe("/api/generated/pptx");
        const body = JSON.parse(String(init?.body));
      expect(body.template).toMatchObject({ name: "board-template.pptx", base64: "cHB0eA==", brandProfile: { colors: { accent1: "#336699" } } });
        return new Response(JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/template-output.pptx" }), { status: 200 });
      };

      const result = await executeToolCall("powerpoint", {
        id: "gen-template-1",
        name: "powerpoint_generate_deck_file",
        arguments: { templateAssetName: "board-template.pptx", slides: [{ title: "Board update" }] },
      });
      expect(result.ok).toBe(true);
    } finally {
      globalThis.fetch = previousFetch;
    }
  });

  it("resolves uploaded image names before creating generated Word and Excel files", async () => {
    registerUploadedAsset({
      id: "asset-2",
      name: "logo.png",
      type: "image/png",
      size: 10,
      dataUrl: "data:image/png;base64,bG9nbw==",
      base64: "bG9nbw==",
    });

    const seenPaths: string[] = [];
    const previousFetch = globalThis.fetch;
    try {
      globalThis.fetch = async (input, init) => {
        seenPaths.push(String(input));
        const body = JSON.parse(String(init?.body));
        if (String(input) === "/api/generated/docx") {
          expect(body.sections[0].images[0]).toMatchObject({ assetName: "logo.png", name: "logo.png", type: "image/png", base64: "bG9nbw==" });
        }
        if (String(input) === "/api/generated/xlsx") {
          expect(body.sheets[0].images[0]).toMatchObject({ assetName: "logo.png", name: "logo.png", type: "image/png", base64: "bG9nbw==" });
        }
        return new Response(JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/test" }), { status: 200 });
      };

      const wordResult = await executeToolCall("word", {
        id: "gen-word-1",
        name: "word_generate_document_file",
        arguments: { title: "Logo Brief", sections: [{ heading: "Brand", images: [{ assetName: "logo.png" }] }] },
      });
      const excelResult = await executeToolCall("excel", {
        id: "gen-excel-1",
        name: "excel_generate_workbook_file",
        arguments: { sheets: [{ name: "Brand", rows: [["Asset"]], images: [{ assetName: "logo.png", cell: "B2" }] }] },
      });

      expect(wordResult.ok).toBe(true);
      expect(excelResult.ok).toBe(true);
      expect(seenPaths).toEqual(["/api/generated/docx", "/api/generated/xlsx"]);
    } finally {
      globalThis.fetch = previousFetch;
    }
  });
  it("imports public web image URLs as reusable assets and resolves them for Word and Excel files", async () => {
    const previousFetch = globalThis.fetch;
    const seenPaths: string[] = [];
    try {
      globalThis.fetch = async (input, init) => {
        seenPaths.push(String(input));
        if (String(input).startsWith("/api/image-asset")) {
          return new Response(JSON.stringify({ name: "web-logo.png", type: "image/png", size: 12, dataUrl: "data:image/png;base64,d2ViLWxvZ28=", base64: "d2ViLWxvZ28=" }), { status: 200 });
        }
        const body = JSON.parse(String(init?.body));
        if (String(input) === "/api/generated/docx") {
          expect(body.images[0]).toMatchObject({ imageUrl: "https://example.test/logo.png", name: "web-logo.png", type: "image/png", base64: "d2ViLWxvZ28=" });
        }
        if (String(input) === "/api/generated/xlsx") {
          expect(body.sheets[0].images[0]).toMatchObject({ assetName: "web-logo.png", name: "web-logo.png", type: "image/png", base64: "d2ViLWxvZ28=" });
        }
        return new Response(JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/test" }), { status: 200 });
      };

      const importResult = await executeToolCall("word", {
        id: "import-web-image-1",
        name: "web_image_import",
        arguments: { imageUrl: "https://example.test/logo.png" },
      });
      expect(importResult.ok).toBe(true);
      expect(importResult.content).toContain("web-logo.png");

      const wordResult = await executeToolCall("word", {
        id: "gen-word-web-1",
        name: "word_generate_document_file",
        arguments: { title: "Web Logo Brief", images: [{ imageUrl: "https://example.test/logo.png" }] },
      });
      const excelResult = await executeToolCall("excel", {
        id: "gen-excel-web-1",
        name: "excel_generate_workbook_file",
        arguments: { sheets: [{ name: "Web Brand", rows: [["Asset"]], images: [{ assetName: "web-logo.png", cell: "B2" }] }] },
      });

      expect(wordResult.ok).toBe(true);
      expect(excelResult.ok).toBe(true);
      expect(seenPaths).toEqual([expect.stringContaining("/api/image-asset"), "/api/generated/docx", "/api/generated/xlsx"]);
    } finally {
      globalThis.fetch = previousFetch;
    }
  });
});

