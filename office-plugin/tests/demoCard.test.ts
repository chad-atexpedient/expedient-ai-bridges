import { describe, expect, it } from "vitest";
import { demoCardFor, encodeDemoCard, isSimpleDemoRequest, parseDemoCard } from "../src/App";

describe("in-chat demo card", () => {
  it("detects simple demo requests without treating explicit creation requests as local-only demos", () => {
    expect(isSimpleDemoRequest("show me a demo")).toBe(true);
    expect(isSimpleDemoRequest("what can CTRL do?")).toBe(true);
    expect(isSimpleDemoRequest("create a demo deck in PowerPoint")).toBe(false);
    expect(isSimpleDemoRequest("make downloadable demo files")).toBe(false);
  });

  it("round-trips the structured demo card payload", () => {
    const card = demoCardFor("powerpoint", "CTRL");
    const encoded = encodeDemoCard(card);
    const parsed = parseDemoCard(encoded);

    expect(parsed?.kind).toBe("ctrl-demo-showcase");
    expect(parsed?.title).toBe("CTRL demo");
    expect(parsed?.subtitle).toContain("PowerPoint");
    expect(parsed?.actions.map((action) => action.id)).toEqual(["live", "artifact"]);
  });
});
