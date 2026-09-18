import assert from "node:assert/strict";
import test from "node:test";

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import JSZip from "jszip";

import { createItemSource } from "../../src/review/itemSource";
import type { ReviewItem } from "../../src/review/model";

Object.assign(globalThis, { DOMParser, XMLSerializer });

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

type Segment = string | { literal: "tab" | "break" };
type Paragraph = { nativeId?: string; segments: Segment[] };

function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function docxFile(paragraphs: Paragraph[]): Promise<{
  file: File;
  arrayBufferReads: () => number;
}> {
  const body = paragraphs
    .map(
      ({ nativeId, segments }) =>
        `<w:p${nativeId ? ` w14:paraId="${nativeId}"` : ""}>${segments
          .map((segment) => {
            if (typeof segment === "string") {
              return `<w:r><w:t>${escapeXml(segment)}</w:t></w:r>`;
            }
            return segment.literal === "tab"
              ? "<w:r><w:tab/></w:r>"
              : "<w:r><w:br/></w:r>";
          })
          .join("")}</w:p>`,
    )
    .join("");
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`,
  );
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const file = new File([Uint8Array.from(bytes).buffer], "review.docx", {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const original = file.arrayBuffer.bind(file);
  let reads = 0;
  Object.defineProperty(file, "arrayBuffer", {
    value: async () => {
      reads += 1;
      return original();
    },
  });
  return { file, arrayBufferReads: () => reads };
}

function comment(
  id: string,
  quote: string,
  anchor: Partial<ReviewItem["anchor"]> = {},
): ReviewItem {
  return {
    id,
    kind: "comment",
    anchor: { quote, ...anchor },
    body: `Comment ${id}`,
  };
}

function redline(
  id: string,
  quote: string,
  replacement: string,
): ReviewItem {
  return { id, kind: "redline", anchor: { quote }, replacement };
}

test("resolves a contextual quote across formatting runs and verifies native paragraph hints", async () => {
  const { file } = await docxFile([
    { nativeId: "ABC123", segments: ["Alpha ", "beta clause", " omega"] },
    { nativeId: "DEF456", segments: ["Alpha beta clause omega"] },
  ]);
  const source = await createItemSource(file, [
    comment("c1", "ha beta cl", {
      prefix: "Alp",
      suffix: "ause",
      paragraphId: "ABC123",
    }),
  ]);

  assert.equal(source.manifest.paragraphs[0].id, "p0");
  assert.equal(source.manifest.paragraphs[0].nativeId, "ABC123");
  const snapshot = source.getSnapshot();
  assert.deepEqual(snapshot.entities, [
    {
      id: "c1",
      kind: "comment",
      body: "Comment c1",
      quote: "ha beta cl",
      status: "open",
    },
  ]);
  assert.deepEqual(snapshot.paragraphs.get("p0")?.texts, [
    {
      id: "t0",
      value: "Alpha ",
      annotations: [
        {
          entityId: "c1",
          kind: "comment",
          start: 3,
          end: 6,
          first: true,
          last: false,
        },
      ],
    },
    {
      id: "t1",
      value: "beta clause",
      annotations: [
        {
          entityId: "c1",
          kind: "comment",
          start: 0,
          end: 7,
          first: false,
          last: true,
        },
      ],
    },
    { id: "t2", value: " omega", annotations: [] },
  ]);
});

test("reports ambiguous, stale-hint, and unsupported break anchors instead of guessing", async () => {
  const { file } = await docxFile([
    { nativeId: "ONE", segments: ["one repeat two repeat"] },
    { nativeId: "BREAK", segments: ["before", { literal: "tab" }, "after"] },
    { nativeId: "LEFT", segments: ["paragraph end"] },
    { nativeId: "RIGHT", segments: ["next paragraph"] },
  ]);
  const source = await createItemSource(file, [
    comment("ambiguous", "repeat"),
    comment("stale", "one", { paragraphId: "MISSING" }),
    comment("tab", "before\tafter", { paragraphId: "BREAK" }),
    comment("cross-paragraph", "endnext", { paragraphId: "LEFT" }),
    comment("contextual", "repeat", { prefix: "one ", suffix: " two" }),
    comment("duplicate", "one"),
    comment("duplicate", "two"),
    { id: "", kind: "comment", anchor: { quote: "one" }, body: "invalid" },
    null as unknown as ReviewItem,
  ]);

  const entities = new Map(
    source.getSnapshot().entities.map((entity) => [entity.id, entity]),
  );
  assert.equal(entities.get("ambiguous")?.status, "invalid");
  assert.match(entities.get("ambiguous")?.reason ?? "", /ambiguous/i);
  assert.equal(entities.get("stale")?.status, "invalid");
  assert.match(entities.get("stale")?.reason ?? "", /paragraph/i);
  assert.equal(entities.get("tab")?.status, "invalid");
  assert.match(entities.get("tab")?.reason ?? "", /tab|break/i);
  assert.equal(entities.get("cross-paragraph")?.status, "invalid");
  assert.match(
    entities.get("cross-paragraph")?.reason ?? "",
    /cross-paragraph/i,
  );
  assert.equal(entities.get("contextual")?.status, "open");
  assert.equal(
    source.getSnapshot().entities.filter((entity) => entity.id === "duplicate")
      .length,
    2,
  );
  assert.ok(
    source
      .getSnapshot()
      .entities.filter((entity) => entity.id === "duplicate")
      .every((entity) => entity.status === "invalid"),
  );
  assert.equal(
    source
      .getSnapshot()
      .entities.filter((entity) => entity.id.startsWith("invalid:"))
      .every((entity) => entity.status === "invalid"),
    true,
  );
});

test("marks every overlapping redline as conflicting while retaining valid comments", async () => {
  const { file } = await docxFile([{ segments: ["abcdefghij"] }]);
  const source = await createItemSource(file, [
    redline("left", "cdef", "LEFT"),
    redline("right", "efgh", "RIGHT"),
    comment("note", "abc"),
  ]);

  const snapshot = source.getSnapshot();
  const entities = new Map(snapshot.entities.map((entity) => [entity.id, entity]));
  assert.equal(entities.get("left")?.status, "conflict");
  assert.equal(entities.get("right")?.status, "conflict");
  assert.match(entities.get("left")?.reason ?? "", /overlap/i);
  assert.equal(entities.get("note")?.status, "open");
  assert.deepEqual(
    snapshot.paragraphs
      .get("p0")
      ?.texts.flatMap((text) => text.annotations.map((value) => value.entityId)),
    ["note"],
  );
});

test("coalesces item updates, invalidates old and new paragraphs, and never rereads the DOCX", async () => {
  const { file, arrayBufferReads } = await docxFile([
    { segments: ["first target"] },
    { segments: ["second target"] },
    { segments: ["untouched"] },
  ]);
  const source = await createItemSource(file, [
    comment("move", "first", { paragraphId: "p0" }),
  ]);
  const untouched = source.getSnapshot().paragraphs.get("p2");
  const changes: Array<{
    revision: number;
    changed: string[];
  }> = [];
  source.subscribe(({ snapshot, changedParagraphIds }) => {
    changes.push({
      revision: snapshot.revision,
      changed: [...changedParagraphIds].sort(),
    });
  });

  source.setItems([comment("move", "second", { paragraphId: "p1" })]);
  source.setItems([comment("move", "target", { paragraphId: "p1" })]);
  await Promise.resolve();

  assert.deepEqual(changes, [{ revision: 2, changed: ["p0", "p1"] }]);
  assert.equal(source.getSnapshot().paragraphs.get("p2"), untouched);
  assert.equal(arrayBufferReads(), 1);

  source.setItems([comment("move", "target", { paragraphId: "p1" })]);
  await Promise.resolve();
  assert.equal(changes.length, 1, "semantic no-op updates must not notify");

  source.setItems([
    {
      ...comment("move", "target", { paragraphId: "p1" }),
      body: "Updated body",
    },
  ]);
  await Promise.resolve();
  assert.deepEqual(changes.at(-1), { revision: 3, changed: ["p1"] });

  source.setItems([]);
  await Promise.resolve();
  assert.deepEqual(changes.at(-1), { revision: 4, changed: ["p1"] });
  assert.equal(source.getSnapshot().entities.length, 0);
  assert.equal(arrayBufferReads(), 1);

  source.setItems([comment("added", "untouched", { paragraphId: "p2" })]);
  await Promise.resolve();
  assert.deepEqual(changes.at(-1), { revision: 5, changed: ["p2"] });
  assert.equal(arrayBufferReads(), 1);

  source.dispose();
  source.setItems([comment("ignored", "untouched")]);
  await Promise.resolve();
  assert.equal(changes.length, 4);
});
