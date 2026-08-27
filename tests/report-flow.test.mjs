import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const page = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("keeps the 40-question questionnaire and Central access control", () => {
  assert.match(page, /const totalQ = 40/);
  assert.match(page, /researchId: 'bloqueios-emocionais'/);
  assert.match(page, /ImpulsoResearchAccess\.start/);
});

test("automatically sends the generated PDF to the Bloqueios Apps Script", () => {
  assert.match(page, /AKfycbwzUpg5e7Hx2qlPqDymrGz1lTkHQJZcOcU2M9TN4dlFD8qL1aXPsa1VkdjX_JikkfFBHw/);
  assert.match(page, /testId: "bloqueios_emocionais_40_v2"/);
  assert.match(page, /pdfBase64: base64/);
  assert.match(page, /sendEmail: true/);
  assert.match(page, /await postResultToScript\(payload\)/);
});

test("does not fall back to manual mailto flow", () => {
  assert.doesNotMatch(page, /mailto:/);
  assert.match(page, /Relatório processado automaticamente/);
  assert.match(page, /doc\.save\(pdfName\)/);
});
