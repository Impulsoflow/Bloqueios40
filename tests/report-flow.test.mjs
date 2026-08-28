import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const page = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("keeps the 40-question questionnaire and Central access control", () => {
  assert.match(page, /const TOTAL_Q = QUESTIONS\.length/);
  assert.match(page, /researchId:"bloqueios-emocionais"/);
  assert.match(page, /ImpulsoResearchAccess\.start/);
});

test("automatically sends the generated PDF to the Bloqueios Apps Script", () => {
  assert.match(page, /const WEB_APP_URL = "https:\/\/script\.google\.com\/macros\/s\/[^\"]+\/exec"/);
  assert.match(page, /testId:"bloqueios_emocionais_40_v4"/);
  assert.match(page, /pdfBase64,pdfFileName:pdfName/);
  assert.match(page, /pdfBase64,pdfFileName:pdfName,sendEmail:true/);
  assert.match(page, /await postResultToScript\(payload\)/);
});

test("does not fall back to manual mailto flow", () => {
  assert.doesNotMatch(page, /mailto:/);
  assert.match(page, /Relatório processado automaticamente/);
  assert.match(page, /doc\.save\(pdfName\)/);
});

test("only reports success after a readable backend confirmation", () => {
  assert.doesNotMatch(page, /mode:"no-cors"/);
  assert.match(page, /data=await response\.json\(\)/);
  assert.match(page, /data\.ok!==true/);
});

test("places the detailed ten-dimension map before the deep explanations", () => {
  const mapPosition = page.indexOf('// MAPA COMPLETO LOGO NO INÍCIO');
  const detailPosition = page.indexOf('// DETALHE DOS 3 PRINCIPAIS');
  assert.ok(mapPosition > 0 && detailPosition > mapPosition);
  assert.match(page, /drawDimensionSummary\(\(index\+1\)\+"\. "\+d\.label,value,d\.summary,y\)/);
  assert.match(page, /Percentuais e resumo interpretativo de cada dimensão/);
});
