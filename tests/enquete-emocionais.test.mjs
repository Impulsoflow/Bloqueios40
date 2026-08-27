import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const html = fs.readFileSync(new URL("../enquete-emocionais/index.html", import.meta.url), "utf8");

test("Enquete Emocionais é independente e contém as 40 perguntas", () => {
  assert.match(html, /<title>Pesquisa Enquete Emocionais<\/title>/);
  const match = html.match(/const QUESTIONS=(\[[\s\S]*?\]);\n\s*const TOTAL_Q=/);
  assert.ok(match, "QUESTIONS não encontrado");
  const questions = JSON.parse(match[1]);
  assert.equal(questions.length, 40);
  assert.equal(new Set(questions.map(q => q.id)).size, 40);
});

test("preserva 10 áreas com 4 perguntas cada", () => {
  const match = html.match(/const QUESTIONS=(\[[\s\S]*?\]);\n\s*const TOTAL_Q=/);
  const questions = JSON.parse(match[1]);
  const counts = {};
  questions.forEach(q => counts[q.category] = (counts[q.category] || 0) + 1);
  assert.equal(Object.keys(counts).length, 10);
  Object.values(counts).forEach(count => assert.equal(count, 4));
});

test("mistura perguntas e alternativas sem mudar os valores", () => {
  assert.match(html, /order=shuffleCopy\(QUESTIONS\.map\(q=>q\.id\)\)/);
  assert.match(html, /optionOrders\[q\.id\]=shuffleCopy\(q\.options\)/);
  assert.match(html, /scores\[q\.category\]\+=Number\(answers\[q\.id\]\|\|0\)/);
});

test("usa layout responsivo de respostas inspirado no DISC", () => {
  assert.match(html, /\.opt\{[^}]*flex:0 1 calc\(20% - 10px\)/);
  assert.match(html, /@media\(max-width:480px\)[\s\S]*\.opt\{[^}]*calc\(50% - 4px\)/);
  assert.match(html, /position:fixed;top:0;left:0;right:0/);
});

test("não usa o backend quebrado do Bloqueios40", () => {
  assert.doesNotMatch(html, /AKfycbwzUpg5e7Hx2qlPqDymrGz1lTkHQJZcOcU2M9TN4dlFD8qL1aXPsa1VkdjX_JikkfFBHw/);
  assert.doesNotMatch(html, /envio automático não foi confirmado/i);
});
