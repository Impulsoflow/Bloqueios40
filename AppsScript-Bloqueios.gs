/**
 * Backend Google Apps Script — Mapa de Bloqueios Emocionais (40 questões)
 * Padrão Impulso: salva em planilha, evita duplicidade, envia PDF e registra status.
 *
 * PLANILHA: Mapa Impulso — Resultados
 * ABA: Bloqueios Emocionais
 *
 * Para publicar:
 * 1. Cole este arquivo no projeto Apps Script do Bloqueios.
 * 2. Implantar > Gerenciar implantações > Editar.
 * 3. Executar como: você.
 * 4. Quem pode acessar: qualquer pessoa.
 * 5. Criar nova versão e implantar.
 */

const CONFIG = Object.freeze({
  spreadsheetId: "1CrQJx-M-u5F9G6RQ7X6fThwaF5OU2gcRPlkrRRp8WPQ",
  sheetName: "Bloqueios Emocionais",
  apiKey: "2026Impulso$",
  adminEmail: "impulsoflow@gmail.com",
  senderName: "Instituto Impulso Coaching de Liderança",
  maxPdfBase64Length: 8500000
});

const HEADERS = [
  "Recebido em",
  "Respondido em",
  "Instrumento",
  "Nome",
  "E-mail",
  "WhatsApp",
  "Idade",
  "Profissão",
  "Segurança e Origem Emocional",
  "Rejeição e Pertencimento",
  "Inadequação e Valor Pessoal",
  "Ameaça e Antecipação",
  "Culpa e Hiperresponsabilidade",
  "Autoexigência e Perfeccionismo",
  "Escassez e Segurança",
  "Evitação e Procrastinação",
  "Exposição e Vulnerabilidade",
  "Impotência e Locus Externo",
  "Média Total",
  "Resultados (JSON)",
  "Arquivo PDF",
  "E-mail enviado",
  "Erro",
  "Submission ID"
];

function doGet() {
  return jsonResponse_({
    ok: true,
    service: "Mapa de Bloqueios Emocionais — Impulso",
    version: "2026-08-27-v3",
    sheet: CONFIG.sheetName
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const payload = parsePayload_(e);
    validatePayload_(payload);

    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(CONFIG.sheetName) || spreadsheet.insertSheet(CONFIG.sheetName);
    ensureHeaders_(sheet);

    const submissionId = clean_(payload.submissionId) || Utilities.getUuid();
    const existingRow = findSubmissionRow_(sheet, submissionId);
    let row = existingRow;

    if (!row) {
      const scores = getScores_(payload);
      sheet.appendRow([
        new Date(),
        parseDateOrNow_(payload.respondidoEm),
        clean_(payload.instrumento) || "Bloqueios Emocionais",
        clean_(payload.nome),
        clean_(payload.email).toLowerCase(),
        clean_(payload.whatsapp),
        clean_(payload.idade),
        clean_(payload.profissao),
        scores.Pais,
        scores.Rejeicao,
        scores.Inadequação e Valor Pessoal,
        scores.Depressao,
        scores.Culpa e Hiperresponsabilidade,
        scores.Criticas,
        scores.Escassez e Segurança,
        scores.Procrastinacao,
        scores.Exposição e Vulnerabilidade,
        scores.Impotência e Locus Externo,
        Number(payload.mediaTotal || 0),
        JSON.stringify(scores),
        clean_(payload.pdfFileName),
        "PENDENTE",
        "",
        submissionId
      ]);
      row = sheet.getLastRow();
      SpreadsheetApp.flush();
    }

    if (String(sheet.getRange(row, 22).getValue()).toUpperCase() === "SIM") {
      return jsonResponse_({
        ok: true,
        saved: true,
        emailSent: true,
        duplicate: true,
        submissionId: submissionId,
        row: row
      });
    }

    try {
      sendClientReport_(payload);
      sheet.getRange(row, 22, 1, 2).setValues([["SIM", ""]]);
      SpreadsheetApp.flush();
      notifyAdmin_(payload, submissionId, row);

      return jsonResponse_({
        ok: true,
        saved: true,
        emailSent: true,
        duplicate: Boolean(existingRow),
        submissionId: submissionId,
        row: row
      });
    } catch (emailError) {
      const message = errorMessage_(emailError);
      sheet.getRange(row, 22, 1, 2).setValues([["NÃO", message]]);
      SpreadsheetApp.flush();

      return jsonResponse_({
        ok: false,
        saved: true,
        emailSent: false,
        submissionId: submissionId,
        row: row,
        message: "Resultado salvo, mas o e-mail falhou: " + message
      });
    }
  } catch (error) {
    return jsonResponse_({
      ok: false,
      saved: false,
      emailSent: false,
      message: errorMessage_(error)
    });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function parsePayload_(e) {
  if (!e) throw new Error("Requisição vazia.");

  if (e.parameter && e.parameter.payload) {
    try { return JSON.parse(e.parameter.payload); }
    catch (error) { throw new Error("Payload de formulário inválido."); }
  }

  if (!e.postData || !e.postData.contents) {
    throw new Error("Requisição sem conteúdo.");
  }

  try { return JSON.parse(e.postData.contents); }
  catch (error) { throw new Error("JSON inválido."); }
}

function validatePayload_(payload) {
  if (!payload || clean_(payload.apiKey) !== CONFIG.apiKey) {
    throw new Error("Chave de acesso inválida.");
  }
  if (!clean_(payload.nome)) throw new Error("Nome não informado.");

  const email = clean_(payload.email).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error("E-mail inválido.");
  }
  if (!clean_(payload.submissionId)) throw new Error("Submission ID não informado.");
  if (!clean_(payload.pdfBase64)) throw new Error("PDF não recebido.");
  if (clean_(payload.pdfBase64).length > CONFIG.maxPdfBase64Length) {
    throw new Error("PDF excede o tamanho máximo permitido.");
  }
  if (payload.sendEmail !== true) throw new Error("Envio de e-mail não autorizado.");
}

function ensureHeaders_(sheet) {
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }

  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsUpdate = HEADERS.some(function(header, index) {
    return String(current[index] || "").trim() !== header;
  });

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#0b63ce")
      .setFontColor("#ffffff");
  }
}

function findSubmissionRow_(sheet, submissionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const match = sheet
    .getRange(2, 24, lastRow - 1, 1)
    .createTextFinder(submissionId)
    .matchEntireCell(true)
    .findNext();

  return match ? match.getRow() : 0;
}

function getScores_(payload) {
  if (payload.scores && typeof payload.scores === "object") return payload.scores;

  if (clean_(payload.scoresJson)) {
    try {
      const parsed = JSON.parse(payload.scoresJson);
      if (parsed && typeof parsed === "object") return parsed;
    } catch (ignore) {}
  }

  return {
    Pais: Number(payload.pais || 0),
    Rejeicao: Number(payload.rejeicao || 0),
    Inadequação e Valor Pessoal: Number(payload.autoimagem || 0),
    Depressao: Number(payload.depressao || 0),
    Culpa e Hiperresponsabilidade: Number(payload.culpa || 0),
    Criticas: Number(payload.criticas || 0),
    Escassez e Segurança: Number(payload.escassez || 0),
    Procrastinacao: Number(payload.procrastinacao || 0),
    Exposição e Vulnerabilidade: Number(payload.timidez || 0),
    Impotência e Locus Externo: Number(payload.vitimismo || 0)
  };
}

function sendClientReport_(payload) {
  const rawBase64 = clean_(payload.pdfBase64).replace(/^data:application\/pdf;base64,/, "");
  const bytes = Utilities.base64Decode(rawBase64);
  const fileName = sanitizeFilename_(payload.pdfFileName || "Bloqueios_Emocionais.pdf");
  const pdfBlob = Utilities.newBlob(bytes, MimeType.PDF, fileName);

  const clientName = clean_(payload.nome);
  const media = Number(payload.mediaTotal || 0);
  const subject = "Seu resultado — Bloqueios Emocionais";
  const plainText =
    "Olá, " + clientName + "!\n\n" +
    "Seu relatório da Mapa de Bloqueios Emocionais está anexado a este e-mail.\n" +
    "Média total: " + media + "%\n\n" +
    "Instituto Impulso Coaching de Liderança";

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;color:#0f172a;line-height:1.6">' +
      '<h2 style="color:#063b7a">Olá, ' + escapeHtml_(clientName) + '!</h2>' +
      '<p>Seu relatório da <strong>Mapa de Bloqueios Emocionais</strong> está anexado.</p>' +
      '<p><strong>Média total:</strong> ' + media + '%</p>' +
      '<p>Use este material como apoio de autoconhecimento e desenvolvimento.</p>' +
      '<p style="color:#475569">Instituto Impulso Coaching de Liderança<br>' +
      'A mudança pode acontecer em um instante.</p>' +
    '</div>';

  GmailApp.sendEmail(
    clean_(payload.email).toLowerCase(),
    subject,
    plainText,
    {
      htmlBody: htmlBody,
      attachments: [pdfBlob],
      name: CONFIG.senderName
    }
  );
}

function notifyAdmin_(payload, submissionId, row) {
  try {
    const subject = "Novo resultado — Bloqueios Emocionais";
    const body =
      "Participante: " + clean_(payload.nome) + "\n" +
      "E-mail: " + clean_(payload.email).toLowerCase() + "\n" +
      "WhatsApp: " + clean_(payload.whatsapp) + "\n" +
      "Média Total: " + Number(payload.mediaTotal || 0) + "%\n" +
      "Linha: " + row + "\n" +
      "Submission ID: " + submissionId;

    GmailApp.sendEmail(CONFIG.adminEmail, subject, body, { name: CONFIG.senderName });
  } catch (ignore) {
    console.log("Aviso administrativo não enviado: " + errorMessage_(ignore));
  }
}

function parseDateOrNow_(value) {
  if (!clean_(value)) return new Date();
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? new Date() : parsed;
}

function sanitizeFilename_(value) {
  const name = clean_(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_");
  return /\.pdf$/i.test(name) ? name : (name || "Bloqueios_Emocionais") + ".pdf";
}

function clean_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}

function escapeHtml_(value) {
  return clean_(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function errorMessage_(error) {
  return error && error.message ? String(error.message) : String(error || "Erro desconhecido.");
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
