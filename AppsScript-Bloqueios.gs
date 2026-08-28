/**
 * Backend Google Apps Script — Pesquisa Bloqueios Emocionais (40 questões)
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
  maxPdfBase64Length: 8500000,
  driveFolderName: "Relatorios_Bloqueios_Emocionais_Impulso"
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
  "Índice Geral",
  "Resultados (JSON)",
  "Arquivo PDF",
  "E-mail enviado",
  "Erro",
  "Submission ID",
  "Relatório PDF",
  "Link Relatório PDF"
];

function doGet() {
  return jsonResponse_({
    ok: true,
    service: "Mapa de Bloqueios Emocionais — Impulso",
    version: "2026-08-27-v4",
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
        clean_(payload.instrumento) || "Mapa de Bloqueios Emocionais",
        clean_(payload.nome),
        clean_(payload.email).toLowerCase(),
        clean_(payload.whatsapp),
        clean_(payload.idade),
        clean_(payload.profissao),
        scores.Pais,
        scores.Rejeicao,
        scores.Autoimagem,
        scores.Depressao,
        scores.Culpa,
        scores.Criticas,
        scores.Escassez,
        scores.Procrastinacao,
        scores.Timidez,
        scores.Vitimismo,
        Number(payload.mediaTotal || 0),
        clean_(payload.analysisJson) || JSON.stringify(scores),
        clean_(payload.pdfFileName),
        "PENDENTE",
        "",
        submissionId,
        "",
        ""
      ]);
      row = sheet.getLastRow();
      SpreadsheetApp.flush();
    }

    let reportLink = clean_(sheet.getRange(row, 26).getValue());
    let pdfBlob = null;

    if (!reportLink) {
      pdfBlob = createPdfBlob_(payload);
      reportLink = storePdfBlobAndGetLink_(pdfBlob);
      setReportLink_(sheet, row, reportLink);
      SpreadsheetApp.flush();
    }

    if (String(sheet.getRange(row, 22).getValue()).toUpperCase() === "SIM") {
      return jsonResponse_({
        ok: true,
        saved: true,
        emailSent: true,
        duplicate: true,
        reportLink: reportLink,
        submissionId: submissionId,
        row: row
      });
    }

    try {
      if (!pdfBlob) pdfBlob = createPdfBlob_(payload);
      sendClientReport_(payload, pdfBlob, reportLink);
      sheet.getRange(row, 22, 1, 2).setValues([["SIM", ""]]);
      SpreadsheetApp.flush();
      notifyAdmin_(payload, submissionId, row, reportLink);

      return jsonResponse_({
        ok: true,
        saved: true,
        emailSent: true,
        duplicate: Boolean(existingRow),
        reportLink: reportLink,
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
        reportLink: reportLink,
        submissionId: submissionId,
        row: row,
        message: "Resultado e PDF salvos, mas o e-mail falhou: " + message
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
    Autoimagem: Number(payload.autoimagem || 0),
    Depressao: Number(payload.depressao || 0),
    Culpa: Number(payload.culpa || 0),
    Criticas: Number(payload.criticas || 0),
    Escassez: Number(payload.escassez || 0),
    Procrastinacao: Number(payload.procrastinacao || 0),
    Timidez: Number(payload.timidez || 0),
    Vitimismo: Number(payload.vitimismo || 0)
  };
}

function getOrCreateReportsFolder_() {
  const folders = DriveApp.getFoldersByName(CONFIG.driveFolderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.driveFolderName);
}

function createPdfBlob_(payload) {
  const rawBase64 = clean_(payload.pdfBase64).replace(/^data:application\/pdf;base64,/, "");
  const bytes = Utilities.base64Decode(rawBase64);
  const fileName = sanitizeFilename_(payload.pdfFileName || "Mapa_Bloqueios_Emocionais.pdf");
  return Utilities.newBlob(bytes, MimeType.PDF, fileName);
}

function storePdfBlobAndGetLink_(pdfBlob) {
  const folder = getOrCreateReportsFolder_();
  const file = folder.createFile(pdfBlob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    console.log("Compartilhamento por link não habilitado pela política da conta: " + errorMessage_(sharingError));
  }
  return file.getUrl();
}

function setReportLink_(sheet, row, reportLink) {
  if (!reportLink) return;

  try {
    const rich = SpreadsheetApp.newRichTextValue()
      .setText("Abrir relatório")
      .setLinkUrl(reportLink)
      .build();
    sheet.getRange(row, 25).setRichTextValue(rich);
  } catch (error) {
    sheet.getRange(row, 25).setValue(reportLink);
  }

  sheet.getRange(row, 26).setValue(reportLink);
}

function sendClientReport_(payload, pdfBlob, reportLink) {
  const clientName = clean_(payload.nome);
  const media = Number(payload.mediaTotal || 0);
  const perfilPrimario = clean_(payload.perfilPrimario);
  const subject = "Seu relatório — Mapa de Bloqueios Emocionais";

  const plainText =
    "Olá, " + clientName + "!\n\n" +
    "Seu relatório detalhado do Mapa de Bloqueios Emocionais está anexado a este e-mail.\n" +
    "Padrão predominante: " + perfilPrimario + "\n" +
    "Índice geral: " + media + "%\n\n" +
    "Link permanente do relatório: " + reportLink + "\n\n" +
    "Instituto Impulso IE™ de Liderança\n" +
    "A mudança pode acontecer em um instante.";

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;background:#f4f9ff;padding:24px;border-radius:16px;color:#0f172a">' +
      '<div style="background:linear-gradient(135deg,#0b4aa2,#062b62);padding:24px;border-radius:14px;text-align:center">' +
        '<h1 style="color:#fff;margin:0;font-size:22px">Instituto Impulso IE™ de Liderança</h1>' +
        '<p style="color:#cfe6ff;margin:6px 0 0;font-size:13px">A mudança pode acontecer em um instante</p>' +
      '</div>' +
      '<div style="background:#fff;padding:24px;border-radius:14px;margin-top:16px">' +
        '<h2 style="color:#063b7a;margin:0 0 12px">Olá, ' + escapeHtml_(clientName) + '!</h2>' +
        '<p style="color:#334155">Seu <strong>Relatório Detalhado do Mapa de Bloqueios Emocionais</strong> está anexado.</p>' +
        '<p style="color:#334155"><strong>Padrão predominante:</strong> ' + escapeHtml_(perfilPrimario) + '<br>' +
        '<strong>Índice geral:</strong> ' + media + '%</p>' +
        '<p style="color:#334155">O relatório aprofunda seus três padrões mais fortes e apresenta crenças associadas, medos, mecanismos de proteção, compromissos ocultos, grandes pressupostos, ciclo do bloqueio, perguntas de coaching e experimentos seguros de mudança.</p>' +
        '<div style="text-align:center;margin:24px 0">' +
          '<a href="' + escapeHtml_(reportLink) + '" style="display:inline-block;background:linear-gradient(45deg,#0b63ce,#084aa0);color:#fff;padding:13px 22px;border-radius:10px;text-decoration:none;font-weight:bold">Abrir relatório PDF</a>' +
        '</div>' +
        '<p style="font-size:12px;color:#64748b">Este material é educacional e voltado ao autoconhecimento e coaching; não substitui avaliação psicológica ou diagnóstico clínico.</p>' +
      '</div>' +
    '</div>';

  GmailApp.sendEmail(
    clean_(payload.email).toLowerCase(),
    subject,
    plainText,
    {
      htmlBody: htmlBody,
      attachments: [pdfBlob],
      name: CONFIG.senderName,
      replyTo: CONFIG.adminEmail
    }
  );
}

function notifyAdmin_(payload, submissionId, row, reportLink) {
  try {
    const subject = "Novo resultado — Mapa de Bloqueios Emocionais";
    const body =
      "Participante: " + clean_(payload.nome) + "\n" +
      "E-mail: " + clean_(payload.email).toLowerCase() + "\n" +
      "WhatsApp: " + clean_(payload.whatsapp) + "\n" +
      "Padrão predominante: " + clean_(payload.perfilPrimario) + "\n" +
      "Índice Geral: " + Number(payload.mediaTotal || 0) + "%\n" +
      "Linha: " + row + "\n" +
      "Relatório PDF: " + clean_(reportLink) + "\n" +
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
