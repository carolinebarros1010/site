/* ======================================================
 *  SALVAR EVOLUÇÃO — FEMFLOW (VERSÃO FINAL ESTÁVEL)
 * ------------------------------------------------------
 * Finalidade:
 * - Registrar evolução de carga (peso, reps, séries, PSE)
 * - Atualizar último peso por exercício
 *
 * Regras estruturais:
 * ✅ NÃO avança DiaPrograma
 * ✅ NÃO altera Fase ou Dia do Ciclo na planilha
 * ✅ Fase é calculada dinamicamente a partir do startDate
 *    (manual ou fisiológico), usando o motor oficial
 *
 * Fonte da verdade:
 * - Tempo (startDate / manualStart) define fase e dia
 * - Programa só avança em treino ou descanso
 * - Evolução é evento neutro no ciclo
 *
 * Segurança:
 * - Sessão validada (_assertSession_)
 * - Device lock ativo
 *
 * Compatível com:
 * - Perfil regular
 * - Perfil irregular
 * - Perfil energético / DIU / menopausa
 *
 * FemFlow Cycle Engine • 2025
 * ====================================================== */

/* ============================================================
 * 🌸 BLOCO 1 — INFRAESTRUTURA BASE
 * ============================================================ */


/* ============================================================
 * 🔹 1) PADRÃO DE RESPOSTA — JSON
 * ============================================================ */
function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 * 🔹 2) UTILITÁRIOS DE PLANILHA
 * ============================================================ */
function _sheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

/**
 * Garante que a aba existe e tem o cabeçalho correto
 * ✅ Atualiza a linha 1 (não insere nova linha)
 * ✅ Expande colunas se precisar
 */
function ensureSheet(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    return sh;
  }

  // Se existir mas estiver vazia
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    return sh;
  }

  // Garante que tem colunas suficientes
  const maxCols = sh.getMaxColumns();
  if (maxCols < header.length) {
    sh.insertColumnsAfter(maxCols, header.length - maxCols);
  }

  // Lê o header atual (linha 1) no tamanho do header novo
  const firstRow = sh.getRange(1, 1, 1, header.length).getValues()[0];

  // Se for diferente, sobrescreve a linha 1 (sem empurrar dados)
  const diff = header.some((h, i) => (firstRow[i] || "").toString().trim() !== h);

  if (diff) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }

  return sh;
}

/* ============================================================
 * 🔹 3) GERADOR DE IDs (unificado)
 * ============================================================ */
function gerarID() {
  const ts = Utilities.formatDate(new Date(), "GMT-3", "yyMMdd");
  const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return "FF-" + ts + "-" + rand;
}

/* ============================================================
 * 🔹 4) LOG INTERNO — Upgrade / Eventos
 * ============================================================ */
function _logUpgrade(entry) {
  const ss = SpreadsheetApp.getActive();
  let log = ss.getSheetByName("Logs");
  if (!log) log = ss.insertSheet("Logs");

  log.appendRow([
    new Date(),
    entry.id || "",
    entry.nivel || "",
    entry.origem || "",
    entry.status || "",
    entry.obs || ""
  ]);
}



/* ============================================================
 * 🔧 parseBody_ — aceita JSON, e.parameter, querystring
 * ============================================================ */
function parseBody_(e) {
  const raw  = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
  const type = (e && e.postData && e.postData.type) ? String(e.postData.type).toLowerCase() : "";

  // 1) tenta JSON
  try {
    if (raw && (type.includes("application/json") || raw.trim().startsWith("{"))) {
      return JSON.parse(raw);
    }
  } catch (_) {}

  // 1.1) payload=<json> (alguns serviços mandam assim)
  try {
    if (raw && raw.startsWith("payload=")) {
      const p = decodeURIComponent(raw.substring("payload=".length));
      if (p.trim().startsWith("{")) return JSON.parse(p);
    }
  } catch (_) {}

  // 2) fallback: parâmetros já parseados
  if (e && e.parameter && Object.keys(e.parameter).length) {
    return Object.assign({}, e.parameter);
  }

  // 3) fallback: querystring manual
  if (raw && raw.includes("=")) {
    const obj = {};
    raw.split("&").forEach(kv => {
      const parts = kv.split("=");
      const k = decodeURIComponent(parts[0] || "");
      const v = decodeURIComponent(parts.slice(1).join("=") || "");
      obj[k] = v;
    });
    return obj;
  }

  return {};
}






/* ======================================================
 * 🟦 CADASTRO FEMFLOW 2025 — com pontuação de anamnese
 * ====================================================== */
function _calcularPontuacaoAnamnese(anamneseJSON) {
  if (!anamneseJSON) return 0;
  try {
    const obj = JSON.parse(anamneseJSON);
    let score = 0;
    const keys = Object.keys(obj);
    for (let i = 0; i < keys.length; i++) {
      const n = Number(obj[keys[i]]);
      if (!isNaN(n)) score += n;
    }
    return score;
  } catch (e) {
    return 0;
  }
}


/* ======================================================
 * 🟦 LEAD PARCIAL
 * ====================================================== */
function _registrarLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Leads") || ss.insertSheet("Leads");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Data", "Nome", "Email", "Telefone", "Origem"]);
  }

  sheet.appendRow([
    new Date().toISOString(),
    data.nome || "",
    data.email || "",
    data.telefone || "",
    data.origem || "Anamnese FemFlow"
  ]);

  return { status: "ok", msg: "Lead parcial salvo", email: data.email };
}


/* ======================================================
 * 🔹 HISTÓRICO COMPARTILHADO
 * ====================================================== */
function _historico(id, n) {
  n = Math.max(1, Math.min(Number(n) || 30, 120));
  const out = { diario: [], treinos: [] };
  const ss = SpreadsheetApp.getActive();
  const diario = ss.getSheetByName("Diario");

  if (diario) {
    const vals = diario.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1 && out.diario.length < n; i--) {
      if (String(vals[i][0]).trim() === String(id).trim()) {
        out.diario.push({
          data: vals[i][1],
          fase: vals[i][2],
          semana: vals[i][3],
          treino: vals[i][4],
          tipo: vals[i][5],
          descanso: !!vals[i][6],
          obs: vals[i][7] || ""
        });
      }
    }
  }

  const treinos = ss.getSheetByName("Treinos");
  if (treinos) {
    const valsT = treinos.getDataRange().getValues();
    for (let j = valsT.length - 1; j >= 1 && out.treinos.length < n; j--) {
      if (String(valsT[j][0]).trim() === String(id).trim()) {
        out.treinos.push({
          data: valsT[j][1],
          fase: valsT[j][2],
          diaPrograma: valsT[j][3],
          pse: Number(valsT[j][4]) || null
        });
      }
    }
  }

  out.diario.reverse();
  out.treinos.reverse();
  return out;
}

function _normFase(s) {
  return String(s || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().trim();
}

/* ======================================================
 * 🔹 Atualizações simples
 * ====================================================== */
function atualizarEnfase(id, enfase) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "sheet_not_found" };

  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) {
      sh.getRange(i + 1, 13).setValue(String(enfase || "").toLowerCase());
      return { status: "ok", id: id, enfase: enfase };
    }
  }
  return { status: "notfound", id: id };
}

/* ======================================================
 * 🔹 ÚLTIMO PESO
 * ====================================================== */
function getUltimoPeso_(data) {
  const id = String(data.id || "").trim();
  const exercicio = String(data.exercicio || "").trim();
  const chave = exercicio.toLowerCase().trim();

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("UltimosPesos");
  if (!sh) return { status: "ok", peso: "" };

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id && rows[i][1] === chave) {
      return { status: "ok", peso: rows[i][2] || "" };
    }
  }

  return { status: "ok", peso: "" };
}

/* ======================================================
 * 🔹 SET NÍVEL
 * ====================================================== */
function setnivel(id, nivel) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "sheet_not_found" };

  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) {
      sh.getRange(i + 1, 9).setValue(String(nivel || "iniciante").toLowerCase());
      Logger.log("📈 NÍVEL atualizado → " + id + " → " + nivel);
      return { status: "ok", id: id, nivel: nivel };
    }
  }

  return { status: "notfound", id: id };
}
/* ======================================================
 * 🧾 LOG DO CICLO — FEMFLOW
 * ====================================================== */
function logCiclo_(id, evento, origem, antes, depois, obs) {
  const sh = _sheet("LogsCiclo");
  if (!sh) return;

  sh.appendRow([
    new Date(),               // Data
    String(id || ""),         // ID
    String(evento || ""),     // Evento
    String(origem || ""),     // Origem (setciclo / sync / validar / manual)
    antes !== undefined ? JSON.stringify(antes) : "",
    depois !== undefined ? JSON.stringify(depois) : "",
    String(obs || "")
  ]);
}
