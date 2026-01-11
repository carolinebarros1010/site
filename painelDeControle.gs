/* ======================================================
 * 📊 FEMFLOW — PAINEL DE CONTROLE (READ ONLY)
 * ------------------------------------------------------
 * Funções EXCLUSIVAS de leitura e agregação
 * Usadas por dashboards administrativos
 * NÃO escreve em planilhas
 * NÃO chama IA
 * NÃO altera estado
 * ====================================================== */

/* ------------------------------
 * DASHBOARD SAC (GLOBAL)
 * ------------------------------ */
function sacDashboardResumo_() {
  const sh = _sheet("SAC_METRICS"); // ajuste se o nome for outro
  if (!sh) {
    return { status: "error", msg: "sheet_not_found" };
  }

  const data = sh.getDataRange().getValues();
  if (data.length <= 1) {
    return { status: "ok", total: 0 };
  }

  const headers = data[0];
  const rows = data.slice(1);

  const idx = {
    categoria: headers.indexOf("Categoria"),
    rota: headers.indexOf("Rota"),
    custo: headers.indexOf("CustoUSD"),
    lang: headers.indexOf("Lang")
  };

  const resumo = {
    status: "ok",
    total: rows.length,
    porCategoria: {},
    porRota: {},
    porLang: {},
    custoTotalUSD: 0
  };

  rows.forEach(r => {
    const cat = r[idx.categoria] || "desconhecida";
    const rota = r[idx.rota] || "desconhecida";
    const lang = r[idx.lang] || "pt";
    const custo = Number(r[idx.custo] || 0);

    resumo.porCategoria[cat] = (resumo.porCategoria[cat] || 0) + 1;
    resumo.porRota[rota] = (resumo.porRota[rota] || 0) + 1;
    resumo.porLang[lang] = (resumo.porLang[lang] || 0) + 1;
    resumo.custoTotalUSD += custo;
  });

  resumo.custoTotalUSD = Number(resumo.custoTotalUSD.toFixed(4));
  return resumo;
}

/* ------------------------------
 * DASHBOARD POR ALUNA
 * ------------------------------ */
function dashboardAluna_(id) {
  if (!id) return { status: "error", msg: "missing_id" };

  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "sheet_alunas_not_found" };

  const values = sh.getDataRange().getValues();
  const headers = values[0];

  let row = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(id).trim()) {
      row = values[i];
      break;
    }
  }

  if (!row) return { status: "notfound", id };

  // --- dados estruturais da aluna
  const aluna = {
    id: row[0],
    nome: row[1],
    email: row[2],
    produto: row[5],
    ativa: !!row[7],
    dataCompra: row[6] || null,

    nivel: row[8],
    enfase: row[12],
    cicloDuracao: row[9],
    fase: row[13],
    diaCiclo: row[14],
    perfilHormonal: row[19],

    diaPrograma: row[COL_DIA_PROGRAMA] || 1,
    dataInicioPrograma: row[COL_DATA_INICIO_PROGRAMA] || null,
    ultimaAtividade: row[COL_ULTIMA_ATIVIDADE] || null
  };

  // --- blocos agregados
  const sac = dashboardSACPorAluna_(id);

  return {
    status: "ok",
    aluna,
    sac
  };
}


function dashboardSACPorAluna_(id) {
  const sh = _sheet("SAC_METRICS");
  if (!sh) return { total: 0 };

  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { total: 0 };

  const headers = data[0];
  const rows = data.slice(1);

  const idx = {
    id: headers.indexOf("AlunoId"),
    categoria: headers.indexOf("Categoria"),
    rota: headers.indexOf("Rota"),
    status: headers.indexOf("Status"),
    custo: headers.indexOf("CustoUSD"),
    lang: headers.indexOf("Lang"),
    data: headers.indexOf("Data")
  };

  const out = {
    total: 0,
    porCategoria: {},
    porRota: {},
    custoTotalUSD: 0,
    ultimaRota: null,
    ultimaData: null
  };

  rows.forEach(r => {
    if (String(r[idx.id]).trim() !== String(id).trim()) return;

    out.total++;

    const cat = r[idx.categoria] || "desconhecida";
    const rota = r[idx.rota] || "desconhecida";
    const custo = Number(r[idx.custo] || 0);
    const data = r[idx.data];

    out.porCategoria[cat] = (out.porCategoria[cat] || 0) + 1;
    out.porRota[rota] = (out.porRota[rota] || 0) + 1;

    out.custoTotalUSD += custo;

    if (!out.ultimaData || (data && data > out.ultimaData)) {
      out.ultimaData = data;
      out.ultimaRota = rota;
    }
  });

  out.custoTotalUSD = Number(out.custoTotalUSD.toFixed(4));
  return out;
}


/* ------------------------------
 * DASHBOARD OPERACIONAL
 * ------------------------------ */
function dashboardAlertas_() {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "sheet_alunas_not_found" };

  const values = sh.getDataRange().getValues();
  const alerts = [];

  const hoje = new Date();

  for (let i = 1; i < values.length; i++) {
    const r = values[i];

    const id = r[0];
    const nome = r[1];
    const ativa = !!r[7];
    const ultimaAtividade = r[COL_ULTIMA_ATIVIDADE];

    let diasSemTreinar = null;
    if (ultimaAtividade instanceof Date && !isNaN(ultimaAtividade)) {
      diasSemTreinar = Math.floor((hoje - ultimaAtividade) / 86400000);
    }

    // ⚠️ ALERTA 1 — Inatividade
    if (ativa && diasSemTreinar !== null && diasSemTreinar >= 7) {
      alerts.push({
        tipo: "inatividade",
        id,
        nome,
        diasSemTreinar,
        nivel: r[8],
        fase: r[13]
      });
    }

    // ⚠️ ALERTA 2 — SAC humano recorrente
    const sac = dashboardSACPorAluna_(id);
    const humanoCount = sac.porRota?.humano || 0;
if (sac.total >= 3 && humanoCount >= 2) {

      alerts.push({
        tipo: "sac_humano_recorrente",
        id,
        nome,
        sacTotal: sac.total,
        humano: sac.porRota.humano
      });
    }

    // ⚠️ ALERTA 3 — Ativa sem atividade registrada
    if (ativa && !ultimaAtividade) {
      alerts.push({
        tipo: "sem_atividade_registrada",
        id,
        nome
      });
    }
  }

  return {
    status: "ok",
    total: alerts.length,
    alertas: alerts
  };
}


/* ------------------------------
 * HELPERS INTERNOS (READ)
 * ------------------------------ */
function _countBy_(rows, idx) {
  const out = {};
  rows.forEach(r => {
    const key = r[idx] || "desconhecido";
    out[key] = (out[key] || 0) + 1;
  });
  return out;
}

/* ======================================================
 * 📊 DASHBOARD — METRICS GERAIS (SAC + IA + ALUNAS)
 * ====================================================== */
function dashboardMetrics_() {

  const ss = SpreadsheetApp.getActive();

  const shAlunas   = ss.getSheetByName("Alunas");
  const shSacLog   = ss.getSheetByName("SAC_LOG");
  const shMetrics  = ss.getSheetByName("SAC_METRICS");

  const tz = Session.getScriptTimeZone();
  const hojeISO = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  /* ===============================
     👤 ALUNAS
  =============================== */
  let totalAlunas = 0;
  let ativas = 0;
  let inativas = 0;

  if (shAlunas) {
    const rows = shAlunas.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      totalAlunas++;
      const ativa = rows[i][7] === true; // coluna ATIVA
      ativa ? ativas++ : inativas++;
    }
  }

  /* ===============================
     🎧 SAC LOG
  =============================== */
  let sacTotal = 0;
  let sacHoje = 0;
  const rotaCount = { auto: 0, ia_low: 0, humano: 0 };

  if (shSacLog) {
    const rows = shSacLog.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      sacTotal++;

      const rota = String(rows[i][5] || "").toLowerCase();
      if (rotaCount[rota] !== undefined) rotaCount[rota]++;

      const ts = rows[i][0];
      if (ts instanceof Date) {
        const tsISO = Utilities.formatDate(ts, tz, "yyyy-MM-dd");
        if (tsISO === hojeISO) sacHoje++;
      }
    }
  }

  /* ===============================
     🤖 IA — TOKENS & CUSTO
  =============================== */
  let tokensTotal = 0;
  let custoTotal = 0;
  let custoHoje = 0;

  if (shMetrics) {
    const rows = shMetrics.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {

      const tokens = Number(rows[i][6]) || 0;
      const custo  = Number(rows[i][7]) || 0;
      const data   = rows[i][0];

      tokensTotal += tokens;
      custoTotal  += custo;

      if (data instanceof Date) {
        const dISO = Utilities.formatDate(data, tz, "yyyy-MM-dd");
        if (dISO === hojeISO) custoHoje += custo;
      }
    }
  }

  /* ===============================
     🚨 ALERTAS (OPCIONAL)
  =============================== */
  let alertasTotal = 0;
  const shAlertas = ss.getSheetByName("ALERTAS");
  if (shAlertas) {
    alertasTotal = Math.max(shAlertas.getLastRow() - 1, 0);
  }

  /* ===============================
     ✅ RETORNO FINAL
  =============================== */
  return {
    status: "ok",

    alunos: {
      total: totalAlunas,
      ativos: ativas,
      inativos: inativas
    },

    sac: {
      total: sacTotal,
      hoje: sacHoje,
      rotas: rotaCount
    },

    alertas: {
      total: alertasTotal
    },

    ia: {
      tokens_total: tokensTotal,
      custo_usd_total: Number(custoTotal.toFixed(4)),
      custo_usd_hoje: Number(custoHoje.toFixed(4))
    }
  };
}
