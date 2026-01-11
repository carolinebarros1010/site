/* ============================================================
 * 🌙 STARGATE — CÁLCULO REAL DO CICLO FEMFLOW 2025
 * ============================================================ */
function calcularCicloReal({
  startDate,
  cicloDuracao,
  perfilHormonal,
  nivel,
  faseSalva,
  diaCicloSalvo
}) {
  const hoje = new Date();

  // 🛑 1. Corrigir datas ruins
  let inicio = new Date(startDate);
  if (isNaN(inicio.getTime()) || inicio.getFullYear() < 1990) {
    inicio = new Date();
  }

  const diff = Math.floor((hoje - inicio) / 86400000);

  const length = Number(cicloDuracao) > 0 ? Number(cicloDuracao) : 28;

  const pos = ((diff % length) + length) % length;
  const dia = pos + 1;

  const perfil = perfilHormonal
  ? String(perfilHormonal).toLowerCase()
  : null;

  if (!perfil) {
  // Não inventar perfil
  return {
    fase: faseSalva || "menstrual",
    dia: diaCicloSalvo || 1
  };
}


  const nivelNorm = (nivel || "iniciante").toLowerCase();

 // ✅ PERFIS ENERGÉTICO / MENOPAUSA / DIU HORMONAL
// Agora seguem o MESMO ciclo fisiológico de 28 dias (treinos no Firebase = 1..28).
if (perfil === "energetico" || perfil === "menopausa" || perfil === "diu_hormonal") {
  const d = ((diff % cicloDuracao) + cicloDuracao) % cicloDuracao + 1;

  if (d <= 5)  return { fase: "menstrual", dia: d };
  if (d <= 13) return { fase: "follicular", dia: d };
  if (d <= 17) return { fase: "ovulatory", dia: d };
  return { fase: "luteal", dia: Math.max(18, d) };
}


/**
 * PERFIL IRREGULAR — REGRA FEMFLOW
 * --------------------------------
 * DiaCiclo NÃO é cronológico.
 * É ancorado no início fisiológico da fase:
 * - Menstrual   → dia 1
 * - follicular   → dia 6
 * - Ovulatória  → dia 14
 * - Lútea       → dia 18
 *
 * O tempo (diff) serve apenas para identificar a fase.
 */

if (perfil === "irregular") {

  const d = ((diff % 28) + 28) % 28 + 1;

  if (d <= 5) {
    return { fase: "menstrual", dia: d };
  }

  if (d <= 13) {
    return { fase: "follicular", dia: d };
  }

  if (d <= 17) {
    return { fase: "ovulatory", dia: d };
  }

  // 🔒 GARANTIA: lútea nunca abaixo de 18
  return {
    fase: "luteal",
    dia: Math.max(18, d)
  };
}



return { fase: faseSalva || "follicular", dia: diaCicloSalvo || 1 };

}

/* ======================================================
 * 🌸 SET CICLO — OPÇÃO A (STARTDATE RETROATIVO)
 * ------------------------------------------------------
 * Objetivo:
 * - Recebe diaCicloInicial (1..28)
 * - Calcula DataInicio real de forma retroativa
 * - Atualiza IMEDIATAMENTE:
 *   • DataInicio
 *   • Fase (N)
 *   • DiaCiclo (O)
 *
 * Decisões:
 * ✅ PerfilHormonal sempre "regular"
 * ✅ ManualStart SEMPRE limpo
 * ✅ VALIDAR passa a ser corretivo, não primário
 * ====================================================== */
function setCiclo_(data) {
  
    const sh = ensureSheet(SHEET_ALUNAS, HEADER_ALUNAS);
  if (!sh) return { status: "error", msg: "sheet_not_found" };

  const id = String(data.id || "").trim();
  if (!id) return { status: "error", msg: "missing_id" };

   // ✅ 1) SALVAR PERFIL HORMONAL (se veio do front)
  const perfil = (data.perfilHormonal || data.perfil || "").toString().toLowerCase().trim();
  if (perfil) {
    setPerfilHormonal(id, perfil); // escreve coluna T (19)
  }

  const values = sh.getDataRange().getValues();

  /* ===============================
     Helpers locais
  =============================== */
  const _today0 = () => {
    const t = new Date();
    t.setHours(0, 0, 0, 0);
    return t;
  };

  const _toDateSafe = (d) => {
    const dt = new Date(d);
    if (!(dt instanceof Date) || isNaN(dt.getTime()) || dt.getFullYear() < 1990) return null;
    dt.setHours(0, 0, 0, 0);
    return dt;
  };

  const _clamp = (n, min, max) => Math.max(min, Math.min(max, n));

  const fasePorDia = (dia) => {
    if (dia <= 5)  return "menstrual";
    if (dia <= 13) return "follicular";
    if (dia <= 17) return "ovulatory";
    return "luteal";
  };

  /* ===============================
     Loop de busca
  =============================== */
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[0]).trim() !== id) continue;

    const linha = i + 1;

    /* ===============================
       1) CicloDuracao (J)
    =============================== */
    const cicloDuracao = _clamp(
      Number(data.cicloDuracao) || Number(r[9]) || 28,
      21,
      35
    );
    sh.getRange(linha, 10).setValue(cicloDuracao);

    /* ===============================
       2) Dia do ciclo (intenção explícita)
    =============================== */
    const diaRaw = data.diaCicloInicial ?? data.diaCiclo ?? null;
    const diaCicloFinal = _clamp(Number(diaRaw) || 1, 1, cicloDuracao);

    /* ===============================
       3) DataInicio retroativa (K)
    =============================== */
    let dataInicioFinal = null;

    if (diaRaw !== null && diaRaw !== undefined && String(diaRaw).trim() !== "") {
      const base = _today0();
      base.setDate(base.getDate() - (diaCicloFinal - 1));
      dataInicioFinal = base;
    } else if (data.dataInicio) {
      dataInicioFinal = _toDateSafe(data.dataInicio);
    }

    if (dataInicioFinal) {
      sh.getRange(linha, 11).setValue(dataInicioFinal);
    }

    /* ===============================
       4) Fase fisiológica (N)
    =============================== */
    const faseFinal = fasePorDia(diaCicloFinal);
    sh.getRange(linha, 14).setValue(faseFinal);

    /* ===============================
       5) DiaCiclo (O)
    =============================== */
    sh.getRange(linha, 15).setValue(diaCicloFinal);

        /* ===============================
       7) Limpar ManualStart (U)
    =============================== */
    sh.getRange(linha, 21).clearContent();

    /* ===============================
       8) DiaPrograma (V)
    =============================== */
    sh.getRange(linha, 22).setValue(Number(data.diaPrograma) || 1);

    /* ===============================
       9) DataInicioPrograma
    =============================== */
    if (!r[COL_DATA_INICIO_PROGRAMA]) {
      sh.getRange(linha, COL_DATA_INICIO_PROGRAMA + 1).setValue(new Date());
    }

    /* ===============================
       Retorno
    =============================== */
    return {
      status: "ok",
      id,
      cicloDuracao,
      dataInicio: dataInicioFinal ? dataInicioFinal.toISOString() : null,
      fase: faseFinal,
      diaCiclo: diaCicloFinal,
      perfilHormonal: perfil || r[19] || null,

      manualCleared: true
    };
  }

  return { status: "notfound" };
}


/* ======================================================
 * 🔹 Motor de Treino HÍBRIDO (resumo — usado pelo front)
 * ====================================================== */
function _resolverPerfil(id) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return null;

  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[0]).trim() === String(id).trim()) {
      return {
        id: r[0],
        nome: r[1],
        email: r[2],
        produto: r[5] || "",
        ativo: !!r[7],
        nivel: String(r[8] || "iniciante").toLowerCase(),
        cicloDuracao: Number(r[9] || 28),
        dataInicio: r[10] || new Date(),
        link_planilha: r[11] || "",
        enfase: _norm(r[12] || "nenhuma"),
        fase: _norm(r[13] || "follicular"),
        diaCiclo: Number(r[14] || 1)
      };
    }
  }
  return null;
}
function resolverDiaTreino({ perfilHormonal, nivel, diaCiclo, diaPrograma }) {

  // PERFIL ENERGÉTICO
  if (perfilHormonal === "energetico") {

    const faseAlta = {
      iniciante: "luteal",
      intermediaria: "follicular",
      avancada: "ovulatory"
    };

    return {
      fase: faseAlta[nivel] || "follicular",
      diaTreino: ((diaPrograma - 1) % 7) + 1,
      fonte: "programa"
    };
  }

  // PERFIS BIOLÓGICOS
  return {
    fase: fasePorDiaCiclo(diaCiclo),
    diaTreino: diaCiclo,
    fonte: "ciclo"
  };
}

/* ======================================================
 * 🧬 Perfil Hormonal — (estava faltando)
 * Coluna 20 (T) = índice 19 no array
 * ====================================================== */
function setPerfilHormonal(id, perfil) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "sheet_not_found" };

  const idNorm = String(id || "").trim();
  const p = String(perfil || "").toLowerCase().trim();

  if (!idNorm) return { status: "error", msg: "missing_id" };
  if (!p) return { status: "error", msg: "missing_perfil" };

  // ✅ PRIMEIRO declarar vals
  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== idNorm) continue;

    const perfilAtual = String(vals[i][19] || "").toLowerCase();

    // 🛡️ BLINDAGEM ANTI-RESET
    if (perfilAtual && perfilAtual !== "regular" && p === "regular") {
      return {
        status: "ignored",
        motivo: "anti_reset_regressivo",
        perfilMantido: perfilAtual
      };
    }

    // ✅ grava perfil corretamente
    sh.getRange(i + 1, 20).setValue(p);

    // regra: se virou regular, limpa manual start
    if (p === "regular") {
      sh.getRange(i + 1, 21).clearContent();
    }

    return { status: "ok", id: idNorm, perfilHormonal: p };
  }

  return { status: "notfound", id: idNorm };
}


function fasePorDiaCiclo_(dia) {
  const d = Number(dia) || 1;
   if (dia <= 5)  return "menstrual";
  if (dia <= 13) return "follicular";
  if (dia <= 17) return "ovulatory";
  return "luteal";
}

/* ======================================================
 * 🔄 FASE ATUAL — FONTE ÚNICA
 * Coluna N (Fase) = índice 13
 * ====================================================== */
function calcularEFixarFase_(id) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return null;

  const idNorm = String(id || "").trim();
  if (!idNorm) return null;

  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== idNorm) continue;

    /* ============================
       🔎 ESTADO HORMONAL
    ============================ */
    const perfilHormonal = String(vals[i][19] || "").toLowerCase(); // T
    const diaCiclo       = Number(vals[i][14] || 1);               // O
    const cicloDuracao   = Number(vals[i][9]  || 28);              // J
    const manualStart    = vals[i][20];                            // U
    const nivel          = String(vals[i][8]  || "").toLowerCase();// I

    let fase = "menstrual"; // fallback seguro

    /* ============================
       🧬 PERFIL REGULAR / DIU
    ============================ */
    if (perfilHormonal === "regular" || perfilHormonal === "diu") {
      const d = Math.max(1, Math.min(diaCiclo, cicloDuracao));

      if (d <= 5) fase = "menstrual";
      else if (d <= 13) fase = "follicular";
      else if (d <= 17) fase = "ovulatory";
      else fase = "luteal";
    }

    /* ============================
       🔀 PERFIL IRREGULAR
    ============================ */
    else if (perfilHormonal === "irregular") {
      // aqui a fase já foi induzida no onboarding
      if (diaCiclo <= 5) fase = "menstrual";
      else if (diaCiclo <= 13) fase = "follicular";
      else if (diaCiclo <= 17) fase = "ovulatory";
      else fase = "luteal";
    }

    /* ============================
       🔋 PERFIL ENERGÉTICO / MENOPAUSA
    ============================ */
  // regra única por dia para TODOS os perfis
if (diaCiclo <= 5) fase = "menstrual";
else if (diaCiclo <= 13) fase = "follicular";
else if (diaCiclo <= 17) fase = "ovulatory";
else fase = "luteal";


    /* ============================
       ✋ MANUAL START (override)
    ============================ */
    if (manualStart instanceof Date) {
      fase = vals[i][13] || fase; // mantém fase manual já definida
    }

    /* ============================
       ✍️ ESCREVER FASE
    ============================ */
    sh.getRange(i + 1, 14).setValue(fase); // coluna N

    return fase;
  }

  return null;
}
