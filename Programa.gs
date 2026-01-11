/* ======================================================
 * 🔹 DiaPrograma
 * ====================================================== */
function getDiaPrograma_(id) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", diaPrograma: 1 };

  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) {
      const diaP = Number(vals[i][COL_DIA_PROGRAMA] || 1);
      return { status: "ok", diaPrograma: diaP };
    }
  }
  return { status: "notfound", diaPrograma: 1 };
}

function setDiaPrograma_(id, dia) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error" };

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(id).trim()) {
      sh.getRange(i + 1, COL_DIA_PROGRAMA + 1).setValue(Number(dia));
      return { status: "ok", diaPrograma: Number(dia) };
    }
  }
  return { status: "notfound" };
}
function avancarDiaPrograma_(shA, rowIndex, motivo) {
  if (!["treino", "descanso"].includes(motivo)) return;

  const atual = Number(shA.getRange(rowIndex + 1, COL_DIA_PROGRAMA + 1).getValue() || 1);
  shA.getRange(rowIndex + 1, COL_DIA_PROGRAMA + 1).setValue(atual + 1);
  shA.getRange(rowIndex + 1, COL_ULTIMA_ATIVIDADE + 1).setValue(new Date());
}
