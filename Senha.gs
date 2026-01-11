/* ============================================================
 * 🔹 5) NORMALIZADOR DE STRINGS (unificado)
 * ============================================================ */
function _norm(s) {
  return (s || "")
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

/* ======================================================
 * 🔐 SESSÃO ÚNICA + DEVICE LOCK (FemFlow 2025)
 * ====================================================== */
function _generateSessionToken_() {
  return Utilities.getUuid() + "-" + Date.now();
}

function _newSession_() {
  const token = Utilities.getUuid();
  const expira = Date.now() + (1000 * 60 * 60 * 24 * 30); // 30 dias
  return { token, expira };
}

function _ensureDeviceId_(data, opts) {
  opts = opts || {};
  const d = String((data && data.deviceId) || "").trim();
  if (d) return d;
  return opts.allowGenerate ? Utilities.getUuid() : "";
}

function _getHeaderMap_(sh) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  header.forEach((h, i) => map[String(h || "").trim()] = i + 1); // 1-based
  return map;
}

function _findRowById_(sh, id) {
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) return i + 1; // 1-based row
  }
  return -1;
}

/**
 * ✅ _assertSession_ (UNIFICADO e COMPATÍVEL)
 * - Aceita (id, deviceId, sessionToken)
 * - Compat: se vier só (id, sessionToken), valida token/exp e tenta validar device se já existir
 */
function _assertSession_(id, deviceId, sessionToken) {
  // compat: (id, sessionToken)
  if (sessionToken === undefined) {
    sessionToken = deviceId;
    deviceId = "";
  }

  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { ok: false, msg: "sheet_not_found" };

  const rows = sh.getDataRange().getValues();
  const idNorm = String(id || "").trim();
  const devIn  = String(deviceId || "").trim();
  const tkIn   = String(sessionToken || "").trim();

  if (!idNorm || !tkIn) return { ok: false, msg: "missing_session" };

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== idNorm) continue;

    const devDB = String(rows[i][COL_DEVICE_ID] || "").trim();
    const tkDB  = String(rows[i][COL_SESSION_TOKEN] || "").trim();
    const exp   = Number(rows[i][COL_SESSION_EXP] || 0);

    if (!tkDB || !exp) return { ok: false, msg: "no_session" };
    if (tkDB !== tkIn) return { ok: false, msg: "token_invalid" };
    if (Date.now() > exp) return { ok: false, msg: "token_expired" };

    // ✅ Device lock REAL (anti-compartilhamento)
    let deviceUpdated = false;

    if (devDB) {
      // se já existe device vinculado, exige match
      if (!devIn) return { ok: false, msg: "device_required" };
      if (devDB !== devIn) return { ok: false, msg: "device_mismatch" };
    } else {
      // migração segura: se não tinha device salvo, salva o primeiro que vier
      if (devIn) {
        sh.getRange(i + 1, COL_DEVICE_ID + 1).setValue(devIn);
        deviceUpdated = true;
      }
    }

    return { ok: true, row: i + 1, deviceUpdated };
  }

  return { ok: false, msg: "id_not_found" };
}

/* ======================================================
 * 🔐 LOGIN FEMFLOW 2025 — acesso mensal ao app
 * ====================================================== */
function _hashSenha(senha) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    senha,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(digest);
}

function _fazerLogin(data) {
  const sh = ensureSheet(SHEET_ALUNAS, HEADER_ALUNAS);
  if (!sh) return { status: "error", msg: "Aba Alunas não encontrada." };

  const email = String(data.email || "").toLowerCase().trim();
  const senha = String(data.senha || "").trim();

  // ✅ em login, deviceId deve vir do app (não gera aleatório)
  const deviceId = _ensureDeviceId_(data, { allowGenerate: false });
  if (!deviceId) return { status: "error", msg: "device_required" };

  if (!email || !senha) {
    return { status: "error", msg: "E-mail e senha são obrigatórios." };
  }

  const valores = sh.getDataRange().getValues();
  const hashDigitada = _hashSenha(senha);

  for (let i = 1; i < valores.length; i++) {
    const row = valores[i];

    const id         = row[0];
    const nome       = row[1];
    const emailDB    = String(row[2] || "").toLowerCase().trim();
    const senhaHash  = String(row[4] || "").trim();
    const produto    = row[5];
    const dataCompra = row[6];
    const ativa      = !!row[7];
    const nivel      = row[8];
    const ciclo      = row[9];
    const inicio     = row[10];
    const enfase     = row[12];
    const fase       = row[13];
    const diaCiclo   = row[14];
    const perfilHormonal = row[19] || "regular";

    if (emailDB !== email) continue;

    if (senhaHash !== hashDigitada) {
      return { status: "error", msg: "Senha incorreta." };
    }

    // assinatura expirada
    if (dataCompra) {
      const diff = (new Date() - new Date(dataCompra)) / 86400000;
      if (diff > 30) {
        sh.getRange(i + 1, 8).setValue(false); // LicencaAtiva
        return { status: "expired", msg: "Sua assinatura expirou.", email: email, id: id };
      }
    }

    if (!ativa) {
      return { status: "inactive", msg: "Assinatura inativa.", email: email, id: id };
    }

    // 🔒 DEVICE LOCK
    const deviceDB = String(row[COL_DEVICE_ID] || "").trim();
    if (!deviceDB) {
      sh.getRange(i + 1, COL_DEVICE_ID + 1).setValue(deviceId);
    } else if (deviceDB !== deviceId) {
      return {
        status: "blocked",
        msg: "Este acesso está vinculado a outro dispositivo.",
        reason: "device_mismatch"
      };
    }

    // 🔐 Sessão única
    const sessionToken = _generateSessionToken_();
    const sessionExp = Date.now() + (1000 * 60 * 60 * 24 * 30);

    sh.getRange(i + 1, COL_SESSION_TOKEN + 1).setValue(sessionToken);
    sh.getRange(i + 1, COL_SESSION_EXP + 1).setValue(sessionExp);

    // garante DiaPrograma
    if (!row[COL_DIA_PROGRAMA]) {
      sh.getRange(i + 1, COL_DIA_PROGRAMA + 1).setValue(1);
    }

    return {
      status: "ok",
      id: id,
      nome: nome,
      email: emailDB,
      licencaAtiva: true,
      nivel: nivel,
      enfase: enfase,
      fase: fase,
      diaCiclo: diaCiclo,
      perfilHormonal: perfilHormonal,
      produto: produto,
      personal: row[COL_ACESSO_PERSONAL] === true,
      ciclo_duracao: ciclo,
      data_inicio: inicio,

      deviceId: deviceId,
      sessionToken: sessionToken,
      sessionExpira: sessionExp
    };
  }

  return { status: "error", msg: "E-mail não encontrado." };
}

function _loginOuCadastro(data) {
  const sh = ensureSheet(SHEET_ALUNAS, HEADER_ALUNAS);

  const nome      = String(data.nome || "").trim();
  const email     = String(data.email || "").toLowerCase().trim();
  const telefone  = String(data.telefone || "").trim();
  const senha     = String(data.senha || "").trim();
  const anamnese  = data.anamnese || "";

  if (!nome || !email || !senha) {
    return { status: "error", msg: "Nome, e-mail e senha são obrigatórios." };
  }

  const senhaHash = _hashSenha(senha);
  const valores = sh.getDataRange().getValues();

  let pont = _calcularPontuacaoAnamnese(anamnese);
  let nivelDetectado = "iniciante";
  if (pont >= 13 && pont < 23) nivelDetectado = "intermediaria";
  if (pont >= 23) nivelDetectado = "avancada";

  // Atualizar aluna existente
  for (let i = 1; i < valores.length; i++) {
    const row = valores[i];
    const emailDB = String(row[2] || "").toLowerCase().trim();

    if (emailDB === email) {
      const linha = i + 1;

      sh.getRange(linha, 2).setValue(nome);
      sh.getRange(linha, 3).setValue(email);
      sh.getRange(linha, 4).setValue(telefone);
      sh.getRange(linha, 5).setValue(senhaHash);

      sh.getRange(linha, 9).setValue(nivelDetectado);
      sh.getRange(linha, 16).setValue(pont);
      sh.getRange(linha, 17).setValue(anamnese);

      // Corrigir DataInicio inválida (col 11)
      const dataIni = row[10];
      if (!dataIni || !(dataIni instanceof Date) || dataIni.getFullYear() < 1990) {
        sh.getRange(linha, 11).setValue(new Date());
      }

      // garantir DiaPrograma
      if (!row[COL_DIA_PROGRAMA]) sh.getRange(linha, COL_DIA_PROGRAMA + 1).setValue(1);

      return { status: "ok", id: row[0], email, nivel: nivelDetectado, pontuacao: pont };
    }
  }

  // Novo cadastro
  const novoID = gerarID();
  const hoje = new Date();

  sh.appendRow([
    novoID,                 // ID
    nome,                   // Nome
    email,                  // Email
    telefone,               // Telefone
    senhaHash,              // SenhaHash
    "acesso_app",           // Produto
    hoje,                   // DataCompra
    true,                   // LicencaAtiva
    nivelDetectado,         // Nivel
    Number(data.cicloDuracao) || 28, // CicloDuracao
    hoje,                   // DataInicio
    "",                     // LinkPlanilha
    "nenhuma",              // Enfase
    "",           // Fase
    "",                      // DiaCiclo
    pont,                   // Pontuacao
    anamnese,               // AnamneseJSON
    "",                     // TokenReset
    "",                     // TokenExpira
    (data.perfilHormonal || "regular"), // PerfilHormonal
    "",                     // CicloStartDateManual
    1,                      // DiaPrograma
    "",                     // DeviceId
    "",                     // SessionToken
    "",                     // SessionExpira
    "",                     // data
    ""                      // ultima

  ]);

  return { status: "created", id: novoID, email, nivel: nivelDetectado, pontuacao: pont };
}

/* ======================================================
 * 🔁 RESET PROGRAMA (corrigido para não sobrescrever colunas erradas)
 * ====================================================== */
function resetPrograma_(id) {
  try {
    const sh = ensureSheet(SHEET_ALUNAS, HEADER_ALUNAS);
    const rows = sh.getDataRange().getValues();
    const idNorm = String(id || "").trim();
    if (!idNorm) return { status: "erro", msg: "ID inválido" };

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() !== idNorm) continue;

      const rowNum = i + 1;
      const agora = new Date();

      // ✅ DataInicioPrograma (nova coluna) = agora
sh.getRange(rowNum, COL_DATA_INICIO_PROGRAMA + 1).setValue(agora);

// ✅ Reset diaPrograma
sh.getRange(rowNum, COL_DIA_PROGRAMA + 1).setValue(1);

// ❌ NÃO resetar fase/diaCiclo aqui
// ❌ NÃO mexer em DataInicio aqui


      return { status: "ok", reset: true };
    }

    return { status: "erro", msg: "ID não encontrado" };
  } catch (e) {
    return { status: "erro", msg: e.toString() };
  }
}

/* ======================================================
 * 🔹 RESET DE SENHA
 * ====================================================== */
function _gerarTokenReset(id) {
  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return null;

  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) {
      const token = Utilities.getUuid();
      const expira = Date.now() + 1000 * 60 * 60; // 1h

      sh.getRange(i + 1, 18).setValue(token);
      sh.getRange(i + 1, 19).setValue(expira);

      return { token: token, expira: expira };
    }
  }

  return null;
}

function _resetSenha(data) {
  const id = String(data.id || "").trim();
  const token = String(data.token || "").trim();
  const nova = String(data.novaSenha || "").trim();

  if (!id || !token || !nova) {
    return { status: "error", msg: "Dados incompletos" };
  }

  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "Planilha não encontrada." };

  const vals = sh.getDataRange().getValues();

  let rowFound = -1;
  let tokenArmazenado = "";
  let expira = 0;

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === id) {
      rowFound = i + 1;
      tokenArmazenado = vals[i][17];  // TokenReset
      expira = Number(vals[i][18]);   // TokenExpira
      break;
    }
  }

  if (rowFound < 0) return { status: "notfound", msg: "ID não encontrado." };
  if (!tokenArmazenado || tokenArmazenado !== token) return { status: "invalid", msg: "Token inválido." };
  if (Date.now() > expira) return { status: "expired", msg: "Link expirado." };

  const hash = _hashSenha(nova);
  sh.getRange(rowFound, 5).setValue(hash);

  sh.getRange(rowFound, 18).setValue("");
  sh.getRange(rowFound, 19).setValue("");

  return { status: "ok", msg: "Senha atualizada." };
}

function _enviarEmailReset(email, id, token) {
  const link =
    "https://carolinebarros1010.github.io/myflowlife/femflow/app/reset.html" +
    "?id=" + encodeURIComponent(id) +
    "&token=" + encodeURIComponent(token) +
    "&email=" + encodeURIComponent(email);

  const html =
    "<p>Olá! 💫</p>" +
    "<p>Você solicitou redefinição de senha no FemFlow.</p>" +
    "<p>Clique abaixo para criar uma nova senha:</p>" +
    '<p><a href="' + link + '" target="_blank">🔒 Redefinir senha</a></p>' +
    "<p>Se não foi você, ignore este e-mail.</p>" +
    "<p><small>O link expira em 1 hora.</small></p>";

  MailApp.sendEmail({
    to: email,
    subject: "Redefinir senha • FemFlow",
    htmlBody: html
  });

  return { status: "sent", link: link };
}

function _solicitarResetSenha(data) {
  const email = String(data.email || "").trim().toLowerCase();
  if (!email) return { status: "error", msg: "E-mail obrigatório." };

  const sh = _sheet(SHEET_ALUNAS);
  if (!sh) return { status: "error", msg: "Planilha não encontrada." };

  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][2] || "").toLowerCase() === email) {
      const id = vals[i][0];
      const tk = _gerarTokenReset(id);
      const r = _enviarEmailReset(email, id, tk.token);

      return {
        status: "ok",
        msg: "Link enviado.",
        email: email,
        id: id,
        tokenDebug: r.link
      };
    }
  }

  return { status: "notfound", msg: "E-mail não localizado." };
}

/* ======================================================
 * 🔐 RESET DEVICE (admin)
 * ====================================================== */
function resetDevice_(data) {
  const token = String(data.token || "");
  const id = String(data.id || "").trim();

  if (token !== SECURITY_TOKEN) return { status: "unauthorized" };
  if (!id) return { status: "error", msg: "missing_id" };

  const sh = ensureSheet(SHEET_ALUNAS, HEADER_ALUNAS);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === id) {
      sh.getRange(i + 1, COL_DEVICE_ID + 1).setValue("");
      sh.getRange(i + 1, COL_SESSION_TOKEN + 1).setValue("");
      sh.getRange(i + 1, COL_SESSION_EXP + 1).setValue("");
      return { status: "ok", reset: true };
    }
  }

  return { status: "notfound" };
}
