/* ======================================================
 * 🌸 FEMFLOW — API CENTRAL (v2.8 Unificada, corrigida 2025)
 * ====================================================== */
function doPost(e) {

  // 🔒 Preflight — pode chegar sem postData
  if (!e || !e.postData) {
    console.log("⚠️ doPost sem postData");
    return _json({});
  }

  // 📦 Parse do body
  const data = parseBody_(e);

  // 🧪 DEBUG CONFIÁVEL (WebApp)
  console.log("📥 ACTION:", data.action);
  console.log("📥 perfilHormonal:", data.perfilHormonal);
  console.log("📥 perfilInterno:", data.perfilInterno);

  const action = (data.action || "").toString().toLowerCase();
  let resposta = { status: "ignored" };

  try {
    switch (action) {

      /* ===========================
         🔐 LOGIN
      ============================ */
      case "login":
        resposta = _fazerLogin(data);
        break;

      /* ===========================
         📝 CADASTRO COMPLETO
      ============================ */
      case "register":
      case "loginoucadastro":
      case "enviarcadastro":
        resposta = _loginOuCadastro(data);
        break;

      /* ===========================
         🟦 LEAD PARCIAL
      ============================ */
      case "leadparcial":
        resposta = _registrarLead(data);
        break;

      /* ===========================
         SAC
      ============================ */
      case "sac_abrir":
        resposta = sacAbrir_(data);
        break;

      /* ===========================
         🔄 RECUPERAR SENHA
      ============================ */
      case "solicitarreset":
        resposta = _solicitarResetSenha(data);
        break;

      case "resetsenha":
        resposta = _resetSenha(data);
        break;

      case "resetdevice":
        resposta = resetDevice_(data);
        break;

      /* ===========================
         🌸 SALVAR ÊNFASE
      ============================ */
      case "setenfase":
        resposta = atualizarEnfase(data.id, data.enfase);
        break;

      /* ===========================
         🌙 SALVAR FASE @deprecated
      ============================ */
      case "setfase":
        resposta = atualizarFase(data.id, data.fase);
        break;

      /* ===========================
         🧬 PERFIL HORMONAL
      ============================ */
      case "setperfilhormonal":
        resposta = setPerfilHormonal(
          data.id,
          data.perfilHormonal || data.perfil
        );
        break;

      /* ===========================
         🔢 DIA DO CICLO @deprecated
      ============================ */
      case "setdiaciclo":
        resposta = atualizarDiaCiclo(data.id, data.dia);
        break;

      /* ===========================
         🔁 REINICIAR PROGRAMA
      ============================ */
      case "resetprograma":
        resposta = resetPrograma_(data.id);
        break;

      /* ===========================
         💾 SALVAR TREINO
      ============================ */
      case "salvartreino":
        resposta = salvarTreino_(data);
        break;

      /* ===========================
         📈 SALVAR EVOLUÇÃO
      ============================ */
      case "salvarevolucao":
        resposta = salvarEvolucao_(data);
        break;

      /* ===========================
         📈 ÚLTIMO PESO
      ============================ */
      case "getultimopeso":
        resposta = getUltimoPeso_(data);
        break;

      /* ===========================
         📆 START MANUAL DO CICLO
      ============================ */
      case "setmanualstart":
        resposta = setmanualstart(data.id, data.startDate);
        break;

      /* ===========================
         📆 CICLO START (antigo)
      ============================ */
      case "setciclostart":
        resposta = atualizarCicloStart(data.id, data.startDate);
        break;

      /* ===========================
         🎚️ ALTERAR NÍVEL
      ============================ */
      case "setnivel":
        resposta = setnivel(data.id, data.nivel);
        break;

      /* ===========================
         🔄 SYNC CICLO
      ============================ */
      case "sync":
        resposta = sync(data.id || (e.parameter && e.parameter.id));
        break;

      /* ===========================
         🌸 SET CICLO (onboarding)
      ============================ */
      case "setciclo":
        resposta = setCiclo_(data);
        break;

      /* ===========================
         😴 SALVAR DESCANSO
      ============================ */
      case "salvardescanso":
        resposta = salvarDescanso_(data);
        break;

      /* ===========================
         📌 DIA DO PROGRAMA
      ============================ */
      case "setdiaprograma":
        resposta = setDiaPrograma_(data.id, data.diaPrograma);
        break;

      /* ===========================
         🛒 HOTMART / DEFAULT
      ============================ */
      default:
        resposta = _pareceHotmart_(data)
          ? _processarHotmart(data)
          : { status: "ignored", msg: "unknown_action", action };
        break;
    }

  } catch (err) {
    console.log("❌ ERRO doPost:", err);
    resposta = { status: "error", msg: err.toString() };
  }

  return _json(resposta);
}

