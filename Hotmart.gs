/* ======================================================
 * 🔹 PROCESSAR HOTMART • MAPA DE PRODUTOS
 * ====================================================== */
function mapearProduto(productName) {
  const fallback = { slug: "geral", nivel: "iniciante", fase: "follicular", enfase: "nenhuma" };
  if (!productName) return fallback;

  const prodNorm = _norm(productName);

  if (prodNorm === "treino_personal") {
  return {
    slug: "addon_personal",
    nivel: null,
    fase: null,
    enfase: null,
    acesso_personal: true
  };
}

  if (prodNorm === "femflow_premium") {
    return { slug: "premium", nivel: "iniciante", fase: "follicular", enfase: "geral" };
  }
  if (prodNorm === "acesso_app") {
    return { slug: "app", nivel: "iniciante", fase: "follicular", enfase: "nenhuma" };
  }

  const sh = _sheet("Produtos");
  if (!sh) return fallback;

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const pHot = rows[i][0];
    const slug = rows[i][1];
    const nivel = rows[i][2];
    const fase = rows[i][3];
    const enfase = rows[i][4];
    const ativo = rows[i][5];

    if (ativo !== "✅") continue;
    if (_norm(pHot) === prodNorm) {
      return {
        slug: slug || fallback.slug,
        nivel: _norm(nivel || fallback.nivel),
        fase: _norm(fase || fallback.fase),
        enfase: _norm(enfase || fallback.enfase)
      };
    }
  }

  return fallback;
}

function _processarHotmart(data) {

  /* ======================================================
     1) EVENTO
  ====================================================== */
  let eventoRaw = String(data.event || data.Event || data.event_name || "");
  if (!eventoRaw && data.data && data.data.event_name) {
    eventoRaw = String(data.data.event_name || "");
  }

  const evento = eventoRaw.toUpperCase().replace(/\s+/g, "_").trim();

  Logger.log("📬 Hotmart evento: " + evento);
  Logger.log("📦 Payload keys: " + Object.keys(data || {}).join(","));

  /* ======================================================
     2) BUYER / PRODUCT
  ====================================================== */
  let buyer = {};
  if (data.data && data.data.buyer) buyer = data.data.buyer;
  else if (data.buyer) buyer = data.buyer;

  let product = {};
  if (data.data && data.data.product) product = data.data.product;
  else if (data.product) product = data.product;

  const email = String(
    buyer.email ||
    data.email ||
    data["buyer.email"] ||
    data["buyer[email]"] ||
    ""
  ).toLowerCase().trim();

  if (!email) {
    return { status: "error", msg: "hotmart_missing_email", evento };
  }

  const nome = String(
    buyer.name ||
    data.name ||
    data["buyer.name"] ||
    data["buyer[name]"] ||
    ""
  ).trim();

  const telefone = String(
    buyer.phone_number ||
    buyer.phone ||
    data.phone ||
    data["buyer.phone_number"] ||
    data["buyer[phone_number]"] ||
    ""
  ).trim();

  let productName = String(
    product.name ||
    product.product_name ||
    product.title ||
    (data.data && data.data.product && data.data.product.name) ||
    data.product_name ||
    data["product.name"] ||
    data["product[name]"] ||
    ""
  ).trim();

  /* ======================================================
     3) NORMALIZAÇÃO DE PRODUTO
  ====================================================== */
  const prodNorm = _norm(productName);

  if (prodNorm.includes("personal")) {
    productName = "treino_personal";
  } else if (prodNorm.includes("premium")) {
    productName = "femflow_premium";
  } else if (prodNorm.includes("acesso") || prodNorm.includes("assinatura")) {
    productName = "acesso_app";
  }

  const isAddonPersonal = productName === "treino_personal";
  const isPlanoBase =
    productName === "acesso_app" || productName === "femflow_premium";

  /* ======================================================
     4) PLANILHA
  ====================================================== */
  const sh = ensureSheet(SHEET_ALUNAS, HEADER_ALUNAS);
  const values = sh.getDataRange().getValues();

  function findRowByEmail(em) {
    for (let i = 1; i < values.length; i++) {
      if (_norm(values[i][2]) === _norm(em)) return i + 1;
    }
    return -1;
  }

  /* ======================================================
     5) COMPRA APROVADA
  ====================================================== */
  if (
    evento.includes("PURCHASE_APPROVED") ||
    evento.includes("SUBSCRIPTION_APPROVED")
  ) {

    const row = findRowByEmail(email);
    let idAluno = "";

    if (row > 0) {
      idAluno = values[row - 1][0];

      if (isAddonPersonal) {
        if (typeof COL_ACESSO_PERSONAL === "number") {
          sh.getRange(row, COL_ACESSO_PERSONAL + 1).setValue(true);
        }
        sh.getRange(row, 8).setValue(true); // LicencaAtiva
      } else {
        sh.getRange(row, 6).setValue(productName);
        sh.getRange(row, 7).setValue(new Date());
        sh.getRange(row, 8).setValue(true);
      }

      if (telefone) sh.getRange(row, 4).setValue(telefone);

      const diaProg = values[row - 1][COL_DIA_PROGRAMA];
      if (!diaProg) sh.getRange(row, COL_DIA_PROGRAMA + 1).setValue(1);

    } else {
      idAluno = gerarID();

      sh.appendRow([
        idAluno,
        nome,
        email,
        telefone,
        "",
        "acesso_app",
        new Date(),
        true,
        "iniciante",
        28,
        new Date(),
        "",
        "nenhuma",
        "follicular",
        1,
        "",
        "",
        "",
        "",
        "regular",
        "",
        1,
        "",
        "",
        ""
      ]);

      const newRow = sh.getLastRow();

      if (isAddonPersonal && typeof COL_ACESSO_PERSONAL === "number") {
        sh.getRange(newRow, COL_ACESSO_PERSONAL + 1).setValue(true);
      }
    }

    return {
      status: "ok",
      produto: isAddonPersonal ? "addon_personal" : productName,
      acesso_personal: isAddonPersonal,
      evento
    };
  }

  /* ======================================================
     6) CANCELAMENTO / CHARGEBACK
  ====================================================== */
  if (
    evento.includes("SUBSCRIPTION_CANCELLATION") ||
    evento.includes("CHARGEBACK")
  ) {

    const row = findRowByEmail(email);
    if (row <= 0) {
      return { status: "notfound", email, evento };
    }

    if (isAddonPersonal) {
      if (typeof COL_ACESSO_PERSONAL === "number") {
        sh.getRange(row, COL_ACESSO_PERSONAL + 1).setValue(false);
      }
      return { status: "personal_inativo", email, evento };
    }

    if (isPlanoBase) {
      sh.getRange(row, 8).setValue(false); // LicencaAtiva

      if (typeof COL_ACESSO_PERSONAL === "number") {
        sh.getRange(row, COL_ACESSO_PERSONAL + 1).setValue(false);
      }

const COL_ACESSO_FOLLOWME = 31; // ajuste para a coluna real
      if (typeof COL_ACESSO_FOLLOWME === "number") {
        sh.getRange(row, COL_ACESSO_FOLLOWME + 1).setValue("");
      }

      return { status: "plano_base_inativo", email, evento };
    }

    return { status: "cancelamento_ignorado", email, evento };
  }

  /* ======================================================
     7) FALLBACK
  ====================================================== */
  return { status: "ignored", evento };
}

function _pareceHotmart_(data) {
  if (!data) return false;
  return !!(
    data.event || data.Event || data.event_name ||
    (data.data && (data.data.event_name || data.data.buyer || data.data.product)) ||
    data.buyer || data.product ||
    data["data.event_name"] || data["buyer[email]"] || data["buyer.email"]
  );
}
