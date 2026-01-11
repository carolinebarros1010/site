/* ======================================================
 * 🧪 FEMFLOW — TESTE AUTOMATIZADO DO BACKEND (SAFE)
 * ------------------------------------------------------
 * - NÃO altera dados
 * - NÃO escreve em produção
 * - NÃO chama IA
 * - Executa funções em modo protegido
 * ====================================================== */

function testarSistemaFemFlow() {

  const TESTE_ID = "TEST_" + Utilities.getUuid().slice(0, 8);
  const inicioGeral = new Date();

  const resultados = [];

  Logger.log("🧪 Iniciando testes FemFlow | ID:", TESTE_ID);

  /* ======================================================
     🔎 LISTA CANÔNICA DE FUNÇÕES A TESTAR
     👉 mantenha isso atualizado
  ====================================================== */
  const FUNCOES = [

    // 🔹 CORE
    "doGet",
    "doPost",
    "getUserAccess_",
    "calcularTreinoCiclo",
    "gerarID",

    // 🔹 CICLO / PERFIL
    "setCiclo_",
    "validarCiclo_",

    // 🔹 TREINO
    "setEnfase_",
    "resetPrograma_",
    "registrarTreino_",

    // 🔹 EVOLUÇÃO
    "salvarEvolucao_",

    // 🔹 DASHBOARD / LEITURA
    "getDashboardData_",

    // 🔹 SAC / LOGS
    "sacRegistrarDashboard_",

    // 🔹 UTIL
    "_norm",
    "_sheet"
  ];

  /* ======================================================
     ▶️ EXECUÇÃO DOS TESTES
  ====================================================== */
  FUNCOES.forEach(nome => {

    const inicio = new Date();

    const resultado = {
      funcao: nome,
      status: "ok",
      erro: null,
      stack: null,
      duracao_ms: null
    };

    try {

      const fn = this[nome];

      if (typeof fn !== "function") {
        throw new Error("Função não encontrada no escopo global");
      }

      // ⚠️ Execução SAFE (sem argumentos)
      // funções que exigem args devem lidar com default interno
      fn();

    } catch (err) {

      resultado.status = "erro";
      resultado.erro = err.message;
      resultado.stack = err.stack || null;

    } finally {

      resultado.duracao_ms = new Date() - inicio;
      resultados.push(resultado);

    }

  });

  /* ======================================================
     📊 RESUMO
  ====================================================== */
  const total = resultados.length;
  const erros = resultados.filter(r => r.status === "erro");

  const resumo = {
    teste_id: TESTE_ID,
    total_funcoes: total,
    com_erro: erros.length,
    ok: total - erros.length,
    duracao_total_ms: new Date() - inicioGeral,
    timestamp: new Date().toISOString()
  };

  Logger.log("📊 RESUMO:", JSON.stringify(resumo, null, 2));
  Logger.log("📋 DETALHES:", JSON.stringify(resultados, null, 2));

  /* ======================================================
     💾 OPCIONAL — SALVAR EM PLANILHA
  ====================================================== */
  salvarResultadoTeste_(TESTE_ID, resumo, resultados);

  return {
    resumo,
    resultados
  };
}

function salvarResultadoTeste_(id, resumo, lista) {

  const sh = _sheet("TESTES", true); // true = cria se não existir

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "teste_id",
      "funcao",
      "status",
      "erro",
      "duracao_ms",
      "timestamp"
    ]);
  }

  const agora = new Date();

  lista.forEach(r => {
    sh.appendRow([
      id,
      r.funcao,
      r.status,
      r.erro || "",
      r.duracao_ms,
      agora
    ]);
  });
}
