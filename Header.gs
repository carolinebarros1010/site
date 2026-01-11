/* Nome padrão da aba principal */
const SHEET_ALUNAS = "Alunas";

/* Token de upgrade seguro */
const SECURITY_TOKEN = "Bmc082849$";

/* HÍBRIDO → TRUE = exercícios vêm do Firebase */
const HYBRID_EXERCISES = true;

const SCRIPT_URL = ScriptApp.getService().getUrl();

/**
 * IMPORTANTE:
 * - PerfilHormonal fica no índice 19 (0-based)
 * - CicloStartDateManual fica no índice 20 (0-based)
 */
const COLUNA_PERFIL_HORMONAL = 19;
const COL_CICLO_DEFINIDO_MANUAL = 20;
const COL_FREE_ENABLED = 27; // AB
const COL_FREE_ENFASES = 28; // AC
const COL_FREE_UNTIL   = 29; // AD
const COL_ACESSO_PERSONAL   = 30; // 🔥 AJUSTE PARA AE REAL

/**
 * ✅ HEADER OFICIAL (corrigido)
 * Inclui DiaPrograma ANTES do Device/Session para não conflitar.
 */
const HEADER_ALUNAS = [
  "ID","Nome","Email","Telefone","SenhaHash","Produto","DataCompra","LicencaAtiva",
  "Nivel","CicloDuracao","DataInicio","LinkPlanilha","Enfase","Fase","DiaCiclo",
  "Pontuacao","AnamneseJSON","TokenReset","TokenExpira","PerfilHormonal",
  "CicloStartDateManual","DiaPrograma","DeviceId","SessionToken","SessionExpira",
  "DataInicioPrograma","UltimaAtividade", "FreeEnabled" , "FreeEnfases", "FreeUntil", "acesso_personal"
];

// índices (0-based) para leitura rápida
const COL_DIA_PROGRAMA   = 21; // col 22 (1-based)
const COL_DEVICE_ID      = 22; // col 23
const COL_SESSION_TOKEN  = 23; // col 24
const COL_SESSION_EXP    = 24; // col 25
const COL_DATA_INICIO_PROGRAMA = 25; // ajuste conforme posição real
const COL_ULTIMA_ATIVIDADE     = 26; // ajuste conforme posição real

