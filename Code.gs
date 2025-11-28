/***** CONFIGURAÇÃO BÁSICA *****/
const GERAL_SHEET_NAME   = 'Geral Interconsultas';
const SAIDAS_SHEET_NAME  = 'Saídas';
const EMERG_SHEET_NAME   = 'Emergência';

/**
 * IMPORTANTE:
 * Ajuste os índices de colunas abaixo conforme sua planilha.
 * A contagem é 1 = coluna A, 2 = B, etc.
 *
 * Estes índices são um CHUTE razoável baseado nas conversas.
 * Se algo vier vazio no webapp, é só ajustar aqui.
 */

/** GERAL INTERCONSULTAS (fonte principal) */
const COL_GERAL_PRONT      = 5;  // Ex: coluna E = nº prontuário
const COL_GERAL_NOME       = 6;  // Ex: F = nome
const COL_GERAL_IDADE      = 7;  // Ex: G = idade
const COL_GERAL_CLINICA    = 8;  // Ex: H = clínica atual
const COL_GERAL_DIAG       = 9;  // Ex: I = diagnóstico
const COL_GERAL_ORIGEM     = 10; // Ex: J = origem
const COL_GERAL_MUNICIPIO  = 11; // se existir
const COL_GERAL_ADMISSAO   = 12; // data de admissão
const COL_GERAL_SOLIC      = 13; // data solicitação
const COL_GERAL_PROG       = 14; // data programação
const COL_GERAL_RESP       = 15; // data resposta

/** SAÍDAS */
const COL_SAIDAS_PRONT     = 1;  // A = nº prontuário
const COL_SAIDAS_DATA      = 13; // M = data saída
const COL_SAIDAS_DESFECHO  = 17; // Q = desfecho
const COL_SAIDAS_LOCAL     = 18; // R (se tiver)

/** EMERGÊNCIA */
const COL_EMERG_PRONT      = 1;  // A = nº prontuário
const COL_EMERG_DATA       = 15; // O = data passagem na emergência (ajuste se precisar)

/**
 * Desfechos que consideramos "finais" (paciente não é mais internado ativo)
 */
const FINAL_DESFECHOS = [
  'Residência',
  'Óbito',
  'Outro hospital',
  'Alta',
  'Alta hospitalar',
  'Transferência interna',
  'Transferência externa'
];

/***** WEBAPP *****/
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cuidados Paliativos – Pacientes Atuais')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Retorna a lista de pacientes ATIVOS (internados) para o webapp.
 * Usa apenas: Geral Interconsultas, Saídas e Emergência.
 */
function getActivePatients() {
  const ss = SpreadsheetApp.getActive();

  const geral = ss.getSheetByName(GERAL_SHEET_NAME);
  const saidas = ss.getSheetByName(SAIDAS_SHEET_NAME);
  const emerg = ss.getSheetByName(EMERG_SHEET_NAME);

  if (!geral || !saidas || !emerg) {
    throw new Error('Verifique se as abas "Geral Interconsultas", "Saídas" e "Emergência" existem.');
  }

  const geralData = getValuesWithoutEmptyTail(geral);
  const saidasData = getValuesWithoutEmptyTail(saidas);
  const emergData = getValuesWithoutEmptyTail(emerg);

  // Mapas auxiliares
  const saidasMap = buildSaidasMap(saidasData);
  const emergMap = buildEmergMap(emergData);

  const pacientes = [];

  // Assumindo que a primeira linha é cabeçalho
  for (let i = 1; i < geralData.length; i++) {
    const row = geralData[i];

    const pront = (row[COL_GERAL_PRONT - 1] || '').toString().trim();
    if (!pront) continue;

    const nome = safeString(row[COL_GERAL_NOME - 1]);
    const idade = row[COL_GERAL_IDADE - 1] || '';
    const clinica = safeString(row[COL_GERAL_CLINICA - 1]);
    const diag = safeString(row[COL_GERAL_DIAG - 1]);
    const origem = safeString(row[COL_GERAL_ORIGEM - 1]);
    const municipio = safeString(row[COL_GERAL_MUNICIPIO - 1]);

    const dataAdmissao = parseDateOrNull(row[COL_GERAL_ADMISSAO - 1]);
    const dataSolic = parseDateOrNull(row[COL_GERAL_SOLIC - 1]);
    const dataProg = parseDateOrNull(row[COL_GERAL_PROG - 1]);
    const dataResp = parseDateOrNull(row[COL_GERAL_RESP - 1]);

    // Verifica saídas desse prontuário
    const saidasPaciente = saidasMap.get(pront) || [];
    const final = getFinalDesfecho(saidasPaciente);

    // Se tiver desfecho final => não está mais internado
    if (final) {
      continue;
    }

    // Última emergência
    const dataEmerg = emergMap.get(pront) || null;

    // Cálculos de tempo
    const tempoAdmSolic = (dataAdmissao && dataSolic) ? (dataSolic.getTime() - dataAdmissao.getTime()) : null;
    const tempoSolicResp = (dataSolic && dataResp) ? (dataResp.getTime() - dataSolic.getTime()) : null;

    const tempoAdmSolicLabel = tempoAdmSolic != null ? formatDuration(tempoAdmSolic) : '';
    const tempoSolicRespLabel = tempoSolicResp != null ? formatDuration(tempoSolicResp) : '';

    // Flag atraso > 48h (para tempo Solic -> Resp)
    const flagAtraso = tempoSolicResp != null && (tempoSolicResp > 48 * 3600 * 1000);

    const paciente = {
      prontuario: pront,
      nome: nome,
      idade: idade ? idade.toString() : '',
      clinicaAtual: clinica,
      diagnostico: diag,
      origem: origem,
      municipio: municipio,

      dataAdmissao: dataAdmissao ? dataAdmissao.getTime() : null,
      dataAdmissaoLabel: formatDateForLabel(dataAdmissao),
      dataSolicitacao: dataSolic ? dataSolic.getTime() : null,
      dataSolicitacaoLabel: formatDateForLabel(dataSolic),
      dataProgramacao: dataProg ? dataProg.getTime() : null,
      dataProgramacaoLabel: formatDateForLabel(dataProg),
      dataResposta: dataResp ? dataResp.getTime() : null,
      dataRespostaLabel: formatDateForLabel(dataResp),

      dataEmergencia: dataEmerg ? dataEmerg.getTime() : null,
      dataEmergenciaLabel: formatDateForLabel(dataEmerg),

      tempoAdmSolicMs: tempoAdmSolic,
      tempoAdmSolicLabel: tempoAdmSolicLabel,
      tempoSolicRespostaMs: tempoSolicResp,
      tempoSolicRespostaLabel: tempoSolicRespLabel,

      desfechoLabel: 'ENCONTRA-SE INTERNADO',

      flagAtraso: flagAtraso
    };

    pacientes.push(paciente);
  }

  // Ordena por nome, só para ficar bonito
  pacientes.sort((a, b) => (a.nome || '').localeCompare(b.nome || ''));

  return pacientes;
}

/***** FUNÇÕES AUXILIARES *****/

/**
 * Remove "rabo" de linhas totalmente vazias do final.
 */
function getValuesWithoutEmptyTail(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();

  let lastNonEmptyRow = values.length - 1;
  for (; lastNonEmptyRow >= 0; lastNonEmptyRow--) {
    if (!rowIsEmpty(values[lastNonEmptyRow])) break;
  }

  return values.slice(0, lastNonEmptyRow + 1);
}

function rowIsEmpty(row) {
  return row.every(cell => cell === '' || cell === null);
}

function safeString(value) {
  return value == null ? '' : value.toString().trim();
}

/**
 * Constrói um mapa: prontuário -> array de saídas
 */
function buildSaidasMap(data) {
  const map = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pront = (row[COL_SAIDAS_PRONT - 1] || '').toString().trim();
    if (!pront) continue;

    const dataSaida = parseDateOrNull(row[COL_SAIDAS_DATA - 1]);
    const desfecho = safeString(row[COL_SAIDAS_DESFECHO - 1]);
    const local = safeString(row[COL_SAIDAS_LOCAL - 1]);

    const obj = {
      dataSaida,
      desfecho,
      local
    };

    if (!map.has(pront)) {
      map.set(pront, []);
    }
    map.get(pront).push(obj);
  }
  return map;
}

/**
 * Constrói um mapa: prontuário -> ÚLTIMA data de emergência
 */
function buildEmergMap(data) {
  const map = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pront = (row[COL_EMERG_PRONT - 1] || '').toString().trim();
    if (!pront) continue;

    const d = parseDateOrNull(row[COL_EMERG_DATA - 1]);
    if (!d) continue;

    const atual = map.get(pront);
    if (!atual || d.getTime() > atual.getTime()) {
      map.set(pront, d);
    }
  }
  return map;
}

/**
 * Dado um array de saídas, verifica se existe um desfecho "final".
 */
function getFinalDesfecho(saidasArr) {
  if (!saidasArr || !saidasArr.length) return null;
  for (let i = 0; i < saidasArr.length; i++) {
    const s = saidasArr[i];
    const d = (s.desfecho || '').toString().trim().toLowerCase();
    const isFinal = FINAL_DESFECHOS.some(fd => d === fd.toLowerCase());
    if (isFinal) return s;
  }
  return null;
}

/**
 * Tenta interpretar um valor como Date.
 */
function parseDateOrNull(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  // Se vier como string dd/mm/aaaa
  if (typeof value === 'string') {
    const parts = value.split(/[\/\-]/);
    if (parts.length === 3) {
      const d = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      const y = parseInt(parts[2], 10);
      const dt = new Date(y, m, d);
      if (!isNaN(dt)) return dt;
    }
  }
  return null;
}

/**
 * Formata duração (ms) em:
 * - < 1h => "X min"
 * - >= 1h e < 24h => "X h"
 * - >= 24h => "Xd Yh"
 */
function formatDuration(ms) {
  if (ms == null) return '';
  if (ms < 0) return '';

  const totalMinutes = Math.round(ms / 60000);
  if (totalMinutes < 60) {
    return totalMinutes + ' min';
  }

  const totalHours = Math.floor(ms / 3600000);
  if (totalHours < 24) {
    return totalHours + ' h';
  }

  const days = Math.floor(totalHours / 24);
  const hours = totalHours % 24;
  if (hours === 0) {
    return days + ' d';
  }
  return days + ' d ' + hours + ' h';
}

/**
 * Formata data no padrão dd/mm/aaaa.
 */
function formatDateForLabel(date) {
  if (!date) return '';
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}
