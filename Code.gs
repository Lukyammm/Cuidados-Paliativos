/*******************************************************************************************
 *  CP · HUC – WEBAPP BACKEND (versão Apple Health)
 *  Compatível com HTML fornecido por Luky
 *  Funções expostas ao front-end:
 *     - searchPacientePorProntuario
 *     - adicionarOuAtualizarAcompanhamento
 *     - listarAcompanhamentosAtivos
 *     - moverParaSerieHistorica
 *******************************************************************************************/


/**************************************
 * BASES OFICIAIS
 **************************************/
const ABA_GERAL   = 'Geral Interconsultas';
const ABA_SAIDAS  = 'Saídas';
const ABA_EMERG   = 'Emergência';

const ABA_ATIVOS  = 'Acompanhamento_Ativo';
const ABA_HIST    = 'Acompanhamento_Historico';
const ABA_LOGS    = 'CP_Logs';

/**************************************
 * DESFECHOS DEFINITIVOS
 **************************************/
const DESFECHOS_FINAIS = [
  "Óbito",
  "Obito",
  "Alta",
  "Alta Hospitalar",
  "Outro hospital",
  "Residência",
  "Residencia",
  "Transferência interna",
  "Transferência externa"
];


function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("CP · Painel de Acompanhamento");
}


/*******************************************************************************************
 * GARANTIR ABAS
 *******************************************************************************************/
function garantirAbas() {
  const ss = SpreadsheetApp.getActive();

  if (!ss.getSheetByName(ABA_ATIVOS)) {
    const sh = ss.insertSheet(ABA_ATIVOS);
    sh.getRange("A1:L1").setValues([[
      "Prontuario",
      "Nome",
      "Idade",
      "Clinica",
      "Origem",
      "Municipio",
      "Data_Insercao",
      "Inserido_Por",
      "Status_CP",
      "Desfecho_Detectado",
      "Data_Desfecho",
      "Local_Desfecho"
    ]]);
  }

  if (!ss.getSheetByName(ABA_HIST)) {
    const sh = ss.insertSheet(ABA_HIST);
    sh.getRange("A1:L1").setValues([[
      "Prontuario",
      "Nome",
      "Idade",
      "Clinica",
      "Origem",
      "Municipio",
      "Data_Insercao",
      "Data_Encerramento",
      "Desfecho",
      "Local_Desfecho",
      "Movido_Por",
      "Movido_Em"
    ]]);
  }

  if (!ss.getSheetByName(ABA_LOGS)) {
    const sh = ss.insertSheet(ABA_LOGS);
    sh.getRange("A1:E1").setValues([[
      "Timestamp",
      "Usuario",
      "Acao",
      "Prontuario",
      "Detalhes"
    ]]);
  }
}


/*******************************************************************************************
 * LOG
 *******************************************************************************************/
function logAcao(acao, pront, detalhes) {
  const ss = SpreadsheetApp.getActive();
  garantirAbas();
  const sh = ss.getSheetByName(ABA_LOGS);
  sh.appendRow([
    new Date(),
    Session.getActiveUser().getEmail() || "desconhecido",
    acao,
    pront,
    detalhes || ""
  ]);
}


/*******************************************************************************************
 * FUNÇÃO 1 — Buscar paciente por prontuário
 *******************************************************************************************/
function searchPacientePorProntuario(pront) {
  garantirAbas();

  const ss = SpreadsheetApp.getActive();
  const shG = ss.getSheetByName(ABA_GERAL);
  const shS = ss.getSheetByName(ABA_SAIDAS);
  const shE = ss.getSheetByName(ABA_EMERG);
  const shA = ss.getSheetByName(ABA_ATIVOS);

  const prontStr = String(pront).trim();

  let geral = shG.getDataRange().getValues();
  let saidas = shS.getDataRange().getValues();
  let emerg = shE.getDataRange().getValues();
  let ativos = shA.getDataRange().getValues();

  let encontradoGeral = null;
  let ultimaSaida = null;
  let ultimaEmerg = null;

  for (let i = 1; i < geral.length; i++) {
    if (String(geral[i][4]).trim() == prontStr) {
      encontradoGeral = geral[i];
      break;
    }
  }

  for (let i = 1; i < saidas.length; i++) {
    if (String(saidas[i][0]).trim() == prontStr) {
      const d = parseDate(saidas[i][13]);
      if (!ultimaSaida || d > ultimaSaida.data) {
        ultimaSaida = {
          data: d,
          label: formatDate(d),
          desfecho: saidas[i][16] || "",
          local: saidas[i][14] || ""
        };
      }
    }
  }

  for (let i = 1; i < emerg.length; i++) {
    if (String(emerg[i][0]).trim() == prontStr) {
      const d = parseDate(emerg[i][14]);
      if (!ultimaEmerg || d > ultimaEmerg) ultimaEmerg = d;
    }
  }

  const jaAcompanhado = ativos.some(r => String(r[0]).trim() == prontStr);

  if (!encontradoGeral && !ultimaSaida && !ultimaEmerg) {
    return { encontrado: false };
  }

  const nome = encontradoGeral ? encontradoGeral[5] : "";
  const dataNasc = encontradoGeral ? parseDate(encontradoGeral[6]) : null;
  const idade = dataNasc ? calcIdade(dataNasc) : "";

  const clinica = encontradoGeral ? encontradoGeral[3] : "";
  const origem = encontradoGeral ? encontradoGeral[2] : "";
  const dataSolic = encontradoGeral ? parseDate(encontradoGeral[0]) : null;
  const dataProg = encontradoGeral ? parseDate(encontradoGeral[9]) : null;
  const dataResp = encontradoGeral ? parseDate(encontradoGeral[10]) : null;
  const sexo = "";  
  const municipio = "";

  return {
    encontrado: true,
    prontuario: prontStr,
    nome,
    idade,
    sexo,
    municipio,
    clinica,
    origem,
    especialidade: encontradoGeral ? encontradoGeral[8] : "",

    dataSolicitacaoLabel: formatDate(dataSolic),
    dataProgramacaoLabel: formatDate(dataProg),
    dataRespostaLabel: formatDate(dataResp),

    dataEmergenciaLabel: ultimaEmerg ? formatDate(ultimaEmerg) : "–",

    dataSaidaLabel: ultimaSaida ? ultimaSaida.label : "–",
    desfechoSaida: ultimaSaida ? ultimaSaida.desfecho : "",

    jaAcompanhado
  };
}


/*******************************************************************************************
 * FUNÇÃO 2 — Adicionar / atualizar acompanhamento
 *******************************************************************************************/
function adicionarOuAtualizarAcompanhamento(pront) {
  garantirAbas();

  const ss = SpreadsheetApp.getActive();
  const shA = ss.getSheetByName(ABA_ATIVOS);
  let dados = shA.getDataRange().getValues();

  const prontStr = String(pront).trim();

  const linhasEncontradas = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim() == prontStr) {
      linhasEncontradas.push(i + 1); // índice real na planilha
    }
  }

  if (linhasEncontradas.length) {
    // Mantém o primeiro registro e elimina duplicatas
    shA.getRange(linhasEncontradas[0], 9).setValue("Ativo");
    shA.getRange(linhasEncontradas[0], 12).setValue(new Date());

    if (linhasEncontradas.length > 1) {
      const duplicates = linhasEncontradas.slice(1).sort((a, b) => b - a);
      duplicates.forEach(idx => shA.deleteRow(idx));
      logAcao(
        "Limpar duplicados",
        prontStr,
        `Removidas ${duplicates.length} duplicatas ao atualizar.`
      );
    }

    logAcao("Atualizou", prontStr, "Atualização do acompanhamento");
    return linhasEncontradas.length > 1
      ? "Acompanhamento atualizado e duplicatas removidas."
      : "Acompanhamento atualizado.";
  }

  const info = searchPacientePorProntuario(prontStr);
  if (!info || !info.encontrado) return "Paciente não encontrado.";

  shA.appendRow([
    info.prontuario,
    info.nome,
    info.idade,
    info.clinica,
    info.origem,
    info.municipio,
    new Date(),
    Session.getActiveUser().getEmail() || "desconhecido",
    "Ativo",
    "",
    "",
    "",
  ]);

  logAcao("Novo", prontStr, "Inserido no acompanhamento");
  return "Paciente adicionado ao acompanhamento.";
}


/*******************************************************************************************
 * FUNÇÃO 3 — Listar acompanhamentos ativos
 *******************************************************************************************/
function listarAcompanhamentosAtivos() {
  garantirAbas();

  const ss = SpreadsheetApp.getActive();

  const geral = ss.getSheetByName(ABA_GERAL).getDataRange().getValues();
  const saidas = ss.getSheetByName(ABA_SAIDAS).getDataRange().getValues();
  const emerg = ss.getSheetByName(ABA_EMERG).getDataRange().getValues();
  const atvs = ss.getSheetByName(ABA_ATIVOS).getDataRange().getValues();

  const contagemPorPront = {};
  for (let i = 1; i < atvs.length; i++) {
    const prontTmp = String(atvs[i][0]).trim();
    if (!prontTmp) continue;
    if (!contagemPorPront[prontTmp]) contagemPorPront[prontTmp] = [];
    contagemPorPront[prontTmp].push(i + 1); // linha real na planilha
  }

  let lista = [];

  for (let i = 1; i < atvs.length; i++) {
    const row = atvs[i];
    const pront = String(row[0]).trim();
    if (!pront) continue;

    const linhaPlanilha = i + 1;

    const geralInfo = buscarNaGeral(geral, pront);
    const saidaInfo = buscarSaida(saidas, pront);
    const emergInfo = buscarEmerg(emerg, pront);

    const idade = geralInfo.dataNasc ? calcIdade(geralInfo.dataNasc) : row[2];

    let precisaMover = false;
    let desfechoLabel = "";
    let dataDesfechoLabel = "";

    if (saidaInfo && saidaInfo.desfecho) {
      if (DESFECHOS_FINAIS.some(d => d.toLowerCase() == saidaInfo.desfecho.toLowerCase())) {
        precisaMover = true;
        desfechoLabel = saidaInfo.desfecho;
        dataDesfechoLabel = formatDate(saidaInfo.dataSaida);
      }
    }

    const tempoSolicResp = gerarDiferenca(geralInfo.dataSolic, geralInfo.dataResp);
    const statusCp = row[8] || "Ativo";
    const dataInsercao = row[6] ? parseDate(row[6]) || row[6] : null;

    lista.push({
      prontuario: pront,
      nome: geralInfo.nome || row[1],
      idade,
      clinica: geralInfo.clinica,
      origem: geralInfo.origem,
      municipio: geralInfo.municipio,

      dataSolicitacaoLabel: formatDate(geralInfo.dataSolic),
      dataProgramacaoLabel: formatDate(geralInfo.dataProg),
      dataRespostaLabel: formatDate(geralInfo.dataResp),

      dataInsercaoLabel: formatDate(dataInsercao),
      inseridoPor: row[7] || "",
      statusCp,

      dataEmergenciaLabel: emergInfo ? formatDate(emergInfo) : "–",

      tempoSolicRespLabel: formatTempo(tempoSolicResp),
      tempoSolicRespMs: tempoSolicResp,
      flagAtraso: tempoSolicResp ? tempoSolicResp > 1000 * 60 * 60 * 48 : false,

      precisaMoverSerie: precisaMover,
      desfechoDetectadoLabel: desfechoLabel,
      dataDesfechoLabel: dataDesfechoLabel,

      rowIndex: linhaPlanilha,
      duplicado: (contagemPorPront[pront] || []).length > 1,
      linhasDuplicadas: contagemPorPront[pront] || []
    });
  }

  return lista;
}


/*******************************************************************************************
 * FUNÇÃO 4 — Mover para série histórica
 *******************************************************************************************/
function moverParaSerieHistorica(pront) {
  garantirAbas();

  const ss = SpreadsheetApp.getActive();
  const shA = ss.getSheetByName(ABA_ATIVOS);
  const shH = ss.getSheetByName(ABA_HIST);

  let dados = shA.getDataRange().getValues();
  let alvo = null;
  let idx = -1;

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim() == pront) {
      alvo = dados[i];
      idx = i;
      break;
    }
  }

  if (!alvo) return "Paciente não encontrado no acompanhamento.";

  shH.appendRow([
    alvo[0], // pront
    alvo[1], // nome
    alvo[2], // idade
    alvo[3], // clinica
    alvo[4], // origem
    alvo[5], // municipio
    alvo[6], // data insercao
    new Date(),
    alvo[9], // desfecho
    alvo[11], // local
    Session.getActiveUser().getEmail() || "desconhecido",
    new Date()
  ]);

  shA.deleteRow(idx + 1);

  logAcao("Mover", pront, "Movido para série histórica");
  return "Movido para a série histórica.";
}


/*******************************************************************************************
 * FUNÇÃO 5 — Excluir acompanhamento (apoio a duplicados)
 *******************************************************************************************/
function excluirAcompanhamentoPorLinha(rowIndex) {
  garantirAbas();

  const idx = parseInt(rowIndex, 10);
  if (!idx || idx < 2) return "Linha inválida para exclusão.";

  const ss = SpreadsheetApp.getActive();
  const shA = ss.getSheetByName(ABA_ATIVOS);
  const lastRow = shA.getLastRow();

  if (idx > lastRow) return "Registro não encontrado.";

  const row = shA.getRange(idx, 1, 1, shA.getLastColumn()).getValues()[0];
  const pront = String(row[0]).trim() || "desconhecido";

  shA.deleteRow(idx);
  logAcao("Excluir", pront, `Registro removido da linha ${idx}.`);
  return "Registro excluído com sucesso.";
}


/*******************************************************************************************
 * FUNÇÕES AUXILIARES
 *******************************************************************************************/
function buscarNaGeral(geral, pront) {
  for (let i = 1; i < geral.length; i++) {
    if (String(geral[i][4]).trim() == pront) {
      return {
        nome: geral[i][5],
        dataNasc: parseDate(geral[i][6]),
        clinica: geral[i][3],
        origem: geral[i][2],
        municipio: "",
        dataSolic: parseDate(geral[i][0]),
        dataProg: parseDate(geral[i][9]),
        dataResp: parseDate(geral[i][10])
      };
    }
  }
  return {
    nome: "",
    dataNasc: null,
    clinica: "",
    origem: "",
    municipio: "",
    dataSolic: null,
    dataProg: null,
    dataResp: null
  };
}

function buscarSaida(saidas, pront) {
  let ultima = null;

  for (let i = 1; i < saidas.length; i++) {
    if (String(saidas[i][0]).trim() == pront) {
      const d = parseDate(saidas[i][13]);
      if (!ultima || d > ultima.dataSaida) {
        ultima = {
          dataSaida: d,
          desfecho: saidas[i][16] || "",
          local: saidas[i][14] || ""
        };
      }
    }
  }
  return ultima;
}

function buscarEmerg(emerg, pront) {
  let ultima = null;

  for (let i = 1; i < emerg.length; i++) {
    if (String(emerg[i][0]).trim() == pront) {
      const d = parseDate(emerg[i][14]);
      if (!ultima || d > ultima) ultima = d;
    }
  }
  return ultima;
}

/********************* DATE UTILS ************************/
function parseDate(v) {
  if (!v) return null;

  if (v instanceof Date) return v;

  if (typeof v === "number") {
    return new Date(v);
  }

  if (typeof v === "string") {
    const p = v.split(/[\/\-]/);
    if (p.length === 3) {
      const d = parseInt(p[0]);
      const m = parseInt(p[1]) - 1;
      const y = parseInt(p[2]);
      return new Date(y, m, d);
    }
  }
  return null;
}

function formatDate(dt) {
  if (!dt) return "–";
  let d = dt.getDate().toString().padStart(2, "0");
  let m = (dt.getMonth() + 1).toString().padStart(2, "0");
  let y = dt.getFullYear();
  return `${d}/${m}/${y}`;
}

function calcIdade(n) {
  const hoje = new Date();
  let idade = hoje.getFullYear() - n.getFullYear();
  const m = hoje.getMonth() - n.getMonth();
  if (m < 0 || (m === 0 && hoje.getDate() < n.getDate())) idade--;
  return idade;
}

function gerarDiferenca(a, b) {
  if (!a || !b) return null;
  return b.getTime() - a.getTime();
}

function formatTempo(ms) {
  if (!ms) return null;

  const min = Math.floor(ms / 60000);
  if (min < 60) return `${min} min`;

  const h = Math.floor(ms / 3600000);
  if (h < 24) return `${h} h`;

  const d = Math.floor(h / 24);
  const hh = h % 24;
  return `${d} d ${hh} h`;
}
