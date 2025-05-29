// Função principal para listar todas as OS
function criarTriggersOmie() {
  // Remove triggers existentes para evitar duplicação
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    ScriptApp.deleteTrigger(t);
  }

  const horas = [5, 12, 15]; // Horários desejados

  for (const hora of horas) {
    // Trigger para OS
    ScriptApp.newTrigger("listarTodasOrdensServico")
      .timeBased()
      .atHour(hora)
      .everyDays(1)
      .create();

    // Trigger para Pedidos de Venda
    ScriptApp.newTrigger("listarTodosPedidosVendaOmie_CORRETO")
      .timeBased()
      .atHour(hora)
      .nearMinute(30)      // 30 minutos depois
      .everyDays(1)
      .create();
  }
}

function apagarTodasAsTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    ScriptApp.deleteTrigger(t);
  }
}

function listarTodasOrdensServico() {
  try {
    const { sheet, headers } = inicializarPlanilhaServicos(true);
    const todasOS = buscarTodasOS();

    if (todasOS.length > 0) {
      const mapaProjetos = buscarNomesProjetos(todasOS);
      const mapaClientes = buscarNomesFantasiasClientes(todasOS);
      const mapaVendedores = buscarNomesVendedores(todasOS);
      const dadosFormatados = processarDadosOS(todasOS, mapaProjetos, mapaClientes, mapaVendedores);

      // Apagar linhas antigas em blocos de 1000
      const totalAntigas = sheet.getLastRow() - 1;
      const maxClearBatch = 1000;
      for (let i = 0; i < totalAntigas; i += maxClearBatch) {
        const linhasParaApagar = Math.min(maxClearBatch, totalAntigas - i);
        sheet.getRange(2 + i, 1, linhasParaApagar, headers.length).clearContent();
      }

      // Inserir novas linhas em blocos
      const batchSize = 1000;
      for (let i = 0; i < dadosFormatados.length; i += batchSize) {
        const batch = dadosFormatados.slice(i, i + batchSize);
        sheet.getRange(i + 2, 1, batch.length, headers.length).setValues(batch);
      }

      formatarPlanilha(sheet, headers);
      SpreadsheetApp.flush(); // Garante que tudo foi salvo no Sheets
      Logger.log('Processamento finalizado');
      registrarUltimaAtualizacao();
      return;
    } else {
      console.log("Nenhuma OS encontrada.");
      return;
    }

  } catch (error) {
    console.error("Erro no processamento principal:", error);
  }
}

function registrarUltimaAtualizacao() {
  const ss = SpreadsheetApp.openById("1peZYbOOT8g_5cntCehgTh0W9fo4i46S6aaMHXDl5B_8");
  const abaConfiguracoes = ss.getSheetByName("Config");
  if (!abaConfiguracoes) return;

  const agora = new Date();
  abaConfiguracoes.getRange("G6").setValue("Última atualização OS: " + agora.toLocaleString("pt-BR"));

}


function buscarTodasOS() {
  const urlOS = "https://app.omie.com.br/api/v1/servicos/os/";
  const { appKey, appSecret } = getCredentials();
  let pagina = 1;
  let todasOS = [];

  do {
    const payloadOS = {
      app_key: appKey,
      app_secret: appSecret,
      call: "ListarOS",
      param: [{
        pagina: pagina,
        registros_por_pagina: 500,
        apenas_importado_api: "N"
      }]
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payloadOS),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(urlOS, options);
      const data = JSON.parse(response.getContentText());
      if (data.osCadastro && data.osCadastro.length > 0) {
        todasOS = todasOS.concat(data.osCadastro);
        console.log(`Página ${pagina}: ${data.osCadastro.length} OS`);
      } else {
        break;
      }
    } catch (error) {
      console.error(`Erro na página ${pagina}:`, error);
      throw new Error(`Erro na página ${pagina}: ${error.message}`);
    }

    pagina++;
  } while (pagina <= 10);

  return todasOS;
}

function buscarNomesProjetos(ordens) {
  return buscarNomesGenerico(ordens, "Cache_Projetos", "https://app.omie.com.br/api/v1/geral/projetos/", "nCodProj", "ConsultarProjeto", "codigo");
}

function buscarNomesFantasiasClientes(ordens) {
  return buscarNomesGenerico(ordens, "Cache_Clientes", "https://app.omie.com.br/api/v1/geral/clientes/", "nCodCli", "ConsultarCliente", "codigo_cliente_omie");
}

function buscarNomesVendedores(ordens) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cache_Vendedores");
  const url = "https://app.omie.com.br/api/v1/geral/vendedores/";
  const { appKey, appSecret } = getCredentials();
  const mapa = {};
  const usados = new Set();

  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const [codigo, nome] = dados[i];
    mapa[codigo] = nome;
    usados.add(codigo);
  }

  const codigos = [...new Set(ordens.map(o => o.Cabecalho?.nCodVend).filter(c => c && !usados.has(c)))];
  let pagina = 1;
  let encontrouTodos = false;

  while (!encontrouTodos && pagina <= 20) {
    const payload = {
      app_key: appKey,
      app_secret: appSecret,
      call: "ListarVendedores",
      param: [{ pagina, registros_por_pagina: 50 }]
    };
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      const cadastros = json.cadastro;
      if (!cadastros || cadastros.length === 0) break;

      cadastros.forEach(v => {
        const codigo = v.codigo;
        const nome = v.nome || "Sem Nome";
        if (!mapa[codigo]) {
          mapa[codigo] = nome;
          sheet.appendRow([codigo, nome]);
          usados.add(codigo);
        }
      });

      const faltando = codigos.filter(cod => !mapa[cod]);
      if (faltando.length === 0) encontrouTodos = true;
    } catch (e) {
      console.error(`Erro na página ${pagina} ao buscar vendedores:`, e);
      break;
    }

    pagina++;
  }
  return mapa;
}

function buscarNomesGenerico(ordens, aba, url, campoCodigo, call, paramName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aba);
  const { appKey, appSecret } = getCredentials();
  const mapa = {};
  const usados = new Set();

  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const [codigo, nome] = dados[i];
    mapa[codigo] = nome;
    usados.add(codigo);
  }

  const codigos = [...new Set(ordens.map(o => o.Cabecalho?.[campoCodigo] || o.InformacoesAdicionais?.[campoCodigo]).filter(c => c && !usados.has(c)))];

  codigos.forEach(cod => {
    try {
      const payload = {
        app_key: appKey,
        app_secret: appSecret,
        call,
        param: [{ [paramName]: cod }]
      };
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());
      const nome = data.nome || data.nome_fantasia || "Não encontrado";
      mapa[cod] = nome;
      sheet.appendRow([cod, nome]);
    } catch (e) {
      console.error(`Erro ao buscar código ${cod} na aba ${aba}:`, e);
      mapa[cod] = "Erro";
    }
  });
  return mapa;
}

function processarDadosOS(dados, projetos, clientes, vendedores) {
  const linhas = [];
  dados.forEach(os => {
    const cab = os?.Cabecalho || {};
    const infoAd = os.InformacoesAdicionais || {};
    const infoCadastro = os.InfoCadastro || {}; 
    const cancelada = infoCadastro.cCancelada === "N" ? "NÃO" : "SIM";
    const etapa = getDescricaoEtapa(cab.cEtapa, infoAd.nCodCC);
    const numPedido = infoAd.cNumPedido || "NÃO PREENCHIDO";
    const projeto = projetos[infoAd.nCodProj] || "NÃO PREENCHIDO";
    const cliente = clientes[cab.nCodCli] || "";
    const servicos = os.servicos?.servico || os.ServicosPrestados || [];
    const vendedor = vendedores[cab.nCodVend] || "NÃO INFORMADO";
    const dataemissao = infoCadastro.dDtFat || "AINDA NÃO FINALIZADA";
    const dataprev = cab.dDtPrevisao || "";
    const tipofat = infoAd.cNumContrato || "NÃO PREENCHIDO";
    const obs = os.Observacoes?.cObsOS || "";

    servicos.forEach(servico => {
      const tipo = getTipoServico(servico.nCodServico);
      const qtd = parseFloat(servico.nQtde || 0);
      const val = parseFloat(servico.nValUnit || 0);
      const total = qtd * val;

      linhas.push([
        cab.cNumOS || "ERRO",
        infoCadastro.dDtInc || "",
        etapa,
        numPedido,
        servico.nValorDesconto || "0.00",
        cab.nValorTotal || "0.00",
        cancelada,
        tipo,
        Math.round(qtd),
        val,
        projeto,
        cliente,
        total,
        vendedor,
        dataemissao,
        dataprev,
        tipofat,
        obs
      ]);
    });
  });
  return linhas;
}

function inicializarPlanilhaServicos(preservarDados = false) {
  const spreadsheet = SpreadsheetApp.openById("1peZYbOOT8g_5cntCehgTh0W9fo4i46S6aaMHXDl5B_8");
  const aba = "Serviço";
  let sheet = spreadsheet.getSheetByName(aba);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(aba);
  } else if (!preservarDados) {
    sheet.getRange("A2:Z").clearContent();
  }

  const headers = [
    "NUMERO DA ORDEM", "DATA ABERTURA", "ETAPA", "NUMERO DO PEDIDO", "VALOR DESCONTO",
    "TOTAL", "CANCELADAS", "TIPO", "QTD", "VALOR UNITÁRIO",
    "NOME PROJETO", "NOME CLIENTE", "VALOR TOTAL", "VENDEDOR", "DATA DE EMISSÃO",
    "PREVISÃO DE FATURAMENTO", "Nº CONTRATO VENDA", "OBSERVAÇÕES"
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const currentHeaders = headerRange.getValues()[0];
  if (currentHeaders.every(cell => cell === "")) {
    headerRange.setValues([headers]);
  }

  return { sheet, headers };
}

function formatarPlanilha(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat("R$ #,##0.00");
    sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat("R$ #,##0.00");
    sheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat("R$ #,##0.00");
    sheet.getRange(2, 13, lastRow - 1, 1).setNumberFormat("R$ #,##0.00");
    if (sheet.getLastRow() < 200) {
      sheet.autoResizeColumns(1, headers.length);
    }
  }
}

function getDescricaoEtapa(etapa, codContaCorrente) {
  switch (etapa?.toString()) {
    case "00": return "1. ORÇAMENTO";
    case "20": return "2. EM EXECUÇÃO";
    case "30": return "3. EXECUTADA";
    case "50": return "4. PROCESSO DE GARANTIA";
    case "60":
      if (codContaCorrente == 2041409931) return "5. FINALIZADO (INTERNO)";
      if (codContaCorrente == 1969919786) return "5.1 FATURADO (NFS)";
      return "5.1 FATURADO (NFS)";
    default: return "ERRO";
  }
}

function getTipoServico(cod) {
  switch (cod) {
    case 1979758762: return "HORA";
    case 1975974257: return "KM";
    case 2209673817: return "SERVIÇO";
    default: return "ERRO";
  }
}

function getCredentials() {
  return {
    appKey: "SUA_KEY",
    appSecret: "SUA_APPSECRET"
  };
}

