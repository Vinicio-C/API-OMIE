// Função principal para puxar todos os pedidos de venda (todas as páginas) numa execução só
function listarTodosPedidosVendaOmie_CORRETO() {
  try {

    const { sheet, headers } = inicializarPlanilha();

    const mapaClientes = buscarTodosClientesOmie_CACHE();
    const todosPedidos = [];

    let pagina = 1;
    while (true) {
      const pedidosPagina = buscarPedidosPagina(pagina);
      if (pedidosPagina.length === 0) break;

      todosPedidos.push(...pedidosPagina);
      console.log(`Página ${pagina}: ${pedidosPagina.length} pedidos obtidos, total acumulado: ${todosPedidos.length}`);
      pagina++;
      Utilities.sleep(300); // Para evitar limites da API
    }

    const mapaProjetos = buscarNomesProjetosVenda_CACHE(todosPedidos);

    const dadosFormatados = processarDadosPedidos(todosPedidos, mapaClientes, mapaProjetos);

    // Verifica consistência das linhas formatadas
    const linhasComErros = dadosFormatados.filter(linha => linha.length !== headers.length);
    if (linhasComErros.length > 0) {
      throw new Error("Inconsistência no número de colunas nos dados formatados.");
    }

    escreverDadosNaPlanilha(sheet, headers, dadosFormatados);
    if (dadosFormatados.length > 0) formatarPlanilhaVenda(sheet);

    console.log(`✅ Total de pedidos extraídos: ${todosPedidos.length}`);
    registrarUltimaAtualizacaoPedidos();
  } catch (error) {
    console.error("Erro no processamento principal:", error);
    throw error;
  } 
}

function registrarUltimaAtualizacaoPedidos() {
  const ss = SpreadsheetApp.openById("1peZYbOOT8g_5cntCehgTh0W9fo4i46S6aaMHXDl5B_8");
  const abaConfiguracoes = ss.getSheetByName("Config");
  if (!abaConfiguracoes) return;

  const agora = new Date();
  abaConfiguracoes.getRange("G11").setValue("Última atualização Pedidos de Venda: " + agora.toLocaleString("pt-BR"));

}

// Inicializa a planilha "Pedidos", limpa conteúdo e escreve cabeçalho
function inicializarPlanilha() {
  const spreadsheet = SpreadsheetApp.openById("1peZYbOOT8g_5cntCehgTh0W9fo4i46S6aaMHXDl5B_8");
  const nomeAba = "Pedidos";
  let sheet = spreadsheet.getSheetByName(nomeAba);
  if (!sheet) sheet = spreadsheet.insertSheet(nomeAba);
  sheet.clearContents();

  const headers = [
    "NUMERO DA VENDA",
    "DATA ABERTURA",
    "ETAPA",
    "NUMERO DO PEDIDO",
    "VALOR DESCONTO",
    "TOTAL",
    "CANCELADAS",
    "CLIENTE",
    "NOME PROJETO",
    "DATA DE EMISSÃO"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return { sheet, headers };
}

// Função que busca uma página específica de pedidos da API Omie
function buscarPedidosPagina(pagina) {
  const url = "https://app.omie.com.br/api/v1/produtos/pedido/";
  const { appKey, appSecret } = getCredentials();

  const payload = {
    app_key: appKey,
    app_secret: appSecret,
    call: "ListarPedidos",
    param: [{ pagina, registros_por_pagina: 500, apenas_importado_api: "N" }]
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(resp.getContentText());
  const pedidos = json.pedido_venda_produto || [];
  return pedidos;
}

// Cache para clientes - carrega da aba e busca novos na API, atualizando cache
function buscarTodosClientesOmie_CACHE() {
  const cacheSheet = getOrCreateSheet("Cache_Clientes");
  const cache = {};

  const dadosvenda = cacheSheet.getDataRange().getValues();
  if (dadosvenda.length === 0) cacheSheet.appendRow(["codigo_cliente", "nome_fantasia"]);

  dadosvenda.slice(1).forEach(linha => {
    const codigo = linha[0]?.toString();
    const nome = linha[1];
    if (codigo) cache[codigo] = nome;
  });

  const url = "https://app.omie.com.br/api/v1/geral/clientes/";
  const { appKey, appSecret } = getCredentials();

  let pagina = 1;
  const novos = [];

  while (true) {
    const payload = {
      app_key: appKey,
      app_secret: appSecret,
      call: "ListarClientes",
      param: [{ pagina, registros_por_pagina: 500 }]
    };

    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(resp.getContentText());
    const clientes = json.clientes_cadastro || [];
    if (!clientes.length) break;

    for (const cli of clientes) {
      const codigo = cli.codigo_cliente_omie?.toString();
      const nome = cli.nome_fantasia || "Sem nome";

      if (codigo && !cache[codigo]) {
        cache[codigo] = nome;
        novos.push([codigo, nome]);
      }
    }

    if (clientes.length < 500) break;
    pagina++;
    Utilities.sleep(1000);
  }

  if (novos.length > 0) {
    const ultimaLinha = cacheSheet.getLastRow();
    cacheSheet.getRange(ultimaLinha + 1, 1, novos.length, 2).setValues(novos);
  }

  return cache;
}

// Cache para projetos - carrega da aba e busca novos na API, atualizando cache
function buscarNomesProjetosVenda_CACHE(pedidos) {
  const cacheSheet = getOrCreateSheet("Cache_Projetos");
  const cache = {};

  const dadosvenda = cacheSheet.getDataRange().getValues();
  dadosvenda.slice(1).forEach(linha => {
    const codigo = linha[0];
    const nome = linha[1];
    if (codigo) cache[codigo] = nome;
  });

  const url = "https://app.omie.com.br/api/v1/geral/projetos/";
  const { appKey, appSecret } = getCredentials();

  // Extrai todos os códigos únicos de projeto dos pedidos
  const codigosProjeto = [...new Set(pedidos.map(p => p.informacoes_adicionais?.codProj || p.informacoes_adicionais?.cProjeto).filter(Boolean))];

  const novos = [];
  for (const codigoStr of codigosProjeto) {
    const codigo = parseInt(codigoStr);
    if (!codigo || cache[codigo]) continue;

    const payload = {
      app_key: appKey,
      app_secret: appSecret,
      call: "ConsultarProjeto",
      param: [{ codigo }]
    };

    try {
      const resp = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const dadosvenda = JSON.parse(resp.getContentText());
      const nome = dadosvenda.nome || "Projeto sem nome";
      cache[codigo] = nome;
      novos.push([codigo, nome]);
    } catch (e) {
      console.warn(`Erro ao buscar projeto ${codigo}: ${e.message}`);
      cache[codigo] = "Erro na consulta";
      novos.push([codigo, "Erro na consulta"]);
    }

    Utilities.sleep(400);
  }

  if (novos.length > 0) {
    const ultimaLinha = cacheSheet.getLastRow();
    cacheSheet.getRange(ultimaLinha + 1, 1, novos.length, 2).setValues(novos);
  }

  return cache;
}

// Utilitário para criar ou obter aba
function getOrCreateSheet(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.appendRow(["codigo", "nome"]);
  }
  return aba;
}

// Processa os pedidos para o formato da planilha
function processarDadosPedidos(pedidos, mapaClientes, mapaProjetos) {
  return pedidos.map(p => {
    const c = p.cabecalho || {};
    const info = p.infoCadastro || {};
    const add = p.informacoes_adicionais || {};
    const tot = p.total_pedido || {};

    const etapa = c.etapa || "00";
    const descEtapa = getDescricaoEtapaVenda(etapa);
    const cancel = info.cancelado === "S";

    const codCli = (c.codigo_cliente || "").toString();
    const nomeCli = mapaClientes[codCli] || "Cliente não encontrado";

    const codProj = parseInt(add.codProj || add.cProjeto) || null;
    const nomProj = codProj ? (mapaProjetos[codProj] || "Projeto não encontrado") : "NÃO INFORMADO";

    return [
      c.numero_pedido || "ERRO",
      info.dInc || c.data_previsao || "",
      descEtapa,
      add.numero_pedido_cliente || "NÃO PREENCHIDO",
      tot.valor_descontos || 0,
      tot.valor_total_pedido || 0,
      cancel ? "SIM" : "NÃO",
      nomeCli,
      nomProj,
      info.dFat || "NÃO FATURADO"
    ].map(v => v == null ? "" : String(v).trim());
  });
}

// Escreve os dados formatados na planilha, a partir da linha 2
function escreverDadosNaPlanilha(sheet, headers, dadosvenda) {
  if (dadosvenda.length === 0) return;
  sheet.getRange(2, 1, dadosvenda.length, headers.length).setValues(dadosvenda);
}

// Formata as colunas da planilha "Pedidos"
function formatarPlanilhaVenda(sheet) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr <= 1) return;

  try {
    sheet.getRange(2, 5, lr - 1, 2).setNumberFormat('R$ #,##0.00');  // VALOR DESCONTO e TOTAL
    sheet.getRange(2, 2, lr - 1, 1).setNumberFormat('dd/MM/yyyy');   // DATA ABERTURA
    sheet.autoResizeColumns(1, lc);
  } catch (e) {
    console.error("Erro na formatação:", e.message);
  }
}

// Mapeia código de etapa para descrição
function getDescricaoEtapaVenda(codigoEtapa) {
  const map = {
    "00": "ORÇAMENTO",
    "10": "APROVADO",
    "20": "SEPARAR ESTOQUE",
    "50": "FATURAR...",
    "60": "FATURADO"
  };
  return map[codigoEtapa] || "ETAPA DESCONHECIDA";
}

// Credenciais da API Omie (coloque as suas)
function getCredentials() {
  return {
    appKey: "SUA_KEY",
    appSecret: "SUA_APPSECRET"
  };
}
