/**
 * Carrega a interface principal do web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Análise de Capabilidade do Processo');
}

/**
 * Processa os dados, grava na planilha e calcula os índices de capabilidade.
 * @param {string} featureName Nome da característica.
 * @param {string} rawDataText Valores colados no textarea, um por linha.
 * @param {number} lsl Limite inferior de especificação.
 * @param {number} usl Limite superior de especificação.
 * @param {number} subgroupSize Tamanho do subgrupo.
 * @return {Object} Objeto com os resultados para exibição no front-end.
 */
function processAndCalculate(featureName, rawDataText, lsl, usl, subgroupSize) {
  // Quebra o texto colado por linhas, converte em número e remove entradas inválidas.
  var data = rawDataText.split(/\r?\n/)
      .map(function(line) { return parseFloat(line.toString().trim().replace(',', '.')); })
      .filter(function(value) { return !isNaN(value); });

  if (data.length === 0) {
    throw new Error('Nenhum valor numérico válido foi informado.');
  }

  // Garante valores numéricos para os parâmetros recebidos da interface.
  lsl = parseFloat(lsl);
  usl = parseFloat(usl);
  subgroupSize = parseInt(subgroupSize, 10);

  if (isNaN(lsl) || isNaN(usl)) {
    throw new Error('Informe LSL e USL válidos.');
  }
  if (isNaN(subgroupSize) || subgroupSize < 2) {
    throw new Error('Informe um tamanho de subgrupo inteiro maior ou igual a 2.');
  }

  // Persiste os dados na aba "Dados" da planilha ativa, sobrescrevendo o conteúdo anterior.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Dados');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Dados');
  }
  sheet.getRange('A:A').clearContent();
  sheet.getRange(1, 1).setValue(featureName || 'Característica');
  sheet.getRange(2, 1, data.length, 1).setValues(data.map(function(value) { return [value]; }));

  var stats = calculateCapability(data, lsl, usl, subgroupSize);
  stats.featureName = featureName;
  stats.data = data;
  stats.lsl = lsl;
  stats.usl = usl;
  stats.subgroupSize = subgroupSize;

  return stats;
}

/**
 * Calcula média, desvios padrão overall e within e os índices de capabilidade.
 */
function calculateCapability(data, lsl, usl, subgroupSize) {
  var n = data.length;
  var mean = data.reduce(function(sum, value) { return sum + value; }, 0) / n;

  // Desvio padrão overall (amostral).
  var sigmaOverall = Math.sqrt(data.reduce(function(sum, value) {
    return sum + Math.pow(value - mean, 2);
  }, 0) / (n - 1));

  // Calcula o desvio padrão "within" usando desvios dos subgrupos.
  var sigmaWithin = calculateWithinSigma(data, subgroupSize);

  // Evita divisão por zero caso a variabilidade seja zero.
  if (sigmaOverall === 0) {
    sigmaOverall = Number.EPSILON;
  }
  if (sigmaWithin === 0) {
    sigmaWithin = Number.EPSILON;
  }

  // Cálculo dos índices Cp, Cpk, Pp, Ppk.
  var cp = (usl - lsl) / (6 * sigmaWithin);
  var cpk = Math.min((usl - mean) / (3 * sigmaWithin), (mean - lsl) / (3 * sigmaWithin));
  var pp = (usl - lsl) / (6 * sigmaOverall);
  var ppk = Math.min((usl - mean) / (3 * sigmaOverall), (mean - lsl) / (3 * sigmaOverall));

  return {
    n: n,
    mean: mean,
    sigmaOverall: sigmaOverall,
    sigmaWithin: sigmaWithin,
    cp: cp,
    cpk: cpk,
    pp: pp,
    ppk: ppk
  };
}

/**
 * Calcula o desvio padrão pooled com base nos subgrupos consecutivos.
 * Implementação similar ao método utilizado por softwares estatísticos para a variabilidade "within".
 */
function calculateWithinSigma(data, subgroupSize) {
  var groups = [];
  for (var i = 0; i + subgroupSize <= data.length; i += subgroupSize) {
    groups.push(data.slice(i, i + subgroupSize));
  }

  if (groups.length === 0) {
    return 0;
  }

  var numerator = 0;
  var denominator = 0;

  groups.forEach(function(group) {
    var gMean = group.reduce(function(sum, value) { return sum + value; }, 0) / group.length;
    var variance = group.reduce(function(sum, value) {
      return sum + Math.pow(value - gMean, 2);
    }, 0);
    numerator += variance;
    denominator += (group.length - 1);
  });

  if (denominator === 0) {
    return 0;
  }

  return Math.sqrt(numerator / denominator);
}
