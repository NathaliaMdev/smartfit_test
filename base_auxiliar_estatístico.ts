
class Movimentacao {
    codigo_beneficio: string;
    descricao_beneficio: string;
    plano: string; //nome do plano
    codigo_plano: string;
    valor_inicial: string;
    valor_entrada: string;
    valor_saida: string;
    ano: string;
    mes: string;
    semestre: string;

  
    constructor(codigo_beneficio: string, descricao_beneficio: string, plano: string, codigo_plano: string, valor_inicial: string, valor_entrada: string, valor_saida: string, numero_linha: string | number, mes: string, ano: string, semestre: string) {
      this.codigo_beneficio = codigo_beneficio.toString();
      this.descricao_beneficio = descricao_beneficio.toString();
      this.plano = plano.toString()
      this.codigo_plano = codigo_plano.toString();
      this.valor_inicial = valor_inicial.toString();
      this.valor_inicial = isNaN(parseFloat(this.valor_inicial)) ? "0" : this.valor_inicial;
      this.valor_entrada = valor_entrada.toString();
      this.valor_entrada = isNaN(parseFloat(this.valor_entrada)) ? "0" : this.valor_entrada;
      this.valor_saida = valor_saida.toString();
      this.valor_saida = isNaN(parseFloat(this.valor_saida)) ? "0" : this.valor_saida;
  
      this.mes = mes.toString();
      this.ano = ano.toString();
      if (ano) {
        const ano = this.ano.toString();
      }
      this.semestre = semestre.toString();
    }
  }
  
  
  
  function main(workbook: ExcelScript.Workbook) {
  
    let lista_movimentacao: Movimentacao[] = [];
    let mes_atual: number = 0;
    const worksheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    const lista_meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
    const lista_planos = ["BD", "CENIBRA", "VALE FERTILIZANTES", "VALE MAIS", "VALIA PREV", "MOSAIC MAIS", "ABONO COMPLEMENTAÇÃO", "PREVALER", "MOSAIC 1", "MOSAIC 2"]
    const lista_codigo_plano = ["1973000156", "1995002356", "2012000274", "1999005211", "2000008283", "2020000229", "2020001438", "2019002329", "2011002192", "2011002265"]
    let ano_celula = worksheets[0].getRange("D3");
    const ano = ano_celula.getUsedRange().getValues();
    if(ano < 2023){
      ano_celula.setFormulaLocal("Digite um ano válido")
    }
    let semestre_celula = worksheets[0].getRange("D4");
    const semestre = semestre_celula.getUsedRange().getValues();
    if (semestre > 2 || semestre < 1){
        semestre_celula.setFormulaLocal('Digite 1 ou 2 para classificar o semestre')
    }
   

  
  
  
    for (let sheetIndex = 2; sheetIndex < worksheets.length; sheetIndex++) {
      for (let mes = 0; mes <= lista_meses.length; mes++) {
  
        let atual_worksheet = worksheets[sheetIndex];
        let intervalo: ExcelScript.Range = atual_worksheet.getRange("A4:BG25");
        let used = intervalo.getUsedRange();
  
        const nome_atual: string = worksheets[sheetIndex].getName();
        if (lista_meses[mes] == "JAN" && nome_atual == "JAN") {
          let usedRange = used.getValues();
          let mes_atual = "1";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "FEV" && nome_atual == "FEV") {
          let usedRange = used.getValues();
          let mes_atual = "2";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "MAR" && nome_atual == "MAR") {
          let usedRange = used.getValues();
          let mes_atual = "3";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "ABR" && nome_atual == "ABR") {
          let usedRange = used.getValues();
          let mes_atual = "4";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "MAI" && nome_atual == "MAI") {
          let usedRange = used.getValues();
          let mes_atual = "5";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "JUN" && nome_atual == "JUN") {
          let usedRange = used.getValues();
          let mes_atual = "6";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "JUL" && nome_atual == "JUL") {
          let usedRange = used.getValues();
          let mes_atual = "7";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "AGO" && nome_atual == "AGO") {
          let usedRange = used.getValues();
          let mes_atual = "8";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "SET" && nome_atual == "SET") {
          let usedRange = used.getValues();
          let mes_atual = "9";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "OUT" && nome_atual == "OUT") {
          let usedRange = used.getValues();
          let mes_atual = "10";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "NOV" && nome_atual == "NOV") {
          let usedRange = used.getValues();
          let mes_atual = "11";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
        if (lista_meses[mes] == "DEZ" && nome_atual == "DEZ") {
          let usedRange = used.getValues();
          let mes_atual = "12";
          usedRange.forEach((valor, index) => {
            //BD
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[0], lista_codigo_plano[0], valor[10], valor[11], valor[12], index, mes_atual, ano, semestre))
            //CENIBRA
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[1], lista_codigo_plano[1], valor[15], valor[16], valor[17], index, mes_atual, ano, semestre))
            //ValeFertilizantes
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[2], lista_codigo_plano[2], valor[20], valor[21], valor[22], index, mes_atual, ano, semestre))
            //ValeMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[3], lista_codigo_plano[3], valor[25], valor[26], valor[27], index, mes_atual, ano, semestre))
            //ValiaPrev
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[4], lista_codigo_plano[4], valor[30], valor[31], valor[32], index, mes_atual, ano, semestre))
            //MosaicMais
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[5], lista_codigo_plano[5], valor[35], valor[36], valor[37], index, mes_atual, ano, semestre))
            //Abono Complementacao 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[6], lista_codigo_plano[6], valor[40], valor[41], valor[42], index, mes_atual, ano, semestre))
            //Prevaler 
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[7], lista_codigo_plano[7], valor[45], valor[46], valor[47], index, mes_atual, ano, semestre))
            //Mosaic 1
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[8], lista_codigo_plano[8], valor[50], valor[51], valor[52], index, mes_atual, ano, semestre))
            //Mosaic 2
            lista_movimentacao.push(new Movimentacao(valor[0], valor[1], lista_planos[9], lista_codigo_plano[9], valor[55], valor[56], valor[57], index, mes_atual, ano, semestre))
          })
        }
      }
    }
  
  
  
  
    function Create_BaseAuxiliar_Sheet() {
  
  
      let existingWorksheet: ExcelScript.Worksheet | null = null;
      try {
        existingWorksheet = workbook.getWorksheet("Base Auxiliar");
      } catch (error) {
        console.log("A planilha buscada existe")
      }
  
      if (existingWorksheet) {
        let range_used = existingWorksheet.getUsedRange()
        if (range_used) {
          range_used.delete(Excel.DeleteShiftDirection.up);
          // Itera sobre as linhas da planilha
  
          let coluna_base_auxilar: ExcelScript.Range = existingWorksheet.getRange("A1");
          coluna_base_auxilar.setFormulaLocal("Codigo Beneficio");
          coluna_base_auxilar = existingWorksheet.getRange("B1");
          coluna_base_auxilar.setFormulaLocal("Descrição Beneficio");
          coluna_base_auxilar = existingWorksheet.getRange("C1");
          coluna_base_auxilar.setFormulaLocal("Plano");
          coluna_base_auxilar = existingWorksheet.getRange("D1");
          coluna_base_auxilar.setFormulaLocal("Codigo Plano");
          coluna_base_auxilar = existingWorksheet.getRange("E1");
          coluna_base_auxilar.setFormulaLocal("Inicial");
          coluna_base_auxilar = existingWorksheet.getRange("F1");
          coluna_base_auxilar.setFormulaLocal("Entrada");
          coluna_base_auxilar = existingWorksheet.getRange("G1");
          coluna_base_auxilar.setFormulaLocal("Saida");
          coluna_base_auxilar = existingWorksheet.getRange("H1");
          coluna_base_auxilar.setFormulaLocal("Ano");
          coluna_base_auxilar = existingWorksheet.getRange("I1");
          coluna_base_auxilar.setFormulaLocal("Mes");
          coluna_base_auxilar = existingWorksheet.getRange("J1");
          coluna_base_auxilar.setFormulaLocal("Semestre");
          coluna_base_auxilar = existingWorksheet.getRange("K1");
          coluna_base_auxilar.setFormulaLocal("Observacao");
          coluna_base_auxilar = existingWorksheet.getRange("L1");

  
  
  
          for (let indice in lista_movimentacao) {
            let movimentacao = lista_movimentacao[indice]
            let index_sheet = Number.parseInt(indice)
  
            let cell_values = existingWorksheet.getRange(`A${index_sheet + 2}`)
            cell_values.setValue(movimentacao.codigo_beneficio)
  
            cell_values = existingWorksheet.getRange(`B${index_sheet + 2}`)
            cell_values.setValue(movimentacao.descricao_beneficio)
            cell_values = existingWorksheet.getRange(`C${index_sheet + 2}`)
            cell_values.setValue(movimentacao.plano)
  
            cell_values = existingWorksheet.getRange(`D${index_sheet + 2}`)
            cell_values.setValue(movimentacao.codigo_plano)
  
            cell_values = existingWorksheet.getRange(`E${index_sheet + 2}`)
            cell_values.setValue(movimentacao.valor_inicial)
  
            cell_values = existingWorksheet.getRange(`F${index_sheet + 2}`)
            cell_values.setValue(movimentacao.valor_entrada)
  
            cell_values = existingWorksheet.getRange(`G${index_sheet + 2}`)
            cell_values.setValue(movimentacao.valor_saida)
  
            cell_values = existingWorksheet.getRange(`H${index_sheet + 2}`)
            cell_values.setValue(movimentacao.ano)
  
            cell_values = existingWorksheet.getRange(`I${index_sheet + 2}`)
            cell_values.setValue(movimentacao.mes)
  
            cell_values = existingWorksheet.getRange(`J${index_sheet + 2}`)
            cell_values.setValue(movimentacao.semestre)


          }
  
        }
  
      } else {
        let new_base_auxiliar_sheet: ExcelScript.Worksheet = workbook.addWorksheet("Base Auxiliar");
        let new_coluna_base_auxilar: ExcelScript.Range = new_base_auxiliar_sheet.getRange("A1");
        new_coluna_base_auxilar.setFormulaLocal("Codigo Beneficio");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("B1");
        new_coluna_base_auxilar.setFormulaLocal("Descricao Beneficio");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("C1");
        new_coluna_base_auxilar.setFormulaLocal("Plano");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("D1");
        new_coluna_base_auxilar.setFormulaLocal("Codigo Plano");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("E1");
        new_coluna_base_auxilar.setFormulaLocal("Inicial");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("F1");
        new_coluna_base_auxilar.setFormulaLocal("Entrada");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("G1");
        new_coluna_base_auxilar.setFormulaLocal("Saida");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("H1");
        new_coluna_base_auxilar.setFormulaLocal("Ano");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("I1");
        new_coluna_base_auxilar.setFormulaLocal("Mes");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("J1");
        new_coluna_base_auxilar.setFormulaLocal("Semestre");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("K1");
        new_coluna_base_auxilar.setFormulaLocal("Observação");
        new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("L1");
  
  
  
        for (let indice in lista_movimentacao) {
          let movimentacao = lista_movimentacao[indice]
          let index_sheet = Number.parseInt(indice)
  
          let cell_values = new_base_auxiliar_sheet.getRange(`A${index_sheet + 2}`)
          cell_values.setValue(movimentacao.codigo_beneficio)
  
          cell_values = new_base_auxiliar_sheet.getRange(`B${index_sheet + 2}`)
          cell_values.setValue(movimentacao.descricao_beneficio)
  
          cell_values = new_base_auxiliar_sheet.getRange(`C${index_sheet + 2}`)
          cell_values.setValue(movimentacao.plano)
  
          cell_values = new_base_auxiliar_sheet.getRange(`D${index_sheet + 2}`)
          cell_values.setValue(movimentacao.codigo_plano)
  
          cell_values = new_base_auxiliar_sheet.getRange(`E${index_sheet + 2}`)
          cell_values.setValue(movimentacao.valor_inicial)
  
          cell_values = new_base_auxiliar_sheet.getRange(`F${index_sheet + 2}`)
          cell_values.setValue(movimentacao.valor_entrada)
  
          cell_values = new_base_auxiliar_sheet.getRange(`G${index_sheet + 2}`)
          cell_values.setValue(movimentacao.valor_saida)
  
          cell_values = new_base_auxiliar_sheet.getRange(`H${index_sheet + 2}`)
          cell_values.setValue(movimentacao.ano)
  
          cell_values = new_base_auxiliar_sheet.getRange(`I${index_sheet + 2}`)
          cell_values.setValue(movimentacao.mes)
  
          cell_values = new_base_auxiliar_sheet.getRange(`J${index_sheet + 2}`)
          cell_values.setValue(movimentacao.semestre)

  
        }
  
  
      }
  
  
    }
    Create_BaseAuxiliar_Sheet()
  }