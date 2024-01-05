
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
  
  
  function transformTable(mes: number, values: string[][], lista_movimentacao: Movimentacao[], lista_planos: string[], lista_codigo_plano: string[], ano: string, semestre: string) {
    const intervalos = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55]; // Índices dos valores para cada plano

    for (let i = 0; i < lista_planos.length; i++) {
        const intervaloInicial = intervalos[i];
        const intervaloFinal = intervaloInicial + 2;
      
        values.forEach((valor, index) => {
            lista_movimentacao.push(new Movimentacao(
                valor[0], valor[1],
                lista_planos[i], lista_codigo_plano[i],
                valor[intervaloInicial], valor[intervaloInicial + 1], valor[intervaloFinal],
                index, mes.toString(), ano, semestre
            ));
        });
    }
}

function Create_BaseAuxiliar_Sheet(lista_movimentacao:Movimentacao[],workbook: ExcelScript.Workbook) {
  
  
  let existingWorksheet: ExcelScript.Worksheet | null = null;
  try {
    existingWorksheet = workbook.getWorksheet("Base Auxiliar");
  } catch (error) {
    console.log("A planilha buscada existe")
  }

  if (existingWorksheet) {
    existingWorksheet.delete()
  } 
  
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




const mes_str_map: {str:number} = {
      "JAN": 0,
      "FEV": 1,
      "MAR": 2,
      "ABR": 3, 
      "MAI": 4, 
      "JUN": 5, 
      "JUL": 6, 
      "AGO": 7, 
      "SET": 8, 
      "OUT": 9, 
      "NOV": 10, 
      "DEZ": 11
 }

function main(workbook: ExcelScript.Workbook) {
    let hora_inicio = new Date()
    let lista_movimentacao: Movimentacao[] = [];
    const worksheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
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
    console.log(`${new Date().toISOString()} iniciando preenchimento ${new Date() - hora_inicio}`)
    hora_inicio = new Date()

    for (let sheetIndex = 2; sheetIndex < worksheets.length; sheetIndex++) {
      
      const nome_atual: string = worksheets[sheetIndex].getName();
      if(!(nome_atual in mes_str_map)){
        continue
      }
      let atual_worksheet = worksheets[sheetIndex];
      let intervalo: ExcelScript.Range = atual_worksheet.getRange("A4:BG25");
      let used = intervalo.getUsedRange();
      const usedRange = used.getValues();

      
      let mes_atual:number = mes_str_map[nome_atual] +1
      transformTable(mes_atual, usedRange, lista_movimentacao, lista_planos, lista_codigo_plano, ano, semestre);
      
      
    }  
    
    console.log(`${new Date().toISOString()} terminou preenchimento ${new Date() - hora_inicio}`)

    hora_inicio = new Date()
    Create_BaseAuxiliar_Sheet(lista_movimentacao,workbook)

    console.log(`${new Date().toISOString()} finalizou a criacao aux ${new Date() - hora_inicio}`)

  }