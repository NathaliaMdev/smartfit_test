class Movimentacao {
    codigo_beneficio: string;
    descricao_beneficio: string;
    plano: string; //nome do plano
    codigo_plano: string;
    codigo_faixaetaria: string;
    faixa_etaria: string;
    valor_masculino: string;
    valor_feminino: string;
    ano: string;
    mes: string;
  
  
    constructor(codigo_beneficio: string, descricao_beneficio: string, plano: string, codigo_plano: string, codigo_faixaetaria: string, faixa_etaria: string, valor_masculino: string, valor_feminino: string, ano: string) {
      this.codigo_beneficio = codigo_beneficio;
      this.descricao_beneficio = descricao_beneficio;
      this.plano = plano;
      this.codigo_plano = codigo_plano;
      this.codigo_faixaetaria = codigo_faixaetaria;
      this.faixa_etaria = faixa_etaria;
      this.valor_masculino = valor_masculino;
      this.valor_masculino = isNaN(parseFloat(this.valor_masculino)) ? "0" : this.valor_masculino;
      this.valor_feminino = valor_feminino;
      this.valor_feminino = isNaN(parseFloat(this.valor_feminino)) ? "0" : this.valor_feminino;
      this.ano = ano.toString();
      if (ano) {
        const ano = this.ano.toString();
      }
      this.mes = "12";
    }
  
  }
  const lista_planos: {string, number} = {
    "Abono Complementação" : 2020001438,
    "BD" : 1973000156,
    "Cenibra" : 1995002356,
    "Prev Mosaic I" : 2011002192,
    "Prev Mosaic II" : 2011002265,
    "Mosaic Mais" : 2020000229,
    "Vale Fertilizantes" : 2012000274,
    "Vale Mais" : 1999005211,
    "Valiaprev" : 2000008283,
    "Prevaler" : 2019002329,
  }

 

  const codigo_faixaetaria = [
    "41100",
    "41200",
    "41300",
    "41400",
    "41500",
    "41600",
    "41700",
  ]

  const codigo_faixaetaria2 = [
    "42100",
    "42200",
    "42300",
    "42400",
    "42500",
    "42600",
    "42700",
  ]

  const codigo_faixaetaria3 = [
    "43100",
    "43200",
    "43300",
    "43400",
    "43500",
    "43600",
    "43700",
  ]

  const codigo_beneficio = [
    "31000",
    "32000",
    "33000",
  ]


  const faixa_etaria = [
    "ATE_24",
    "ENTRE_25_34",
    "ENTRE_35_54",
    "ENTRE_55_64",
    "ENTRE_65_74",
    "ENTRE_75_84",
    "MAIS_85"
  ]

  const descricao_beneficio = [
    "Participantes (totalizador)",
    "Assistidos - Aposentados",
    "Beneficiários de Pensão",
  ]
  
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
    new_coluna_base_auxilar.setFormulaLocal("codigo faixa-etária");
    new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("F1");
    new_coluna_base_auxilar.setFormulaLocal("faixa-etária");
    new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("G1");
    new_coluna_base_auxilar.setFormulaLocal("Masculino");
    new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("H1");
    new_coluna_base_auxilar.setFormulaLocal("Feminino");
    new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("I1");
    new_coluna_base_auxilar.setFormulaLocal("Competencia");
    new_coluna_base_auxilar = new_base_auxiliar_sheet.getRange("J1");
    new_coluna_base_auxilar.setFormulaLocal("Mes");
    


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
      cell_values.setValue(movimentacao.codigo_faixaetaria)

      cell_values = new_base_auxiliar_sheet.getRange(`F${index_sheet + 2}`)
      cell_values.setValue(movimentacao.faixa_etaria)

      cell_values = new_base_auxiliar_sheet.getRange(`G${index_sheet + 2}`)
      cell_values.setValue(movimentacao.valor_masculino)

      cell_values = new_base_auxiliar_sheet.getRange(`H${index_sheet + 2}`)
      cell_values.setValue(movimentacao.valor_feminino)

      cell_values = new_base_auxiliar_sheet.getRange(`I${index_sheet + 2}`)
      cell_values.setValue(movimentacao.ano)

      cell_values = new_base_auxiliar_sheet.getRange(`J${index_sheet + 2}`)
      cell_values.setValue(movimentacao.mes)

    }

}


function transformTable(sexoidade: string[][], aposentado: string[][], beneficiodepensao: string[][], lista_movimentacao: Movimentacao[], ano: string, nome_plano : string) {
  
  // Itera sobre o array "sexoidade"
  sexoidade.forEach((valor, index) => {
    lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], nome_plano, lista_planos[nome_plano], codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()));
  });

  // Itera sobre o array "aposentado"
  aposentado.forEach((valor, index) => {
    lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], nome_plano, lista_planos[nome_plano], codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()));
  });

  // Itera sobre o array "beneficiodepensao"
  beneficiodepensao.forEach((valor, index) => {
    lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], nome_plano, lista_planos[nome_plano], codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()));
  });
  
  

}



function main(workbook: ExcelScript.Workbook) {
  let hora_inicio = new Date()
    let lista_movimentacao: Movimentacao[] = [];
    const worksheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    //Pega o ano na primeira worksheet
    const ano_celula = worksheets[0].getRange("D3");
    const ano = ano_celula.getUsedRange().getValue().toString();
  
    console.log(`${new Date().toISOString()} Entrou no loop da sheet  ${new Date() - hora_inicio}`)
    for (let sheetIndex = 1; sheetIndex <= 10; sheetIndex++) {
        let nome_planilha = worksheets[sheetIndex].getName()
        let intervalo_sexoidade: ExcelScript.Range = worksheets[sheetIndex].getRange("A6:C12");
        let used_sexoidade = intervalo_sexoidade.getUsedRange();
        let intervalo_aposentado: ExcelScript.Range = worksheets[sheetIndex].getRange("A17:C23");
        let used_aposentado = intervalo_aposentado.getUsedRange();
        let intervalo_beneficiodepensao: ExcelScript.Range = worksheets[sheetIndex].getRange("A28:C34");
        let used_beneficiodepensao = intervalo_beneficiodepensao.getUsedRange();

        let sexoidade = used_sexoidade.getValues();
        let aposentado = used_aposentado.getValues();
        let beneficiodepensao = used_beneficiodepensao.getValues();

    
        
        transformTable(sexoidade,aposentado,beneficiodepensao,lista_movimentacao,ano, nome_planilha);
      
      }
  
    console.log(`${new Date().toISOString()} terminou preenchimento ${new Date() - hora_inicio}`)

    hora_inicio = new Date()

    Create_BaseAuxiliar_Sheet(lista_movimentacao,workbook)

    console.log(`${new Date().toISOString()} finalizou a criacao aux ${new Date() - hora_inicio}`)

  
    } 