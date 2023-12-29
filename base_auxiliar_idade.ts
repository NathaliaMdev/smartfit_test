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
  
  function main(workbook: ExcelScript.Workbook) {
    let lista_movimentacao: Movimentacao[] = [];
    const worksheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    const lista_planos = [
      "BD",
      "Cenibra",
      "Vale Fertilizantes",
      "Vale Mais",
      "Valiaprev",
      "Mosaic Mais",
      "Abono Complementação",
      "Prevaler",
      "Prev Mosaic I",
      "Prev Mosaic II"]
  
    const lista_codigo_plano = [
      "2020001438",
      "1973000156",
      "1995002356",
      "2011002192",
      "2011002265",
      "2020000229",
      "2012000274",
      "1999005211",
      "2000008283",
      "2019002329"]
  
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
    //Pega o ano na primeira worksheet
    const ano_celula = worksheets[0].getRange("D3");
    const ano = ano_celula.getUsedRange().getValues();
  
    function iteraDados() {
  
      for (let sheetIndex = 1; sheetIndex < worksheets.length; sheetIndex++) {
        let atual_worksheet = worksheets[sheetIndex];
        let intervalo_sexoidade: ExcelScript.Range = atual_worksheet.getRange("A6:C12");
        let used_sexoidade = intervalo_sexoidade.getUsedRange();
        let intervalo_aposentado: ExcelScript.Range = atual_worksheet.getRange("A17:C23");
        let used_aposentado = intervalo_aposentado.getUsedRange();
        let intervalo_beneficiodepensao: ExcelScript.Range = atual_worksheet.getRange("A28:C34");
        let used_beneficiodepensao = intervalo_beneficiodepensao.getUsedRange();
        let nome_atual: string = worksheets[sheetIndex].getName();
  
        let sexoidade = used_sexoidade.getValues();
        let aposentado = used_aposentado.getValues();
        let beneficiodepensao = used_beneficiodepensao.getValues();
        if (sheetIndex > 10) {
          return
        }
  
        for (let plano = 0; plano < lista_planos.length; plano++) {
          for (let c_plano = 0; c_plano < lista_codigo_plano.length; c_plano++) {
            if (lista_planos[plano] === "Abono Complementação" && nome_atual === "Abono Complementação" && lista_codigo_plano[c_plano] === "2020001438") {
  
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
  
            if (lista_planos[plano] === "BD" && nome_atual === "BD" && lista_codigo_plano[c_plano] === "1973000156") {
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
  
            }
  
            if (lista_planos[plano] === "Cenibra" && nome_atual === "Cenibra" && lista_codigo_plano[c_plano] === "1995002356") {
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
  
            }
  
            if (lista_planos[plano] === "Prev Mosaic I" && nome_atual === "Prev Mosaic I" && lista_codigo_plano[c_plano] === "2011002192") {
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
            if (lista_planos[plano] === "Prev Mosaic II" && nome_atual === "Prev Mosaic II" && lista_codigo_plano[c_plano] === "2011002265") {
  
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
  
            if (lista_planos[plano] === "Mosaic Mais" && nome_atual === "Mosaic Mais" && lista_codigo_plano[c_plano] === "2020000229") {
  
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
  
            if (lista_planos[plano] === "Vale Fertilizantes" && nome_atual === "Vale Fertilizantes" && lista_codigo_plano[c_plano] === "2012000274") {
  
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
  
            if (lista_planos[plano] === "Vale Mais" && nome_atual === "Vale Mais" && lista_codigo_plano[c_plano] === "1999005211") {
  
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
            if (lista_planos[plano] === "Valiaprev" && nome_atual === "Valiaprev" && lista_codigo_plano[c_plano] === "2000008283") {
  
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
  
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
            }
            
            if (lista_planos[plano] === "Prevaler" && nome_atual === "Prevaler" && lista_codigo_plano[c_plano] === "2019002329") {
            
              let plano_atual: string = lista_planos[plano]
              let codigo_plano: string = lista_codigo_plano[c_plano]
        
              sexoidade.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[0], descricao_beneficio[0], plano_atual, codigo_plano, codigo_faixaetaria[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria[index + 1]
                faixa_etaria[index + 1]
  
              })
              aposentado.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[1], descricao_beneficio[1], plano_atual, codigo_plano, codigo_faixaetaria2[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria2[index + 1]
                faixa_etaria[index + 1]
              })
  
              beneficiodepensao.forEach((valor, index) => {
                lista_movimentacao.push(new Movimentacao(codigo_beneficio[2], descricao_beneficio[2], plano_atual, codigo_plano, codigo_faixaetaria3[index], faixa_etaria[index], valor[1].toString(), valor[2].toString(), ano.toString()))
                codigo_faixaetaria3[index + 1]
                faixa_etaria[index + 1]
              })
  
  
            }
  
  
          }
  
        }
  
      }
  
      return lista_movimentacao
    }
  
    function Create_BaseAuxiliar_Sheet() {
      iteraDados()
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
          coluna_base_auxilar.setFormulaLocal("codigo faixa-etária");
          coluna_base_auxilar = existingWorksheet.getRange("F1");
          coluna_base_auxilar.setFormulaLocal("faixa-etária");
          coluna_base_auxilar = existingWorksheet.getRange("G1");
          coluna_base_auxilar.setFormulaLocal("Masculino");
          coluna_base_auxilar = existingWorksheet.getRange("H1");
          coluna_base_auxilar.setFormulaLocal("Feminino");
          coluna_base_auxilar = existingWorksheet.getRange("I1");
          coluna_base_auxilar.setFormulaLocal("Competencia");
          coluna_base_auxilar = existingWorksheet.getRange("J1");
          coluna_base_auxilar.setFormulaLocal("Mês");
    
  
  
  
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
            cell_values.setValue(movimentacao.codigo_faixaetaria)
  
            cell_values = existingWorksheet.getRange(`F${index_sheet + 2}`)
            cell_values.setValue(movimentacao.faixa_etaria)
  
            cell_values = existingWorksheet.getRange(`G${index_sheet + 2}`)
            cell_values.setValue(movimentacao.valor_masculino)
  
            cell_values = existingWorksheet.getRange(`H${index_sheet + 2}`)
            cell_values.setValue(movimentacao.valor_feminino)
  
            cell_values = existingWorksheet.getRange(`I${index_sheet + 2}`)
            cell_values.setValue(movimentacao.ano)
  
            cell_values = existingWorksheet.getRange(`J${index_sheet + 2}`)
            cell_values.setValue(movimentacao.mes)
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
  
  
    }
    Create_BaseAuxiliar_Sheet()
  }
  