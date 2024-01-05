
class Movimentacao {
    //Declara o tipo das variáveis serão colocadas no mapa 
    codigo_beneficio: string;
    codigo_plano: string;
    inicial: string;
    entrada: string;
    saida: string;
    ano: string;
    mes: string;
    semestre: string;
    observacao: string;
    valor_final: string;
    numero_linha: number;
  
    //Cria um construtor, que fornecerá acesso as variáveis e a posição da linha. 
    constructor(arr_movimentacao: string | number[], numero_linha) {
  
      this.codigo_beneficio = arr_movimentacao[0].toString();
      this.codigo_plano = arr_movimentacao[3].toString();
      this.ano = arr_movimentacao[7].toString();
      this.mes = arr_movimentacao[8].toString();
      this.semestre = arr_movimentacao[9].toString();
      this.observacao = arr_movimentacao[10].toString();
      this.numero_linha = numero_linha;
  
  
      //regra do código 24000 | não permite a entrada de dados, devendo seus valores permanecerem NULOS ou ZERADOS
      if (this.codigo_beneficio == "24000") {
        this.inicial = "0";
        this.entrada = "0";
        this.saida = "0";
      } else {
        this.inicial = arr_movimentacao[4].toString();
        this.inicial = isNaN(parseFloat(this.inicial)) ? "0" : this.inicial;
        this.entrada = arr_movimentacao[5].toString();
        this.entrada = isNaN(parseFloat(this.entrada)) ? "0" : this.entrada;
        this.saida = arr_movimentacao[6].toString();
        this.saida = isNaN(parseFloat(this.saida)) ? "0" : this.saida;
      }
  
      //regra do código 13000,15000,16000, 23000 | o valor inicial do primeiro mes sempre será 0. 
      if ((this.codigo_beneficio == "13000" || this.codigo_beneficio == "15000" || this.codigo_beneficio == "16000" || this.codigo_beneficio == "23000" || this.codigo_beneficio == "24100" || this.codigo_beneficio == "24200") && this.mes == "1") {
        this.inicial = "0";
      }
  
      // regra do código 24100 | Primeiro mês de entrada precisa ter o valor inicial igual a 0, e as entradas sempre serão de valor 0; 
      if (this.codigo_beneficio == "24100") {
        this.entrada = "0";
      }
      // regra do código 24200 | A saída sempre será 0, precisa ser alterada depois pelo arquivo retificador.
      if (this.codigo_beneficio == "24200") {
        this.saida = "0";
      }
  
      if (this.codigo_beneficio == "24100") {
        this.valor_final = (Number.parseInt(this.inicial) + Number.parseInt(this.saida)).toString()
  
      }
      else if (this.codigo_beneficio == "24200") {
        this.valor_final = (Number.parseInt(this.inicial) + Number.parseInt(this.entrada)).toString();
  
      } else {
        this.valor_final = (Number.parseInt(this.inicial) + Number.parseInt(this.entrada) - Number.parseInt(this.saida)).toString();
      }
    }
  }
  
  function transformTable(mes: number, values: string[][], lista_Movimentacao_Base: Movimentacao_Base[], lista_planos: string[], lista_codigo_plano: string[], ano: string, semestre: string) {
      const intervalos = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55]; // Índices dos valores para cada plano
  
      for (let i = 0; i < lista_planos.length; i++) {
          const intervaloInicial = intervalos[i];
          const intervaloFinal = intervaloInicial + 2;
        
          values.forEach((valor, index) => {
              lista_Movimentacao_Base.push(new Movimentacao_Base(
                  valor[0], valor[1],
                  lista_planos[i], lista_codigo_plano[i],
                  valor[intervaloInicial], valor[intervaloInicial + 1], valor[intervaloFinal],
                  index, mes.toString(), ano, semestre
              ));
          });
      }
  }
  
  function Reune_Dados_Base_Auxiliar(lista_Movimentacao_Base:Movimentacao_Base[],workbook: ExcelScript.Workbook) {
    
    
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
  
  
  
      for (let indice in lista_Movimentacao_Base) {
        let Movimentacao_Base = lista_Movimentacao_Base[indice]
        let index_sheet = Number.parseInt(indice)
  
        let cell_values = new_base_auxiliar_sheet.getRange(`A${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.codigo_beneficio)
  
        cell_values = new_base_auxiliar_sheet.getRange(`B${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.descricao_beneficio)
  
        cell_values = new_base_auxiliar_sheet.getRange(`C${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.plano)
  
        cell_values = new_base_auxiliar_sheet.getRange(`D${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.codigo_plano)
  
        cell_values = new_base_auxiliar_sheet.getRange(`E${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.valor_inicial)
  
        cell_values = new_base_auxiliar_sheet.getRange(`F${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.valor_entrada)
  
        cell_values = new_base_auxiliar_sheet.getRange(`G${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.valor_saida)
  
        cell_values = new_base_auxiliar_sheet.getRange(`H${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.ano)
  
        cell_values = new_base_auxiliar_sheet.getRange(`I${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.mes)
  
        cell_values = new_base_auxiliar_sheet.getRange(`J${index_sheet + 2}`)
        cell_values.setValue(Movimentacao_Base.semestre)
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
  
   function cria_xml(workbook: ExcelScript.Workbook) {
    // Obtém a planilha "Base Auxiliar" do Excel
    const base_auxiliar: ExcelScript.Worksheet = workbook.getWorksheet("Base Auxiliar");
  
  
    // Lista para armazenar instâncias da classe Movimentacao
    let list_movimentacao: Movimentacao[] = [];
  
    let primeira_linha = true;
    // Obtém os valores da planilha "Base Auxiliar"
    let usedRange = base_auxiliar.getUsedRange().getValues();
    var mes_inicial = 1;
    var mes_final = 6;
    var semestre = 1;
    // Itera sobre as linhas da planilha
    usedRange.forEach((linha, index) => {
      if (primeira_linha) {
        primeira_linha = false;
        return
      }
      // Cria instâncias da classe Movimentacao e adiciona à lista
      list_movimentacao.push(new Movimentacao(linha, index))
    });
  
    let semestreUm = list_movimentacao.some(mov => mov.semestre == "1");
  
    if (semestreUm) {
      mes_inicial = 1;
      mes_final = 6;
    } else {
      mes_inicial = 7;
      mes_final = 12;
      semestre = 2;
  
    }
  
    // Função para filtrar movimentações por mês e plano consolidado
    function filtra_codigo_por_mes_e_plano_consolidado() {
      const codigo_beneficio = ['11100', '11200', '12000', '13000', '14000', '15000', '16000', '17000', '21000', '22000', '23000', '24000', '24100', '24200', '31100', '31200', '31300', '33000', '34000']
      let mapa_consolidado_mes = new Map<string, MovimentacaoConsolidado[]>()
      for (let m = mes_inicial; m <= mes_final; m++) {
        //configura a chave do mapa   
        mapa_consolidado_mes.set(m.toString(), []);
  
        for (let c: number = 0; c < codigo_beneficio.length; c++) {
          let movimentos_mes = list_movimentacao.filter(mov => mov.codigo_beneficio === codigo_beneficio[c] && mov.mes === m.toString())
          let movimento_consolidado: MovimentacaoConsolidado = new MovimentacaoConsolidado(codigo_beneficio[c])
          movimentos_mes.forEach(movimentacao => {
            //soma todas os valores filtrados no objeto  
            movimento_consolidado.mes = movimentacao.mes
            movimento_consolidado.inicial += Number.parseInt(movimentacao.inicial);
            movimento_consolidado.entrada += Number.parseInt(movimentacao.entrada);
            movimento_consolidado.saida += Number.parseInt(movimentacao.saida);
            if (movimentacao.observacao !== "" && movimentacao.observacao !== undefined && movimentacao.observacao !== null) {
              movimento_consolidado.observacao += movimentacao.observacao + " ";
            }
            //somatório específico "24100 || Que também serve para validar os somatórios iniciais e finais"
            if (movimentacao.codigo_beneficio == "24100") {
              movimento_consolidado.final += Number.parseInt(movimentacao.inicial) + Number.parseInt(movimentacao.saida);
              //somatório específico "24200 Que também serve para validar os somatórios iniciais e finais"
            } else if (movimentacao.codigo_beneficio == "24200") {
              movimento_consolidado.final += Number.parseInt(movimentacao.inicial) + Number.parseInt(movimentacao.entrada);
            } else {
              // validacao de valores finais com valores de entrada do mês seguinte  
              movimento_consolidado.final = movimento_consolidado.inicial + movimento_consolidado.entrada - movimento_consolidado.saida;
            }
          });
  
  
          if (movimento_consolidado.observacao !== "" && movimento_consolidado !== null) {
            movimento_consolidado.observacao = movimento_consolidado.observacao.substring(0, movimento_consolidado.observacao.length + 1)
          }
          
          if (movimento_consolidado.codigo_beneficio == "14000") {
            // Encontrar o movimento correspondente ao código 33000
            let movimento_33000 = mapa_consolidado_mes.get(m.toString())!.find(mov => mov.codigo_beneficio == '33000');
  
            if (movimento_33000) {
              // Verificar se o final do movimento_consolidado é igual a 0
              if (movimento_consolidado.final === 0) {
                // Atribuir o valor final de movimento_consolidado ao movimento_33000.final
                movimento_33000.final = movimento_consolidado.final;
              }
            }
          }
          mapa_consolidado_mes.get(m.toString())!.push(movimento_consolidado)
         
  
      }
          let movimento_11100 = mapa_consolidado_mes.get(m.toString())!.find(mov => mov.codigo_beneficio == '11100');
          let movimento_11200 = mapa_consolidado_mes.get(m.toString())!.find(mov => mov.codigo_beneficio == '11200');
  
          let movimento_consolidado_11000 = new MovimentacaoConsolidado("11000");
          movimento_consolidado_11000.inicial +=  Number.parseInt(movimento_11100?.inicial) + Number.parseInt(movimento_11200?.inicial)
          movimento_consolidado_11000.entrada += Number.parseInt(movimento_11100?.entrada) + Number.parseInt(movimento_11200?.entrada)
          movimento_consolidado_11000.saida += Number.parseInt(movimento_11100?.saida) + Number.parseInt(movimento_11200?.saida)
          movimento_consolidado_11000.final += Number.parseInt(movimento_11100?.final) + Number.parseInt(movimento_11200?.final)
          movimento_consolidado_11000.observacao += movimento_11100.observacao + " " + movimento_11200.observacao
          mapa_consolidado_mes.get(m.toString()).push(movimento_consolidado_11000)
  
          let movimento_consolidado_32000 = new MovimentacaoConsolidado("32000")
                
          movimento_consolidado_32000.inicial = movimento_consolidado_11000.inicial;
          movimento_consolidado_32000.entrada = movimento_consolidado_11000.entrada;
          movimento_consolidado_32000.saida = movimento_consolidado_11000.saida;
          movimento_consolidado_32000.final = movimento_consolidado_11000.final;
          movimento_consolidado_32000.observacao = movimento_consolidado_11000.observacao;
    
          mapa_consolidado_mes.get(m.toString())!.push(movimento_consolidado_32000);
        
                
          // preenche o valor de 31000 com os valores 31100, 31200 e 31300
          let movimento_31100 = mapa_consolidado_mes.get(m.toString())!.find(mov => mov.codigo_beneficio == '31100');
          let movimento_31200 = mapa_consolidado_mes.get(m.toString())!.find(mov => mov.codigo_beneficio == '31200');
          let movimento_31300 = mapa_consolidado_mes.get(m.toString())!.find(mov => mov.codigo_beneficio == '31300');
  
          let movimento_consolidado_31000 = new MovimentacaoConsolidado("31000")
          movimento_consolidado_31000.inicial +=  Number.parseInt(movimento_31100?.inicial) + Number.parseInt(movimento_31200?.inicial)+ Number.parseInt(movimento_31300?.inicial)
          movimento_consolidado_31000.entrada += Number.parseInt(movimento_31100?.entrada) + Number.parseInt(movimento_31200?.entrada)+ Number.parseInt(movimento_31300?.entrada)
          movimento_consolidado_31000.saida += Number.parseInt(movimento_31100?.saida) + Number.parseInt(movimento_31200?.saida)+ Number.parseInt(movimento_31300?.saida)
          movimento_consolidado_31000.final += Number.parseInt(movimento_consolidado_31000?.inicial) + Number.parseInt(movimento_consolidado_31000?.entrada) - Number.parseInt(movimento_consolidado_31000?.saida)
          movimento_consolidado_31000.observacao += movimento_31100.observacao + " " + movimento_31200.observacao + " " + movimento_31300.observacao
          movimento_consolidado_31000.mes = m;
          mapa_consolidado_mes.get(m.toString()).push(movimento_consolidado_31000)
  
  
  
        mapa_consolidado_mes.get(m.toString())?.sort((a, b) => a.codigo_beneficio.localeCompare(b.codigo_beneficio))
    }
      return mapa_consolidado_mes
    }
  
    function root_father_xml(balancete_estatistico: string[]) {
      let ano = list_movimentacao[0].ano
      const worksheet:ExcelScript.Worksheet = workbook.getWorksheets()[0];
      const email_celula: ExcelScript.Range = worksheet.getRange("D2");
      const email_range: ExcelScript.Range = email_celula.getUsedRange();
      const email = email_range.getValues()[0][0];
  
  
      // Constrói o XML final
      const xml: string = `
                  <balancetes-estatisticos xmlns="http://arquivosemestral.xml.modelo.comum.estatistico.dataprev.gov.br">
                      <entidade>2083</entidade>
                      <ano>${ano}</ano>
                      <semestre>${semestre}</semestre>
                      <email>${email}</email>
                          ${balancete_estatistico.join("")}
              
                  </balancetes-estatisticos>
        
                  `
      return xml
  
  
    }
  
    function consolidado_plano_estatistico_xml(mapa_consolidado_mes: Map<string, MovimentacaoConsolidado[]>, ) {
      let lista_balancete_estatistico_xml: string[] = [];
      let planos_beneficios_mes = filtra_plano_beneficio_cnpb();
  
      mapa_consolidado_mes.forEach((valor, mes, map) => {
        lista_balancete_estatistico_xml.push(`
                                <balancete-estatistico mes="${mes}">
                                <consolidado>
                                ${valor.map(mov => {
          if (mov.observacao === null || mov.observacao === undefined || mov.observacao.trim() === "") {
            return `
                                  <movimentacao codigo-beneficio="${mov.codigo_beneficio}">
                                      <inicial>${mov.inicial}</inicial>
                                      <entradas>${mov.entrada}</entradas>
                                      <saidas>${mov.saida}</saidas>
                                  </movimentacao>`
          }
          return `
                                      <movimentacao codigo-beneficio="${mov.codigo_beneficio}">
                                            <inicial>${mov.inicial}</inicial>
                                            <entradas>${mov.entrada}</entradas>
                                            <saidas>${mov.saida}</saidas>
                                            <observacao>${mov.observacao}</observacao>
                                        </movimentacao>`
        }).join("")}
                                </consolidado>
                                ${create_cnpb_child(planos_beneficios_mes.get(mes)!).join("")}
                    </balancete-estatistico> 
                                `
  
        );
      });
      return lista_balancete_estatistico_xml;
    };
  
    function filtra_plano_beneficio_cnpb() {
      const codigo_plano_cnpb = ["1995002356", "2000008283", "1999005211", "1973000156", "2012000274", "2019002329", "2020000229", "2020001438", "2011002192", "2011002265"]
      let lista_filtrada: Movimentacao[] = []
      let mapa_mes_beneficio = new Map<string, Map<string, Movimentacao[]>>()
      for (let m = mes_inicial; m <= mes_final; m++) {
        mapa_mes_beneficio.set(m.toString(), new Map<string, Movimentacao[]>())
        for (let codigo_plano of codigo_plano_cnpb) {
          mapa_mes_beneficio.get(m.toString())?.set(codigo_plano, [])
          lista_filtrada = list_movimentacao.filter(movimento => movimento.codigo_plano.toString() === codigo_plano && movimento.mes == m.toString())
          mapa_mes_beneficio.get(m.toString())?.get(codigo_plano)?.push(...lista_filtrada)
          lista_filtrada.forEach(mov => {
            if (mov.codigo_beneficio == "11000") {
              let codigoBeneficio3200 = lista_filtrada.find(mov => mov.codigo_beneficio == '32000')
              if (codigoBeneficio3200) {
                codigoBeneficio3200.entrada = mov.entrada;
                codigoBeneficio3200.saida = mov.saida;
                codigoBeneficio3200.inicial = mov.inicial;
                codigoBeneficio3200.valor_final = mov.valor_final;
              }
            }
          })
        }
      }
  
      return mapa_mes_beneficio
    }
  
    function create_cnpb_child(movimentacao_beneficio: Map<string, Movimentacao[]>): string[] {
      let lista_beneficio_movimentacao: string[] = []
      movimentacao_beneficio.forEach((valor, chave, map) => {
        lista_beneficio_movimentacao.push(`<plano-beneficio cnpb="${chave}">
                         ${valor.map(mov => {
          if (mov.observacao === null || mov.observacao === undefined || mov.observacao.trim() === "") {
            return `
                                <movimentacao codigo-beneficio="${mov.codigo_beneficio}">
                                      <inicial>${mov.inicial}</inicial>
                                      <entradas>${mov.entrada}</entradas>
                                      <saidas>${mov.saida}</saidas>
                                  </movimentacao>`
          } return `
                                      <movimentacao codigo-beneficio="${mov.codigo_beneficio}">
                                            <inicial>${mov.inicial}</inicial>
                                            <entradas>${mov.entrada}</entradas>
                                            <saidas>${mov.saida}</saidas>
                                            <observacao>${mov.observacao}</observacao>
                                        </movimentacao>`
        }).join("")}   
                          </plano-beneficio>`)
      });
      return lista_beneficio_movimentacao
  
  
    }
    function criarOuLimparPlanilhaErros() {
  
      let existingWorksheet: ExcelScript.Worksheet | null = null;
      try {
        existingWorksheet = workbook.getWorksheet("erros");
      } catch (error) {
        console.log("A planilha buscada existe")
      }
  
      
      if (existingWorksheet) {
        existingWorksheet.delete()
      } 
  
        // A planilha "Abaxml" não existe, crie uma nova planilha
        let newWorksheet_xml: ExcelScript.Worksheet = workbook.addWorksheet("erros");
        let coluna_erros: ExcelScript.Range = newWorksheet_xml.getRange("A1");
        coluna_erros.setFormulaLocal("Mês");
        coluna_erros = newWorksheet_xml.getRange("B1");
        coluna_erros.setFormulaLocal("Código Plano");
        coluna_erros = newWorksheet_xml.getRange("C1");
        coluna_erros.setFormulaLocal("Código Benefício");
        coluna_erros = newWorksheet_xml.getRange("D1");
        coluna_erros.setFormulaLocal("Erros");
  
      
    }
  
    function criaTabelaXML(xmlContent: string, cellsPerRow: number) {
      let existingWorksheet: ExcelScript.Worksheet | null = null;
  
      try {
        existingWorksheet = workbook.getWorksheet("tabelaxml");
      } catch (error) {
        console.log("A planilha não existe")
      }
  
      if (existingWorksheet) {
        existingWorksheet.delete()
      } 
  
        let newWorksheet_xml: ExcelScript.Worksheet = workbook.addWorksheet("tabelaxml");
        const xmlParts = xmlContent.match(new RegExp(`.{1,${cellsPerRow}}`, 'g'));
  
        // Preencha as células na coluna A com partes do XML
        xmlParts.forEach((part, index) => {
          const cell = newWorksheet_xml.getRange(`A${index + 1}`);
          cell.setValue(part);
        });
  
      
  
    }
  
    function confere_se_a_regra_e_seguida(mapa_consolidado_mes: Map<string, MovimentacaoConsolidado[]>, mapa_mes_beneficio: Map<string, Map<string, Movimentacao[]>>) {
      let lista_erros: string[] = [];
      let movimentacoes_erro: MovimentacaoErro[] = []
  
  
      mapa_consolidado_mes.forEach((movimentacoes, mes) => {
        let mes_atual = movimentacoes;
        let proximo_mes = Number.parseInt(mes) + 1;
        if (proximo_mes == 13) {
          return;
        }
        let movimentos_do_proximo_mes = mapa_consolidado_mes.get(proximo_mes.toString());
  
        let mes_atual_filtrado = mes_atual.forEach((movimentacao) => {
          let codigo = movimentacao.codigo_beneficio;
          let valor_inicial = movimentacao.inicial;
          let valor_entrada = movimentacao.entrada;
          let valor_saida = movimentacao.saida;
          let valor_final = movimentacao.final;
          let mes = movimentacao.mes;
  
  
          if (codigo == "11000") {
            let codigo_11100 = mes_atual.find(mov => mov.codigo_beneficio == "11100")
            let codigo_11200 = mes_atual.find(mov => mov.codigo_beneficio == "11200")
  
            if (codigo_11100 && codigo_11200) {
              let total = codigo_11100.final + codigo_11200.final
              if (valor_final != total) {
                let erro = `No mês ${mes}, a soma dos valores finais código de benefício 11100 (${codigo_11100.final}) + 11200 (${codigo_11200.final}) = total (${total}) deveria ser igual ao valor de 11000 (${valor_final}).`
                lista_erros.push(erro);
              }
            }
          }
  
  
          if (codigo == "14000") {
            let codigo_33000 = mes_atual.find(mov => mov.codigo_beneficio == '33000');
  
            // Verifica se encontrou o código 33000
            if (codigo_33000) {
              if (valor_inicial > codigo_33000.inicial) {
                let erro = `Valor inicial do código de benefício (14000) é maior que valor inicial do código de benefício (33000) para o mês (${mes}) : (${valor_inicial} > ${codigo_33000.inicial})`;
                lista_erros.push(erro)
              }
  
              if (valor_entrada > codigo_33000.entrada) {
                let erro = `Valor de entrada do código de benefício (14000) é maior que valor entrada do código de benefício (33000) para o mês (${mes}) : (${valor_entrada} > ${codigo_33000.entrada})`;
                lista_erros.push(erro)
              }
  
              if (valor_saida > codigo_33000.saida) {
                let erro = `Valor de saida do código de benefício (14000) é maior que valor saida do código de benefício (33000) para o mês (${mes}) : (${valor_saida} > ${codigo_33000.saida})`;
                lista_erros.push(erro)
              }
  
              if (valor_final > codigo_33000.final) {
                let erro = `O valor do código de benefício (14000) é maior que valor final do código de benefício (33000) para o mês (${mes})`;
                lista_erros.push(erro)
              }
              if (valor_final == 0 && codigo_33000.final != 0) {
                let erro = `O valor final do código de benefício (14000) é igual a 0, e o código de benefício (33000) é diferente 0, ambos precisam ter o valor 0. Verifique no mês: ${mes}`
                lista_erros.push(erro);
              }
              if (valor_final != 0 && codigo_33000.final == 0) {
                let erro = `O valor final do código de benefício (14000) é diferente de 0, e o código de benefício (33000) é igual a 0, ambos precisam ter o valor 0. Verifique no mês: ${mes}`
                lista_erros.push(erro);
              }
              else {
                return
              }
            }
          }
  
  
  
          if (codigo == "31000") {
            let codigo_31100 = mes_atual.find(mov => mov.codigo_beneficio == '31100')
            let codigo_31200 = mes_atual.find(mov => mov.codigo_beneficio == '31200')
            let codigo_31300 = mes_atual.find(mov => mov.codigo_beneficio == '31300')
  
  
            if (codigo_31100 && codigo_31200 && codigo_31300) {
              let total_dos_codigos = codigo_31100.final + codigo_31200.final + codigo_31300.final;
  
              if (valor_final !== total_dos_codigos) {
                let erro = `O valor código de benefício 31000 (${valor_final}) precisa ser igual ao somatório dos valores finais entre 31100 (${codigo_31100.final}),31200 (${codigo_31200.final}) e 31300 (${codigo_31300.final}) = ${total_dos_codigos}, verifique no mês:${mes}`
                lista_erros.push(erro);
              }
  
  
            }
          }
  
  
        });
  
      });
  
  
      mapa_mes_beneficio.forEach((mapa_beneficio, mes) => {
        mapa_beneficio.forEach((movimentacoes, codigo_plano) => {
          movimentacoes.forEach(mov => {
            let mes_atual = mes;
            let mes_seguinte = Number.parseInt(mes) + 1;
            if (mes_seguinte == 13) {
              return
            }
  
            let movimentacoes_plano_proximo_mes = mapa_mes_beneficio.get(mes_seguinte.toString())
            if (movimentacoes_plano_proximo_mes) {
  
              let movimentos_do_proximo_mes = movimentacoes_plano_proximo_mes?.get(codigo_plano);
  
              let movimento_mes_seguinte = movimentos_do_proximo_mes?.find(movimento_proximo_mes => movimento_proximo_mes.codigo_beneficio == mov.codigo_beneficio);
  
              if (movimento_mes_seguinte) {
                if (mov.valor_final != movimento_mes_seguinte.inicial) {
  
                  let erro = `Há uma inconsistência na comparação entre os meses ${mov.mes} e ${movimento_mes_seguinte.mes}, relacionada ao código do plano ${mov.codigo_plano} e código do benefício ${mov.codigo_beneficio}. O valor final do mês ${mov.mes} (${mov.valor_final}) difere do valor inicial do mês ${movimento_mes_seguinte.mes} (${movimento_mes_seguinte.inicial})`
                  movimentacoes_erro.push(new MovimentacaoErro(erro, mov.mes, mov.codigo_beneficio, mov.codigo_plano))
  
                }
              }
            }
  
  
          });
  
  
        });
  
      });
  
  
  
      let existingWorksheet = workbook.getWorksheet("erros")
  
  
      for (let indice in movimentacoes_erro) {
  
        let erro_movimentacao = movimentacoes_erro[indice]
  
        let int_index = Number.parseInt(indice);
  
  
        let cell_erros = existingWorksheet.getRange(`A${int_index + 2}`)
        cell_erros.setValue(erro_movimentacao.mes)
  
  
        cell_erros = existingWorksheet.getRange(`B${int_index + 2}`)
        cell_erros.setValue(erro_movimentacao.codigo_plano)
  
  
        cell_erros = existingWorksheet.getRange(`C${int_index + 2}`)
        cell_erros.setValue(erro_movimentacao.codigo_beneficio)
  
  
        cell_erros = existingWorksheet.getRange(`D${int_index + 2}`)
        cell_erros.setValue(erro_movimentacao.erro)
  
  
  
      }
  
      for (let i = 0; i <= lista_erros.length; i++) {
        let existingWorksheet = workbook.getWorksheet("erros")
        let cell_erros = existingWorksheet.getRange(`E${i + 2}`)
        cell_erros.setValue(lista_erros[i])
  
      }
  
  
  
  
  
    }
  
    criarOuLimparPlanilhaErros()
    
  
    let mapa_mes_beneficio = filtra_plano_beneficio_cnpb()
    let mapa_consolidado = filtra_codigo_por_mes_e_plano_consolidado()
    let balancete_estatistico = consolidado_plano_estatistico_xml(mapa_consolidado)
    let xmlFullBody = root_father_xml(balancete_estatistico)
    confere_se_a_regra_e_seguida(mapa_consolidado, mapa_mes_beneficio)
    criaTabelaXML(xmlFullBody, 500)
  
  
  }
  
  
  function main(workbook: ExcelScript.Workbook) {
      let hora_inicio = new Date()
      let lista_Movimentacao_Base: Movimentacao_Base[] = [];
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
        transformTable(mes_atual, usedRange, lista_Movimentacao_Base, lista_planos, lista_codigo_plano, ano, semestre);
        
        
      }  
      
      console.log(`${new Date().toISOString()} terminou preenchimento ${new Date() - hora_inicio}`)
  
      hora_inicio = new Date()
      Reune_Dados_Base_Auxiliar(lista_Movimentacao_Base,workbook)
  
      console.log(`${new Date().toISOString()} finalizou a criacao aux ${new Date() - hora_inicio}`)
  
      cria_xml(workbook)
      
  }
  
  
  