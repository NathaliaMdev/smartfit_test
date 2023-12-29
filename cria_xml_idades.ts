
class MovimentacaoPorBeneficio {
    codigo_plano: string;
    codigo_beneficio: string;
    faixa_etaria: string;
    valor_masculino: string;
    valor_feminino: string;
    mes: string;
    ano: string;
  
    constructor(codigo_plano: string) {
      this.codigo_plano = codigo_plano;
      this.codigo_beneficio = "";
      this.valor_masculino = ""
      this.valor_feminino = "";
    }
  }
  
  class MovimentacaoConsolidado {
  
    codigo_beneficio: string;
    faixa_etaria: string;
    valor_masculino: number;
    valor_feminino: number;
    mes: string;
    ano: string;
  
  
    constructor(codigo_beneficio: string, faixa_etaria: string) {
      this.codigo_beneficio = codigo_beneficio;
      this.faixa_etaria = faixa_etaria;
      this.valor_masculino = 0;
      this.valor_feminino = 0;
  
    }
  
  }
  
  class Movimentacao {
    //Declara o tipo das variáveis serão colocadas no mapa 
    codigo_beneficio: string;
    descricao_beneficio: string;
    plano: string;
    codigo_plano: string;
    codigo_faixa_etaria: string;
    faixa_etaria: string;
    valor_masculino: string;
    valor_feminino: string;
    mes: string;
    ano: string;

  
  
    //Cria um construtor, que fornecerá acesso as variáveis e a posição da linha. 
    constructor(arr_movimentacao: string | number[], numero_linha) {
  
      this.codigo_beneficio = arr_movimentacao[0].toString();
      this.descricao_beneficio = arr_movimentacao[1].toString();
      this.plano = arr_movimentacao[2].toString();
      this.codigo_plano = arr_movimentacao[3].toString();
      this.codigo_faixa_etaria = arr_movimentacao[4].toString();
      this.faixa_etaria = arr_movimentacao[5].toString();
      this.valor_masculino = arr_movimentacao[6].toString();
      this.valor_feminino = arr_movimentacao[7].toString();
      this.ano = arr_movimentacao[8].toString();
      this.mes = arr_movimentacao[9].toString();

  
    }
  
  }
  
  function main(workbook: ExcelScript.Workbook) {
    CriaOuLimpaPlanilhaXML()
    const base_auxiliar: ExcelScript.Worksheet = workbook.getWorksheet("Base Auxiliar")
    
  
    let lista_movimentacao: Movimentacao[] = [];
    let primeira_linha = true;
    let usedRange = base_auxiliar.getUsedRange().getValues();
  
    usedRange.forEach((linha, index) => {
      if (primeira_linha) {
        primeira_linha = false;
        return
      }
      // Cria instâncias da classe Movimentacao e adiciona à lista
      lista_movimentacao.push(new Movimentacao(linha, index))
  
    });
  
    function preenche_mapa_consolidado_faixaetaria(mapa_consolidado_por_codigo_beneficio: Map<string, MovimentacaoConsolidado[]>, codigo_beneficio: string, lista_numeros_faixaetaria: string[]): void {
      mapa_consolidado_por_codigo_beneficio.set(codigo_beneficio, []);
      for (let faixa_etaria of lista_numeros_faixaetaria) {
        let lista_movimentacao_por_codigos_faixa_etaria = lista_movimentacao.filter(mov => mov.codigo_faixa_etaria === faixa_etaria && mov.codigo_beneficio === codigo_beneficio)
        let movimento_consolidado = new MovimentacaoConsolidado(codigo_beneficio, faixa_etaria)
        lista_movimentacao_por_codigos_faixa_etaria.forEach(mov => {
          movimento_consolidado.mes = mov.mes;
          movimento_consolidado.ano = mov.ano;
          movimento_consolidado.faixa_etaria = mov.faixa_etaria;
          movimento_consolidado.valor_feminino += Number.parseInt(mov.valor_feminino);
          movimento_consolidado.valor_masculino += Number.parseInt(mov.valor_masculino);
        });
        mapa_consolidado_por_codigo_beneficio.get(codigo_beneficio)!.push(movimento_consolidado)
      }
    }
  
    function filtra_por_codigo_de_beneficio_e_faixa_etaria_consolidado() {
      let mapa_consolidado_por_codigo_beneficio = new Map<string, MovimentacaoConsolidado[]>()
  
  
      preenche_mapa_consolidado_faixaetaria(mapa_consolidado_por_codigo_beneficio,
        "31000",
        ["41100", "41200", "41300", "41400", "41500", "41600", "41700"]
      )
  
      preenche_mapa_consolidado_faixaetaria(mapa_consolidado_por_codigo_beneficio,
        "32000",
        ["42100", "42200", "42300", "42400", "42500", "42600", "42700"]
      )
  
      preenche_mapa_consolidado_faixaetaria(mapa_consolidado_por_codigo_beneficio,
        "33000",
        ["43100", "43200", "43300", "43400", "43500", "43600", "43700"]
      )
  
      return mapa_consolidado_por_codigo_beneficio
  
    }
  
  
    function consolidado_faixas_etarias() {
      let mapa_consolidado = filtra_por_codigo_de_beneficio_e_faixa_etaria_consolidado()
      let lista_consolidado_xml: string[] = [];
      mapa_consolidado.forEach((valor, codigo_beneficio, map) => {
  
        lista_consolidado_xml.push(`
                    <populacao-beneficio codigo-beneficio="${codigo_beneficio}">
                    ${valor.map(mov => {
          return `
                        <faixa-etaria tipo = "${mov.faixa_etaria}">
                            <masculino>${mov.valor_masculino}</masculino>
                            <feminino>${mov.valor_feminino}</feminino>
                        </faixa-etaria>`
        }).join("")}
                  </populacao-beneficio>`
        )
  
      });
      return lista_consolidado_xml.join("")
    }
  
  
    function filtra_por_plano_beneficio_cnpb() {
      const lista_de_codigo_plano = ["2020001438", "1973000156", "1995002356", "2011002192", "2011002265", "2020000229", "2012000274", "1999005211", "2000008283", "2019002329"];
      const codigo_beneficio_numeros = ["31000", "32000", "33000"]
      let lista_filtrada: Movimentacao[] = []
      let mapa_filtrado_por_plano_beneficio = new Map<string, Map<string, Movimentacao[]>>()
      //let codigo_beneficio of codigo_beneficio_numeros
      for (let codigo_plano of lista_de_codigo_plano) {
        mapa_filtrado_por_plano_beneficio.set(codigo_plano, new Map<string, Movimentacao[]>())
        for (let codigo_beneficio of codigo_beneficio_numeros) {
          mapa_filtrado_por_plano_beneficio.get(codigo_plano)?.set(codigo_beneficio, []);
          lista_filtrada = lista_movimentacao.filter(mov => mov.codigo_plano === codigo_plano && mov.codigo_beneficio === codigo_beneficio)
          mapa_filtrado_por_plano_beneficio.get(codigo_plano)?.get(codigo_beneficio)?.push(...lista_filtrada)
        }
      }
      return mapa_filtrado_por_plano_beneficio
    }
  
  
    function create_create_cnpb_child() {
      let beneficio_filtrado = filtra_por_plano_beneficio_cnpb()
      let lista_movimentacao: string[] = [];
  
      beneficio_filtrado.forEach((mapa_plano_beneficio, codigo_plano, mapa) => {
        lista_movimentacao.push(`
          <plano-beneficio cnpb="${codigo_plano}">
            ${Array.from(mapa_plano_beneficio).map(([codigo_beneficio, movimentacao]) => `
              <populacao-beneficio codigo-beneficio="${codigo_beneficio}">
                ${movimentacao.map(mov => `
                  <faixa-etaria tipo="${mov.faixa_etaria}">
                    <masculino>${mov.valor_masculino}</masculino>
                    <feminino>${mov.valor_feminino}</feminino>
                  </faixa-etaria>
                `).join('')}
              </populacao-beneficio>
            `).join('')}
          </plano-beneficio>
        `);
      });
  
      return lista_movimentacao.join("");
    }
  
  
  
    function root_father_xml(email:string,competencia:string) {
      let cnpbs = create_create_cnpb_child()
      let consolidados = consolidado_faixas_etarias()
      

      const xml: string = `
                  <balancete-sexo-idade xmlns="http://sexoidadeporplano.xml.modelo.comum.estatistico.dataprev.gov.br">
                      <entidade>2083</entidade>
                      <competencia>${competencia}</competencia>
                      <email>${email}</email>
                      <consolidado>
                          ${consolidados}
                      </consolidado>
                         ${cnpbs}
                  </balancete-sexo-idade>
                  `
      return xml
    }
  
  
    function CriaOuLimpaPlanilhaXML() {
      let existingWorksheet: ExcelScript.Worksheet | null = null;
  
      try {
        existingWorksheet = workbook.getWorksheet("tabelaxml");
      } catch (error) {
        console.log("A planilha não existe")
      }
      if (existingWorksheet) {
        let range_used = existingWorksheet.getUsedRange()
        if (range_used) {
          range_used.delete(Excel.DeleteShiftDirection.up);
        }
      } else {
        // A planilha "Abaxml" não existe, crie uma nova planilha
        let newWorksheet_xml: ExcelScript.Worksheet = workbook.addWorksheet("tabelaxml");
      }
    }
  
    function criaTabelaXML() {
      const worksheet:ExcelScript.Worksheet = workbook.getWorksheets()[0];
      const email_celula: ExcelScript.Range = worksheet.getRange("D2");
      const email_range: ExcelScript.Range = email_celula.getUsedRange();
      const email = email_range.getValues()[0][0];
      const competencia_celula: ExcelScript.Range = worksheet.getRange("D3");
      const competencia_range: ExcelScript.Range = competencia_celula.getUsedRange();
      const competencia = competencia_range.getValues()[0][0]; 

      let xml = root_father_xml(email,competencia);
      let cellsPerRow = 500;
  
      let existingWorksheet: ExcelScript.Worksheet | null = null;
  
      try {
        existingWorksheet = workbook.getWorksheet("tabelaxml");
      } catch (error) {
        console.log("A planilha não existe")
      }
  
      if (existingWorksheet) {
        // A planilha "Abaxml" já existe, use a mesma aba e sobrescreva a célula A1
        const xmlParts = xml.match(new RegExp(`.{1,${cellsPerRow}}`, 'g'));
  
        // Preencha as células na coluna A com partes do XML
        xmlParts.forEach((part, index) => {
          const cell = existingWorksheet.getRange(`A${index + 1}`);
          cell.setValue(part);
        });
  
  
      } else {
        // A planilha "Abaxml" não existe, crie uma nova planilha
        let newWorksheet_xml: ExcelScript.Worksheet = workbook.addWorksheet("tabelaxml");
        const xmlParts = xml.match(new RegExp(`.{1,${cellsPerRow}}`, 'g'));
  
        // Preencha as células na coluna A com partes do XML
        xmlParts.forEach((part, index) => {
          const cell = newWorksheet_xml.getRange(`A${index + 1}`);
          cell.setValue(part);
        });
  
      }
    }
  
    criaTabelaXML()
  
  }
  
  
  
  
  
  
  
  
  
  