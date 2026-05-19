const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  PageNumber,
  LevelFormat,
  TabStopType,
} = require("docx");
const fs = require("fs");
const path = require("path");

const BLUE_DARK = "1A3A6B";
const BLUE_MID = "2E5FA3";
const BLUE_LIGHT = "D0DCF0";
const GRAY_LIGHT = "F2F4F8";
const GRAY_MID = "D9DDE6";
const AMBER = "8B5E00";
const AMBER_BG = "FFF8E7";
const RED = "8B1A1A";
const RED_BG = "FFF0F0";
const GREEN = "1A5C2A";
const GREEN_BG = "F0FFF4";
const WHITE = "FFFFFF";
const TEXT_MAIN = "1A1A2E";
const TEXT_MUTED = "4A4A6A";

const border = { style: BorderStyle.SINGLE, size: 1, color: GRAY_MID };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE_MID, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: BLUE_DARK })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: BLUE_MID })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 80, after: 80, line: 320 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: TEXT_MAIN, ...opts })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60, line: 300 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: TEXT_MAIN })],
  });
}

function space(n = 1) {
  return new Paragraph({ children: [new TextRun({ text: "", size: n === 1 ? 12 : 20 })] });
}

function alertBox(title, lines, fillColor, titleColor) {
  const cellChildren = [
    new Paragraph({
      spacing: { before: 60, after: 80 },
      children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: titleColor })],
    }),
    ...lines.map(
      (line) =>
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { before: 40, after: 40 },
          children: [new TextRun({ text: line, font: "Arial", size: 20, color: TEXT_MAIN })],
        }),
    ),
  ];

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: fillColor, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 180, right: 180 },
            children: cellChildren,
          }),
        ],
      }),
    ],
  });
}

function sectionTable(rows) {
  const colW = [3000, 6360];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colW,
    rows: rows.map(
      (row, index) =>
        new TableRow({
          children: [
            new TableCell({
              borders,
              width: { size: colW[0], type: WidthType.DXA },
              shading: { fill: index === 0 ? BLUE_DARK : BLUE_LIGHT, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: row[0],
                      font: "Arial",
                      size: 20,
                      bold: true,
                      color: index === 0 ? WHITE : BLUE_DARK,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              borders,
              width: { size: colW[1], type: WidthType.DXA },
              shading: { fill: index === 0 ? GRAY_LIGHT : WHITE, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: row[1],
                      font: "Arial",
                      size: 20,
                      color: TEXT_MAIN,
                      bold: index === 0,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
    ),
  });
}

function reqTable(items) {
  const colW = [400, 5760, 3200];
  const header = new TableRow({
    tableHeader: true,
    children: ["No", "Requisito / Questionamento", "Base Legal / Normativa"].map(
      (heading, cellIndex) =>
        new TableCell({
          borders,
          width: { size: colW[cellIndex], type: WidthType.DXA },
          shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: heading, font: "Arial", size: 20, bold: true, color: WHITE }),
              ],
            }),
          ],
        }),
    ),
  });

  const dataRows = items.map(
    (item, index) =>
      new TableRow({
        children: [
          new TableCell({
            borders,
            width: { size: colW[0], type: WidthType.DXA },
            shading: { fill: index % 2 === 0 ? GRAY_LIGHT : WHITE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: String(index + 1).padStart(2, "0"),
                    font: "Arial",
                    size: 20,
                    bold: true,
                    color: BLUE_MID,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders,
            width: { size: colW[1], type: WidthType.DXA },
            shading: { fill: index % 2 === 0 ? GRAY_LIGHT : WHITE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: item[0], font: "Arial", size: 20, color: TEXT_MAIN })],
              }),
            ],
          }),
          new TableCell({
            borders,
            width: { size: colW[2], type: WidthType.DXA },
            shading: { fill: index % 2 === 0 ? GRAY_LIGHT : WHITE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: item[1],
                    font: "Arial",
                    size: 19,
                    color: TEXT_MUTED,
                    italics: true,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
  );

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colW,
    rows: [header, ...dataRows],
  });
}

const blocks = {
  classificacao: [
    ["Qual e o material base do cateter: latex, PVC, silicone ou poliuretano?", "RDC 185/2001; norma tecnica aplicavel ao material"],
    ["A impregnacao por NPAg e apresentada como inovacao tecnologica principal?", "RDC 185/2001 - dossie tecnico completo"],
    ["Ha precedente de registro na ANVISA para produto similar?", "Consulta a base ANVISA e DATAVISA"],
    ["O requerente pretende registro por equivalencia tecnica? Justificar a impossibilidade quando houver tecnologia nova.", "RDC 185/2001; RDC 56/2001"],
    ["Ha nota tecnica GGTPS/ANVISA ou orientacao vigente para nanotecnologia em dispositivos medicos aplicavel ao caso?", "ISO/TR 10993-22; orientacoes GGTPS"],
  ],
  caracterizacao: [
    ["Qual e o tamanho medio das NPAg e o indice de polidispersao? Fornecer distribuicao granulometrica completa.", "ISO/TR 10993-22; FDA Nanotechnology Guidance 2014"],
    ["Qual e a morfologia das nanoparticulas? Apresentar imagens de microscopia eletronica de transmissao.", "ISO/TR 10993-22:2017, item 6.2"],
    ["Qual e a concentracao superficial de prata e a concentracao total por unidade do produto?", "RDC 56/2001 - Requisito essencial 10"],
    ["Qual e o metodo de impregnacao: coating superficial, incorporacao na matriz polimerica ou conjugacao quimica?", "RDC 185/2001 - descricao tecnologica"],
    ["Apresentar cinetica de liberacao de ions Ag+ in vitro em simulador de urina nos tempos 1h, 24h, 72h, 7 dias e ate o fim da vida util.", "ISO/TR 10993-22; ASTM E2696"],
    ["Apresentar estabilidade das nanoparticulas apos esterilizacao por EtO, radiacao gama ou vapor umido.", "ISO 10993-12:2021; RDC 16/2013"],
    ["Qual e o potencial zeta das NPAg e qual a estabilidade coloidal residual?", "ISO/TR 10993-22:2017"],
    ["Ha agentes estabilizantes como citrato, PVP ou PEG? Apresentar biocompatibilidade individual.", "ISO 10993-1:2018"],
  ],
  biocompatibilidade: [
    ["Citotoxicidade com linhagem celular urotelial humana, nao apenas fibroblastos L929.", "ISO 10993-5:2009"],
    ["Sensibilizacao por GPMT ou ensaio de Buehler.", "ISO 10993-10:2021"],
    ["Irritacao e reatividade intracutanea.", "ISO 10993-10; ISO 10993-23"],
    ["Toxicidade sistemica aguda.", "ISO 10993-11:2017"],
    ["Toxicidade subcronica ou subaguda, essencial para uso repetido.", "ISO 10993-11:2017"],
    ["Genotoxicidade: Ames, micronucleo in vitro e micronucleo in vivo.", "ISO 10993-3:2014"],
    ["Implantacao, se houver potencial de permanencia tecidual de NPAg.", "ISO 10993-6:2016"],
    ["Toxicocinetica de nanoparticulas com distribuicao e acumulo em orgaos-alvo, especialmente rins.", "ISO/TR 10993-22:2017, item 8"],
    ["Avaliacao de argiria local e sistemica por deposicao de prata, incluindo estudo de biodistribuicao.", "ISO/TR 10993-22; literatura toxicologica"],
    ["Confirmar se os testes foram feitos com o produto final apos esterilizacao e nao apenas com NPAg isoladas.", "ISO 10993-12:2021"],
  ],
  antimicrobiana: [
    ["Quais sao os claims exatos: reduz colonizacao bacteriana, previne ITU ou possui acao antimicrobiana?", "RDC 185/2001, Anexo I"],
    ["Os testes foram realizados na superficie do cateter ou apenas com extratos?", "ASTM E2180; ISO 22196 adaptado"],
    ["Foi realizado teste de aderencia bacteriana em superficie plastica flexivel?", "ISO 22196 adaptado; ASTM E2180"],
    ["O painel contempla E. coli, Klebsiella pneumoniae, Pseudomonas aeruginosa, Enterococcus faecalis, Candida albicans e Staphylococcus epidermidis?", "CLSI; ABNT NBR ISO 20776"],
    ["Foram realizados testes de formacao e inibicao de biofilme?", "ASTM E2196; protocolo CDC Biofilm"],
    ["Qual e a duracao da atividade antimicrobiana ao longo da vida util do produto?", "RDC 56/2001 - requisito essencial de desempenho"],
    ["Foram feitos estudos de resistencia bacteriana apos exposicao subinibitoria as NPAg?", "OMS - resistencia antimicrobiana; literatura NPAg"],
    ["Existe estudo clinico prospectivo, randomizado e controlado demonstrando reducao de ITU?", "CONSORT 2010; RDC 204/2017"],
  ],
  bpf: [
    ["Confirmar conformidade com ABNT NBR ISO 8669-2 ou norma especifica aplicavel ao tipo de sonda.", "ABNT NBR ISO 8669-2"],
    ["Documentar requisitos dimensionais, resistencia a tracao, vedacao do balao se aplicavel e ausencia de rebarbas.", "ABNT NBR ISO 8669-2"],
    ["Descrever processo de esterilizacao e impacto comprovado sobre as NPAg.", "ISO 11135; ISO 11137; RDC 16/2013"],
    ["Apresentar estudo de estabilidade pos-esterilizacao com tamanho, concentracao, morfologia e atividade antimicrobiana residual.", "ISO 10993-12; RDC 16/2013"],
    ["Apresentar CBPF vigente do fabricante, a ser verificado no DATAVISA.", "RDC 16/2013"],
    ["Apresentar prazo de validade e estudo de estabilidade acelerada e em tempo real.", "RDC 185/2001; ASTM F1980"],
    ["Monitorar concentracao superficial de prata e possivel oxidacao a Ag2O ao longo da validade.", "RDC 56/2001; estabilidade de nanomateriais"],
    ["Demonstrar rastreabilidade de lote das NPAg, fornecedor e certificado de analise.", "RDC 16/2013 - rastreabilidade"],
  ],
  rotulagem: [
    ["Rotulagem deve mencionar explicitamente a presenca de nanoparticulas de prata.", "RDC 185/2001; RDC 36/2015"],
    ["IFU deve incluir populacao-alvo, contraindicacao para alergia a prata e tempo maximo de permanencia.", "RDC 185/2001; RDC 36/2015"],
    ["Incluir advertencias sobre descarte adequado por se tratar de resíduo de servico de saude com componente nanomaterial.", "CONAMA 358/2005; RDC ANVISA 222/2018"],
    ["Apresentar plano de tecnovigilancia pos-registro com monitoramento ativo de eventos adversos.", "RDC 67/2009; RDC 204/2017"],
    ["Monitorar hipersensibilidade a prata, argiria local e eventos em pacientes renais, oncologicos ou imunossuprimidos.", "Tecnovigilancia ANVISA"],
  ],
};

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: TEXT_MAIN } } },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: BLUE_MID },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1260, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Table({
              width: { size: 9206, type: WidthType.DXA },
              columnWidths: [6000, 3206],
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      borders: noBorders,
                      width: { size: 6000, type: WidthType.DXA },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "AGENCIA NACIONAL DE VIGILANCIA SANITARIA",
                              font: "Arial",
                              size: 18,
                              bold: true,
                              color: BLUE_DARK,
                            }),
                          ],
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "Gerencia-Geral de Tecnologia de Produtos para Saude - GGTPS",
                              font: "Arial",
                              size: 16,
                              color: TEXT_MUTED,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      borders: noBorders,
                      width: { size: 3206, type: WidthType.DXA },
                      verticalAlign: VerticalAlign.CENTER,
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          children: [
                            new TextRun({
                              text: "DOCUMENTO CONFIDENCIAL",
                              font: "Arial",
                              size: 16,
                              bold: true,
                              color: RED,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: GRAY_MID, space: 2 } },
              tabStops: [{ type: TabStopType.RIGHT, position: 9206 }],
              children: [
                new TextRun({
                  text: "Parecer Tecnico Preliminar - Sonda Uretral com Nanoparticulas de Prata",
                  font: "Arial",
                  size: 16,
                  color: TEXT_MUTED,
                }),
                new TextRun({ text: "\tPag. ", font: "Arial", size: 16, color: TEXT_MUTED }),
                new PageNumber({ font: "Arial", size: 16, color: TEXT_MUTED }),
              ],
            }),
          ],
        }),
      },
      children: [
        new Paragraph({
          spacing: { before: 480, after: 120 },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "AGENCIA NACIONAL DE VIGILANCIA SANITARIA",
              font: "Arial",
              size: 36,
              bold: true,
              color: BLUE_DARK,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 240, after: 240 },
          children: [
            new TextRun({
              text: "PARECER TECNICO PRELIMINAR",
              font: "Arial",
              size: 32,
              bold: true,
              color: BLUE_DARK,
            }),
          ],
        }),
        sectionTable([
          ["Campo", "Informacao"],
          ["Tipo de documento", "Parecer Tecnico Preliminar - Uso Interno"],
          ["Produto avaliado", "Sonda Uretral de Alivio com Nanoparticulas de Prata"],
          ["Classificacao preliminar", "Produto para Saude - Artigo Medico-Hospitalar"],
          ["Classe de risco estimada", "Classe II / III - sujeito a confirmacao"],
          ["Referencia normativa principal", "RDC 185/2001 | RDC 56/2001 | RDC 27/2011"],
          ["Data de elaboracao", new Date().toLocaleDateString("pt-BR", { day: "2-digit", month: "long", year: "numeric" })],
        ]),
        space(2),
        alertBox(
          "Natureza do documento",
          [
            "Este parecer e de uso interno e destina-se a discussao tecnica com o fabricante/requerente.",
            "Nao constitui decisao administrativa definitiva e nao gera direitos subjetivos ao registro.",
            "Esta sujeito a alteracao apos analise documental completa do dossie tecnico.",
          ],
          AMBER_BG,
          AMBER,
        ),
        new Paragraph({ children: [new TextRun({ text: "" })], pageBreakBefore: true }),
        heading1("1. Sumario Executivo"),
        para("O presente parecer tecnico preliminar foi elaborado no ambito da analise de pedido de registro de sonda uretral de alivio impregnada com nanoparticulas de prata. O produto incorpora tecnologia inovadora que determina exigencias regulatorias mais abrangentes do que as aplicaveis a cateteres uretrais convencionais."),
        para("A presenca de nanoparticulas de prata como componente funcional ativo aciona, de forma cumulativa, os arcaboucos regulatorios de biocompatibilidade de nanomateriais, toxicologia, eficacia antimicrobiana e controle de qualidade de manufatura com insumo nanotecnologico."),
        alertBox(
          "Ponto critico de atencao",
          [
            "Dados de toxicidade de prata ionica convencional ou prata bulk nao substituem dados de nanoparticulas de prata.",
            "A extrapolacao de dados bibliograficos de outras formulacoes de NPAg nao e regulatoriamente admissivel.",
            "Todos os testes de biocompatibilidade devem ser realizados com o produto final apos esterilizacao.",
          ],
          RED_BG,
          RED,
        ),
        heading1("2. Enquadramento Regulatorio"),
        heading2("2.1 Base legal aplicavel"),
        sectionTable([
          ["Normativa", "Aplicacao ao produto"],
          ["RDC 185/2001", "Registro de produtos para saude - requisitos de dossie tecnico"],
          ["RDC 56/2001", "Requisitos essenciais de seguranca e eficacia"],
          ["RDC 16/2013", "Boas Praticas de Fabricacao"],
          ["ISO 10993-1:2018", "Avaliacao biologica de dispositivos medicos"],
          ["ISO/TR 10993-22:2017", "Orientacoes para avaliacao de nanomateriais"],
          ["ABNT NBR ISO 8669-2", "Cateteres uretrais - requisitos de desempenho e ensaios"],
        ]),
        heading2("2.2 Classificacao e perguntas iniciais ao requerente"),
        para("O produto e apresentado como sonda uretral de alivio impregnada com nanoparticulas de prata. A relatoria verifica inicialmente o material base do cateter, a existencia de precedentes de registro na ANVISA e a possibilidade de enquadramento como tecnologia nova, sem registro por simples equivalencia tecnica."),
        reqTable(blocks.classificacao),
        heading1("3. Bloco I - Caracterizacao das Nanoparticulas de Prata"),
        para("Este e o no central da analise. A variacao em parametros nanotecnologicos pode alterar completamente o perfil toxicologico, a reatividade biologica, a migracao, a liberacao ionica e a eficacia antimicrobiana das NPAg."),
        reqTable(blocks.caracterizacao),
        para("Pesquisa conduzida neste bloco: consulta a PubMed e Embase para comparar tamanhos e concentracoes de NPAg documentados em literatura. Nanoparticulas abaixo de 10 nm tendem a penetrar membranas celulares com maior facilidade, enquanto particulas acima de 100 nm podem se comportar mais proximamente a prata coloidal convencional. Essa diferenca e regulatoriamente relevante."),
        heading1("4. Bloco II - Biocompatibilidade e Toxicologia"),
        para("Este e o bloco mais critico. A prata nanoparticulada apresenta propriedades toxicologicas distintas da prata ionica convencional ou da prata metalica bulk. Por isso, nao se aceita extrapolacao direta de dados de prata nao-nano para NPAg."),
        para("Dado que a sonda uretral tem contato prolongado com mucosa urogenital e pode envolver sangue em pacientes cateterizados, aplica-se a matriz de contato da ISO 10993-1 para mucosa com duracao prolongada maior que 24 horas."),
        alertBox(
          "Exigencia inegociavel - ISO 10993-12",
          [
            "Todos os extratos devem ser preparados a partir do produto final apos esterilizacao.",
            "Dados de biocompatibilidade de NPAg isoladas nao substituem os dados do produto final.",
          ],
          AMBER_BG,
          AMBER,
        ),
        reqTable(blocks.biocompatibilidade),
        para("Pesquisa conduzida: consulta aos documentos tecnicos da ECHA sobre risco de nanoparticulas de prata, documentos da WHO sobre nanotecnologia e saude, guidance da FDA sobre nanotecnologia de 2014 e uso da ISO 10993-1 de 2020, alem de registros SCENIHR sobre prata nanoparticulada."),
        heading1("5. Bloco III - Eficacia Antimicrobiana"),
        para("A razao de existir das NPAg no produto e reduzir infeccao do trato urinario associada a cateter. Essa reivindicacao precisa ser sustentada com rigor cientifico e metodologico."),
        reqTable(blocks.antimicrobiana),
        para("Pesquisa conduzida: revisoes sistematicas e metanalises sobre cateteres impregnados com prata versus cateteres convencionais, consulta a Cochrane, NEJM, Lancet Infectious Diseases, diretrizes ABNT e CLSI, e verificacao de produtos com oxido de prata ou liga de prata ja registrados como precedentes."),
        heading1("6. Bloco IV - Desempenho do Dispositivo e BPF"),
        reqTable(blocks.bpf),
        heading1("7. Bloco V - Rotulagem e Vigilancia Pos-Mercado"),
        reqTable(blocks.rotulagem),
        heading1("8. Pesquisas e Consultas do Relator"),
        bullet("Consulta ao SINAES e ao banco de tecnovigilancia da ANVISA para eventos adversos com produtos similares."),
        bullet("Consulta ao DATAVISA para situacao do CBPF e verificacao de restricoes sanitarias."),
        bullet("Consulta ao FDA MAUDE e ao EUDAMED para recalls, alertas e seguranca pos-mercado de cateteres com prata."),
        bullet("Revisao de literatura dos ultimos 5 anos em PubMed e Embase com termos sobre silver nanoparticles, urinary catheter, safety e biocompatibility."),
        bullet("Verificacao de precedentes regulatorios internacionais e exigencias associadas a marcacao CE."),
        bullet("Consulta a GGMON e GGTPS para alinhamento interno sobre monitoramento e nanomateriais."),
        bullet("Consulta ao ISO/TR 10993-22:2017 como principal referencial tecnico internacional para nanomateriais em dispositivos medicos."),
        heading1("9. Sintese e Posicionamento do Relator"),
        alertBox(
          "Posicionamento tecnico",
          [
            "A presenca de NPAg em contato direto e prolongado com mucosa urogenital vascularizada eleva substancialmente o grau de exigencia regulatoria.",
            "O dossie padrao de sonda uretral convencional e necessario, mas insuficiente.",
            "O requerente deve apresentar modulo nanotoxicológico completo e independente, com dados proprios do produto final.",
            "Dados extrapolados de literatura ou de prata nao-nano nao sustentam a seguranca do produto.",
          ],
          GREEN_BG,
          GREEN,
        ),
        para("A existencia de nanoparticulas de prata em dispositivo que permanecera em contato direto e prolongado com mucosa urogenital, potencialmente em pacientes com comprometimento renal, oncologicos ou imunossuprimidos, exige abordagem regulatoria conservadora e documentacao propria do produto final."),
        para("A ANVISA tem o dever legal, nos termos da Lei 6.360/1976 e da Lei 9.782/1999, de assegurar que o produto seja seguro, eficaz e de qualidade antes de chegar ao paciente. A inovacao tecnologica nao reduz esse dever; ela o amplifica."),
        heading1("10. Referencias Normativas e Documentos-Base"),
        bullet("ANVISA. RDC 185/2001 - Registro de Produtos para Saude."),
        bullet("ANVISA. RDC 56/2001 - Requisitos Essenciais de Seguranca e Eficacia."),
        bullet("ANVISA. RDC 16/2013 - Boas Praticas de Fabricacao de Produtos Medicos."),
        bullet("ISO 10993-1:2018 - Biological evaluation of medical devices."),
        bullet("ISO/TR 10993-22:2017 - Guidance on nanomaterials."),
        bullet("ABNT NBR ISO 8669-2 - Cateteres uretrais."),
        space(2),
        new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: GRAY_MID, space: 8 } },
          spacing: { before: 240, after: 60 },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: `Brasilia, ${new Date().toLocaleDateString("pt-BR", { day: "2-digit", month: "long", year: "numeric" })}`,
              font: "Arial",
              size: 20,
              color: TEXT_MUTED,
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const outputPath = path.join(__dirname, "Parecer_Tecnico_ANVISA_Sonda_NPAg.docx");
  fs.writeFileSync(outputPath, buffer);
  console.log(`OK: ${outputPath}`);
});
