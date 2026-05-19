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
  caracterizacao: [
    ["Qual e o tamanho medio das NPAg e o indice de polidispersao? Fornecer distribuicao granulometrica completa.", "ISO/TR 10993-22; FDA Nanotechnology Guidance 2014"],
    ["Qual e a morfologia das nanoparticulas? Apresentar imagens de microscopia eletronica de transmissao.", "ISO/TR 10993-22:2017, item 6.2"],
    ["Qual e a concentracao superficial de prata e a concentracao total por unidade do produto?", "RDC 56/2001 - Requisito essencial 10"],
    ["Apresentar cinetica de liberacao de ions Ag+ in vitro em simulador de urina.", "ISO/TR 10993-22; ASTM E2696"],
    ["Apresentar estudo de caracterizacao das NPAg antes e apos o processo de esterilizacao.", "ISO 10993-12:2021; RDC 16/2013"],
  ],
  biocompatibilidade: [
    ["Citotoxicidade com linhagem celular urotelial humana, nao apenas fibroblastos L929.", "ISO 10993-5:2009"],
    ["Sensibilizacao por GPMT ou ensaio de Buehler.", "ISO 10993-10:2021"],
    ["Toxicidade sistemica aguda por via intravenosa e intraperitoneal.", "ISO 10993-11:2017"],
    ["Genotoxicidade: Ames, micronucleo in vitro e micronucleo in vivo.", "ISO 10993-3:2014"],
    ["Estudo de biodistribuicao com quantificacao de prata em orgaos-alvo.", "ISO/TR 10993-22:2017, item 8"],
  ],
  antimicrobiana: [
    ["Quais sao os claims de eficacia declarados na rotulagem proposta?", "RDC 185/2001, Anexo I"],
    ["Os testes foram realizados com o dispositivo final ou apenas com extratos?", "ASTM E2180; ISO 22196 adaptado"],
    ["Qual e o painel de microrganismos testados?", "Diretrizes CLSI; ABNT NBR ISO 20776"],
    ["Foram realizados testes de formacao e inibicao de biofilme?", "ASTM E2196; protocolo CDC Biofilm"],
    ["Existe estudo clinico prospectivo demonstrando reducao de ITU?", "CONSORT 2010; RDC 204/2017"],
  ],
  bpf: [
    ["Confirmar conformidade com ABNT NBR ISO 8669-2.", "ABNT NBR ISO 8669-2"],
    ["Apresentar estudo de estabilidade acelerada e em tempo real.", "RDC 185/2001; ASTM F1980"],
    ["Apresentar CBPF vigente do fabricante.", "RDC 16/2013"],
    ["Demonstrar homogeneidade de distribuicao das NPAg ao longo do cateter.", "RDC 56/2001 - Requisito essencial 3"],
  ],
  rotulagem: [
    ["Rotulagem deve mencionar explicitamente a presenca de nanoparticulas de prata.", "RDC 185/2001; RDC 36/2015"],
    ["IFU deve incluir contraindicacao para alergia a prata e tempo maximo de permanencia.", "RDC 185/2001; ABNT NBR 7500"],
    ["Apresentar plano de tecnovigilancia pos-registro.", "RDC 67/2009; RDC 204/2017"],
    ["Apresentar plano de gestao de residuos contendo nanoparticulas.", "CONAMA 358/2005; RDC ANVISA 222/2018"],
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
        heading1("3. Bloco I - Caracterizacao das Nanoparticulas de Prata"),
        para("A variacao em parametros nanotecnologicos altera de forma significativa o perfil toxicologico, a reatividade biologica e a taxa de liberacao ionica das NPAg."),
        reqTable(blocks.caracterizacao),
        heading1("4. Bloco II - Biocompatibilidade e Toxicologia"),
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
        heading1("5. Bloco III - Eficacia Antimicrobiana"),
        para("A alegacao de acao antimicrobiana constitui a principal justificativa tecnologica para a presenca de NPAg no produto."),
        reqTable(blocks.antimicrobiana),
        heading1("6. Bloco IV - Desempenho do Dispositivo e BPF"),
        reqTable(blocks.bpf),
        heading1("7. Bloco V - Rotulagem e Vigilancia Pos-Mercado"),
        reqTable(blocks.rotulagem),
        heading1("8. Pesquisas e Consultas do Relator"),
        bullet("SINAES, DATAVISA e base ANVISA de registros para eventos adversos, CBPF e precedentes regulatorios."),
        bullet("FDA MAUDE, EUDAMED, documentos FDA, SCENIHR e ECHA para referencia internacional."),
        bullet("PubMed, Embase e Cochrane para literatura cientifica sobre cateteres de prata e nanoparticulas."),
        heading1("9. Sintese e Posicionamento do Relator"),
        alertBox(
          "Posicionamento tecnico",
          [
            "O produto apresenta perfil de inovacao que demanda dossie tecnico superior ao exigido para sondas convencionais.",
            "A insuficiencia em qualquer bloco tecnico determinara exigencia ou indeferimento motivado.",
            "Recomenda-se reuniao tecnica previa com o requerente para alinhamento do escopo do dossie.",
          ],
          GREEN_BG,
          GREEN,
        ),
        para("A inovacao tecnologica representada pelas nanoparticulas de prata nao reduz as exigencias regulatorias; ela as amplifica, pois introduz variaveis toxicologicas e de desempenho sem precedente consolidado na literatura regulatoria brasileira."),
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
