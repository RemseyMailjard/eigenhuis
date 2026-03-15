// ============================================================
// Copilot in Dynamics 365 CE — Word Document Generator
// Maakt: Trainershandleiding + Deelnemerswerkboek
// Vereniging Eigen Huis | NL
// ============================================================

"use strict";
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
  ShadingType,
  PageBreak,
  Header,
  Footer,
  TabStopPosition,
  TabStopType,
  UnderlineType,
} = require("docx");
const fs = require("fs");

// ─── Kleuren (hex zonder #) ───────────────────────────────
const PURPLE = "3B2785";
const ORANGE = "F07800";
const LIGHT = "EDE8F9";
const NAVY = "1A1A4E";
const GRAY = "DDDDDD";

// ─── Helper functies ──────────────────────────────────────

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [
      new TextRun({
        text,
        bold: true,
        color: PURPLE,
        size: 36,
        font: "Calibri",
      }),
    ],
    spacing: { before: 400, after: 200 },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [
      new TextRun({ text, bold: true, color: NAVY, size: 28, font: "Calibri" }),
    ],
    spacing: { before: 300, after: 150 },
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [
      new TextRun({
        text,
        bold: true,
        color: ORANGE,
        size: 24,
        font: "Calibri",
      }),
    ],
    spacing: { before: 200, after: 100 },
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    children: [
      new TextRun({ text, font: "Calibri", size: 22, color: NAVY, ...opts }),
    ],
    spacing: { before: 80, after: 80 },
  });
}

function italic(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: "Calibri",
        size: 22,
        color: NAVY,
        italics: true,
      }),
    ],
    spacing: { before: 80, after: 80 },
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    bullet: { level },
    children: [new TextRun({ text, font: "Calibri", size: 22, color: NAVY })],
    spacing: { before: 60, after: 60 },
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbered-list", level },
    children: [new TextRun({ text, font: "Calibri", size: 22, color: NAVY })],
    spacing: { before: 60, after: 60 },
  });
}

function divider() {
  return new Paragraph({
    border: { bottom: { color: PURPLE, size: 6, style: BorderStyle.SINGLE } },
    children: [],
    spacing: { before: 200, after: 200 },
  });
}

function spacer() {
  return new Paragraph({ children: [], spacing: { before: 120, after: 0 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function labeledBox(label, text) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            shading: { color: PURPLE, type: ShadingType.SOLID },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: label,
                    bold: true,
                    color: "FFFFFF",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
          }),
          new TableCell({
            width: { size: 80, type: WidthType.PERCENTAGE },
            shading: { color: LIGHT, type: ShadingType.SOLID },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text, font: "Calibri", size: 22, color: NAVY }),
                ],
              }),
            ],
            margins: { top: 100, bottom: 100, left: 150, right: 100 },
          }),
        ],
      }),
    ],
  });
}

function promptBox(label, promptText) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { color: ORANGE, type: ShadingType.SOLID },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: label,
                    bold: true,
                    color: "FFFFFF",
                    font: "Calibri",
                    size: 20,
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            shading: { color: "F9F5FF", type: ShadingType.SOLID },
            children: promptText.map(
              (line) =>
                new Paragraph({
                  children: [
                    new TextRun({
                      text: line,
                      font: "Courier New",
                      size: 20,
                      color: NAVY,
                    }),
                  ],
                  spacing: { before: 50, after: 50 },
                }),
            ),
            margins: { top: 100, bottom: 100, left: 200, right: 100 },
          }),
        ],
      }),
    ],
  });
}

function compareTable(left, right, leftRows, rightRows) {
  const rows = [
    new TableRow({
      children: [
        new TableCell({
          shading: { color: "AA0000", type: ShadingType.SOLID },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: left,
                  bold: true,
                  color: "FFFFFF",
                  font: "Calibri",
                  size: 22,
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
        }),
        new TableCell({
          shading: { color: "005500", type: ShadingType.SOLID },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: right,
                  bold: true,
                  color: "FFFFFF",
                  font: "Calibri",
                  size: 22,
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          ],
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
        }),
      ],
    }),
  ];

  const maxLen = Math.max(leftRows.length, rightRows.length);
  for (let i = 0; i < maxLen; i++) {
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            shading: { color: "FFF5F5", type: ShadingType.SOLID },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: leftRows[i] || "",
                    font: "Calibri",
                    size: 20,
                    color: "880000",
                    italics: true,
                  }),
                ],
              }),
            ],
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
          }),
          new TableCell({
            shading: { color: "F0FFF0", type: ShadingType.SOLID },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: rightRows[i] || "",
                    font: "Calibri",
                    size: 20,
                    color: "003300",
                  }),
                ],
              }),
            ],
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
          }),
        ],
      }),
    );
  }

  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}

// ══════════════════════════════════════════════════════════════════════════
// TRAINERSHANDLEIDING
// ══════════════════════════════════════════════════════════════════════════

function buildTrainershandleiding() {
  const children = [];

  // Titelpagina
  children.push(
    spacer(),
    spacer(),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Trainershandleiding",
          bold: true,
          color: PURPLE,
          font: "Calibri",
          size: 56,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Copilot in Dynamics 365 CE",
          bold: true,
          color: ORANGE,
          font: "Calibri",
          size: 36,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Vereniging Eigen Huis  |  Training 2026",
          font: "Calibri",
          size: 24,
          color: NAVY,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 800 },
    }),
    labeledBox(
      "Doelgroep",
      "Klantadviseurs, servicemedewerkers, casebehandelaars, marketingmedewerkers, CRM-beheerders",
    ),
    spacer(),
    labeledBox(
      "Duur",
      "2 – 3 uur (afhankelijk van groepsgrootte en demo-tijd)",
    ),
    spacer(),
    labeledBox(
      "Vereist",
      "Laptop met toegang tot Dynamics 365 CE en Copilot-licentie, of projectiescherm voor trainer-demo",
    ),
    pageBreak(),
  );

  // Gebruik van deze handleiding
  children.push(
    heading1("Gebruik van deze handleiding"),
    body(
      "Deze handleiding begeleidt je als trainer door alle 7 modules. Per module vind je:",
    ),
    bullet("Leerdoel en tijdsindicatie"),
    bullet("Spreekpunten en achtergrondinformatie"),
    bullet("Exacte stappen voor de live demonstratie"),
    bullet("Verwachte vragen van deelnemers + modelantwoorden"),
    bullet("Promptvoorbeelden (kopieerklaar voor Copilot)"),
    spacer(),
    body("Hanteer deze aanpak per module:"),
    numbered("Introduceer het onderwerp (2–3 min)"),
    numbered("Geef achtergrond en spreekpunten (5 min)"),
    numbered("Voer de live demo uit (5–7 min)"),
    numbered("Deelnemers voeren de praktijkopdracht uit (10 min)"),
    numbered("Nabespreking in de groep (3–5 min)"),
    divider(),
  );

  // Module 1
  children.push(
    heading1("Module 01 — Introductie Copilot in Dynamics"),
    labeledBox(
      "Leerdoel",
      "Deelnemers begrijpen wat Copilot is, hoe het verschilt van andere AI-tools en wat het kan doen binnen Dynamics 365 CE.",
    ),
    spacer(),
    labeledBox("Tijdsindicatie", "20 minuten (10 min uitleg + 10 min demo)"),
    spacer(),
    heading2("Spreekpunten"),
    bullet(
      "AI wordt steeds gewoner in de werkplek — dit is voor jullie relevant",
    ),
    bullet(
      "Copilot is geen vervanging voor jullie werk. Het is een slimme assistent die je sneller laat werken",
    ),
    bullet(
      "Het grote verschil met ChatGPT: Dynamics Copilot werkt met échte klantdata uit het CRM",
    ),
    bullet(
      "Privacy: Copilot stuurt geen klantdata door naar externe servers — het werkt binnen Microsoft's beveiligde omgeving",
    ),
    bullet("Je blijft altijd verantwoordelijk voor wat je verstuurt"),
    spacer(),
    heading2("Demo: Copilot in actie"),
    numbered("Open Dynamics 365 CE in de browser"),
    numbered("Open een bestaande case (gebruik een testcase of demo-omgeving)"),
    numbered("Klik op het Copilot-icoon rechts in het scherm"),
    numbered("Klik op 'Samenvatten' en wijs de output aan op het scherm"),
    numbered("Klik op 'Antwoord voorstellen' en bespreek het concept"),
    spacer(),
    heading2("Verwachte vragen"),
    labeledBox("Vraag", "Is Copilot altijd correct?"),
    spacer(),
    body(
      "Antwoord: Nee. Copilot genereert op basis van beschikbare data, maar kan fouten maken of context missen. Controleer altijd de output voordat je iets verstuurt.",
      { italics: true },
    ),
    spacer(),
    labeledBox("Vraag", "Wat als Copilot de verkeerde taal gebruikt?"),
    spacer(),
    body(
      "Antwoord: Je kunt in de prompt aangeven welke taal gewenst is. Bijv: 'Schrijf dit antwoord in het Nederlands, formeel.'",
      { italics: true },
    ),
    divider(),
  );

  // Module 2
  children.push(
    heading1("Module 02 — Copilot in klachtafhandeling"),
    labeledBox(
      "Leerdoel",
      "Deelnemers kunnen Copilot gebruiken om cases te samenvatten en vervolgstappen te genereren.",
    ),
    spacer(),
    labeledBox(
      "Tijdsindicatie",
      "25 minuten (5 min theorie + 7 min demo + 10 min opdracht + 3 min nabespreking)",
    ),
    spacer(),
    heading2("Spreekpunten"),
    bullet(
      "Bij VEH komen dagelijks veel cases binnen over aannemers, isolatie, conflicten, etc.",
    ),
    bullet(
      "Een medewerker besteedt gemiddeld 5–10 minuten aan het doorlezen van een case voordat hij/zij reageert",
    ),
    bullet("Copilot kan dit terugbrengen naar 30–60 seconden"),
    bullet(
      "Belangrijk: de samenvatting is een hulpmiddel — altijd even checken op correctheid",
    ),
    spacer(),
    heading2("Demo-stappen"),
    numbered("Open een klachtcase met minimaal 3 berichten"),
    numbered("Klik op Copilot → 'Samenvatten'"),
    numbered("Wijs de samenvatting aan en bespreek wat Copilot herkende"),
    numbered("Klik op 'Vervolgstappen voorstellen' — bespreek de suggesties"),
    spacer(),
    heading2("Promptvoorbeelden module 02"),
    promptBox("Prompt: case samenvatten", [
      "Vat de kern van deze klacht samen in maximaal 3 bullets.",
      "Noem:",
      "- het probleem",
      "- wat het lid al probeert heeft",
      "- wat het lid als oplossing verwacht",
    ]),
    spacer(),
    promptBox("Prompt: vervolgstap", [
      "Je bent adviseur bij Vereniging Eigen Huis.",
      "Stel drie mogelijke vervolgstappen voor op basis van deze case.",
      "Houd rekening met: de rechten van het lid als consument,",
      "de rol van VEH als adviseur (niet de uitvoerder).",
    ]),
    spacer(),
    heading2("Verwachte vragen"),
    labeledBox("Vraag", "Wat als Copilot details mist?"),
    spacer(),
    body(
      "Antwoord: Vraag Copilot om de samenvatting te verfijnen: 'Voeg ook toe wat het lid eerder heeft ondernomen.' Copilot gebruikt alleen data die in de case staat — ontbrekende info kan niet worden aangevuld.",
      { italics: true },
    ),
    divider(),
  );

  // Module 3
  children.push(
    heading1("Module 03 — E-mailafhandeling met Copilot"),
    labeledBox(
      "Leerdoel",
      "Deelnemers kunnen Copilot inzetten om e-mails te samenvatten en conceptantwoorden te genereren en aan te passen.",
    ),
    spacer(),
    labeledBox(
      "Tijdsindicatie",
      "25 minuten (5 min theorie + 7 min demo + 10 min opdracht + 3 min nabespreking)",
    ),
    spacer(),
    heading2("Spreekpunten"),
    bullet(
      "E-mail is nog steeds het meest gebruikte communicatiekanaal bij VEH",
    ),
    bullet(
      "Copilot leest de e-mail en stelt een antwoord voor dat past bij de vraag",
    ),
    bullet("Toon is aanpasbaar: professioneel, warm, kort, uitgebreid"),
    bullet(
      "Let op: Copilot kent de persoonlijke situatie van het lid niet altijd — voeg context toe via de prompt",
    ),
    spacer(),
    heading2("Demo-stappen"),
    numbered("Open een inkomende e-mail in Dynamics over isolatiesubsidie"),
    numbered("Klik op Copilot → 'Samenvatten'"),
    numbered("Klik op 'Antwoord genereren'"),
    numbered("Bespreek het concept — wat klopt? Wat ontbreekt?"),
    numbered(
      "Pas aan: gebruik prompt 'Maak dit antwoord eenvoudiger, voor een leek'",
    ),
    spacer(),
    heading2("Promptvoorbeelden module 03"),
    promptBox("Prompt: e-mail beantwoorden", [
      "Je bent medewerker bij Vereniging Eigen Huis.",
      "Schrijf een professioneel en begripvol antwoord op de onderstaande e-mail.",
      "Gebruik een vriendelijke maar formele toon.",
      "Sluit af met een concrete volgende stap of advies.",
    ]),
    spacer(),
    promptBox("Prompt: toon aanpassen", [
      "Herschrijf het bovenstaande antwoord eenvoudiger.",
      "Vermijd vakjargon.",
      "Schrijf alsof je het uitlegt aan iemand zonder technische kennis.",
      "Maximaal 150 woorden.",
    ]),
    divider(),
  );

  // Module 4
  children.push(
    heading1("Module 04 — Marketing en communicatie"),
    labeledBox(
      "Leerdoel",
      "Deelnemers kunnen Copilot gebruiken voor het genereren en herschrijven van communicatiecontent.",
    ),
    spacer(),
    labeledBox(
      "Tijdsindicatie",
      "20 minuten (5 min theorie + 5 min demo + 10 min opdracht)",
    ),
    spacer(),
    heading2("Spreekpunten"),
    bullet(
      "VEH stuurt regelmatig nieuwsbrieven, adviesteksten en lidmaatschapsberichten",
    ),
    bullet(
      "Copilot versnelt het schrijfproces — van leeg scherm naar bruikbaar concept in 30 seconden",
    ),
    bullet("Het eindresultaat vereist altijd inhoudelijke controle"),
    bullet(
      "Doelgroepgericht schrijven: Copilot kan de tekst aanpassen voor verschillende groepen",
    ),
    spacer(),
    heading2("Demo-stappen"),
    numbered("Open de marketingmodule in Dynamics"),
    numbered("Activeer Copilot en geef het thema: 'energiebesparing woningen'"),
    numbered("Laat Copilot 5 onderwerpen genereren"),
    numbered("Kies onderwerp 1 en laat een nieuwsbriefartikel schrijven"),
    numbered("Pas de lengte aan: 'Maak dit korter, maximaal 100 woorden'"),
    spacer(),
    heading2("Promptvoorbeelden module 04"),
    promptBox("Prompt: nieuwsbrief schrijven", [
      "Schrijf een kort nieuwsbriefartikel van 150 woorden over",
      "energiebesparing in woningen voor VEH-leden.",
      "Neem minimaal 3 concrete tips op.",
      "Sluit af met een call-to-action: verwijs naar de VEH website.",
      "Gebruik een vriendelijke, informatieve toon.",
    ]),
    divider(),
  );

  // Module 5
  children.push(
    heading1("Module 05 — Marketingsegmentatie met Copilot"),
    labeledBox(
      "Leerdoel",
      "Deelnemers kunnen Copilot gebruiken om segmenten te bouwen op basis van gedragsdata in het CRM.",
    ),
    spacer(),
    labeledBox(
      "Tijdsindicatie",
      "20 minuten (5 min theorie + 5 min demo + 10 min opdracht)",
    ),
    spacer(),
    heading2("Spreekpunten"),
    bullet(
      "VEH heeft duizenden leden met diverse interesses en gedragspatronen",
    ),
    bullet(
      "Segmentatie zorgt voor relevantere communicatie en hogere betrokkenheid",
    ),
    bullet(
      "Copilot vertaalt gewone taal ('leden die geïnteresseerd zijn in isolatie') naar concrete filtercriteria",
    ),
    bullet(
      "Controleer altijd of de segmentcriteria kloppen met de data die beschikbaar is",
    ),
    spacer(),
    heading2("Demo-stappen"),
    numbered("Open Dynamics → Marketing → Segmenten"),
    numbered("Klik op 'Nieuw segment'"),
    numbered("Activeer Copilot: 'Help mij een segment te maken'"),
    numbered(
      "Beschrijf: 'Leden die in 2025 minimaal één e-mail over isolatie openden'",
    ),
    numbered("Bekijk de gegenereerde filtercriteria en pas zo nodig aan"),
    spacer(),
    heading2("Promptvoorbeelden module 05"),
    promptBox("Prompt: segment definiëren", [
      "Maak een marketingsegment van alle leden van VEH die:",
      "- in de afgelopen 6 maanden minimaal één nieuwsbrief openden",
      "- en minstens één keer contact opnamen over energiebesparing of isolatie.",
      "Welke filtercriteria zijn hiervoor nodig in Dynamics?",
    ]),
    divider(),
  );

  // Module 6
  children.push(
    heading1("Module 06 — Prompttechnieken"),
    labeledBox(
      "Leerdoel",
      "Deelnemers begrijpen hoe een effectieve prompt is opgebouwd en kunnen zwakke prompts verbeteren.",
    ),
    spacer(),
    labeledBox(
      "Tijdsindicatie",
      "20 minuten (10 min theorie/vergelijking + 10 min oefening)",
    ),
    spacer(),
    heading2("Spreekpunten"),
    bullet(
      "De prompt is de 'instructie' die je aan Copilot geeft — kwaliteit bepaalt alles",
    ),
    bullet("4 elementen van een goede prompt: context, taak, structuur, toon"),
    bullet("Itereren is normaal: verfijn je prompt op basis van het resultaat"),
    bullet("Je hoeft geen programmeur te zijn — gewone taal werkt het beste"),
    spacer(),
    heading2("Vergelijking: zwak vs. sterk"),
    compareTable(
      "Zwakke prompt",
      "Sterke prompt",
      [
        "Schrijf een antwoord",
        "Maak een segment",
        "Samenvatten",
        "Schrijf een nieuwsbrief",
      ],
      [
        "Je bent adviseur bij VEH. Schrijf een professioneel antwoord met begrip, uitleg en vervolgstappen.",
        "Selecteer leden die in 2025 minimaal één e-mail openden over isolatie.",
        "Vat de kernklacht samen in 3 bullets. Noem ook het gewenste resultaat van het lid.",
        "Schrijf een nieuwsbrief (150 woorden) over energiebesparing, actiegericht, voor huiseigenaren.",
      ],
    ),
    spacer(),
    heading2("Oefening: prompts verbeteren"),
    body(
      "Laat deelnemers individueel in 5 minuten betere versies schrijven van:",
    ),
    numbered('"Schrijf een nieuwsbrief"'),
    numbered('"Analyseer deze case"'),
    numbered('"Maak een segment van leden"'),
    numbered('"Beantwoord deze e-mail"'),
    spacer(),
    body(
      "Bespreek de antwoorden in de groep. Er zijn meerdere goede oplossingen —  focus op: bevat de prompt context, taak, structuur en toon?",
    ),
    divider(),
  );

  // Module 7
  children.push(
    heading1("Module 07 — Eindcase"),
    labeledBox(
      "Leerdoel",
      "Deelnemers passen alle geleerde vaardigheden toe in een realistische casesituatie.",
    ),
    spacer(),
    labeledBox(
      "Tijdsindicatie",
      "25 minuten (5 min intro + 15 min zelfstandig werken + 5 min nabespreking)",
    ),
    spacer(),
    heading2("Scenario"),
    italic(
      '"Ik heb een conflict met mijn aannemer over de isolatie van mijn woning. Hij heeft het werk niet goed afgerond en reageert niet meer op mijn berichten."',
    ),
    spacer(),
    heading2("Stappen voor deelnemers"),
    numbered("Laat Copilot de klacht samenvatten in 3 bullets"),
    numbered(
      "Laat Copilot het probleem analyseren en juridische aspecten benoemen",
    ),
    numbered(
      "Genereer een professionele antwoordmail met begrip en vervolgstappen",
    ),
    numbered("Herschrijf het antwoord voor een lid zonder juridische kennis"),
    spacer(),
    heading2("Spreekpunten nabespreking"),
    bullet("Wat werkte goed? Welke prompts gaven de beste resultaten?"),
    bullet("Waar moest je Copilot bijsturen?"),
    bullet("Wat zou je zonder Copilot anders/langer hebben gedaan?"),
    spacer(),
    heading2("Aandachtspunten Copilot verantwoord gebruiken"),
    bullet(
      "Copilot is een hulpmiddel — jij blijft verantwoordelijk voor het eindresultaat",
    ),
    bullet("Controleer altijd de inhoud vóór je iets verstuurt"),
    bullet("Wees voorzichtig met gevoelige informatie in prompts"),
    bullet(
      "AI begrijpt emotie en context niet altijd — herschrijf empathische onderdelen zelf",
    ),
    bullet(
      "Gebruik Copilot voor routinetaken — complexe situaties vragen menselijk inzicht",
    ),
    divider(),
  );

  // Promptcheatsheet
  children.push(
    heading1("Bijlage: Promptcheatsheet Dynamics Copilot"),
    heading2("Structuurformule"),
    labeledBox(
      "Formule",
      "[Rol/context] + [Duidelijke taak] + [Gewenste structuur] + [Toon/stijl]",
    ),
    spacer(),
    heading2("Productie-prompts per taak"),
    heading3("Case samenvatten"),
    promptBox("Kopieer & plak", [
      "Vat de kern van deze case samen in maximaal 3 bullets.",
      "Noem: het probleem, wat het lid eerder ondernam, en wat het lid verwacht.",
    ]),
    spacer(),
    heading3("E-mail beantwoorden"),
    promptBox("Kopieer & plak", [
      "Je bent medewerker bij Vereniging Eigen Huis.",
      "Schrijf een professioneel en begripvol antwoord op de onderstaande e-mail.",
      "Sluit af met een concreet advies of volgende stap.",
    ]),
    spacer(),
    heading3("Toon aanpassen"),
    promptBox("Kopieer & plak", [
      "Herschrijf het bovenstaande antwoord eenvoudiger.",
      "Geen vakjargon. Schrijf alsof je het uitlegt aan iemand zonder technische kennis.",
      "Maximaal 150 woorden.",
    ]),
    spacer(),
    heading3("Nieuwsbrief schrijven"),
    promptBox("Kopieer & plak", [
      "Schrijf een kort nieuwsbriefartikel van 150 woorden over [onderwerp] voor VEH-leden.",
      "Neem minimaal 3 concrete tips op.",
      "Sluit af met een call-to-action.",
      "Gebruik een vriendelijke, informatieve toon.",
    ]),
    spacer(),
    heading3("Segment definiëren"),
    promptBox("Kopieer & plak", [
      "Welke filtercriteria heb ik nodig in Dynamics voor een segment van leden die:",
      "- [criterium 1]",
      "- [criterium 2]",
      "- [criterium optioneel]?",
    ]),
  );

  return new Document({
    numbering: {
      config: [
        {
          reference: "numbered-list",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    sections: [{ children }],
  });
}

// ══════════════════════════════════════════════════════════════════════════
// DEELNEMERSWERKBOEK
// ══════════════════════════════════════════════════════════════════════════

function buildWerkboek() {
  const children = [];

  // Titelpagina
  children.push(
    spacer(),
    spacer(),
    spacer(),
    new Paragraph({
      children: [
        new TextRun({
          text: "Deelnemerswerkboek",
          bold: true,
          color: PURPLE,
          font: "Calibri",
          size: 56,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Copilot in Dynamics 365 CE",
          bold: true,
          color: ORANGE,
          font: "Calibri",
          size: 36,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Vereniging Eigen Huis  |  Training 2026",
          font: "Calibri",
          size: 24,
          color: NAVY,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 800 },
    }),
    labeledBox("Naam deelnemer", "___________________________________________"),
    spacer(),
    labeledBox("Rol / afdeling", "___________________________________________"),
    spacer(),
    labeledBox("Datum training", "___________________________________________"),
    pageBreak(),
  );

  // Welkom
  children.push(
    heading1("Welkom!"),
    body(
      "Dit werkboek begeleidt je door de training Copilot in Dynamics 365 CE. Je vindt hier per module:",
    ),
    bullet("De belangrijkste leerpunten"),
    bullet("De praktijkopdrachten"),
    bullet("Ruimte voor aantekeningen"),
    bullet("Een promptcheatsheet om na de training te gebruiken"),
    spacer(),
    body("Na deze training kun je:"),
    bullet("Begrijpen wat Copilot is en hoe het werkt in Dynamics 365"),
    bullet("AI gebruiken om cases sneller te analyseren"),
    bullet("E-mails en reacties efficiënter schrijven"),
    bullet("Copilot inzetten voor marketing en communicatie"),
    bullet("Effectieve prompts schrijven voor betere AI-resultaten"),
    divider(),
  );

  const modules = [
    {
      num: "01",
      title: "Introductie Copilot in Dynamics",
      punten: [
        "Copilot is een AI-assistent ingebouwd in Dynamics 365",
        "Het werkt direct met jouw klantdata uit het CRM",
        "Dynamics Copilot verschilt van ChatGPT: het gebruikt échte CRM-data",
        "Jij blijft altijd verantwoordelijk — Copilot ondersteunt, beslist niet",
      ],
      opdracht: null,
      aantekeningen: true,
    },
    {
      num: "02",
      title: "Copilot in klachtafhandeling",
      punten: [
        "Copilot kan een case samenvatten in seconden",
        "Vervolgstappen worden automatisch voorgesteld op basis van de context",
        "Controleer altijd of de samenvatting klopt met de werkelijkheid",
      ],
      scenario:
        "Een lid meldt dat een aannemer slecht werk heeft geleverd en niet meer reageert. Het gaat om gevelbekleding die loslaat na isolatiewerkzaamheden.",
      opdracht: [
        "Open de case in Dynamics 365",
        "Laat Copilot de klacht samenvatten in 3 bullets",
        "Laat Copilot mogelijke vervolgstappen voorstellen",
        "Noteer hieronder: wat klopt goed? Wat ontbreekt of klopt niet?",
      ],
      aantekeningen: true,
    },
    {
      num: "03",
      title: "E-mailafhandeling met Copilot",
      punten: [
        "Copilot vat een e-mail samen en stelt een concept-antwoord voor",
        "De toon is aanpasbaar: formeel, begripvol, eenvoudig",
        "Herschrijven kan met een eenvoudige vervolg-prompt",
      ],
      scenario:
        '"Ik wil mijn woning isoleren en ik hoor dat er subsidies zijn. Kunt u mij vertellen wat ik kan aanvragen en hoe dat werkt?"',
      opdracht: [
        "Laat Copilot de vraag samenvatten in één zin",
        "Laat Copilot een informatief antwoord opstellen",
        "Herschrijf het antwoord voor een lid zonder technische kennis",
        "Vergelijk: wat is beter aan de Copilot-versie? Wat heb jij verbeterd?",
      ],
      aantekeningen: true,
    },
    {
      num: "04",
      title: "Marketing en communicatie",
      punten: [
        "Copilot genereert onderwerpen, artikelen en berichten op thema",
        "Inhoudelijke controle blijft altijd noodzakelijk",
        "Toon aanpassen aan doelgroep via de prompt",
      ],
      scenario:
        "Leden willen weten hoe ze energiekosten kunnen besparen in hun woning.",
      opdracht: [
        "Schrijf met Copilot een korte nieuwsbrief van ca. 150 woorden over energiebesparing",
        "Gebruik de prompt uit de cheatsheet of schrijf je eigen versie",
        "Herschrijf één alinea in een toegankelijkere schrijfstijl",
        "Noteer: welke prompt gaf het beste resultaat?",
      ],
      aantekeningen: true,
    },
    {
      num: "05",
      title: "Marketingsegmentatie met Copilot",
      punten: [
        "Copilot herkent patronen in CRM-data",
        "Segmenten aanmaken via gewone taalbeschrijving",
        "Altijd controleren of de criteria overeenkomen met de beschikbare data",
      ],
      scenario:
        "Je wilt een gerichte campagne uitvoeren over energiebesparing voor geïnteresseerde leden.",
      opdracht: [
        "Beschrijf het segment in gewone taal aan Copilot",
        "Voeg minimaal 2 criteria toe (bijv. geopende e-mails + zoekvragen of cases)",
        "Controleer of de filtercriteria logisch zijn",
        "Noteer: welke campagne zou jij op dit segment richten?",
      ],
      aantekeningen: true,
    },
    {
      num: "06",
      title: "Prompttechnieken",
      punten: [
        "Goede prompts bevatten: context + taak + structuur + toon",
        "Itereren is normaal — verfijn je prompt stap voor stap",
        "Je hoeft geen programmeur te zijn: gewone taal werkt het beste",
      ],
      extra: [
        "Verbeter de volgende prompts (schrijf jouw verbeterde versie eronder):",
      ],
      oefenPrompts: [
        { zwak: "Schrijf een nieuwsbrief", ruimte: true },
        { zwak: "Analyseer deze case", ruimte: true },
        { zwak: "Maak een segment van leden", ruimte: true },
        { zwak: "Beantwoord deze e-mail", ruimte: true },
      ],
      aantekeningen: false,
    },
    {
      num: "07",
      title: "Eindcase",
      punten: [
        "Combineer alle vaardigheden in één casus",
        "Zorg voor een goede prompt bij elke stap",
        "Controleer de output kritisch — pas aan waar nodig",
      ],
      scenario:
        '"Ik heb een conflict met mijn aannemer over de isolatie van mijn woning. Hij heeft het werk niet goed afgerond en reageert niet meer op mijn berichten."',
      opdracht: [
        "Stap 1 — Samenvatten: Laat Copilot de klacht samenvatten in 3 bullets",
        "Stap 2 — Analyseren: Laat Copilot het probleem analyseren en juridische aspecten benoemen",
        "Stap 3 — Antwoord schrijven: Genereer een professionele antwoordmail",
        "Stap 4 — Toon aanpassen: Herschrijf het antwoord toegankelijker",
      ],
      aantekeningen: true,
    },
  ];

  for (const mod of modules) {
    children.push(
      heading1(`Module ${mod.num} — ${mod.title}`),
      heading2("Kernpunten"),
    );
    for (const punt of mod.punten) {
      children.push(bullet(punt));
    }

    if (mod.scenario) {
      children.push(spacer(), heading2("Scenario"), italic(mod.scenario));
    }

    if (mod.opdracht) {
      children.push(spacer(), heading2("Praktijkopdracht"));
      mod.opdracht.forEach((t, i) => children.push(numbered(t)));
    }

    if (mod.extra) {
      children.push(spacer(), heading2("Oefening"));
      mod.extra.forEach((e) => children.push(body(e)));
    }

    if (mod.oefenPrompts) {
      for (const p of mod.oefenPrompts) {
        children.push(
          spacer(),
          labeledBox("Zwakke prompt", p.zwak),
          spacer(),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    shading: { color: "F9F5FF", type: ShadingType.SOLID },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Mijn verbeterde prompt:",
                            bold: true,
                            color: PURPLE,
                            font: "Calibri",
                            size: 20,
                          }),
                        ],
                        spacing: { before: 60 },
                      }),
                      new Paragraph({ children: [], spacing: { before: 400 } }),
                    ],
                    margins: { top: 100, bottom: 100, left: 150, right: 100 },
                  }),
                ],
              }),
            ],
          }),
          spacer(),
        );
      }
    }

    if (mod.aantekeningen) {
      children.push(
        spacer(),
        heading2("Mijn aantekeningen"),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { color: "FAFAFA", type: ShadingType.SOLID },
                  children: [
                    new Paragraph({ children: [], spacing: { before: 800 } }),
                  ],
                  margins: { top: 100, bottom: 100, left: 150, right: 100 },
                }),
              ],
            }),
          ],
        }),
      );
    }

    children.push(divider());
  }

  // Promptcheatsheet
  children.push(
    heading1("Promptcheatsheet — Jouw referentiekaart"),
    body("Hang dit op je werkplek of sla het op als referentie."),
    spacer(),
    heading2("Structuurformule"),
    labeledBox(
      "Formule",
      "[Rol] + [Duidelijke taak] + [Structuur] + [Toon/stijl]",
    ),
    spacer(),

    heading3("Case samenvatten"),
    promptBox("", [
      "Vat de kern van deze case samen in maximaal 3 bullets.",
      "Noem: het probleem, wat het lid eerder ondernam, en wat het lid verwacht.",
    ]),
    spacer(),

    heading3("E-mail beantwoorden"),
    promptBox("", [
      "Je bent medewerker bij Vereniging Eigen Huis.",
      "Schrijf een professioneel en begripvol antwoord op de e-mail.",
      "Sluit af met een concreet advies of volgende stap.",
    ]),
    spacer(),

    heading3("Toon aanpassen"),
    promptBox("", [
      "Herschrijf het bovenstaande antwoord eenvoudiger.",
      "Geen vakjargon. Maximaal 150 woorden.",
    ]),
    spacer(),

    heading3("Nieuwsbrief schrijven"),
    promptBox("", [
      "Schrijf een kort nieuwsbriefartikel van 150 woorden over [onderwerp].",
      "Min. 3 concrete tips. Sluit af met een call-to-action.",
    ]),
    spacer(),

    heading3("Segment definiëren"),
    promptBox("", [
      "Welke filtercriteria heb ik nodig voor een segment van leden die:",
      "- [criterium 1]",
      "- [criterium 2]?",
    ]),
    spacer(),

    heading2("Reflectievragen na de training"),
    bullet(
      "Op welk onderdeel van mijn werk kan ik Copilot het meest inzetten?",
    ),
    bullet("Welke taken wil ik de komende week uitproberen met Copilot?"),
    bullet("Op welke punten wil ik mijn prompts verbeteren?"),
    spacer(),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { color: "FAFAFA", type: ShadingType.SOLID },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "Mijn actiepunten voor de komende week:",
                      bold: true,
                      color: PURPLE,
                      font: "Calibri",
                      size: 20,
                    }),
                  ],
                }),
                new Paragraph({ children: [], spacing: { before: 1200 } }),
              ],
              margins: { top: 120, bottom: 120, left: 150, right: 100 },
            }),
          ],
        }),
      ],
    }),
  );

  return new Document({
    numbering: {
      config: [
        {
          reference: "numbered-list",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    sections: [{ children }],
  });
}

// ══════════════════════════════════════════════════════════════════════════
// GENEREER BESTANDEN
// ══════════════════════════════════════════════════════════════════════════

async function main() {
  const handleiding = buildTrainershandleiding();
  const werkboek = buildWerkboek();

  const [buf1, buf2] = await Promise.all([
    Packer.toBuffer(handleiding),
    Packer.toBuffer(werkboek),
  ]);

  fs.writeFileSync("Trainershandleiding_Copilot_D365.docx", buf1);
  console.log(
    "Trainershandleiding aangemaakt: Trainershandleiding_Copilot_D365.docx",
  );

  fs.writeFileSync("Deelnemerswerkboek_Copilot_D365.docx", buf2);
  console.log(
    "Deelnemerswerkboek aangemaakt: Deelnemerswerkboek_Copilot_D365.docx",
  );
}

main().catch((err) => console.error("Fout:", err));
