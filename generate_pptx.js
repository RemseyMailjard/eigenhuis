"use strict";
const PptxGenJS = require("pptxgenjs");

// ─── Kleuren VEH huisstijl ─────────────────────────────────────────────────
const C = {
  PURPLE : "3B2785",
  ORANGE : "F07800",
  WHITE  : "FFFFFF",
  LIGHT  : "F5F5F5",
  DARK   : "1A1A4E",
  GRAY   : "DDDDDD",
  MUTED  : "888888",
  LILAC  : "BEB5D9",
  LAVSUB : "D5CCEE",
};

// ─── Presentatie setup ─────────────────────────────────────────────────────
const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";
pres.author  = "Vereniging Eigen Huis";
pres.title   = "Copilot in Dynamics 365 CE";

const W = 10, H = 5.625;

// ─── Helpers ──────────────────────────────────────────────────────────────

function topBar(s, title, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.65,
    fill: { color: color || C.PURPLE }, line: { color: color || C.PURPLE }
  });
  s.addText(title, {
    x: 0.5, y: 0, w: 9, h: 0.65,
    fontSize: 22, bold: true, color: C.WHITE, fontFace: "Calibri",
    valign: "middle", margin: 0
  });
}

function bulletList(s, bullets, x, y, w, h, fontSize) {
  const items = bullets.map((b, i) => ({
    text: b,
    options: { bullet: true, breakLine: i < bullets.length - 1 }
  }));
  s.addText(items, {
    x, y, w, h,
    fontSize: fontSize || 15, color: C.DARK, fontFace: "Calibri",
    valign: "top", paraSpaceAfter: 6
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE TYPES
// ══════════════════════════════════════════════════════════════════════════

function addCoverSlide() {
  const s = pres.addSlide();
  s.background = { color: C.PURPLE };

  // Orange accent line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.1, w: W, h: 0.1,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  // Decorative right block
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.8, y: 0, w: 2.2, h: H,
    fill: { color: "2D1D6E" }, line: { color: "2D1D6E" }
  });

  // Orange circle accent (top-right)
  s.addShape(pres.shapes.OVAL, {
    x: 8.3, y: 0.3, w: 1.2, h: 1.2,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  s.addText("Copilot in\nDynamics 365 CE", {
    x: 0.7, y: 0.9, w: 6.8, h: 2.2,
    fontSize: 44, bold: true, color: C.WHITE, fontFace: "Calibri",
    align: "left", valign: "top", margin: 0
  });

  s.addText("Training voor Vereniging Eigen Huis", {
    x: 0.7, y: 3.25, w: 7, h: 0.5,
    fontSize: 18, color: C.ORANGE, fontFace: "Calibri",
    bold: false, align: "left", margin: 0
  });

  s.addText("Klantadviseurs  ·  Servicemedewerkers  ·  Casebehandelaars  ·  CRM-beheerders", {
    x: 0.7, y: 3.85, w: 7, h: 0.35,
    fontSize: 11, color: C.LILAC, fontFace: "Calibri", align: "left", margin: 0
  });

  s.addText("Duur: 2 – 3 uur   |   7 modules", {
    x: 0.7, y: 4.25, w: 5, h: 0.3,
    fontSize: 11, color: C.LILAC, fontFace: "Calibri", align: "left", margin: 0
  });

  s.addText("vereniging eigen huis", {
    x: 0.7, y: H - 0.9, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: C.WHITE, fontFace: "Calibri", align: "left", margin: 0
  });
}

function addAgendaSlide() {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, "Programmaoverzicht");

  const modules = [
    ["01", "Introductie Copilot in Dynamics"],
    ["02", "Copilot in klachtafhandeling"],
    ["03", "E-mailafhandeling met Copilot"],
    ["04", "Marketing en communicatie"],
    ["05", "Marketingsegmentatie met Copilot"],
    ["06", "Prompttechnieken"],
    ["07", "Eindcase"],
  ];

  // Two columns: 0-3 left, 4-6 right
  const colX = [0.4, 5.2];
  const startY = 0.95;
  const gapY  = 0.65;

  modules.forEach(([num, title], i) => {
    const col = i >= 4 ? 1 : 0;
    const row = i >= 4 ? i - 4 : i;
    const x = colX[col];
    const y = startY + row * gapY;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.48, h: 0.48,
      fill: { color: C.ORANGE }, line: { color: C.ORANGE }
    });
    s.addText(num, {
      x, y, w: 0.48, h: 0.48,
      fontSize: 13, bold: true, color: C.WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    s.addText(title, {
      x: x + 0.58, y: y + 0.06, w: 4.2, h: 0.38,
      fontSize: 13, color: C.DARK, fontFace: "Calibri",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Total time note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.9, w: 9.2, h: 0.35,
    fill: { color: C.PURPLE }, line: { color: C.PURPLE }
  });
  s.addText("Totale duur: 2 – 3 uur  |  Inclusief demo's en praktijkopdrachten", {
    x: 0.4, y: 4.9, w: 9.2, h: 0.35,
    fontSize: 11, color: C.WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
}

function addSectionSlide(num, title, subtitle) {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };

  // Left purple panel
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 4.1, h: H,
    fill: { color: C.PURPLE }, line: { color: C.PURPLE }
  });

  // Orange accent bottom on left
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.1, w: 4.1, h: 0.1,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  // Module number (huge, orange)
  s.addText(num, {
    x: 0.25, y: 0.3, w: 3.6, h: 1.8,
    fontSize: 110, bold: true, color: C.ORANGE, fontFace: "Calibri",
    align: "left", valign: "top", margin: 0
  });

  // Module title (white, on purple)
  s.addText(title, {
    x: 0.25, y: 2.2, w: 3.6, h: 2.6,
    fontSize: 22, bold: true, color: C.WHITE, fontFace: "Calibri",
    align: "left", valign: "top", margin: 0
  });

  // Right: vertical orange accent
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.1, y: 0, w: 0.06, h: H,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  // Right: description text
  if (subtitle) {
    s.addText(subtitle, {
      x: 4.5, y: 1.2, w: 5.1, h: 3.5,
      fontSize: 15, color: C.DARK, fontFace: "Calibri",
      align: "left", valign: "top", margin: 0
    });
  }
}

function addContentSlide(title, bullets) {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, title);
  bulletList(s, bullets, 0.6, 0.85, 8.8, 4.5, 15);
  return s;
}

function addTwoColSlide(title, leftTitle, leftBullets, rightTitle, rightBullets) {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, title);

  // Left column header
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 0.85, w: 4.1, h: 0.42,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });
  s.addText(leftTitle, {
    x: 0.4, y: 0.85, w: 4.1, h: 0.42,
    fontSize: 14, bold: true, color: C.WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  // Right column header
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 0.85, w: 4.5, h: 0.42,
    fill: { color: C.PURPLE }, line: { color: C.PURPLE }
  });
  s.addText(rightTitle, {
    x: 5.1, y: 0.85, w: 4.5, h: 0.42,
    fontSize: 14, bold: true, color: C.WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  bulletList(s, leftBullets, 0.4, 1.38, 4.1, 3.9, 13);

  // Right bullets in a slightly different style (green italic for strong prompt)
  const rightItems = rightBullets.map((b, i) => ({
    text: b,
    options: { bullet: true, breakLine: i < rightBullets.length - 1 }
  }));
  s.addText(rightItems, {
    x: 5.1, y: 1.38, w: 4.5, h: 3.9,
    fontSize: 13, color: C.DARK, fontFace: "Calibri",
    valign: "top", paraSpaceAfter: 6
  });

  return s;
}

function addDemoSlide(title, steps) {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, title, C.DARK);

  // DEMO badge
  s.addShape(pres.shapes.RECTANGLE, {
    x: 8.1, y: 0.1, w: 1.6, h: 0.45,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });
  s.addText("▶ DEMO", {
    x: 8.1, y: 0.1, w: 1.6, h: 0.45,
    fontSize: 12, bold: true, color: C.WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  // Vertical connector line between steps
  const stepH  = 0.55;
  const startY = 0.9;
  const gap    = 0.78;

  steps.forEach((step, i) => {
    const y = startY + i * gap;

    // Number box
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.48, h: stepH,
      fill: { color: C.ORANGE }, line: { color: C.ORANGE }
    });
    s.addText(`${i + 1}`, {
      x: 0.4, y, w: 0.48, h: stepH,
      fontSize: 15, bold: true, color: C.WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });

    // Step text
    s.addText(step, {
      x: 1.05, y: y + 0.05, w: 8.5, h: stepH - 0.1,
      fontSize: 14, color: C.DARK, fontFace: "Calibri",
      align: "left", valign: "middle", margin: 0
    });

    // Vertical connector (thin rectangle)
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.615, y: y + stepH, w: 0.05, h: gap - stepH,
        fill: { color: C.ORANGE }, line: { color: C.ORANGE }
      });
    }
  });

  return s;
}

function addExerciseSlide(moduleNum, scenario, opdracht) {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };

  // Orange left strip
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: H,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  // Top bar (orange)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.18, y: 0, w: W - 0.18, h: 0.65,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });
  s.addText(`Praktijkopdracht — Module ${moduleNum}`, {
    x: 0.55, y: 0, w: 9.1, h: 0.65,
    fontSize: 20, bold: true, color: C.WHITE, fontFace: "Calibri",
    valign: "middle", margin: 0
  });

  // Scenario
  s.addText("Scenario", {
    x: 0.5, y: 0.85, w: 1.8, h: 0.35,
    fontSize: 13, bold: true, color: C.ORANGE, fontFace: "Calibri", margin: 0
  });
  s.addText(scenario, {
    x: 0.5, y: 1.2, w: 9.1, h: 0.9,
    fontSize: 14, color: C.DARK, fontFace: "Calibri",
    italic: true, margin: 0
  });

  // Divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.18, w: 9.0, h: 0.03,
    fill: { color: C.GRAY }, line: { color: C.GRAY }
  });

  // Opdracht
  s.addText("Jouw opdracht", {
    x: 0.5, y: 2.3, w: 3.0, h: 0.35,
    fontSize: 13, bold: true, color: C.PURPLE, fontFace: "Calibri", margin: 0
  });

  bulletList(s, opdracht, 0.5, 2.7, 9.1, 2.7, 14);

  return s;
}

function addTableSlide(title, rows, colWidths) {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, title);

  s.addTable(rows, {
    x: 0.5, y: 0.85, w: 9.0,
    colW: colWidths,
    border: { pt: 1, color: C.GRAY },
    autoPage: false,
    rowH: 0.6,
  });

  return s;
}

function addClosingSlide() {
  const s = pres.addSlide();
  s.background = { color: C.PURPLE };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.1, w: W, h: 0.1,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  // Decorative right element
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.8, y: 0, w: 2.2, h: H,
    fill: { color: "2D1D6E" }, line: { color: "2D1D6E" }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 8.3, y: 0.3, w: 1.2, h: 1.2,
    fill: { color: C.ORANGE }, line: { color: C.ORANGE }
  });

  s.addText("Klaar!", {
    x: 0.7, y: 0.6, w: 7, h: 1.0,
    fontSize: 54, bold: true, color: C.WHITE, fontFace: "Calibri",
    align: "left", margin: 0
  });

  s.addText("Belangrijke aandachtspunten", {
    x: 0.7, y: 1.8, w: 7, h: 0.4,
    fontSize: 17, bold: true, color: C.ORANGE, fontFace: "Calibri",
    align: "left", margin: 0
  });

  const points = [
    "Copilot is een hulpmiddel — jij blijft verantwoordelijk voor het eindresultaat",
    "Controleer altijd informatie die Copilot genereert op juistheid",
    "Pas de toon en inhoud altijd aan op jouw specifieke situatie",
    "Gebruik AI verantwoord: let op privacy en zorgvuldigheid",
  ];
  bulletList(s, points, 0.7, 2.3, 7, 2.8, 14);
  // Override text color for this slide (bullets are on purple bg)
  const items = points.map((b, i) => ({
    text: b,
    options: { bullet: true, breakLine: i < points.length - 1 }
  }));
  // Overwrite with correct color
  s.addText(items, {
    x: 0.7, y: 2.3, w: 7, h: 2.8,
    fontSize: 14, color: C.LAVSUB, fontFace: "Calibri",
    valign: "top", paraSpaceAfter: 8
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE OPBOUW
// ══════════════════════════════════════════════════════════════════════════

// 1. Titelslide
addCoverSlide();

// 2. Agenda
addAgendaSlide();

// ── MODULE 1 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "01",
  "Introductie\nCopilot in\nDynamics",
  "AI-ondersteuning voor je dagelijkse CRM-werk.\n\nIn dit onderdeel leer je:\n• Wat Copilot is en hoe het werkt\n• Verschil met andere AI-tools\n• Wat Copilot kan doen in Dynamics"
);

addContentSlide("Wat is Copilot?", [
  "Copilot is een AI-assistent, ingebouwd in Microsoft-producten",
  "Copilot begrijpt tekst en context uit jouw CRM-data",
  "Het kan informatie samenvatten, antwoorden genereren en patronen herkennen",
  "Binnen Dynamics 365 werkt Copilot direct met klantdata uit het CRM",
  "Medewerkers houden altijd de controle — Copilot ondersteunt, beslist niet",
]);

addTableSlide("Copilot vs. andere AI-tools", [
  [
    { text: "Tool",              options: { bold: true, color: C.WHITE, fill: { color: C.PURPLE } } },
    { text: "Gebruik",           options: { bold: true, color: C.WHITE, fill: { color: C.PURPLE } } },
    { text: "Werkt met CRM?",    options: { bold: true, color: C.WHITE, fill: { color: C.PURPLE } } },
  ],
  [
    { text: "ChatGPT",           options: { color: C.DARK } },
    { text: "Algemene AI-chatbot",options: { color: C.DARK } },
    { text: "Nee",               options: { color: C.DARK } },
  ],
  [
    { text: "Microsoft 365 Copilot", options: { color: C.DARK } },
    { text: "Word, Outlook, Teams",  options: { color: C.DARK } },
    { text: "Beperkt",               options: { color: C.DARK } },
  ],
  [
    { text: "Dynamics 365 Copilot", options: { bold: true, color: C.PURPLE } },
    { text: "CRM — klanten, cases, e-mail", options: { bold: true, color: C.PURPLE } },
    { text: "Ja — direct vanuit het CRM", options: { bold: true, color: "1A6E38" } },
  ],
], [3.3, 3.5, 2.5]);

addDemoSlide("Demo: Copilot in actie", [
  "Open een bestaande case in Dynamics 365",
  "Klik op het Copilot-icoon (rechts in het scherm)",
  "Kies 'Samenvatten' — bekijk wat Copilot genereert",
  "Kies 'Antwoord voorstellen' — lees het concept door",
  "Pas het concept aan en stuur het bericht",
]);

// ── MODULE 2 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "02",
  "Copilot in\nklacht-\nafhandeling",
  "Snel overzicht bij complexe cases.\n\nVragen van leden, klachten over aannemers, advies over woningproblemen — Copilot helpt je de kern snel te vinden."
);

addContentSlide("Klachtafhandeling met Copilot", [
  "Veel cases bevatten lange beschrijvingen, meerdere e-mails en eerdere contactmomenten",
  "Copilot levert in seconden een overzicht van wat er speelt",
  "Mogelijke vervolgstappen worden automatisch voorgesteld",
  "Jij beoordeelt, past aan en beslist — Copilot levert het ruwe materiaal",
  "Tijdsbesparing: minder lezen, sneller de kern begrijpen",
]);

addDemoSlide("Demo: Case samenvatten", [
  "Open een klachtcase van een lid",
  "Klik op Copilot → 'Samenvatten'",
  "Bekijk de samenvatting: klopt dit met jouw verwachting?",
  "Klik op 'Vervolgstappen voorstellen'",
  "Beoordeel de suggesties en noteer eventuele aanpassingen",
]);

addExerciseSlide(
  "02",
  "Een lid meldt dat een aannemer slecht werk heeft geleverd en niet meer reageert. Het gaat om gevelbekleding die loslaat na isolatiewerkzaamheden.",
  [
    "Open de case in Dynamics 365",
    "Laat Copilot de klacht samenvatten",
    "Laat Copilot mogelijke vervolgstappen voorstellen",
    "Bespreek met een collega: wat klopt wel/niet in de samenvatting?",
  ]
);

// ── MODULE 3 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "03",
  "E-mail-\nafhandeling\nmet Copilot",
  "Sneller en consistenter communiceren.\n\nCopilot analyseert inkomende e-mails, stelt antwoorden voor en past de toon aan — jij stuurt altijd zelf."
);

addContentSlide("E-mailafhandeling met Copilot", [
  "Medewerkers besteden veel tijd aan e-mails lezen en beantwoorden",
  "Copilot vat de inhoud van een e-mail in één of twee zinnen samen",
  "Een conceptantwoord wordt automatisch opgesteld op basis van de context",
  "De toon is aanpasbaar: formeel, begripvol, kort & krachtig",
  "Herschrijf-optie: eenvoudiger, uitgebreider of in andere stijl",
]);

addDemoSlide("Demo: E-mail beantwoorden", [
  "Open een inkomende e-mail over isolatiesubsidie in Dynamics",
  "Klik op Copilot → 'E-mail samenvatten'",
  "Klik op 'Antwoord genereren'",
  "Pas toon aan: 'Begripvol en professioneel'",
  "Herschrijf het antwoord eenvoudiger voor het lid",
]);

addExerciseSlide(
  "03",
  "\"Ik wil mijn woning isoleren en ik hoor dat er subsidies zijn. Kunt u mij vertellen wat ik kan aanvragen en hoe dat werkt?\"",
  [
    "Laat Copilot de vraag samenvatten",
    "Laat Copilot een antwoord schrijven",
    "Pas de toon aan: eenvoudiger en vriendelijker",
    "Vergelijk jouw handmatige versie met de Copilot-versie",
  ]
);

// ── MODULE 4 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "04",
  "Marketing\nen commu-\nnicatie",
  "Content genereren met AI.\n\nCopilot helpt bij nieuwsbrieven, adviesberichten en marketingteksten — consistent én snel."
);

addContentSlide("Marketing met Copilot", [
  "VEH communiceert met leden over advies, diensten en actualiteiten",
  "Copilot genereert onderwerpen voor een nieuwsbrief op basis van een thema",
  "Marketingteksten: Copilot maakt een eerste versie, jij verfijnt",
  "Toon aanpasbaar aan doelgroep: huiseigenaar vs. VvE",
  "Ideeën genereren voor campagnes, berichten of sociale media",
]);

addDemoSlide("Demo: Nieuwsbrief content genereren", [
  "Open Copilot in de marketingmodule van Dynamics",
  "Geef als invoer: thema 'energiebesparing woningen'",
  "Laat Copilot 5 nieuwsbriefonderwerpen genereren",
  "Kies een onderwerp en laat een korte tekst uitwerken",
  "Pas de tekst aan op VEH-huisstijl en toon",
]);

addExerciseSlide(
  "04",
  "Maak een korte nieuwsbrief van maximaal 150 woorden over energiebesparing in woningen voor VEH-leden.",
  [
    "Geef Copilot het thema: energiebesparing",
    "Laat Copilot een nieuwsbriefartikel schrijven",
    "Pas de toon aan: vriendelijk, informatief, actiegericht",
    "Voeg een concrete tip of call-to-action toe",
  ]
);

// ── MODULE 5 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "05",
  "Marketing-\nsegmentatie\nmet Copilot",
  "De juiste boodschap voor de juiste doelgroep.\n\nCopilot herkent patronen in CRM-data en helpt je slimme segmenten te bouwen voor gerichte communicatie."
);

addContentSlide("Segmentatie met Copilot", [
  "CRM-systemen bevatten veel data over gedrag, interesses en contacthistorie",
  "Copilot herkent patronen: wie opent welke berichten? Wie stelde welke vragen?",
  "Segmenten aanmaken op basis van gedragskenmerken",
  "Bijv: leden die de nieuwsbrief openen, of leden met isolatievragen",
  "Betere targeting = hogere relevantie = tevreden leden",
]);

addDemoSlide("Demo: Segment aanmaken", [
  "Open Dynamics → Marketing → Segmenten",
  "Klik op 'Nieuw segment met Copilot'",
  "Beschrijf het gewenste segment in gewone taal",
  "Copilot genereert het segment op basis van beschikbare data",
  "Controleer het segment en activeer het voor een campagne",
]);

addExerciseSlide(
  "05",
  "Maak een segment van leden die in de afgelopen 6 maanden interesse hebben getoond in energiebesparing (via e-mail, cases of website-gedrag).",
  [
    "Beschrijf het segment aan Copilot in gewone taal",
    "Laat Copilot het segment opbouwen",
    "Controleer of de criteria logisch zijn",
    "Bedenk welke campagne je op dit segment zou richten",
  ]
);

// ── MODULE 6 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "06",
  "Prompt-\ntechnieken",
  "Betere vragen = betere antwoorden.\n\nDe kwaliteit van AI-output hangt sterk af van hoe je de vraag stelt. Leer de structuur van een effectieve prompt."
);

addContentSlide("Anatomie van een goede prompt", [
  "Context: wie ben je en wat is de situatie? (bijv. 'Je bent adviseur bij VEH')",
  "Taak: wat moet Copilot doen? Wees concreet en specifiek",
  "Structuur: hoe moet het antwoord eruitzien? (opsomming, alinea's, max. woorden)",
  "Toon: professioneel, begripvol, eenvoudig, formeel?",
  "Itereren: verfijn de prompt als het resultaat niet goed genoeg is",
]);

addTwoColSlide(
  "Promptvergelijking: zwak vs. sterk",
  "Zwakke prompt",
  [
    "Schrijf een antwoord",
    "(Geen context gegeven)",
    "(Geen toon bepaald)",
    "(Geen structuur gevraagd)",
    "Resultaat: generiek en onpersoonlijk",
  ],
  "Sterke prompt",
  [
    "Je bent adviseur bij Vereniging Eigen Huis.",
    "Schrijf een professioneel antwoord op deze klacht:",
    "- Toon begrip voor de situatie van het lid",
    "- Geef een korte uitleg van de volgende stappen",
    "- Sluit positief af (max. 150 woorden)",
  ]
);

addExerciseSlide(
  "06",
  "Hieronder staan drie zwakke prompts. Verbeter ze tot sterke prompts met context, taak en structuur.",
  [
    "Prompt 1: 'Maak een samenvatting' → verbeter",
    "Prompt 2: 'Schrijf een mail' → verbeter",
    "Prompt 3: 'Wat moet ik doen?' → verbeter",
    "Wissel uit met een collega: welke prompt levert het beste resultaat?",
  ]
);

// ── MODULE 7 ──────────────────────────────────────────────────────────────
addSectionSlide(
  "07",
  "Eindcase",
  "Alles samen in één realistische werksituatie.\n\nJe doorloopt het volledige proces: case analyseren, antwoord opstellen, toon aanpassen — met Copilot als jouw assistent."
);

addContentSlide("Eindcase: isolatieconflict", [
  "Scenario: Een lid meldt een conflict met zijn aannemer over isolatiewerkzaamheden",
  "De aannemer werkt niet naar behoren en reageert nauwelijks meer op berichten",
  "Het lid vraagt om advies over zijn rechten en mogelijke vervolgstappen",
  "Jouw taak: analyseren, samenvatten, advies geven en professioneel communiceren",
  "Gebruik Copilot bij élke stap — maar beoordeel en pas aan",
]);

// Eindcase stappenslide (handmatig gebouwd)
{
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, "Stappenplan eindcase");

  const steps = [
    ["1", "Klacht samenvatten",     "Gebruik Copilot: kern van het probleem in 3 zinnen"],
    ["2", "Probleem analyseren",     "Aandachtspunten en rechten van het lid benoemen"],
    ["3", "Antwoordmail genereren", "Professioneel antwoord via Copilot opstellen"],
    ["4", "Toon aanpassen",          "Herschrijven: begripvol, helder en actiegericht"],
  ];

  const stepH  = 0.55;
  const startY = 0.9;
  const gap    = 1.05;

  steps.forEach(([num, title, desc], i) => {
    const y = startY + i * gap;

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.48, h: stepH,
      fill: { color: C.ORANGE }, line: { color: C.ORANGE }
    });
    s.addText(num, {
      x: 0.4, y, w: 0.48, h: stepH,
      fontSize: 16, bold: true, color: C.WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    s.addText(title, {
      x: 1.1, y: y + 0.05, w: 3.3, h: stepH - 0.1,
      fontSize: 16, bold: true, color: C.DARK, fontFace: "Calibri",
      align: "left", valign: "middle", margin: 0
    });
    s.addText(desc, {
      x: 4.6, y: y + 0.08, w: 5.1, h: stepH - 0.16,
      fontSize: 13, color: C.MUTED, fontFace: "Calibri",
      align: "left", valign: "middle", margin: 0
    });

    // Vertical connector
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.615, y: y + stepH, w: 0.05, h: gap - stepH,
        fill: { color: C.ORANGE }, line: { color: C.ORANGE }
      });
    }
  });
}

// Afsluiting
addClosingSlide();

// ─── Opslaan ──────────────────────────────────────────────────────────────
pres
  .writeFile({ fileName: "Presentatie_Copilot_D365.pptx" })
  .then(() => console.log("Presentatie opgeslagen: Presentatie_Copilot_D365.pptx"))
  .catch(err  => { console.error("Fout bij genereren:", err); process.exit(1); });
