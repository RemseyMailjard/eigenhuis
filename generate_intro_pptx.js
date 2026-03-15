"use strict";
const PptxGenJS = require("pptxgenjs");

// ─── Kleuren VEH huisstijl ─────────────────────────────────────────────────
const C = {
  PURPLE: "3B2785",
  ORANGE: "F07800",
  WHITE: "FFFFFF",
  LIGHT: "F5F5F5",
  DARK: "1A1A4E",
  GRAY: "DDDDDD",
  LILAC: "BEB5D9",
  LAVSUB: "D5CCEE",
  DARK2: "2D1D6E",
};

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";
pres.author = "Vereniging Eigen Huis";
pres.title = "Introductie Copilot – Dynamics 365 CE";

const W = 10,
  H = 5.625;

// ─── Helpers ──────────────────────────────────────────────────────────────

function topBar(s, title, color) {
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: W,
    h: 0.65,
    fill: { color: color || C.PURPLE },
    line: { color: color || C.PURPLE },
  });
  s.addText(title, {
    x: 0.5,
    y: 0,
    w: 9,
    h: 0.65,
    fontSize: 22,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    valign: "middle",
    margin: 0,
  });
}

function bulletList(s, bullets, x, y, w, h, fontSize, color) {
  const items = bullets.map((b, i) => ({
    text: b,
    options: { bullet: true, breakLine: i < bullets.length - 1 },
  }));
  s.addText(items, {
    x,
    y,
    w,
    h,
    fontSize: fontSize || 15,
    color: color || C.DARK,
    fontFace: "Calibri",
    valign: "top",
    paraSpaceAfter: 8,
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 1 – TITELSLIDE
// ══════════════════════════════════════════════════════════════════════════
function addCoverSlide() {
  const s = pres.addSlide();
  s.background = { color: C.PURPLE };

  // Orange bottom accent
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: H - 0.1,
    w: W,
    h: 0.1,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Right decorative panel
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.8,
    y: 0,
    w: 2.2,
    h: H,
    fill: { color: C.DARK2 },
    line: { color: C.DARK2 },
  });

  // Orange circle accent top-right
  s.addShape(pres.shapes.OVAL, {
    x: 8.3,
    y: 0.3,
    w: 1.2,
    h: 1.2,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Small orange circle bottom-right
  s.addShape(pres.shapes.OVAL, {
    x: 8.7,
    y: 3.8,
    w: 0.7,
    h: 0.7,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Main title
  s.addText("Introductie\nCopilot", {
    x: 0.7,
    y: 0.7,
    w: 6.8,
    h: 2.3,
    fontSize: 52,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    valign: "top",
    lineSpacingMultiple: 1.1,
    margin: 0,
  });

  // Subtitle
  s.addText("Dynamics 365 CE", {
    x: 0.7,
    y: 3.1,
    w: 6.8,
    h: 0.55,
    fontSize: 22,
    color: C.ORANGE,
    fontFace: "Calibri",
    bold: true,
    align: "left",
    margin: 0,
  });

  // Tagline
  s.addText("Slimmer, sneller en beter communiceren met AI", {
    x: 0.7,
    y: 3.75,
    w: 6.8,
    h: 0.38,
    fontSize: 13,
    color: C.LILAC,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  // VEH label bottom left
  s.addText("vereniging eigen huis", {
    x: 0.7,
    y: H - 0.45,
    w: 5,
    h: 0.3,
    fontSize: 11,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 2 – WAT GA JE LEREN? (3 pillars)
// ══════════════════════════════════════════════════════════════════════════
function addWatGaJeLerenSlide() {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, "Wat ga je leren?");

  // Orange accent line underneath topbar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0.65,
    w: W,
    h: 0.05,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  const pillars = [
    { icon: "⚡", title: "Sneller werken", color: C.DARK2 },
    { icon: "✍️", title: "Beter communiceren", color: C.PURPLE },
    { icon: "🧠", title: "Slimmer werken met AI", color: C.ORANGE },
  ];

  const cardW = 2.8;
  const cardH = 3.6;
  const startX = 0.55;
  const gapX = 0.3;
  const cardY = 0.95;

  pillars.forEach((p, i) => {
    const x = startX + i * (cardW + gapX);

    // Card background
    s.addShape(pres.shapes.RECTANGLE, {
      x,
      y: cardY,
      w: cardW,
      h: cardH,
      fill: { color: C.WHITE },
      line: { color: C.GRAY, pt: 1 },
    });

    // Top color bar
    s.addShape(pres.shapes.RECTANGLE, {
      x,
      y: cardY,
      w: cardW,
      h: 0.08,
      fill: { color: p.color },
      line: { color: p.color },
    });

    // Icon circle
    s.addShape(pres.shapes.OVAL, {
      x: x + cardW / 2 - 0.45,
      y: cardY + 0.25,
      w: 0.9,
      h: 0.9,
      fill: { color: p.color === C.ORANGE ? C.ORANGE : C.ORANGE },
      line: { color: p.color === C.ORANGE ? C.ORANGE : C.ORANGE },
    });
    s.addText(p.icon, {
      x: x + cardW / 2 - 0.45,
      y: cardY + 0.25,
      w: 0.9,
      h: 0.9,
      fontSize: 20,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // Pillar title
    s.addText(p.title, {
      x: x + 0.15,
      y: cardY + 1.3,
      w: cardW - 0.3,
      h: 0.7,
      fontSize: 16,
      bold: true,
      color: p.color,
      fontFace: "Calibri",
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // Thin divider
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.4,
      y: cardY + 2.1,
      w: cardW - 0.8,
      h: 0.03,
      fill: { color: C.GRAY },
      line: { color: C.GRAY },
    });
  });

  // Subtopics per pillar row
  const subtopics = [
    ["Samenvatten van klantcontact", "Overzicht van klantcases"],
    ["Conceptantwoorden genereren", "Toon aanpassen"],
    ["Effectieve prompts schrijven"],
  ];

  pillars.forEach((p, i) => {
    const x = startX + i * (cardW + gapX);
    const subItems = subtopics[i].map((t, j) => ({
      text: t,
      options: { bullet: true, breakLine: j < subtopics[i].length - 1 },
    }));
    s.addText(subItems, {
      x: x + 0.2,
      y: cardY + 2.22,
      w: cardW - 0.4,
      h: 1.2,
      fontSize: 12,
      color: C.DARK,
      fontFace: "Calibri",
      valign: "top",
      paraSpaceAfter: 6,
      margin: 0,
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 3 – PILLAR 1: SNELLER WERKEN
// ══════════════════════════════════════════════════════════════════════════
function addSnellerWerkenSlide() {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };

  // Left panel dark
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 3.8,
    h: H,
    fill: { color: C.DARK2 },
    line: { color: C.DARK2 },
  });

  // Orange accent bottom-left
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: H - 0.1,
    w: 3.8,
    h: 0.1,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Icon
  s.addText("⚡", {
    x: 0.3,
    y: 0.5,
    w: 1.5,
    h: 1.2,
    fontSize: 54,
    align: "left",
    valign: "middle",
    margin: 0,
  });

  // Pillar title on left
  s.addText("Sneller\nwerken", {
    x: 0.3,
    y: 1.7,
    w: 3.2,
    h: 1.5,
    fontSize: 36,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    valign: "top",
    margin: 0,
  });

  // Tag on left
  s.addText("01 van 03", {
    x: 0.3,
    y: H - 0.55,
    w: 3.2,
    h: 0.35,
    fontSize: 11,
    color: C.LILAC,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  // Right orange accent divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.8,
    y: 0,
    w: 0.07,
    h: H,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Right content: context
  s.addText("Wat leer je?", {
    x: 4.2,
    y: 0.5,
    w: 5.4,
    h: 0.45,
    fontSize: 18,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  // Sub-bullet 1 box: Samenvatten
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 1.15,
    w: 5.4,
    h: 1.2,
    fill: { color: C.LIGHT },
    line: { color: C.GRAY, pt: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 1.15,
    w: 0.07,
    h: 1.2,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });
  s.addText("Samenvatten van klantcontact", {
    x: 4.45,
    y: 1.22,
    w: 4.9,
    h: 0.38,
    fontSize: 15,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
  s.addText(
    "Lange gesprekshistorie in één klik samengevat — altijd meteen de kern.",
    {
      x: 4.45,
      y: 1.62,
      w: 4.9,
      h: 0.6,
      fontSize: 12,
      color: C.DARK,
      fontFace: "Calibri",
      align: "left",
      margin: 0,
    },
  );

  // Sub-bullet 2 box: Overzicht klantcases
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 2.55,
    w: 5.4,
    h: 1.2,
    fill: { color: C.LIGHT },
    line: { color: C.GRAY, pt: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 2.55,
    w: 0.07,
    h: 1.2,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });
  s.addText("Overzicht van klantcases", {
    x: 4.45,
    y: 2.62,
    w: 4.9,
    h: 0.38,
    fontSize: 15,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
  s.addText("Direct zicht op openstaande cases, prioriteiten en voortgang.", {
    x: 4.45,
    y: 3.02,
    w: 4.9,
    h: 0.6,
    fontSize: 12,
    color: C.DARK,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  s.addText("Bespaar tijd op routinewerk — focus op wat écht telt.", {
    x: 4.2,
    y: 4.1,
    w: 5.4,
    h: 0.45,
    fontSize: 13,
    italic: true,
    color: C.PURPLE,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 4 – PILLAR 2: BETER COMMUNICEREN
// ══════════════════════════════════════════════════════════════════════════
function addBeterCommunicerenSlide() {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };

  // Left panel purple
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 3.8,
    h: H,
    fill: { color: C.PURPLE },
    line: { color: C.PURPLE },
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: H - 0.1,
    w: 3.8,
    h: 0.1,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  s.addText("✍️", {
    x: 0.3,
    y: 0.5,
    w: 1.5,
    h: 1.2,
    fontSize: 54,
    align: "left",
    valign: "middle",
    margin: 0,
  });

  s.addText("Beter\ncommu-\nniceren", {
    x: 0.3,
    y: 1.7,
    w: 3.2,
    h: 2.0,
    fontSize: 34,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    valign: "top",
    margin: 0,
  });

  s.addText("02 van 03", {
    x: 0.3,
    y: H - 0.55,
    w: 3.2,
    h: 0.35,
    fontSize: 11,
    color: C.LILAC,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.8,
    y: 0,
    w: 0.07,
    h: H,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  s.addText("Wat leer je?", {
    x: 4.2,
    y: 0.5,
    w: 5.4,
    h: 0.45,
    fontSize: 18,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  // Sub-bullet 1: Conceptantwoorden
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 1.15,
    w: 5.4,
    h: 1.2,
    fill: { color: C.LIGHT },
    line: { color: C.GRAY, pt: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 1.15,
    w: 0.07,
    h: 1.2,
    fill: { color: C.PURPLE },
    line: { color: C.PURPLE },
  });
  s.addText("Conceptantwoorden genereren", {
    x: 4.45,
    y: 1.22,
    w: 4.9,
    h: 0.38,
    fontSize: 15,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
  s.addText(
    "Copilot schrijft een eerste versie op basis van de case — jij past aan en verstuurt.",
    {
      x: 4.45,
      y: 1.62,
      w: 4.9,
      h: 0.6,
      fontSize: 12,
      color: C.DARK,
      fontFace: "Calibri",
      align: "left",
      margin: 0,
    },
  );

  // Sub-bullet 2: Toon aanpassen
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 2.55,
    w: 5.4,
    h: 1.2,
    fill: { color: C.LIGHT },
    line: { color: C.GRAY, pt: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 2.55,
    w: 0.07,
    h: 1.2,
    fill: { color: C.PURPLE },
    line: { color: C.PURPLE },
  });
  s.addText("Toon aanpassen", {
    x: 4.45,
    y: 2.62,
    w: 4.9,
    h: 0.38,
    fontSize: 15,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
  s.addText(
    "Formeel, begripvol of eenvoudig — één klik bepaalt de stijl van je bericht.",
    {
      x: 4.45,
      y: 3.02,
      w: 4.9,
      h: 0.6,
      fontSize: 12,
      color: C.DARK,
      fontFace: "Calibri",
      align: "left",
      margin: 0,
    },
  );

  s.addText("Consistente communicatie — ook in drukke periodes.", {
    x: 4.2,
    y: 4.1,
    w: 5.4,
    h: 0.45,
    fontSize: 13,
    italic: true,
    color: C.PURPLE,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 5 – PILLAR 3: SLIMMER WERKEN MET AI
// ══════════════════════════════════════════════════════════════════════════
function addSlimmerWerkenSlide() {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };

  // Left panel orange
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 3.8,
    h: H,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: H - 0.1,
    w: 3.8,
    h: 0.1,
    fill: { color: C.DARK2 },
    line: { color: C.DARK2 },
  });

  s.addText("🧠", {
    x: 0.3,
    y: 0.5,
    w: 1.5,
    h: 1.2,
    fontSize: 54,
    align: "left",
    valign: "middle",
    margin: 0,
  });

  s.addText("Slimmer\nwerken\nmet AI", {
    x: 0.3,
    y: 1.7,
    w: 3.2,
    h: 2.0,
    fontSize: 32,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    valign: "top",
    margin: 0,
  });

  s.addText("03 van 03", {
    x: 0.3,
    y: H - 0.55,
    w: 3.2,
    h: 0.35,
    fontSize: 11,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.8,
    y: 0,
    w: 0.07,
    h: H,
    fill: { color: C.DARK2 },
    line: { color: C.DARK2 },
  });

  s.addText("Wat leer je?", {
    x: 4.2,
    y: 0.5,
    w: 5.4,
    h: 0.45,
    fontSize: 18,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  // Main box: Effectieve prompts schrijven
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 1.15,
    w: 5.4,
    h: 1.5,
    fill: { color: C.LIGHT },
    line: { color: C.GRAY, pt: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.2,
    y: 1.15,
    w: 0.07,
    h: 1.5,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });
  s.addText("Effectieve prompts schrijven", {
    x: 4.45,
    y: 1.22,
    w: 4.9,
    h: 0.38,
    fontSize: 15,
    bold: true,
    color: C.DARK2,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
  s.addText(
    "Leer de anatomie van een goede prompt:\ncontext · taak · structuur · toon",
    {
      x: 4.45,
      y: 1.65,
      w: 4.9,
      h: 0.85,
      fontSize: 12,
      color: C.DARK,
      fontFace: "Calibri",
      align: "left",
      margin: 0,
    },
  );

  // Prompt formula visual
  const labels = ["Context", "Taak", "Structuur", "Toon"];
  const bColors = [C.DARK2, C.PURPLE, C.DARK2, C.ORANGE];
  const boxW = 1.15,
    boxH = 0.55,
    startX = 4.2,
    startY = 2.9,
    gapX = 0.2;

  labels.forEach((lbl, i) => {
    const bx = startX + i * (boxW + gapX);
    s.addShape(pres.shapes.RECTANGLE, {
      x: bx,
      y: startY,
      w: boxW,
      h: boxH,
      fill: { color: bColors[i] },
      line: { color: bColors[i] },
    });
    s.addText(lbl, {
      x: bx,
      y: startY,
      w: boxW,
      h: boxH,
      fontSize: 13,
      bold: true,
      color: C.WHITE,
      fontFace: "Calibri",
      align: "center",
      valign: "middle",
      margin: 0,
    });
    // Plus sign between
    if (i < labels.length - 1) {
      s.addText("+", {
        x: bx + boxW + 0.02,
        y: startY,
        w: 0.18,
        h: boxH,
        fontSize: 16,
        bold: true,
        color: C.DARK,
        fontFace: "Calibri",
        align: "center",
        valign: "middle",
        margin: 0,
      });
    }
  });

  s.addText("Betere vragen = betere antwoorden van Copilot.", {
    x: 4.2,
    y: 3.7,
    w: 5.4,
    h: 0.45,
    fontSize: 13,
    italic: true,
    color: C.ORANGE,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });

  s.addText("Je AI-vaardigheden blijven groeien — ook na deze training.", {
    x: 4.2,
    y: 4.2,
    w: 5.4,
    h: 0.38,
    fontSize: 11,
    color: C.DARK,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 6 – OVERZICHTSSLIDE (recapitulatie van alle 3)
// ══════════════════════════════════════════════════════════════════════════
function addRecapSlide() {
  const s = pres.addSlide();
  s.background = { color: C.LIGHT };
  topBar(s, "Jouw leerdoelen in één oogopslag");

  const rows = [
    [
      {
        text: "Thema",
        options: { bold: true, color: C.WHITE, fill: { color: C.PURPLE } },
      },
      {
        text: "Onderdelen",
        options: { bold: true, color: C.WHITE, fill: { color: C.PURPLE } },
      },
      {
        text: "Resultaat",
        options: { bold: true, color: C.WHITE, fill: { color: C.PURPLE } },
      },
    ],
    [
      { text: "⚡ Sneller werken", options: { bold: true, color: C.DARK2 } },
      { text: "Samenvatten · Klantoverzicht", options: { color: C.DARK } },
      {
        text: "Minder tijd kwijt aan zoeken en lezen",
        options: { color: C.DARK },
      },
    ],
    [
      {
        text: "✍️ Beter communiceren",
        options: { bold: true, color: C.DARK2 },
      },
      { text: "Conceptantwoorden · Toon", options: { color: C.DARK } },
      {
        text: "Consistente en snelle klantcommunicatie",
        options: { color: C.DARK },
      },
    ],
    [
      {
        text: "🧠 Slimmer werken met AI",
        options: { bold: true, color: C.DARK2 },
      },
      { text: "Prompttechnieken", options: { color: C.DARK } },
      {
        text: "Copilot effectiever en gerichter inzetten",
        options: { color: C.DARK },
      },
    ],
  ];

  s.addTable(rows, {
    x: 0.5,
    y: 0.85,
    w: 9.0,
    colW: [2.5, 3.0, 3.5],
    border: { pt: 1, color: C.GRAY },
    rowH: 0.68,
  });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5,
    y: 4.45,
    w: 9.0,
    h: 0.78,
    fill: { color: C.PURPLE },
    line: { color: C.PURPLE },
  });
  s.addText(
    "Na deze training werk jij met AI — niet voor AI. Copilot ondersteunt, jij beslist.",
    {
      x: 0.5,
      y: 4.45,
      w: 9.0,
      h: 0.78,
      fontSize: 14,
      bold: true,
      color: C.WHITE,
      fontFace: "Calibri",
      align: "center",
      valign: "middle",
      margin: 0,
    },
  );
}

// ══════════════════════════════════════════════════════════════════════════
// SLIDE 7 – AFSLUITING / LET'S BEGIN
// ══════════════════════════════════════════════════════════════════════════
function addStartSlide() {
  const s = pres.addSlide();
  s.background = { color: C.PURPLE };

  // Right decorative
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.8,
    y: 0,
    w: 2.2,
    h: H,
    fill: { color: C.DARK2 },
    line: { color: C.DARK2 },
  });

  // Orange oval
  s.addShape(pres.shapes.OVAL, {
    x: 8.3,
    y: 0.3,
    w: 1.2,
    h: 1.2,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Bottom accent
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0,
    y: H - 0.1,
    w: W,
    h: 0.1,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });

  // Big CTA text
  s.addText("Klaar\nom te\nbeginnen?", {
    x: 0.7,
    y: 0.5,
    w: 7,
    h: 3.0,
    fontSize: 52,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    valign: "top",
    lineSpacingMultiple: 1.1,
    margin: 0,
  });

  // CTA pill
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7,
    y: 3.8,
    w: 3.5,
    h: 0.65,
    fill: { color: C.ORANGE },
    line: { color: C.ORANGE },
  });
  s.addText("▶  Start de training", {
    x: 0.7,
    y: 3.8,
    w: 3.5,
    h: 0.65,
    fontSize: 16,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "center",
    valign: "middle",
    margin: 0,
  });

  s.addText("vereniging eigen huis", {
    x: 0.7,
    y: H - 0.45,
    w: 5,
    h: 0.3,
    fontSize: 11,
    bold: true,
    color: C.WHITE,
    fontFace: "Calibri",
    align: "left",
    margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════
// OPBOUW
// ══════════════════════════════════════════════════════════════════════════
addCoverSlide();
addWatGaJeLerenSlide();
addSnellerWerkenSlide();
addBeterCommunicerenSlide();
addSlimmerWerkenSlide();
addRecapSlide();
addStartSlide();

// ══════════════════════════════════════════════════════════════════════════
// OPSLAAN
// ══════════════════════════════════════════════════════════════════════════
pres
  .writeFile({ fileName: "Introductie_Copilot_D365.pptx" })
  .then(() =>
    console.log("Presentatie opgeslagen: Introductie_Copilot_D365.pptx"),
  )
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
