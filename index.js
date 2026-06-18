const express = require("express");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const app = express();
app.use(express.json());

// ─── LOGO (base64 embedded) ───────────────────────────────────────
const LOGO_B64 = "image/png;base64,iVBORw0KGgoAAAANSUhEUgAABdwAAAGQCAYAAAC04KJgAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9ArFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTmlyQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKeDHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAEAAElEQVR4nOz9SbAk2X3fe37PcfeIO+WcVYVCFWoAUIWpSIAEAQLCSGIgSEKiic+MMm1kJutmehtxKS2khcykRW/atNNGC8msTZSZus3U0muOT2SDFEBCAAFiIAaCBIgZqDnHeyPC3c/59+Kc4+4R92ahKvMCmbfy9ynLulMMHh4RHu6/8/f/cYAhIiIiIiIiIiIiIiK3xN/uBRAREREREREREREReTlQ4C4iIiIiIiIiIiIicgwUuIuIiIiIiIiIiIiIHAMF7iIiIiIiIiIiIiIix0CBu4iIiIiIiIiIiIjIMVDgLiIiIiIiIiIiIiJyDBS4i4iIiIiIiIiIiIgcAwXuIiIiIiIiIiIiIiLHQIG7iIiIiIiIiIiIiMgxUOAuIiIiIiIiIiIiInIMFLiLiIiIiIiIiIiIiBwDBe4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIiIsdAgbuIiIiIiIiIiIiIyDFQ4C4iIiIiIiIiIiIicqju9gKIiIiIiLxYjrFiJOav5eeABxfHC9v69ap82bBxfZteeLjuWJdSfu2H68TpTR++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") + "iVBORw0KGgoAAAANSUhEUgAAA+gAAAPoCAYAAABNo9TkAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA3XAAAN1wFCKJt4AAAAGXRFWHRTb2Z0d2FyZQBQYWludC5ORVQgdjMuNS4xMDT0clkAAA==";

// ─── BRAND PALETTE ────────────────────────────────────────────────
const C = {
  white:    "FFFFFF",
  offWhite: "F0F0F0",
  muted:    "666666",
  black:    "000000",
  red:      "FE0000",
  redLight: "1A0000",
  cardBg:   "141414",
  dark2:    "222222",
  dark3:    "333333",
  dark5:    "555555",
};

// ─── ICON HELPER ─────────────────────────────────────────────────
async function iconBase64(iconPath, color = "#FE0000", size = 256) {
  try {
    const [lib, name] = iconPath.split("/");
    const icons = require(`react-icons/${lib}`);
    const Icon = icons[name];
    if (!Icon) throw new Error(`Icon not found: ${iconPath}`);
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Icon, { color, size: String(size) })
    );
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
  } catch (err) {
    // Fallback: red circle
    const fallback = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${size/2}" cy="${size/2}" r="${size/3}" fill="${color}"/></svg>`;
    const buf = await sharp(Buffer.from(fallback)).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
  }
}

// ─── SAFE ICON NAME ───────────────────────────────────────────────
// Claude sometimes returns names without the "Fa" prefix — fix them
function safeFaIcon(name) {
  if (!name) return "FaRocket";
  const cleaned = name.trim();
  // Already correct format
  if (/^Fa[A-Z]/.test(cleaned)) return cleaned;
  // Missing prefix
  if (/^[A-Z]/.test(cleaned)) return "Fa" + cleaned;
  // Lowercase — capitalise first char and prefix
  return "Fa" + cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
}

// ─── SHADOW PRESET ───────────────────────────────────────────────
const mkShadow = () => ({
  type: "outer", color: "000000",
  blur: 8, offset: 2, angle: 45, opacity: 0.5,
});

// ─── REUSABLE ELEMENTS ───────────────────────────────────────────
function redLine(slide, x, y, w = 1.0) {
  slide.addShape("rect", {
    x, y, w, h: 0.04,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
  });
}

function topBar(slide, label) {
  // Black bar
  slide.addShape("rect", {
    x: 0, y: 0, w: 10, h: 0.62,
    fill: { color: C.black }, line: { color: C.black, width: 0 },
  });
  // Red underline
  slide.addShape("rect", {
    x: 0, y: 0.60, w: 10, h: 0.04,
    fill: { color: C.red }, line: { color: C.red, width: 0 },
  });
  // Logo
  slide.addImage({ data: LOGO_B64, x: 0.4, y: 0.06, w: 2.4, h: 0.5 });
  // Section label
  if (label) {
    slide.addText(label, {
      x: 6.5, y: 0.18, w: 3.3, h: 0.26,
      fontSize: 7.5, fontFace: "Arial", color: C.dark5,
      bold: true, charSpacing: 3, align: "right", margin: 0,
    });
  }
}

function statCard(slide, stat, label, x, y) {
  const w = 2.85, h = 1.28;
  slide.addShape("roundRect", {
    x, y, w, h,
    fill: { color: C.cardBg },
    line: { color: C.red, width: 1.2 },
    rectRadius: 0.06,
    shadow: mkShadow(),
  });
  // Dynamic font size for stat
  const statLen = String(stat).length;
  const statSize = statLen > 8 ? 14 : statLen > 5 ? 18 : 20;
  slide.addText(stat, {
    x: x + 0.1, y: y + 0.08, w: w - 0.2, h: 0.76,
    fontSize: statSize, fontFace: "Arial", color: C.red,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
  slide.addText(label, {
    x: x + 0.1, y: y + 0.88, w: w - 0.2, h: 0.30,
    fontSize: 9, fontFace: "Arial", color: C.muted,
    align: "center", margin: 0,
  });
}

// ─── CITY SKYLINE HELPER ─────────────────────────────────────────
function addSkyline(slide, startX, baseY, buildings, opacity = 1) {
  buildings.forEach(b => {
    slide.addShape("rect", {
      x: startX + b.ox, y: baseY - b.h,
      w: b.w, h: b.h,
      fill: { color: C.dark2 },
      line: { color: C.dark2, width: 0 },
    });
  });
}

const SKYLINE_RIGHT = [
  { ox: 0.0,  h: 3.3, w: 0.15 },
  { ox: 0.25, h: 2.9, w: 0.22 },
  { ox: 0.56, h: 4.1, w: 0.17 },
  { ox: 0.82, h: 2.7, w: 0.26 },
  { ox: 1.17, h: 3.2, w: 0.20 },
  { ox: 1.46, h: 2.4, w: 0.28 },
  { ox: 1.82, h: 2.8, w: 0.20 },
  { ox: 2.10, h: 2.2, w: 0.30 },
  { ox: 2.48, h: 2.6, w: 0.22 },
  { ox: 2.78, h: 2.1, w: 0.28 },
  { ox: 3.14, h: 2.4, w: 0.20 },
];

// ─── DECK GENERATOR ──────────────────────────────────────────────
async function generateDeck(data) {
  const {
    company_name = "Client",
    website = "",
    audit_score = "N/A",
    headline = "Here is exactly what we found on your website.",
    problem_1 = "Problem 1",
    problem_2 = "Problem 2",
    problem_3 = "Problem 3",
    solution_1 = "Solution 1",
    solution_2 = "Solution 2",
    solution_3 = "Solution 3",
    solution_icon_1,
    solution_icon_2,
    solution_icon_3,
    proof_stat_1 = "620+",
    proof_label_1 = "Direct reservations in 30 days",
    proof_stat_2 = "24K",
    proof_label_2 = "New YouTube subscribers in 60 days",
    proof_stat_3 = "200+",
    proof_label_3 = "WhatsApp leads in 60 days",
    offer_line = "We will fix this in 30 days or you don't pay.",
    cta_line = "BOOK YOUR FREE STRATEGY CALL",
    industry = "",
  } = data;

  // Resolve icons
  const icons = [
    "fa/" + safeFaIcon(solution_icon_1 || "FaRocket"),
    "fa/" + safeFaIcon(solution_icon_2 || "FaChartLine"),
    "fa/" + safeFaIcon(solution_icon_3 || "FaBullseye"),
  ];

  // Preload solution icons (white on dark cards)
  const sLoaded = await Promise.all(icons.map(i => iconBase64(i, "#FFFFFF", 128)));

  // Preload problem icons
  const pIcons = ["fa/FaExclamationTriangle", "fa/FaTimesCircle", "fa/FaExclamationCircle"];
  const pLoaded = await Promise.all(pIcons.map(i => iconBase64(i, "#FE0000")));

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // ────────────────────────────────────────────────
  // SLIDE 1 — COVER
  // ────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };

    // Logo
    s.addImage({ data: LOGO_B64, x: 0.5, y: 0.35, w: 3.2, h: 0.72 });
    redLine(s, 0.5, 1.22, 1.5);

    s.addText("BRAND GROWTH AGENCY · DUBAI", {
      x: 0.5, y: 1.34, w: 5.5, h: 0.24,
      fontSize: 8, fontFace: "Arial", color: C.dark5,
      charSpacing: 3, margin: 0,
    });

    // Dynamic headline font size
    const hlLen = String(headline).length;
    const hlSize = hlLen > 80 ? 22 : hlLen > 60 ? 26 : 30;

    s.addText(headline, {
      x: 0.5, y: 1.75, w: 5.8, h: 1.80,
      fontSize: hlSize, fontFace: "Arial", color: C.white,
      bold: true, margin: 0,
    });

    s.addText(company_name.toUpperCase(), {
      x: 0.5, y: 3.68, w: 5.5, h: 0.35,
      fontSize: 12, fontFace: "Arial", color: C.red,
      bold: true, charSpacing: 3, margin: 0,
    });

    if (website) {
      s.addText(website, {
        x: 0.5, y: 4.03, w: 5.5, h: 0.26,
        fontSize: 9, fontFace: "Arial", color: C.muted, margin: 0,
      });
    }

    // Audit score badge
    s.addShape("roundRect", {
      x: 0.5, y: 4.42, w: 1.65, h: 0.55,
      fill: { color: C.red }, line: { color: C.red, width: 0 }, rectRadius: 0.05,
    });
    s.addText(`SCORE: ${audit_score}/100`, {
      x: 0.5, y: 4.42, w: 1.65, h: 0.55,
      fontSize: 10, fontFace: "Arial", color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    // Skyline
    addSkyline(s, 6.5, 5.6, SKYLINE_RIGHT);
    s.addShape("rect", {
      x: 6.3, y: 5.54, w: 3.7, h: 0.085,
      fill: { color: C.red }, line: { color: C.red, width: 0 },
    });
  }

  // ────────────────────────────────────────────────
  // SLIDE 2 — DIAGNOSIS
  // ────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "WEBSITE AUDIT — DIAGNOSIS");

    s.addText("What We Found\nOn Your Website", {
      x: 0.5, y: 0.82, w: 7, h: 1.0,
      fontSize: 28, fontFace: "Arial", color: C.white,
      bold: true, margin: 0,
    });
    redLine(s, 0.5, 1.86, 1.0);
    s.addText(company_name.toUpperCase(), {
      x: 0.5, y: 1.98, w: 6, h: 0.28,
      fontSize: 10, fontFace: "Arial", color: C.red,
      bold: true, charSpacing: 2, margin: 0,
    });

    const problems = [problem_1, problem_2, problem_3];
    problems.forEach((prob, i) => {
      const y = 2.45 + i * 0.98;
      s.addShape("roundRect", {
        x: 0.5, y, w: 9.0, h: 0.84,
        fill: { color: C.cardBg },
        line: { color: C.dark3, width: 0.5 },
        rectRadius: 0.06, shadow: mkShadow(),
      });
      s.addImage({ data: pLoaded[i], x: 0.72, y: y + 0.22, w: 0.34, h: 0.34 });
      s.addText(`0${i + 1}`, {
        x: 1.18, y: y + 0.08, w: 0.4, h: 0.30,
        fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0,
      });
      s.addText(prob, {
        x: 1.18, y: y + 0.36, w: 7.9, h: 0.38,
        fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0,
      });
    });
  }

  // ────────────────────────────────────────────────
  // SLIDE 3 — SOLUTION
  // ────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "THE SOLUTION");

    s.addText("How We Fix It", {
      x: 0.5, y: 0.82, w: 7, h: 0.72,
      fontSize: 28, fontFace: "Arial", color: C.white,
      bold: true, margin: 0,
    });
    redLine(s, 0.5, 1.57, 1.0);

    const solutions = [solution_1, solution_2, solution_3];
    solutions.forEach((sol, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("roundRect", {
        x, y: 1.82, w: 2.88, h: 3.45,
        fill: { color: C.cardBg },
        line: { color: C.dark3, width: 0.5 },
        rectRadius: 0.08, shadow: mkShadow(),
      });
      s.addImage({ data: sLoaded[i], x: x + 1.21, y: 2.10, w: 0.46, h: 0.46 });
      s.addText(`0${i + 1}`, {
        x: x + 0.1, y: 2.62, w: 2.68, h: 0.36,
        fontSize: 13, fontFace: "Arial", color: C.red,
        bold: true, align: "center", margin: 0,
      });
      s.addText(sol, {
        x: x + 0.18, y: 3.02, w: 2.52, h: 2.08,
        fontSize: 12, fontFace: "Arial", color: C.offWhite,
        align: "center", margin: 0,
      });
    });
  }

  // ────────────────────────────────────────────────
  // SLIDE 4 — PROOF
  // ────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "PROVEN RESULTS");

    s.addText("What We've Done\nFor Brands Like Yours", {
      x: 0.5, y: 0.82, w: 7, h: 1.0,
      fontSize: 28, fontFace: "Arial", color: C.white,
      bold: true, margin: 0,
    });
    redLine(s, 0.5, 1.86, 1.0);

    statCard(s, proof_stat_1, proof_label_1, 0.50, 2.08);
    statCard(s, proof_stat_2, proof_label_2, 3.57, 2.08);
    statCard(s, proof_stat_3, proof_label_3, 6.64, 2.08);

    // Quote block
    s.addShape("roundRect", {
      x: 0.5, y: 3.55, w: 9.0, h: 1.72,
      fill: { color: C.redLight },
      line: { color: C.red, width: 0.4 },
      rectRadius: 0.07,
    });
    s.addText("\u201C", {
      x: 0.70, y: 3.48, w: 0.6, h: 0.65,
      fontSize: 52, fontFace: "Arial", color: C.red, bold: true, margin: 0,
    });
    s.addText(
      "We don't guess. We diagnose, build, and scale. Every strategy is built on data from your market, not templates from another industry.",
      {
        x: 1.30, y: 3.72, w: 7.8, h: 0.85,
        fontSize: 12, fontFace: "Arial", color: C.offWhite, italic: true, margin: 0,
      }
    );
    s.addText("— ProScaleMEDIA, Dubai", {
      x: 1.30, y: 4.60, w: 5, h: 0.28,
      fontSize: 10, fontFace: "Arial", color: C.red, margin: 0,
    });
  }

  // ────────────────────────────────────────────────
  // SLIDE 5 — CTA
  // ────────────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "NEXT STEP");

    // Skyline right
    addSkyline(s, 5.8, 5.6, SKYLINE_RIGHT);
    s.addShape("rect", {
      x: 5.8, y: 5.54, w: 4.2, h: 0.085,
      fill: { color: C.red }, line: { color: C.red, width: 0 },
    });

    redLine(s, 0.5, 0.82, 1.5);

    s.addText("Ready to\nScale?", {
      x: 0.5, y: 0.98, w: 5.0, h: 1.60,
      fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0,
    });

    s.addText(offer_line, {
      x: 0.5, y: 2.72, w: 5.0, h: 0.72,
      fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0,
    });

    // CTA button
    s.addShape("roundRect", {
      x: 0.5, y: 3.60, w: 3.5, h: 0.68,
      fill: { color: C.red },
      line: { color: C.red, width: 0 },
      rectRadius: 0.05,
      shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 45, opacity: 0.4 },
    });
    s.addText(cta_line, {
      x: 0.5, y: 3.60, w: 3.5, h: 0.68,
      fontSize: 12, fontFace: "Arial", color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0,
      hyperlink: { url: "https://calendly.com/contact-pro-scalemedia/30min" },
    });

    s.addText("📅 calendly.com/contact-pro-scalemedia/30min", {
      x: 0.5, y: 4.48, w: 5.2, h: 0.26,
      fontSize: 9, fontFace: "Arial", color: C.red, margin: 0,
      hyperlink: { url: "https://calendly.com/contact-pro-scalemedia/30min" },
    });

    s.addText("contact@pro-scalemedia.com  ·  pro-scalemedia.com  ·  Dubai, UAE", {
      x: 0.5, y: 4.76, w: 5.5, h: 0.22,
      fontSize: 9, fontFace: "Arial", color: C.dark5, margin: 0,
    });
  }

  // ─── RETURN BUFFER ───────────────────────────────
  return await pres.write({ outputType: "nodebuffer" });
}

// ─── ROUTES ──────────────────────────────────────────────────────
app.get("/health", (req, res) => res.json({ status: "ok", service: "ProScaleMEDIA PPTX API" }));

app.post("/generate", async (req, res) => {
  try {
    const data = req.body;

    if (!data || !data.company_name) {
      return res.status(400).json({ error: "company_name is required" });
    }

    console.log(`[PPTX] Generating deck for: ${data.company_name}`);
    const buffer = await generateDeck(data);

    const safeName = String(data.company_name).replace(/[^a-z0-9]/gi, "_").toUpperCase();
    const filename = `ProScaleMEDIA_${safeName}.pptx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);

    console.log(`[PPTX] ✓ Sent ${buffer.length} bytes — ${filename}`);
  } catch (err) {
    console.error("[PPTX] Error:", err.message);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

// ─── START ───────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`[PPTX] ProScaleMEDIA API running on port ${PORT}`);
});
