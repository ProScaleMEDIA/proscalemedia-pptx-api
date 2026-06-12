const express = require("express");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const app = express();
app.use(express.json());

// ─── LOGO (base64 embedded) ───────────────────────────────────────
const LOGO_B64 = "image/png;base64,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";

// ─── BRAND PALETTE ───────────────────────────────────────────────
const C = {
  white:    "FFFFFF",
  offWhite: "F0F0F0",
  lightGray:"2A2A2A",
  muted:    "666666",
  black:    "000000",
  red:      "FE0000",
  redLight: "1A0000",
  redDim:   "CC0000",
  cardBg:   "141414",
};

// ─── HELPERS ─────────────────────────────────────────────────────
async function iconBase64(iconPath, color = "#FE0000", size = 256) {
  const [lib, name] = iconPath.split("/");
  const icons = require(`react-icons/${lib}`);
  const Icon = icons[name];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Convert an emoji character to a PNG base64 using sharp + SVG text
async function emojiBase64(emoji, size = 200) {
  const svgStr = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}">
    <text x="50%" y="54%" dominant-baseline="middle" text-anchor="middle"
      font-size="${Math.round(size * 0.7)}" font-family="Apple Color Emoji, Segoe UI Emoji, Noto Color Emoji, serif">${emoji}</text>
  </svg>`;
  const buf = await sharp(Buffer.from(svgStr)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const mkShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 2, angle: 45, opacity: 0.5 });

function redLine(slide, x, y, w = 1.0) {
  slide.addShape("rect", { x, y, w, h: 0.04, fill: { color: C.red }, line: { color: C.red, width: 0 } });
}

function topBar(slide, label) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.62, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  slide.addShape("rect", { x: 0, y: 0.6, w: 10, h: 0.04, fill: { color: C.red }, line: { color: C.red, width: 0 } });
  slide.addImage({ data: LOGO_B64, x: 0.4, y: 0.06, w: 2.4, h: 0.5 });
  if (label) {
    slide.addText(label, {
      x: 6.5, y: 0.18, w: 3.3, h: 0.26,
      fontSize: 7.5, fontFace: "Arial", color: "555555",
      bold: true, charSpacing: 3, align: "right", margin: 0,
    });
  }
}

function statCard(slide, stat, label, x, y) {
  const w = 2.85, h = 1.28;
  slide.addShape("roundRect", { x, y, w, h, fill: { color: C.cardBg }, line: { color: C.red, width: 1.2 }, rectRadius: 0.06, shadow: mkShadow() });
  slide.addText(stat, { x: x+0.1, y: y+0.08, w: w-0.2, h: 0.76, fontSize: 20, fontFace: "Arial", color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
  slide.addText(label, { x: x+0.1, y: y+0.88, w: w-0.2, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.muted, align: "center", margin: 0 });
}

// ─── GENERATE FUNCTION ───────────────────────────────────────────
async function generateDeck(data) {
  const {
    company_name, website, audit_score, headline,
    problem_1, problem_2, problem_3,
    solution_1, solution_2, solution_3,
    solution_emoji_1, solution_emoji_2, solution_emoji_3,
    proof_stat_1, proof_label_1,
    proof_stat_2, proof_label_2,
    proof_stat_3, proof_label_3,
    offer_line, cta_line,
  } = data;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // ── SLIDE 1 — COVER ─────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };

    s.addImage({ data: LOGO_B64, x: 0.5, y: 0.35, w: 3.2, h: 0.72 });
    redLine(s, 0.5, 1.22, 1.5);

    s.addText("BRAND GROWTH AGENCY · DUBAI", {
      x: 0.5, y: 1.34, w: 5.5, h: 0.24,
      fontSize: 8, fontFace: "Arial", color: "555555", charSpacing: 3, margin: 0,
    });
    s.addText(headline, {
      x: 0.5, y: 1.75, w: 5.8, h: 1.8,
      fontSize: 30, fontFace: "Arial", color: C.white, bold: true, margin: 0,
    });
    s.addText((company_name || "").toUpperCase(), {
      x: 0.5, y: 3.68, w: 5.5, h: 0.35,
      fontSize: 12, fontFace: "Arial", color: C.red, bold: true, charSpacing: 3, margin: 0,
    });
    s.addText(website || "", {
      x: 0.5, y: 4.03, w: 5.5, h: 0.26,
      fontSize: 9, fontFace: "Arial", color: C.muted, margin: 0,
    });
    s.addShape("roundRect", { x: 0.5, y: 4.42, w: 1.65, h: 0.55, fill: { color: C.red }, line: { color: C.red, width: 0 }, rectRadius: 0.05 });
    s.addText(`SCORE: ${audit_score}/100`, { x: 0.5, y: 4.42, w: 1.65, h: 0.55, fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    const bldgs = [
      { x: 6.5,  y: 2.0, w: 0.15, h: 3.6 }, { x: 6.75, y: 2.6, w: 0.22, h: 3.0 },
      { x: 7.06, y: 1.3, w: 0.17, h: 4.3 }, { x: 7.32, y: 2.8, w: 0.26, h: 2.8 },
      { x: 7.67, y: 2.3, w: 0.2,  h: 3.3 }, { x: 7.96, y: 3.1, w: 0.28, h: 2.5 },
      { x: 8.32, y: 2.7, w: 0.2,  h: 2.9 }, { x: 8.6,  y: 3.4, w: 0.3,  h: 2.2 },
      { x: 8.98, y: 3.0, w: 0.22, h: 2.6 }, { x: 9.28, y: 3.5, w: 0.28, h: 2.1 },
      { x: 9.64, y: 3.2, w: 0.2,  h: 2.4 },
    ];
    bldgs.forEach(b => s.addShape("rect", { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: "222222" }, line: { color: "222222", width: 0 } }));
    s.addShape("rect", { x: 6.3, y: 5.54, w: 3.7, h: 0.085, fill: { color: C.red }, line: { color: C.red, width: 0 } });
  }

  // ── SLIDE 2 — DIAGNOSIS ─────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "WEBSITE AUDIT — DIAGNOSIS");

    s.addText("What We Found\nOn Your Website", { x: 0.5, y: 0.82, w: 7, h: 1.0, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.86, 1.0);
    s.addText((company_name || "").toUpperCase(), { x: 0.5, y: 1.98, w: 6, h: 0.28, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, charSpacing: 2, margin: 0 });

    const pIcons = ["fa/FaExclamationTriangle", "fa/FaTimesCircle", "fa/FaExclamationCircle"];
    const pLoaded = await Promise.all(pIcons.map(i => iconBase64(i, "#FE0000")));
    const problems = [problem_1, problem_2, problem_3];

    problems.forEach((prob, i) => {
      const y = 2.45 + i * 0.98;
      s.addShape("roundRect", { x: 0.5, y, w: 9.0, h: 0.84, fill: { color: C.cardBg }, line: { color: "333333", width: 0.5 }, rectRadius: 0.06, shadow: mkShadow() });
      s.addImage({ data: pLoaded[i], x: 0.72, y: y+0.22, w: 0.34, h: 0.34 });
      s.addText(`0${i+1}`, { x: 1.18, y: y+0.08, w: 0.4, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
      s.addText(prob || "", { x: 1.18, y: y+0.36, w: 7.9, h: 0.38, fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0 });
    });
  }

  // ── SLIDE 3 — SOLUTION ──────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "THE SOLUTION");

    s.addText("How We Fix It", { x: 0.5, y: 0.82, w: 7, h: 0.72, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.57, 1.0);

    const emojis = [
      solution_emoji_1 || "🚀",
      solution_emoji_2 || "📈",
      solution_emoji_3 || "🎯",
    ];
    const emojiImgs = await Promise.all(emojis.map(e => emojiBase64(e, 200)));
    const solutions = [solution_1, solution_2, solution_3];

    solutions.forEach((sol, i) => {
      const x = 0.5 + i * 3.1;
      // Card background
      s.addShape("roundRect", { x, y: 1.82, w: 2.88, h: 3.45, fill: { color: C.cardBg }, line: { color: "333333", width: 0.5 }, rectRadius: 0.08, shadow: mkShadow() });
      // Red accent bar at top of card
      s.addShape("rect", { x, y: 1.82, w: 2.88, h: 0.07, fill: { color: C.red }, line: { color: C.red, width: 0 } });
      // Step number top-left
      s.addText(`0${i+1}`, { x: x+0.14, y: 1.90, w: 0.45, h: 0.30, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
      // Dark circle with red border for the emoji icon
      s.addShape("oval", { x: x+0.94, y: 2.10, w: 0.92, h: 0.92, fill: { color: "1C1C1C" }, line: { color: C.red, width: 1.8 } });
      // Emoji image inside the circle
      s.addImage({ data: emojiImgs[i], x: x+1.04, y: 2.20, w: 0.52, h: 0.52 });
      // Solution text
      s.addText(sol || "", { x: x+0.18, y: 3.22, w: 2.52, h: 1.85, fontSize: 12, fontFace: "Arial", color: C.offWhite, align: "center", valign: "top", margin: 0 });
    });
  }

  // ── SLIDE 4 — PROOF ─────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "PROVEN RESULTS");

    s.addText("What We've Done\nFor Brands Like Yours", { x: 0.5, y: 0.82, w: 7, h: 1.0, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.86, 1.0);

    statCard(s, proof_stat_1 || "", proof_label_1 || "", 0.5,  2.08);
    statCard(s, proof_stat_2 || "", proof_label_2 || "", 3.57, 2.08);
    statCard(s, proof_stat_3 || "", proof_label_3 || "", 6.64, 2.08);

    s.addShape("roundRect", { x: 0.5, y: 3.55, w: 9.0, h: 1.72, fill: { color: C.redLight }, line: { color: C.red, width: 0.4 }, rectRadius: 0.07 });
    s.addText("\u201C", { x: 0.7, y: 3.48, w: 0.6, h: 0.65, fontSize: 52, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s.addText("We don't guess. We diagnose, build, and scale. Every strategy is built on data from your market, not templates from another industry.", { x: 1.3, y: 3.72, w: 7.8, h: 0.85, fontSize: 12, fontFace: "Arial", color: C.offWhite, italic: true, margin: 0 });
    s.addText("— ProScaleMEDIA, Dubai", { x: 1.3, y: 4.6, w: 5, h: 0.28, fontSize: 10, fontFace: "Arial", color: C.red, margin: 0 });
  }

  // ── SLIDE 5 — CTA ───────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "NEXT STEP");

    const bldgs = [
      { x: 5.8,  y: 2.3, w: 0.15, h: 3.3 }, { x: 6.05, y: 2.7, w: 0.22, h: 2.9 },
      { x: 6.36, y: 1.5, w: 0.17, h: 4.1 }, { x: 6.62, y: 2.9, w: 0.26, h: 2.7 },
      { x: 6.97, y: 2.4, w: 0.2,  h: 3.2 }, { x: 7.26, y: 3.2, w: 0.28, h: 2.4 },
      { x: 7.62, y: 2.8, w: 0.2,  h: 2.8 }, { x: 7.9,  y: 3.4, w: 0.3,  h: 2.2 },
      { x: 8.28, y: 3.0, w: 0.22, h: 2.6 }, { x: 8.58, y: 3.6, w: 0.28, h: 2.0 },
      { x: 8.94, y: 3.2, w: 0.2,  h: 2.4 }, { x: 9.22, y: 3.7, w: 0.25, h: 1.9 },
      { x: 9.55, y: 3.3, w: 0.18, h: 2.3 },
    ];
    bldgs.forEach(b => s.addShape("rect", { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: "222222" }, line: { color: "222222", width: 0 } }));
    s.addShape("rect", { x: 5.8, y: 5.54, w: 4.2, h: 0.085, fill: { color: C.red }, line: { color: C.red, width: 0 } });

    redLine(s, 0.5, 0.82, 1.5);
    s.addText("Ready to\nScale?", { x: 0.5, y: 0.98, w: 5.0, h: 1.6, fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    s.addText(offer_line || "", { x: 0.5, y: 2.72, w: 5.0, h: 0.72, fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0 });

    s.addShape("roundRect", { x: 0.5, y: 3.6, w: 3.5, h: 0.68, fill: { color: C.red }, line: { color: C.red, width: 0 }, rectRadius: 0.05, shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 45, opacity: 0.4 } });
    s.addText(cta_line || "", { x: 0.5, y: 3.6, w: 3.5, h: 0.68, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    s.addText("contact@pro-scalemedia.com", { x: 0.5, y: 4.48, w: 5, h: 0.26, fontSize: 10, fontFace: "Arial", color: C.red, margin: 0 });
    s.addText("pro-scalemedia.com  ·  Dubai, UAE", { x: 0.5, y: 4.76, w: 5, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.muted, margin: 0 });
  }

  // ── RETURN AS BUFFER ────────────────────────────
  return await pres.write({ outputType: "nodebuffer" });
}

// ─── ROUTES ──────────────────────────────────────────────────────
app.get("/health", (req, res) => res.json({ status: "ok" }));

app.post("/generate", async (req, res) => {
  try {
    const data = req.body;
    if (!data.company_name) return res.status(400).json({ error: "company_name required" });

    const buffer = await generateDeck(data);
    const filename = `ProScaleMEDIA_${(data.company_name).replace(/[^a-z0-9]/gi, "_")}.pptx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PPTX API running on port ${PORT}`));
