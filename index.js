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
      { x: 6.5,  y: 2.0, w: 0.15, h: 3.6 },
      { x: 6.75, y: 2.6, w: 0.22, h: 3.0 },
      { x: 7.06, y: 1.3, w: 0.17, h: 4.3 },
      { x: 7.32, y: 2.8, w: 0.26, h: 2.8 },
      { x: 7.67, y: 2.3, w: 0.2,  h: 3.3 },
      { x: 7.96, y: 3.1, w: 0.28, h: 2.5 },
      { x: 8.32, y: 2.7, w: 0.2,  h: 2.9 },
      { x: 8.6,  y: 3.4, w: 0.3,  h: 2.2 },
      { x: 8.98, y: 3.0, w: 0.22, h: 2.6 },
      { x: 9.28, y: 3.5, w: 0.28, h: 2.1 },
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

    const solutions = [solution_1, solution_2, solution_3];
    const emojis = [solution_emoji_1, solution_emoji_2, solution_emoji_3];

    solutions.forEach((sol, i) => {
      const x = 0.5 + i * 3.1;
      const cardW = 2.88;
      const cardH = 3.6;
      const cardY = 1.82;

      s.addShape("roundRect", {
        x, y: cardY, w: cardW, h: cardH,
        fill: { color: C.cardBg },
        line: { color: "2A2A2A", width: 0.8 },
        rectRadius: 0.1,
        shadow: mkShadow(),
      });

      s.addShape("rect", {
        x, y: cardY, w: cardW, h: 0.06,
        fill: { color: C.red },
        line: { color: C.red, width: 0 },
      });

      s.addText(`0${i + 1}`, {
        x: x + 0.14, y: cardY + 0.16, w: 0.36, h: 0.3,
        fontSize: 9, fontFace: "Arial", color: C.red,
        bold: true, margin: 0,
      });

      s.addShape("oval", {
        x: x + 0.84, y: cardY + 0.18, w: 1.2, h: 1.2,
        fill: { color: "1A1A1A" },
        line: { color: C.red, width: 1.2 },
      });

      s.addText(emojis[i] || "⚡", {
        x: x + 0.84, y: cardY + 0.22, w: 1.2, h: 1.1,
        fontSize: 32, align: "center", valign: "middle", margin: 0,
      });

      s.addText(sol || "", {
        x: x + 0.18, y: cardY + 1.62, w: cardW - 0.36, h: 1.8,
        fontSize: 11.5, fontFace: "Arial", color: C.offWhite,
        align: "center", margin: 0,
      });
    });
  }

  // ── SLIDE 4 — PROOF ─────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "PROVEN RESULTS");

    s.addText("What We Have Done\nFor Brands Like Yours", { x: 0.5, y: 0.82, w: 7, h: 1.0, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.86, 1.0);

    statCard(s, proof_stat_1 || "", proof_label_1 || "", 0.5,  2.08);
    statCard(s, proof_stat_2 || "", proof_label_2 || "", 3.57, 2.08);
    statCard(s, proof_stat_3 || "", proof_label_3 || "", 6.64, 2.08);

    s.addShape("roundRect", { x: 0.5, y: 3.55, w: 9.0, h: 1.72, fill: { color: C.redLight }, line: { color: C.red, width: 0.4 }, rectRadius: 0.07 });
    s.addText("\u201C", { x: 0.7, y: 3.48, w: 0.6, h: 0.65, fontSize: 52, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s.addText("We don't guess. We diagnose, build, and scale. Every strategy is built on data from your market, not templates from another industry.", { x: 1.3, y: 3.72, w: 7.8, h: 0.85, fontSize: 12, fontFace: "Arial", color: C.offWhite, italic: true, margin: 0 });
    s.addText("ProScaleMEDIA, Dubai", { x: 1.3, y: 4.6, w: 5, h: 0.28, fontSize: 10, fontFace: "Arial", color: C.red, margin: 0 });
  }

  // ── SLIDE 5 — CTA ───────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "NEXT STEP");

    const bldgs = [
      { x: 5.8,  y: 2.3, w: 0.15, h: 3.3 },
      { x: 6.05, y: 2.7, w: 0.22, h: 2.9 },
      { x: 6.36, y: 1.5, w: 0.17, h: 4.1 },
      { x: 6.62, y: 2.9, w: 0.26, h: 2.7 },
      { x: 6.97, y: 2.4, w: 0.2,  h: 3.2 },
      { x: 7.26, y: 3.2, w: 0.28, h: 2.4 },
      { x: 7.62, y: 2.8, w: 0.2,  h: 2.8 },
      { x: 7.9,  y: 3.4, w: 0.3,  h: 2.2 },
      { x: 8.28, y: 3.0, w: 0.22, h: 2.6 },
      { x: 8.58, y: 3.6, w: 0.28, h: 2.0 },
      { x: 8.94, y: 3.2, w: 0.2,  h: 2.4 },
      { x: 9.22, y: 3.7, w: 0.25, h: 1.9 },
      { x: 9.55, y: 3.3, w: 0.18, h: 2.3 },
    ];
    bldgs.forEach(b => s.addShape("rect", { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: "222222" }, line: { color: "222222", width: 0 } }));
    s.addShape("rect", { x: 5.8, y: 5.54, w: 4.2, h: 0.085, fill: { color: C.red }, line: { color: C.red, width: 0 } });

    redLine(s, 0.5, 0.82, 1.5);
    s.addText("Ready to\nScale?", { x: 0.5, y: 0.98, w: 5.0, h: 1.6, fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    s.addText(offer_line || "", { x: 0.5, y: 2.72, w: 5.0, h: 0.72, fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0 });

    s.addShape("roundRect", { x: 0.5, y: 3.6, w: 3.5, h: 0.68, fill: { color: C.red }, line: { color: C.red, width: 0 }, rectRadius: 0.05, shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 45, opacity: 0.4 } });
    s.addText(cta_line || "", { x: 0.5, y: 3.6, w: 3.5, h: 0.68, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0, hyperlink: { url: "https://calendly.com/contact-pro-scalemedia/30min" } });
    s.addText("calendly.com/contact-pro-scalemedia/30min", {
      x: 0.5, y: 4.48, w: 5.2, h: 0.26,
      fontSize: 9, fontFace: "Arial", color: "FE0000", margin: 0,
    });
    s.addText("contact@pro-scalemedia.com  ·  pro-scalemedia.com  ·  Dubai, UAE", {
      x: 0.5, y: 4.76, w: 5.5, h: 0.22,
      fontSize: 9, fontFace: "Arial", color: "555555", margin: 0,
    });
  }

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
