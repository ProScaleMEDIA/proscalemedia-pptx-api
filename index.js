const express = require("express");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const app = express();
app.use(express.json());

// ─── LOGO ─────────────────────────────────────────────────────────
// REPLACE_LOGO_HERE
const LOGO_B64 = "image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";


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
    const fallback = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${size/2}" cy="${size/2}" r="${size/3}" fill="${color}"/></svg>`;
    const buf = await sharp(Buffer.from(fallback)).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
  }
}

// ─── SAFE ICON NAME ───────────────────────────────────────────────
function safeFaIcon(name) {
  if (!name) return "FaRocket";
  const cleaned = name.trim();
  if (/^Fa[A-Z]/.test(cleaned)) return cleaned;
  if (/^[A-Z]/.test(cleaned)) return "Fa" + cleaned;
  return "Fa" + cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
}

const mkShadow = () => ({
  type: "outer", color: "000000",
  blur: 8, offset: 2, angle: 45, opacity: 0.5,
});

function redLine(slide, x, y, w = 1.0) {
  slide.addShape("rect", { x, y, w, h: 0.04, fill: { color: C.red }, line: { color: C.red, width: 0 } });
}

function topBar(slide, label) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.62, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  slide.addShape("rect", { x: 0, y: 0.60, w: 10, h: 0.04, fill: { color: C.red }, line: { color: C.red, width: 0 } });
  slide.addImage({ data: LOGO_B64, x: 0.4, y: 0.06, w: 2.4, h: 0.5 });
  if (label) {
    slide.addText(label, { x: 6.5, y: 0.18, w: 3.3, h: 0.26, fontSize: 7.5, fontFace: "Arial", color: C.dark5, bold: true, charSpacing: 3, align: "right", margin: 0 });
  }
}

function statCard(slide, stat, label, x, y) {
  const w = 2.85, h = 1.28;
  slide.addShape("roundRect", { x, y, w, h, fill: { color: C.cardBg }, line: { color: C.red, width: 1.2 }, rectRadius: 0.06, shadow: mkShadow() });
  const statLen = String(stat).length;
  const statSize = statLen > 8 ? 14 : statLen > 5 ? 18 : 20;
  slide.addText(stat, { x: x+0.1, y: y+0.08, w: w-0.2, h: 0.76, fontSize: statSize, fontFace: "Arial", color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
  slide.addText(label, { x: x+0.1, y: y+0.88, w: w-0.2, h: 0.30, fontSize: 9, fontFace: "Arial", color: C.muted, align: "center", margin: 0 });
}

const SKYLINE = [
  { ox: 0.0, h: 3.3, w: 0.15 }, { ox: 0.25, h: 2.9, w: 0.22 },
  { ox: 0.56, h: 4.1, w: 0.17 }, { ox: 0.82, h: 2.7, w: 0.26 },
  { ox: 1.17, h: 3.2, w: 0.20 }, { ox: 1.46, h: 2.4, w: 0.28 },
  { ox: 1.82, h: 2.8, w: 0.20 }, { ox: 2.10, h: 2.2, w: 0.30 },
  { ox: 2.48, h: 2.6, w: 0.22 }, { ox: 2.78, h: 2.1, w: 0.28 },
  { ox: 3.14, h: 2.4, w: 0.20 },
];

function addSkyline(slide, startX, baseY) {
  SKYLINE.forEach(b => {
    slide.addShape("rect", { x: startX+b.ox, y: baseY-b.h, w: b.w, h: b.h, fill: { color: C.dark2 }, line: { color: C.dark2, width: 0 } });
  });
}

async function generateDeck(data) {
  const {
    company_name = "Client", website = "", audit_score = "N/A",
    headline = "Here is exactly what we found on your website.",
    problem_1 = "Problem 1", problem_2 = "Problem 2", problem_3 = "Problem 3",
    solution_1 = "Solution 1", solution_2 = "Solution 2", solution_3 = "Solution 3",
    solution_icon_1, solution_icon_2, solution_icon_3,
    proof_stat_1 = "620+", proof_label_1 = "Direct reservations in 30 days",
    proof_stat_2 = "24K", proof_label_2 = "New YouTube subscribers in 60 days",
    proof_stat_3 = "200+", proof_label_3 = "WhatsApp leads in 60 days",
    offer_line = "We will fix this in 30 days or you don't pay.",
    cta_line = "BOOK YOUR FREE STRATEGY CALL",
  } = data;

  const icons = [
    "fa/" + safeFaIcon(solution_icon_1 || "FaRocket"),
    "fa/" + safeFaIcon(solution_icon_2 || "FaChartLine"),
    "fa/" + safeFaIcon(solution_icon_3 || "FaBullseye"),
  ];
  const sLoaded = await Promise.all(icons.map(i => iconBase64(i, "#FFFFFF", 128)));
  const pIcons = ["fa/FaExclamationTriangle", "fa/FaTimesCircle", "fa/FaExclamationCircle"];
  const pLoaded = await Promise.all(pIcons.map(i => iconBase64(i, "#FE0000")));

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // SLIDE 1 — COVER
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    s.addImage({ data: LOGO_B64, x: 0.5, y: 0.35, w: 3.2, h: 0.72 });
    redLine(s, 0.5, 1.22, 1.5);
    s.addText("BRAND GROWTH AGENCY · DUBAI", { x: 0.5, y: 1.34, w: 5.5, h: 0.24, fontSize: 8, fontFace: "Arial", color: C.dark5, charSpacing: 3, margin: 0 });
    const hlLen = String(headline).length;
    const hlSize = hlLen > 80 ? 22 : hlLen > 60 ? 26 : 30;
    s.addText(headline, { x: 0.5, y: 1.75, w: 5.8, h: 1.80, fontSize: hlSize, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    s.addText(company_name.toUpperCase(), { x: 0.5, y: 3.68, w: 5.5, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.red, bold: true, charSpacing: 3, margin: 0 });
    if (website) s.addText(website, { x: 0.5, y: 4.03, w: 5.5, h: 0.26, fontSize: 9, fontFace: "Arial", color: C.muted, margin: 0 });
    s.addShape("roundRect", { x: 0.5, y: 4.42, w: 1.65, h: 0.55, fill: { color: C.red }, line: { color: C.red, width: 0 }, rectRadius: 0.05 });
    s.addText(`SCORE: ${audit_score}/100`, { x: 0.5, y: 4.42, w: 1.65, h: 0.55, fontSize: 10, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    addSkyline(s, 6.5, 5.6);
    s.addShape("rect", { x: 6.3, y: 5.54, w: 3.7, h: 0.085, fill: { color: C.red }, line: { color: C.red, width: 0 } });
  }

  // SLIDE 2 — DIAGNOSIS
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "WEBSITE AUDIT — DIAGNOSIS");
    s.addText("What We Found\nOn Your Website", { x: 0.5, y: 0.82, w: 7, h: 1.0, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.86, 1.0);
    s.addText(company_name.toUpperCase(), { x: 0.5, y: 1.98, w: 6, h: 0.28, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, charSpacing: 2, margin: 0 });
    [problem_1, problem_2, problem_3].forEach((prob, i) => {
      const y = 2.45 + i * 0.98;
      s.addShape("roundRect", { x: 0.5, y, w: 9.0, h: 0.84, fill: { color: C.cardBg }, line: { color: C.dark3, width: 0.5 }, rectRadius: 0.06, shadow: mkShadow() });
      s.addImage({ data: pLoaded[i], x: 0.72, y: y+0.22, w: 0.34, h: 0.34 });
      s.addText(`0${i+1}`, { x: 1.18, y: y+0.08, w: 0.4, h: 0.30, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
      s.addText(prob, { x: 1.18, y: y+0.36, w: 7.9, h: 0.38, fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0 });
    });
  }

  // SLIDE 3 — SOLUTION
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "THE SOLUTION");
    s.addText("How We Fix It", { x: 0.5, y: 0.82, w: 7, h: 0.72, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.57, 1.0);
    [solution_1, solution_2, solution_3].forEach((sol, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape("roundRect", { x, y: 1.82, w: 2.88, h: 3.45, fill: { color: C.cardBg }, line: { color: C.dark3, width: 0.5 }, rectRadius: 0.08, shadow: mkShadow() });
      s.addImage({ data: sLoaded[i], x: x+1.21, y: 2.10, w: 0.46, h: 0.46 });
      s.addText(`0${i+1}`, { x: x+0.1, y: 2.62, w: 2.68, h: 0.36, fontSize: 13, fontFace: "Arial", color: C.red, bold: true, align: "center", margin: 0 });
      s.addText(sol, { x: x+0.18, y: 3.02, w: 2.52, h: 2.08, fontSize: 12, fontFace: "Arial", color: C.offWhite, align: "center", margin: 0 });
    });
  }

  // SLIDE 4 — PROOF
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "PROVEN RESULTS");
    s.addText("What We've Done\nFor Brands Like Yours", { x: 0.5, y: 0.82, w: 7, h: 1.0, fontSize: 28, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    redLine(s, 0.5, 1.86, 1.0);
    statCard(s, proof_stat_1, proof_label_1, 0.50, 2.08);
    statCard(s, proof_stat_2, proof_label_2, 3.57, 2.08);
    statCard(s, proof_stat_3, proof_label_3, 6.64, 2.08);
    s.addShape("roundRect", { x: 0.5, y: 3.55, w: 9.0, h: 1.72, fill: { color: C.redLight }, line: { color: C.red, width: 0.4 }, rectRadius: 0.07 });
    s.addText("\u201C", { x: 0.70, y: 3.48, w: 0.6, h: 0.65, fontSize: 52, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s.addText("We don't guess. We diagnose, build, and scale. Every strategy is built on data from your market, not templates from another industry.", { x: 1.30, y: 3.72, w: 7.8, h: 0.85, fontSize: 12, fontFace: "Arial", color: C.offWhite, italic: true, margin: 0 });
    s.addText("— ProScaleMEDIA, Dubai", { x: 1.30, y: 4.60, w: 5, h: 0.28, fontSize: 10, fontFace: "Arial", color: C.red, margin: 0 });
  }

  // SLIDE 5 — CTA
  {
    const s = pres.addSlide();
    s.background = { color: C.black };
    topBar(s, "NEXT STEP");
    addSkyline(s, 5.8, 5.6);
    s.addShape("rect", { x: 5.8, y: 5.54, w: 4.2, h: 0.085, fill: { color: C.red }, line: { color: C.red, width: 0 } });
    redLine(s, 0.5, 0.82, 1.5);
    s.addText("Ready to\nScale?", { x: 0.5, y: 0.98, w: 5.0, h: 1.60, fontSize: 42, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
    s.addText(offer_line, { x: 0.5, y: 2.72, w: 5.0, h: 0.72, fontSize: 13, fontFace: "Arial", color: C.offWhite, margin: 0 });
    s.addShape("roundRect", { x: 0.5, y: 3.60, w: 3.5, h: 0.68, fill: { color: C.red }, line: { color: C.red, width: 0 }, rectRadius: 0.05, shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 45, opacity: 0.4 } });
    s.addText(cta_line, { x: 0.5, y: 3.60, w: 3.5, h: 0.68, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0, hyperlink: { url: "https://calendly.com/contact-pro-scalemedia/30min" } });
    s.addText("calendly.com/contact-pro-scalemedia/30min", { x: 0.5, y: 4.48, w: 5.2, h: 0.26, fontSize: 9, fontFace: "Arial", color: C.red, margin: 0, hyperlink: { url: "https://calendly.com/contact-pro-scalemedia/30min" } });
    s.addText("contact@pro-scalemedia.com  ·  pro-scalemedia.com  ·  Dubai, UAE", { x: 0.5, y: 4.76, w: 5.5, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.dark5, margin: 0 });
  }

  return await pres.write({ outputType: "nodebuffer" });
}

app.get("/health", (req, res) => res.json({ status: "ok", service: "ProScaleMEDIA PPTX API" }));

app.post("/generate", async (req, res) => {
  try {
    const data = req.body;
    if (!data || !data.company_name) return res.status(400).json({ error: "company_name is required" });
    console.log(`[PPTX] Generating deck for: ${data.company_name}`);
    const buffer = await generateDeck(data);
    const safeName = String(data.company_name).replace(/[^a-z0-9]/gi, "_").toUpperCase();
    const filename = `ProScaleMEDIA_${safeName}.pptx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);
    console.log(`[PPTX] Sent ${buffer.length} bytes — ${filename}`);
  } catch (err) {
    console.error("[PPTX] Error:", err.message);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`[PPTX] ProScaleMEDIA API running on port ${PORT}`));
