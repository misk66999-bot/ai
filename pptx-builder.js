const pptxgen = require("pptxgenjs");

// ---- Design tokens ------------------------------------------------------
const PALETTE = {
  bg: "FFFFFF",
  ink: "1B2A41",
  accent: "2E5EAA",
  accentSoft: "EEF3FC",
  accent2: "C9622A",
  accent2Soft: "FBEDE4",
  muted: "6B7686",
};

const FONT = "Arial";
const TOPIC_GLYPHS = ["📘", "💡", "🔍", "🧩", "🎯", "🗺️", "🧪", "✍️", "🎨", "🧠"];

function glyphFor(index) {
  return TOPIC_GLYPHS[index % TOPIC_GLYPHS.length];
}

function toBullets(content) {
  if (!content) return [];
  if (Array.isArray(content)) content = content.join("\n");
  const raw = String(content)
    .split(/\n+|(?<=[.؟!])\s+(?=[أ-ي A-Za-z])/)
    .map((s) => s.replace(/^[-•\s]+/, "").trim())
    .filter(Boolean);
  return raw.length ? raw : [String(content).trim()];
}

function freshShadow(angle = 90, offset = 3) {
  return { type: "outer", color: "1B2A41", opacity: 0.14, blur: 8, offset, angle };
}

// ---- Image fetching (Pexels) --------------------------------------------

async function fetchPexelsImage(query, apiKey) {
  if (!apiKey || !query) return null;
  try {
    const searchRes = await fetch(
      `https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}&per_page=1&orientation=landscape`,
      { headers: { Authorization: apiKey } }
    );
    if (!searchRes.ok) return null;
    const data = await searchRes.json();
    const photo = data.photos && data.photos[0];
    if (!photo) return null;

    const imgRes = await fetch(photo.src.large);
    if (!imgRes.ok) return null;
    const arrayBuffer = await imgRes.arrayBuffer();
    const base64 = Buffer.from(arrayBuffer).toString("base64");
    return `data:image/jpeg;base64,${base64}`;
  } catch (err) {
    console.error("Pexels fetch failed:", err.message);
    return null;
  }
}

// ---- Slide builders ------------------------------------------------------

function buildTitleSlide(pres, lesson) {
  const slide = pres.addSlide();
  slide.background = { color: PALETTE.bg };

  slide.addShape(pres.ShapeType.ellipse, {
    x: 9.6, y: -2.2, w: 6, h: 6,
    fill: { color: PALETTE.accentSoft }, line: { type: "none" },
  });
  slide.addShape(pres.ShapeType.ellipse, {
    x: -1.5, y: 5.2, w: 3.2, h: 3.2,
    fill: { color: PALETTE.accent2Soft }, line: { type: "none" },
  });

  slide.addText(`${lesson.subject || ""}   •   ${lesson.grade || ""}`, {
    x: 0.8, y: 0.7, w: 10, h: 0.5,
    fontFace: FONT, fontSize: 16, bold: true, color: PALETTE.accent,
    align: "right", rtlMode: true,
  });

  slide.addText(lesson.topic || "عنوان الدرس", {
    x: 0.8, y: 2.6, w: 11.7, h: 2,
    fontFace: FONT, fontSize: 48, bold: true, color: PALETTE.ink,
    align: "right", rtlMode: true, margin: 0,
  });

  slide.addText("خطة الدرس والعرض التقديمي", {
    x: 0.8, y: 4.7, w: 8, h: 0.5,
    fontFace: FONT, fontSize: 18, color: PALETTE.muted,
    align: "right", rtlMode: true,
  });
  return slide;
}

function buildBulletSlide(pres, slideData, index) {
  const slide = pres.addSlide();
  slide.background = { color: PALETTE.bg };
  const hasImage = !!slideData.imageData;

  slide.addShape(pres.ShapeType.ellipse, {
    x: 11.2, y: -1.4, w: 3.4, h: 3.4,
    fill: { color: PALETTE.accentSoft }, line: { type: "none" },
  });

  slide.addText(glyphFor(index), {
    x: 0.8, y: 0.55, w: 0.9, h: 0.9, fontSize: 32, align: "center",
  });

  slide.addText(slideData.title || "", {
    x: 1.9, y: 0.55, w: 10.6, h: 0.95,
    fontFace: FONT, fontSize: 30, bold: true, color: PALETTE.ink,
    align: "right", rtlMode: true, margin: 0,
  });

  const textW = hasImage ? 6.9 : 11.0;
  const bullets = toBullets(slideData.content).map((t, i, arr) => ({
    text: t,
    options: {
      bullet: true, breakLine: i !== arr.length - 1,
      color: PALETTE.ink, fontSize: 19, paraSpaceAfter: 12,
    },
  }));

  slide.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 1.75, w: 11.7, h: 5.1, rectRadius: 0.12,
    fill: { color: PALETTE.accentSoft }, line: { type: "none" },
    shadow: freshShadow(),
  });

  slide.addText(bullets, {
    x: 1.2, y: 2.1, w: textW, h: 4.5,
    fontFace: FONT, align: "right", rtlMode: true, valign: "top",
    margin: 0, lineSpacingMultiple: 1.25,
  });

  if (hasImage) {
    slide.addImage({
      data: slideData.imageData,
      x: 8.4, y: 2.1, w: 3.9, h: 4.5,
      sizing: { type: "cover", w: 3.9, h: 4.5 },
      rounding: true,
    });
  }

  addSpeakerNotes(slide, slideData.notes);
  return slide;
}

function buildHighlightSlide(pres, slideData, index) {
  const slide = pres.addSlide();
  slide.background = { color: PALETTE.accent };
  const hasImage = !!slideData.imageData;

  slide.addShape(pres.ShapeType.ellipse, {
    x: -2, y: -2, w: 6, h: 6,
    fill: { color: "3D6FBE" }, line: { type: "none" },
  });

  if (!hasImage) {
    slide.addText(glyphFor(index), {
      x: 5.4, y: 1.0, w: 2.5, h: 1.3, fontSize: 44, align: "center",
    });
  } else {
    slide.addImage({
      data: slideData.imageData,
      x: 4.65, y: 0.6, w: 4.2, h: 2.2,
      sizing: { type: "cover", w: 4.2, h: 2.2 },
      rounding: true,
    });
  }

  slide.addText(slideData.title || "", {
    x: 1.3, y: hasImage ? 3.0 : 2.3, w: 10.7, h: 0.9,
    fontFace: FONT, fontSize: 26, bold: true, color: "FFFFFF",
    align: "center", rtlMode: true,
  });

  const bullets = toBullets(slideData.content).map((t, i, arr) => ({
    text: t,
    options: {
      breakLine: i !== arr.length - 1, color: "FFFFFF",
      fontSize: 20, align: "center",
    },
  }));

  slide.addText(bullets, {
    x: 1.5, y: hasImage ? 4.0 : 3.3, w: 10.3, h: 2.6,
    fontFace: FONT, align: "center", rtlMode: true, valign: "top",
    lineSpacingMultiple: 1.3,
  });

  addSpeakerNotes(slide, slideData.notes);
  return slide;
}

function buildSplitSlide(pres, slideData, index) {
  const slide = pres.addSlide();
  slide.background = { color: PALETTE.bg };
  const hasImage = !!slideData.imageData;

  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 4.2, h: 7.5,
    fill: { color: PALETTE.accent2 }, line: { type: "none" },
  });

  if (hasImage) {
    slide.addImage({
      data: slideData.imageData,
      x: 0.3, y: 0.5, w: 3.6, h: 3.2,
      sizing: { type: "cover", w: 3.6, h: 3.2 },
      rounding: true,
    });
  } else {
    slide.addText(glyphFor(index), {
      x: 0.6, y: 2.9, w: 3, h: 1.5, fontSize: 54, align: "center",
    });
  }

  slide.addText(slideData.title || "", {
    x: 0.4, y: hasImage ? 3.9 : 4.5, w: 3.4, h: 1.8,
    fontFace: FONT, fontSize: 22, bold: true, color: "FFFFFF",
    align: "center", rtlMode: true,
  });

  const bullets = toBullets(slideData.content).map((t, i, arr) => ({
    text: t,
    options: {
      bullet: true, breakLine: i !== arr.length - 1,
      color: PALETTE.ink, fontSize: 18, paraSpaceAfter: 12,
    },
  }));

  slide.addText(bullets, {
    x: 4.7, y: 0.9, w: 7.9, h: 5.8,
    fontFace: FONT, align: "right", rtlMode: true, valign: "top",
    margin: 0, lineSpacingMultiple: 1.25,
  });

  addSpeakerNotes(slide, slideData.notes);
  return slide;
}

function addSpeakerNotes(slide, notes) {
  if (notes) slide.addNotes(String(notes));
}

function buildClosingSlide(pres, lesson) {
  const slide = pres.addSlide();
  slide.background = { color: PALETTE.ink };

  slide.addShape(pres.ShapeType.ellipse, {
    x: 9.8, y: 4.6, w: 5, h: 5,
    fill: { color: "24354F" }, line: { type: "none" },
  });

  slide.addText("شكراً لكم 🌟", {
    x: 0.8, y: 2.6, w: 11.7, h: 1.2,
    fontFace: FONT, fontSize: 44, bold: true, color: "FFFFFF", align: "center",
  });

  slide.addText(lesson.topic || "", {
    x: 0.8, y: 3.9, w: 11.7, h: 0.6,
    fontFace: FONT, fontSize: 18, color: "AAB8CC", align: "center", rtlMode: true,
  });

  return slide;
}

// ---- Layout selection ------------------------------------------------------

function pickLayoutBuilder(index, total, slideData) {
  const bulletCount = toBullets(slideData.content).length;
  if (bulletCount <= 2 && String(slideData.content || "").length < 90) {
    return buildHighlightSlide;
  }
  return index % 2 === 0 ? buildBulletSlide : buildSplitSlide;
}

// ---- Main generation entry ------------------------------------------------

async function generatePptx({ topic, grade, subject, slides }) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  const apiKey = process.env.PEXELS_API_KEY;

  buildTitleSlide(pres, { topic, grade, subject });

  for (let i = 0; i < slides.length; i++) {
    const slideData = slides[i];
    const query = slideData.visual_suggestion || slideData.title || topic;
    const imageData = await fetchPexelsImage(query, apiKey);
    const builder = pickLayoutBuilder(i, slides.length, slideData);
    builder(pres, { ...slideData, imageData }, i);
  }

  buildClosingSlide(pres, { topic });

  return pres.write("nodebuffer");
}

module.exports = { generatePptx };
