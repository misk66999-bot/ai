const pptxgen = require("pptxgenjs");

// 1. دالة جلب رابط الصورة من Pexels باستخدام fetch المدمجة
async function fetchPexelsImageUrl(query) {
  const apiKey = process.env.PEXELS_API_KEY;
  if (!apiKey || !query) return null;

  try {
    const response = await fetch(
      `https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}&per_page=1`,
      {
        headers: { Authorization: apiKey },
        signal: AbortSignal.timeout(3000) // مهلة 3 ثوانٍ
      }
    );

    if (!response.ok) return null;
    const data = await response.json();

    if (data && data.photos && data.photos.length > 0) {
      return data.photos[0].src.medium;
    }
  } catch (error) {
    console.error(`Pexels API Error for query "${query}":`, error.message);
  }
  return null;
}

// 2. دالة بناء العرض التقديمي (PPTX)
async function generatePptx({ topic, grade, subject, slides }) {
  const pptx = new pptxgen();

  // ضبط أبعاد العرض التقديمي (16:9 HD)
  pptx.layout = "LAYOUT_16x9";

  if (Array.isArray(slides)) {
    for (let i = 0; i < slides.length; i++) {
      const slideData = slides[i];
      const slide = pptx.addSlide();

      // عنوان السلايد
      slide.addText(slideData.title || `Slide ${i + 1}`, {
        x: 0.5,
        y: 0.5,
        w: 8.5,
        h: 0.8,
        fontSize: 24,
        bold: true,
        color: "003366",
        align: "right",
        rtl: true
      });

      // نص المحتوى (Bullets)
      slide.addText(slideData.content || "", {
        x: 4.5,
        y: 1.5,
        w: 4.5,
        h: 5.0,
        fontSize: 16,
        color: "333333",
        align: "right",
        rtl: true,
        valign: "top"
      });

      // جلب الصورة
      const searchQuery = slideData.visual_suggestion || slideData.title;
      const imageUrl = await fetchPexelsImageUrl(searchQuery);

      if (imageUrl) {
        slide.addImage({
          path: imageUrl,
          x: 0.5,
          y: 1.5,
          w: 3.8,
          h: 4.5,
          sizing: { type: "contain" }
        });
      }
    }
  }

  const buffer = await pptx.write("nodebuffer");
  return buffer;
}

module.exports = { generatePptx };
