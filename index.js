const pptxgen = require("pptxgenjs");
const axios = require("axios");

// 1. دالة جلب رابط الصورة من Pexels API
async function fetchPexelsImageUrl(query) {
  const apiKey = process.env.PEXELS_API_KEY;
  if (!apiKey || !query) return null;

  try {
    const response = await axios.get("https://api.pexels.com/v1/search", {
      params: { query: query, per_page: 1 },
      headers: { Authorization: apiKey },
      timeout: 3000 // مهلة 3 ثوانٍ لعدم تعطيل السلايد إذا تأخرت الاستجابة
    });

    if (response.data && response.data.photos && response.data.photos.length > 0) {
      return response.data.photos[0].src.medium; // إرجاع رابط الصورة بحجم مناسب
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

  // تكرار على كل السلايدات المعالجة من n8n
  for (let i = 0; i < slides.length; i++) {
    const slideData = slides[i];
    const slide = pptx.addSlide();

    // إضافة عنوان السلايد
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

    // إضافة نص المحتوى (Bullets)
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

    // جلب الصورة من Pexels بناءً على visual_suggestion أو Title
    const searchQuery = slideData.visual_suggestion || slideData.title;
    const imageUrl = await fetchPexelsImageUrl(searchQuery);

    // إذا وُجدت صورة، يتم تضمينها في الشريحة بالجهة اليسرى
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

  // تصدير الملف كـ Buffer لإرساله إلى n8n
  const buffer = await pptx.write("nodebuffer");
  return buffer;
}

module.exports = { generatePptx };
