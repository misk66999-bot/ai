const pptxgen = require("pptxgenjs");

// 1. دالة جلب رابط الصورة من Pexels
async function fetchPexelsImageUrl(query) {
  const apiKey = process.env.PEXELS_API_KEY;
  if (!apiKey || !query) return null;

  try {
    const response = await fetch(
      `https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}&per_page=1`,
      {
        headers: { Authorization: apiKey },
        signal: AbortSignal.timeout(3000)
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
async function generatePptx(payload) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";

  // استخراج البيانات وتجاوز التداخل
  let slidesData = payload.slides;
  if (slidesData && !Array.isArray(slidesData) && Array.isArray(slidesData.slides)) {
    slidesData = slidesData.slides;
  }

  const topic = payload.topic || (payload.slides && payload.slides.topic) || "عرض تقديمي";
  const grade = payload.grade || (payload.slides && payload.slides.grade) || "";
  const subject = payload.subject || (payload.slides && payload.slides.subject) || "";

  // شريحة العنوان
  const titleSlide = pptx.addSlide();
  titleSlide.addText(topic, {
    x: 1.0,
    y: 2.0,
    w: 11.33,
    h: 1.5,
    fontSize: 36,
    bold: true,
    color: "003366",
    align: "center",
    rtl: true
  });

  if (grade || subject) {
    titleSlide.addText(`${subject} ${grade ? '- ' + grade : ''}`, {
      x: 1.0,
      y: 3.8,
      w: 11.33,
      h: 1.0,
      fontSize: 20,
      color: "555555",
      align: "center",
      rtl: true
    });
  }

  // شرائح المحتوى
  if (Array.isArray(slidesData) && slidesData.length > 0) {
    for (let i = 0; i < slidesData.length; i++) {
      const slideData = slidesData[i];
      const slide = pptx.addSlide();

      slide.addText(slideData.title || `الشريحة ${i + 1}`, {
        x: 0.5,
        y: 0.5,
        w: 12.33,
        h: 0.8,
        fontSize: 24,
        bold: true,
        color: "003366",
        align: "right",
        rtl: true
      });

      let contentText = slideData.content;
      if (Array.isArray(contentText)) {
        contentText = contentText.join("\n• ");
      }

      slide.addText(contentText || "", {
        x: 4.8,
        y: 1.5,
        w: 7.0,
        h: 5.0,
        fontSize: 16,
        color: "333333",
        align: "right",
        rtl: true,
        valign: "top"
      });

      const searchQuery = slideData.visual_suggestion || slideData.title;
      const imageUrl = await fetchPexelsImageUrl(searchQuery);

      if (imageUrl) {
        slide.addImage({
          path: imageUrl,
          x: 0.5,
          y: 1.5,
          w: 4.0,
          h: 4.5,
          sizing: { type: "contain" }
        });
      }
    }
  }

  const buffer = await pptx.write("nodebuffer");
  return buffer;
}

// 🛑 السطر الأهم لمنع خطأ TypeError
module.exports = { generatePptx };
