// 2. دالة بناء العرض التقديمي (PPTX)
async function generatePptx(payload) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";

  // فك التداخل: محاولة استخراج slides سواء كانت مباشرة أو مغلفة
  let slidesData = payload.slides;
  
  // إذا كانت slides عبارة عن Object يحتوي بداخله على slides أخرى
  if (slidesData && !Array.isArray(slidesData) && Array.isArray(slidesData.slides)) {
    slidesData = slidesData.slides;
  }

  const topic = payload.topic || (payload.slides && payload.slides.topic) || "عرض تقديمي";
  const grade = payload.grade || (payload.slides && payload.slides.grade) || "";
  const subject = payload.subject || (payload.slides && payload.slides.subject) || "";

  // 1️⃣ شريحة عنوان رئيسية
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

  // 2️⃣ تكرار وإنشاء شرائح الدروس
  if (Array.isArray(slidesData) && slidesData.length > 0) {
    for (let i = 0; i < slidesData.length; i++) {
      const slideData = slidesData[i];
      const slide = pptx.addSlide();

      // عنوان الشريحة
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

      // تحويل محتوى الشريحة لنص متنسق لو جاء كـ Array من Groq
      let contentText = slideData.content;
      if (Array.isArray(contentText)) {
        contentText = contentText.join("\n• ");
      }

      // محتوى الشريحة
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

      // جلب الصورة
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
