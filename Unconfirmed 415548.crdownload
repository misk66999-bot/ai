const fs = require("fs");
const { generatePptx } = require("./pptx-builder");

const sample = {
  topic: "الكهرباء ومصادرها",
  grade: "السادس",
  subject: "العلوم",
  slides: [
    {
      number: 1,
      title: "أهلاً بكم في الدرس",
      content:
        "الصف السادس - المادة العلوم - اسم الدرس علوم الطبيعه - نرحب بالطلاب ونقدم أجواء ممتعة",
      visual_suggestion: "صورة ملونة لطفل يقدم زهرة لشخص معبر",
      notes: "الترحيب بالطلاب وتقديم أجواء ممتعة",
    },
    {
      number: 2,
      title: "نتاجات التعلم",
      content:
        "التعرف على المصادر الرئيسية للكهرباء\nتعلم استخدامات الكهرباء الشائعة\nمعرفة كيفية توليد الكهرباء",
      visual_suggestion: "قائمة بخطوط ملونة متعددة الألوان",
      notes: "تأكد من فهم الطلاب لأهداف الدرس",
    },
    {
      number: 3,
      title: "هل تعلم؟",
      content: "البرق هو شكل طبيعي من الكهرباء!",
      visual_suggestion: "صورة برق في السماء",
      notes: "مثال تشويقي قصير",
    },
    {
      number: 4,
      title: "شرح المفهوم الأساسي",
      content:
        "الكهرباء الساكنة تنتج عن احتكاك الأجسام\nالكهرباء التيارية تنتج عن حركة الإلكترونات\nهناك مصادر متجددة وغير متجددة للكهرباء\nمحطات توليد الطاقة تحول الطاقة الحرارية إلى كهرباء",
      visual_suggestion: "رسم توضيحي لدائرة كهربائية بسيطة",
      notes: "اشرح بالتفصيل مع أمثلة من الحياة اليومية",
    },
  ],
};

generatePptx(sample).then((buffer) => {
  fs.writeFileSync("/home/claude/pptx-service/sample-output.pptx", buffer);
  console.log("Wrote sample-output.pptx, size:", buffer.length);
});
