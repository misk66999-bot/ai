const express = require("express");
const { generatePptx } = require("./pptx-builder");

const app = express();
app.use(express.json());

// 1. مسار اختبار للسيرفر
app.get("/", (req, res) => {
  res.send("PPTX Generation Service is Live!");
});

// 2. مسار إنشاء ملف الباوربوينت
app.post("/generate", async (req, res) => {
  try {
    const { topic, grade, subject, slides } = req.body;
    
    if (!slides || !Array.isArray(slides)) {
      return res.status(400).json({ error: "Missing or invalid slides array" });
    }

    const pptxBuffer = await generatePptx({ topic, grade, subject, slides });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${encodeURIComponent(topic || "lesson")}.pptx"`
    );

    return res.send(pptxBuffer);
  } catch (error) {
    console.error("Error generating PPTX:", error);
    return res.status(500).json({ error: "Internal Server Error", details: error.message });
  }
});

// 🛑 السطر الأهم لحل مشكلة Render Port Binding:
const PORT = process.env.PORT || 10000;
app.listen(PORT, "0.0.0.0", () => {
  console.log(`Server running and listening on port ${PORT}`);
});
