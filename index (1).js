const express = require("express");
const { generatePptx } = require("./pptx-builder");

const app = express();
app.use(express.json({ limit: "5mb" }));

app.get("/", (req, res) => {
  res.json({ status: "ok", service: "lesson-pptx-service" });
});

app.post("/generate", async (req, res) => {
  try {
    const { topic, grade, subject, slides } = req.body || {};
    if (!Array.isArray(slides) || slides.length === 0) {
      return res.status(400).json({ error: "Missing or empty 'slides' array" });
    }

    const buffer = await generatePptx({ topic, grade, subject, slides });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="lesson-presentation.pptx"`
    );
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Generation failed", details: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`lesson-pptx-service listening on port ${PORT}`);
});
