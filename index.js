import express from 'express';
import pptxgen from 'pptxgenjs';

const app = express();
app.use(express.json());

app.post('/generate', async (req, res) => {
  const { slides } = req.body;

  if (!slides || !Array.isArray(slides)) {
    return res.status(400).json({ error: 'Invalid input' });
  }

  const pres = new pptxgen();

  slides.forEach(slideData => {
    const slide = pres.addSlide();
    slide.addText(slideData.title, { x: 0.5, y: 0.3, fontSize: 24, bold: true });

    const bullets = slideData.objects
      .map(obj => obj.text?.value || '')
      .join('\n')
      .split('\n')
      .map(line => line.trim())
      .filter(Boolean);

    slide.addText(bullets, {
      x: 0.5,
      y: 1.2,
      fontSize: 18,
      bullet: true,
      lineSpacingMultiple: 1.2
    });
  });

  const buffer = await pres.stream();

  res.setHeader('Content-Disposition', 'attachment; filename="slides.pptx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.send(buffer);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
