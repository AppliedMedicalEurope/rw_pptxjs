import express from 'express';
import pptxgen from 'pptxgenjs';

const app = express();
app.use(express.json());

app.post('/generate', async (req, res) => {
  const { slides } = req.body;

  if (!slides || !Array.isArray(slides)) {
    return res.status(400).json({ error: 'Missing or invalid slides array' });
  }

  const pres = new pptxgen();

  slides.forEach(slideData => {
    const slide = pres.addSlide();
    slide.addText(slideData.title, { x: 0.5, y: 0.3, fontSize: 24, bold: true });

    const allText = slideData.objects
      .map(obj => obj.text?.value || '')
      .join('\n')
      .split('\n') // for bullets
      .map(line => line.trim())
      .filter(line => line.length > 0);

    // Add bullets below the title
    slide.addText(allText, {
      x: 0.5,
      y: 1.2,
      fontSize: 18,
      bullet: true,
      color: '363636',
      lineSpacingMultiple: 1.2
    });
  });

  const buffer = await pres.stream();

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', 'attachment; filename="presentation.pptx"');
  res.send(buffer);
});

app.listen(process.env.PORT || 3000, () => {
  console.log('Server is running');
});
