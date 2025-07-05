import express from 'express';
import pptxgen from 'pptxgenjs';

const app = express();
app.use(express.json());

app.get('/', (req, res) => {
  res.send('PptxGenJS API is running!');
});

app.post('/generate', async (req, res) => {
  const { slides } = req.body;

  if (!slides || !Array.isArray(slides)) {
    return res.status(400).json({ error: 'Missing or invalid slides array' });
  }

  const pres = new pptxgen();

  slides.forEach(slideData => {
    const slide = pres.addSlide();

    // Title at the top
    slide.addText(slideData.title, {
      x: 0.5,
      y: 0.3,
      fontSize: 24,
      bold: true
    });

    // Extract and flatten text.value from all objects[]
    const bodyText = slideData.objects
      .map(obj => obj.text?.value || '')
      .join('\n')                   // combine multiline strings
      .split('\n')                  // turn into bullets
      .map(line => line.trim())
      .filter(line => line.length > 0);  // skip empty lines

    if (bodyText.length > 0) {
      slide.addText(bodyText, {
        x: 0.7,
        y: 1.2,
        fontSize: 18,
        bullet: true,
        color: '363636',
        lineSpacingMultiple: 1.2
      });
    }
  });

  try {
    const buffer = await pres.stream();
    res.setHeader('Content-Disposition', 'attachment; filename="presentation.pptx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.send(buffer);
  } catch (err) {
    console.error('Error generating PPTX:', err);
    res.status(500).json({ error: 'Failed to generate presentation' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
