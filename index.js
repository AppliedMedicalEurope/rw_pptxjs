import express from 'express';
import pptxgen from 'pptxgenjs';

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

app.post('/generate', async (req, res) => {
  const { slides } = req.body;

  if (!slides || !Array.isArray(slides)) {
    return res.status(400).json({ error: 'Missing or invalid slides array' });
  }

  const pres = new pptxgen();

  slides.forEach(({ title, bullets }) => {
    const slide = pres.addSlide();
    slide.addText(title || 'No Title', { x: 1, y: 0.5, fontSize: 24, bold: true });
    if (bullets && Array.isArray(bullets)) {
      slide.addText(bullets.join('\n'), { x: 1, y: 1.5, fontSize: 18 });
    }
  });

  try {
    const buffer = await pres.stream(); // Get PPTX as Buffer
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="presentation.pptx"');
    res.send(buffer);
  } catch (err) {
    console.error('Error generating PPTX:', err);
    res.status(500).json({ error: 'Failed to generate presentation' });
  }
});

app.get('/', (req, res) => {
  res.send('PPTXGenJS PowerPoint API is running!');
});

app.listen(PORT, () => {
  console.log(`Server is listening on port ${PORT}`);
});
