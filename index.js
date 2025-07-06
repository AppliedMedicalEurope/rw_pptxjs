const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');

const app = express();
const port = process.env.PORT || 3000;

app.use(bodyParser.json({ limit: '10mb' }));

app.get('/', (req, res) => {
  res.send('✅ PPTXGenJS API is running');
});

function applySlideContent(slide, objects) {
  objects.forEach((obj, idx) => {
    try {
      if (obj.text && typeof obj.text.text === 'string') {
        // obj.text: { text: string, options: {} }
        slide.addText(obj.text.text, obj.text.options || {});
      } else if (obj.text && Array.isArray(obj.text)) {
        // obj.text: [ { text, options } ]
        const paragraphs = obj.text.map(t => ({
          text: t.text,
          options: t.options || {}
        }));
        slide.addText(paragraphs, obj.options || {});
      } else if (obj.table && obj.table.rows) {
        slide.addTable(obj.table.rows, obj.options || {});
      } else if (obj.image) {
        slide.addImage(obj.image);
      } else if (obj.rect) {
        slide.addShape('rect', obj.rect);
      } else if (obj.shape && obj.shape.type) {
        slide.addShape(obj.shape.type, obj.shape.options || {});
      } else {
        console.warn(`⚠️ Unknown object at index ${idx}`, obj);
      }
    } catch (err) {
      console.error(`❌ Error rendering object ${idx}:`, err.message);
    }
  });
}

app.post('/generate-pptx', async (req, res) => {
  try {
    const { slides = [], layout } = req.body;

    const pptx = new PptxGenJS();
    if (layout) pptx.layout = layout;

    slides.forEach((slideData, idx) => {
      const slide = pptx.addSlide();
      if (slideData.background) slide.background = { fill: slideData.background };

      if (Array.isArray(slideData.objects)) {
        applySlideContent(slide, slideData.objects);
      } else {
        console.warn(`⚠️ Slide ${idx} missing objects`);
      }
    });

    const base64 = await pptx.write('base64');
    const fileBuffer = Buffer.from(base64, 'base64');
    res.setHeader('Content-Disposition', 'attachment; filename=presentation.pptx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.send(fileBuffer);
  } catch (err) {
    console.error('❌ Error generating PPTX:', err);
    res.status(500).json({ error: 'Failed to generate presentation', details: err.message });
  }
});

app.listen(port, '0.0.0.0', () => {
  console.log(`🚀 PPTXGenJS API listening on http://0.0.0.0:${port}`);
});
