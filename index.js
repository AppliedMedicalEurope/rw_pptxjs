const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');

const app = express();
const port = process.env.PORT || 3000;

app.use(bodyParser.json({ limit: '10mb' }));

// Health check
app.get('/', (req, res) => {
  res.send('🚀 PPTXGenJS API is running on Railway!');
});

// Add slide objects (text, table, image)
function applySlideContent(slide, objects) {
  objects.forEach((obj, idx) => {
    try {
      if (obj.text && Array.isArray(obj.text)) {
        const paragraphs = obj.text.map(t => ({ text: t.text, options: t.options || {} }));
        slide.addText(paragraphs, obj.options || {});
      } else if (obj.table && obj.table.rows) {
        slide.addTable(obj.table.rows, obj.options || {});
      } else if (obj.image) {
        slide.addImage(obj.image);
      } else {
        console.warn(`⚠️ Unknown or unsupported object at index ${idx}:`, obj);
      }
    } catch (err) {
      console.error(`❌ Failed to add object at index ${idx}:`, err.message);
    }
  });
}

// Generate PPTX from custom JSON structure
app.post('/generate-pptx', async (req, res) => {
  try {
    const { slides = [], layout } = req.body;

    const pptx = new PptxGenJS();
    if (layout) {
      pptx.layout = layout;
    }

    slides.forEach((slideData, idx) => {
      const slide = pptx.addSlide();

      // Slide background
      if (slideData.background) {
        slide.background = slideData.background;
      }

      // Slide layout (if needed)
      if (slideData.layout) {
        slide.slideLayout = slideData.layout;
      }

      // Slide content
      if (Array.isArray(slideData.objects)) {
        applySlideContent(slide, slideData.objects);
      } else {
        console.warn(`⚠️ Slide ${idx} missing or invalid 'objects' array.`);
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

// Start the server — needed for Railway
app.listen(port, '0.0.0.0', () => {
  console.log(`🚀 PPTXGenJS API listening on http://0.0.0.0:${port}`);
});
