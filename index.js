const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(bodyParser.json({ limit: '10mb' }));

// Health check
app.get('/', (req, res) => {
  res.send('🚀 PPTXGenJS API is running on Railway!');
});

// Helper to apply slide content
function applySlideContent(slide, elements) {
  elements.forEach(el => {
    const { type, options } = el;

    switch (type) {
      case 'text':
        slide.addText(options.text, options.props || {});
        break;
      case 'image':
        slide.addImage(options);
        break;
      case 'shape':
        slide.addShape(options.shape, options.props || {});
        break;
      case 'chart':
        slide.addChart(options.type, options.data, options.props || {});
        break;
      case 'table':
        slide.addTable(options.data, options.props || {});
        break;
      case 'media':
        slide.addMedia(options);
        break;
      default:
        console.warn(`Unknown element type: ${type}`);
    }
  });
}

// Main route
app.post('/generate-pptx', async (req, res) => {
  try {
    const { slides = [], layout } = req.body;

    const pptx = new PptxGenJS();
    if (layout) {
      pptx.layout = layout;
    }

    slides.forEach(slideData => {
      const slide = pptx.addSlide(slideData.options || {});
      if (slideData.elements) {
        applySlideContent(slide, slideData.elements);
      }
    });

    const base64 = await pptx.write('base64');

    const fileBuffer = Buffer.from(base64, 'base64');
    res.setHeader('Content-Disposition', 'attachment; filename=presentation.pptx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.send(fileBuffer);
  } catch (err) {
    console.error('Error generating PPTX:', err);
    res.status(500).json({ error: 'Failed to generate presentation', details: err.message });
  }
});

// Start the server on 0.0.0.0 for Railway
app.listen(port, '0.0.0.0', () => {
  console.log(`🚀 PPTXGenJS API listening on http://0.0.0.0:${port}`);
});
