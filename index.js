const express = require('express');
const bodyParser = require('body-parser');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Parse JSON payloads
app.use(bodyParser.json({ limit: '10mb' }));

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

    const tmpFilePath = path.join(__dirname, `presentation-${Date.now()}.pptx`);

    // Save and send the file
    await pptx.writeFile({ fileName: tmpFilePath });

    res.download(tmpFilePath, 'presentation.pptx', err => {
      fs.unlink(tmpFilePath, () => {}); // Clean up temp file
      if (err) {
        console.error('Download error:', err);
      }
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to generate presentation' });
  }
});

app.listen(port, () => {
  console.log(`PPTXGenJS API listening at http://localhost:${port}`);
});
