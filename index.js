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
      // Case 1: Simple text string (most common title format)
      if (typeof obj.text === 'string') {
        slide.addText(obj.text, obj.options || {});
      }

      // Case 2: Rich text block (single object with nested text/options)
      else if (obj.text && typeof obj.text.text === 'string') {
        slide.addText(obj.text.text, obj.text.options || {});
      }

      // Case 3: Paragraph array (used for bullet lists or mixed formatting)
      else if (obj.text && Array.isArray(obj.text)) {
        const isBullet = obj.options?.bullet === true;

        const paragraphs = obj.text.map(t => {
          let cleanText = t.text?.trim() || '';
          if (isBullet && cleanText.startsWith('•')) {
            cleanText = cleanText.slice(1).trim();
          }

          return {
            text: cleanText,
            options: {
              ...t.options,
              bullet: isBullet || t.options?.bullet === true
            }
          };
        });

        slide.addText(paragraphs, obj.options || {});
      }

      // Case 4: Table
      else if (obj.table && obj.table.rows) {
        slide.addTable(obj.table.rows, obj.options || {});
      }

      // Case 5: Image
      else if (obj.image) {
        slide.addImage(obj.image);
      }

      // Case 6: Rect
      else if (obj.rect) {
        slide.addShape('rect', obj.rect);
      }

      // Case 7: Shape
      else if (obj.shape && obj.shape.type) {
        slide.addShape(obj.shape.type, obj.shape.options || {});
      }

      // Case 8: Chart
      else if (obj.chart && obj.chart.type && obj.chart.data) {
        slide.addChart(obj.chart.type, obj.chart.data, obj.chart.options || {});
      }

      // Case 9: Media
      else if (obj.media) {
        slide.addMedia(obj.media);
      }

      // Fallback
      else {
        console.warn(`⚠️ Unknown object at index ${idx}:`, obj);
      }
    } catch (err) {
      console.error(`❌ Error rendering object at index ${idx}:`, err.message);
    }
  });
}

app.post('/generate-pptx', async (req, res) => {
  try {
    const { slides = [], layout } = req.body;

    const pptx = new PptxGenJS();

    // Optional global layout
    if (layout && layout.startsWith('LAYOUT_')) {
      pptx.layout = layout;
    }

    slides.forEach((slideData, idx) => {
      const slide = pptx.addSlide();

      // Set background color or image
      if (slideData.background) {
        if (slideData.background.color) {
          slide.background = { fill: slideData.background.color };
        } else {
          slide.background = slideData.background;
        }
      }

      // Slide notes (optional)
      if (slideData.notes) {
        slideData.options = slideData.options || {};
        slideData.options.notes = slideData.notes;
      }

      // Slide objects
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

app.listen(port, '0.0.0.0', () => {
  console.log(`🚀 PPTXGenJS API listening on http://0.0.0.0:${port}`);
});
