const express = require('express');
const bodyParser = require('body-parser');
const PPTX = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(bodyParser.json({ limit: '10mb' }));

// Utility to parse percentage or absolute values
function parsePos(val, total = 10) {
  if (typeof val === 'string' && val.endsWith('%')) {
    return (parseFloat(val) / 100) * total;
  }
  return parseFloat(val);
}

function applyElement(slide, element) {
  if (element.text) {
    const paragraphs = element.text.map(p => ({
      text: p.text,
      options: p.options || {}
    }));

    const options = {
      x: parsePos(element.options?.x || 0),
      y: parsePos(element.options?.y || 0),
      w: parsePos(element.options?.w || 10),
      h: parsePos(element.options?.h || 1),
      align: element.options?.align || 'left',
      bullet: element.options?.bullet || false,
      ...element.options
    };

    slide.addText(paragraphs, options);
  } else if (element.table) {
    const rows = element.table.rows.map(row =>
      row.map(cell => cell.text || "")
    );
    const cellOptions = element.table.rows.map(row =>
      row.map(cell => cell.options || {})
    );

    const options = {
      x: parsePos(element.options?.x || 1),
      y: parsePos(element.options?.y || 1),
      w: parsePos(element.options?.w || 8),
      h: parsePos(element.options?.h || 5),
      ...element.options
    };

    slide.addTable(rows, { ...options, cellOpts: cellOptions });
  }
  // You can add other types like images, charts here
}

app.post('/generate-pptx', async (req, res) => {
  try {
    const input = req.body;
    const pptx = new PPTX();

    if (input.title) pptx.author = input.title;
    if (input.author) pptx.author = input.author;
    if (input.subject) pptx.subject = input.subject;

    input.slides.forEach(slideData => {
      const slide = pptx.addSlide();

      if (slideData.options?.bkgd || slideData.background?.color) {
        slide.background = {
          fill: slideData.options?.bkgd || slideData.background?.color
        };
      }

      const elements = slideData.elements || slideData.objects || [];
      elements.forEach(el => applyElement(slide, el));
    });

    const fileName = `generated-${Date.now()}.pptx`;
    const filePath = path.join('/tmp', fileName);
    await pptx.writeFile({ fileName: filePath });

    res.download(filePath, 'presentation.pptx', err => {
      if (err) console.error('Download error:', err);
      fs.unlink(filePath, () => {});
    });
  } catch (err) {
    console.error('Error generating PPTX:', err);
    res.status(500).json({ error: err.message });
  }
});

// ✅ Port binding for Railway & local
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () =>
  console.log(`PPTXGenJS API running on http://0.0.0.0:${PORT}`)
);
