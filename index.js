const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const rateLimit = require('express-rate-limit');
const multer = require('multer');
const sharp = require('sharp');
const Jimp = require('jimp');
const fetch = require('node-fetch');
const { v4: uuidv4 } = require('uuid');
const PptxGenJS = require('pptxgenjs');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(helmet({
  contentSecurityPolicy: false
}));
app.use(compression());
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // limit each IP to 100 requests per windowMs
  message: 'Too many requests from this IP'
});
app.use('/api/', limiter);

// Multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ 
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Helper function to process images
async function processImage(buffer, options = {}) {
  try {
    let processedBuffer;
    
    if (options.resize) {
      const { width, height } = options.resize;
      processedBuffer = await sharp(buffer)
        .resize(width, height, { fit: 'inside' })
        .jpeg({ quality: 90 })
        .toBuffer();
    } else {
      processedBuffer = buffer;
    }
    
    return `data:image/jpeg;base64,${processedBuffer.toString('base64')}`;
  } catch (error) {
    console.error('Image processing error:', error);
    throw error;
  }
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'healthy', timestamp: new Date().toISOString() });
});

// Root endpoint with API documentation
app.get('/', (req, res) => {
  res.json({
    message: 'PptxGenJS API Server',
    version: '1.0.0',
    endpoints: {
      'GET /health': 'Health check',
      'POST /api/presentation/create': 'Create a basic presentation',
      'POST /api/presentation/advanced': 'Create advanced presentation with all features',
      'POST /api/slide/add-text': 'Add text to presentation',
      'POST /api/slide/add-image': 'Add image to presentation',
      'POST /api/slide/add-chart': 'Add chart to presentation',
      'POST /api/slide/add-table': 'Add table to presentation',
      'POST /api/slide/add-shape': 'Add shape to presentation',
      'POST /api/presentation/template': 'Create presentation from template',
      'POST /api/presentation/bulk': 'Create multiple presentations',
      'POST /api/presentation/merge': 'Merge multiple presentations'
    }
  });
});

// Create basic presentation
app.post('/api/presentation/create', async (req, res) => {
  try {
    const { title = 'My Presentation', slides = [] } = req.body;
    
    const pptx = new PptxGenJS();
    pptx.author = 'PptxGenJS API';
    pptx.company = 'Railway App';
    pptx.title = title;
    
    // Add title slide
    const titleSlide = pptx.addSlide();
    titleSlide.addText(title, {
      x: 1,
      y: 1,
      w: 8,
      h: 2,
      fontSize: 32,
      bold: true,
      color: '0066CC',
      align: 'center'
    });
    
    // Add content slides
    slides.forEach((slideData, index) => {
      const slide = pptx.addSlide();
      
      if (slideData.title) {
        slide.addText(slideData.title, {
          x: 0.5,
          y: 0.5,
          w: 9,
          h: 1,
          fontSize: 24,
          bold: true,
          color: '333333'
        });
      }
      
      if (slideData.content) {
        slide.addText(slideData.content, {
          x: 0.5,
          y: 1.5,
          w: 9,
          h: 4,
          fontSize: 16,
          color: '666666'
        });
      }
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${title}.pptx"`);
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating presentation:', error);
    res.status(500).json({ error: 'Failed to create presentation', details: error.message });
  }
});

// Create advanced presentation with all features
app.post('/api/presentation/advanced', async (req, res) => {
  try {
    const { 
      title = 'Advanced Presentation',
      theme = 'default',
      slides = [],
      masterSlide = null,
      layout = 'LAYOUT_16x9'
    } = req.body;
    
    const pptx = new PptxGenJS();
    
    // Set presentation properties
    pptx.author = 'PptxGenJS API';
    pptx.company = 'Railway App';
    pptx.title = title;
    pptx.subject = 'Generated via API';
    pptx.layout = layout;
    
    // Define master slide if provided
    if (masterSlide) {
      pptx.defineSlideMaster({
        title: masterSlide.name || 'MASTER_SLIDE',
        background: masterSlide.background || { color: 'FFFFFF' },
        objects: masterSlide.objects || []
      });
    }
    
    // Process each slide
    for (const slideData of slides) {
      const slide = pptx.addSlide({ masterName: slideData.masterName });
      
      // Add background
      if (slideData.background) {
        if (slideData.background.color) {
          slide.background = { color: slideData.background.color };
        } else if (slideData.background.image) {
          slide.background = { data: slideData.background.image };
        }
      }
      
      // Add elements
      if (slideData.elements) {
        for (const element of slideData.elements) {
          switch (element.type) {
            case 'text':
              slide.addText(element.text, {
                x: element.x || 0.5,
                y: element.y || 0.5,
                w: element.w || 9,
                h: element.h || 1,
                fontSize: element.fontSize || 16,
                fontFace: element.fontFace || 'Arial',
                color: element.color || '000000',
                bold: element.bold || false,
                italic: element.italic || false,
                underline: element.underline || false,
                align: element.align || 'left',
                valign: element.valign || 'top',
                rotate: element.rotate || 0,
                shadow: element.shadow || null,
                glow: element.glow || null
              });
              break;
              
            case 'image':
              slide.addImage({
                data: element.data,
                x: element.x || 1,
                y: element.y || 1,
                w: element.w || 6,
                h: element.h || 4,
                rotate: element.rotate || 0,
                transparency: element.transparency || 0,
                rounding: element.rounding || false
              });
              break;
              
            case 'shape':
              slide.addShape(element.shape || 'RECTANGLE', {
                x: element.x || 1,
                y: element.y || 1,
                w: element.w || 2,
                h: element.h || 2,
                fill: element.fill || { color: '0066CC' },
                line: element.line || { color: '000000', width: 1 },
                rotate: element.rotate || 0
              });
              break;
              
            case 'table':
              slide.addTable(element.data || [], {
                x: element.x || 0.5,
                y: element.y || 1.5,
                w: element.w || 9,
                h: element.h || 4,
                colW: element.colW || null,
                rowH: element.rowH || null,
                border: element.border || { type: 'solid', color: '666666', pt: 1 },
                fill: element.fill || { color: 'F7F7F7' },
                fontSize: element.fontSize || 12,
                color: element.color || '000000'
              });
              break;
              
            case 'chart':
              slide.addChart(element.chartType || 'COLUMN', element.data || [], {
                x: element.x || 1,
                y: element.y || 1,
                w: element.w || 8,
                h: element.h || 5,
                title: element.title || '',
                showTitle: element.showTitle || false,
                showLegend: element.showLegend || true,
                legendPos: element.legendPos || 'r',
                showPercent: element.showPercent || false,
                dataLabelColor: element.dataLabelColor || '000000',
                dataLabelFontSize: element.dataLabelFontSize || 12
              });
              break;
              
            case 'media':
              if (element.mediaType === 'video') {
                slide.addMedia({
                  type: 'video',
                  data: element.data,
                  x: element.x || 1,
                  y: element.y || 1,
                  w: element.w || 6,
                  h: element.h || 4
                });
              } else if (element.mediaType === 'audio') {
                slide.addMedia({
                  type: 'audio',
                  data: element.data,
                  x: element.x || 1,
                  y: element.y || 1
                });
              }
              break;
              
            case 'notes':
              slide.addNotes(element.text || '');
              break;
          }
        }
      }
    }
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${title}.pptx"`);
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating advanced presentation:', error);
    res.status(500).json({ error: 'Failed to create presentation', details: error.message });
  }
});

// Add text slide
app.post('/api/slide/add-text', async (req, res) => {
  try {
    const { text, options = {} } = req.body;
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    slide.addText(text, {
      x: options.x || 0.5,
      y: options.y || 0.5,
      w: options.w || 9,
      h: options.h || 1,
      fontSize: options.fontSize || 16,
      fontFace: options.fontFace || 'Arial',
      color: options.color || '000000',
      bold: options.bold || false,
      italic: options.italic || false,
      underline: options.underline || false,
      align: options.align || 'left',
      valign: options.valign || 'top',
      rotate: options.rotate || 0,
      shadow: options.shadow || null,
      glow: options.glow || null,
      hyperlink: options.hyperlink || null,
      bullet: options.bullet || false,
      indentLevel: options.indentLevel || 0
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="text-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error adding text slide:', error);
    res.status(500).json({ error: 'Failed to add text slide', details: error.message });
  }
});

// Add image slide (with file upload)
app.post('/api/slide/add-image', upload.single('image'), async (req, res) => {
  try {
    const { options = {} } = req.body;
    let imageData;
    
    if (req.file) {
      imageData = await processImage(req.file.buffer, options);
    } else if (req.body.imageData) {
      imageData = req.body.imageData;
    } else if (req.body.imageUrl) {
      const response = await fetch(req.body.imageUrl);
      const buffer = await response.buffer();
      imageData = await processImage(buffer, options);
    } else {
      return res.status(400).json({ error: 'No image provided' });
    }
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    slide.addImage({
      data: imageData,
      x: options.x || 1,
      y: options.y || 1,
      w: options.w || 6,
      h: options.h || 4,
      rotate: options.rotate || 0,
      transparency: options.transparency || 0,
      rounding: options.rounding || false,
      hyperlink: options.hyperlink || null,
      sizing: options.sizing || null
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="image-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error adding image slide:', error);
    res.status(500).json({ error: 'Failed to add image slide', details: error.message });
  }
});

// Add chart slide
app.post('/api/slide/add-chart', async (req, res) => {
  try {
    const { 
      chartType = 'COLUMN',
      data = [],
      options = {}
    } = req.body;
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    slide.addChart(chartType, data, {
      x: options.x || 1,
      y: options.y || 1,
      w: options.w || 8,
      h: options.h || 5,
      title: options.title || '',
      showTitle: options.showTitle || false,
      showLegend: options.showLegend || true,
      legendPos: options.legendPos || 'r',
      showPercent: options.showPercent || false,
      dataLabelColor: options.dataLabelColor || '000000',
      dataLabelFontSize: options.dataLabelFontSize || 12,
      catAxisTitle: options.catAxisTitle || '',
      valAxisTitle: options.valAxisTitle || '',
      catAxisTitleColor: options.catAxisTitleColor || '000000',
      valAxisTitleColor: options.valAxisTitleColor || '000000',
      catAxisTitleFontSize: options.catAxisTitleFontSize || 12,
      valAxisTitleFontSize: options.valAxisTitleFontSize || 12,
      showCatAxisTitle: options.showCatAxisTitle || false,
      showValAxisTitle: options.showValAxisTitle || false,
      catGridLine: options.catGridLine || null,
      valGridLine: options.valGridLine || null,
      chartColors: options.chartColors || null,
      invertedColors: options.invertedColors || null
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="chart-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error adding chart slide:', error);
    res.status(500).json({ error: 'Failed to add chart slide', details: error.message });
  }
});

// Add table slide
app.post('/api/slide/add-table', async (req, res) => {
  try {
    const { 
      data = [],
      options = {}
    } = req.body;
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    slide.addTable(data, {
      x: options.x || 0.5,
      y: options.y || 1.5,
      w: options.w || 9,
      h: options.h || 4,
      colW: options.colW || null,
      rowH: options.rowH || null,
      border: options.border || { type: 'solid', color: '666666', pt: 1 },
      fill: options.fill || { color: 'F7F7F7' },
      fontSize: options.fontSize || 12,
      color: options.color || '000000',
      align: options.align || 'left',
      valign: options.valign || 'top',
      bold: options.bold || false,
      italic: options.italic || false,
      underline: options.underline || false,
      margin: options.margin || 0.1,
      autoPage: options.autoPage || false,
      autoPageRepeatHeader: options.autoPageRepeatHeader || false,
      autoPageHeaderRows: options.autoPageHeaderRows || 1,
      autoPageLineWeight: options.autoPageLineWeight || 0
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="table-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error adding table slide:', error);
    res.status(500).json({ error: 'Failed to add table slide', details: error.message });
  }
});

// Add shape slide
app.post('/api/slide/add-shape', async (req, res) => {
  try {
    const { 
      shape = 'RECTANGLE',
      options = {}
    } = req.body;
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    slide.addShape(shape, {
      x: options.x || 1,
      y: options.y || 1,
      w: options.w || 2,
      h: options.h || 2,
      fill: options.fill || { color: '0066CC' },
      line: options.line || { color: '000000', width: 1 },
      rotate: options.rotate || 0,
      flipH: options.flipH || false,
      flipV: options.flipV || false,
      shadow: options.shadow || null,
      glow: options.glow || null,
      hyperlink: options.hyperlink || null
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="shape-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error adding shape slide:', error);
    res.status(500).json({ error: 'Failed to add shape slide', details: error.message });
  }
});

// Create presentation from template
app.post('/api/presentation/template', async (req, res) => {
  try {
    const { 
      template = 'business',
      data = {},
      customizations = {}
    } = req.body;
    
    const pptx = new PptxGenJS();
    
    // Apply template-specific settings
    const templates = {
      business: {
        theme: { color: '0066CC', bgColor: 'FFFFFF' },
        font: 'Arial',
        slides: [
          { type: 'title', title: data.title || 'Business Presentation', subtitle: data.subtitle || '' },
          { type: 'agenda', items: data.agenda || [] },
          { type: 'content', title: 'Key Points', content: data.content || [] },
          { type: 'charts', data: data.charts || [] },
          { type: 'conclusion', title: 'Thank You', content: data.conclusion || '' }
        ]
      },
      education: {
        theme: { color: '4CAF50', bgColor: 'F8F9FA' },
        font: 'Calibri',
        slides: [
          { type: 'title', title: data.title || 'Educational Content', subtitle: data.subtitle || '' },
          { type: 'objectives', items: data.objectives || [] },
          { type: 'content', title: 'Learning Material', content: data.content || [] },
          { type: 'quiz', questions: data.quiz || [] },
          { type: 'summary', title: 'Summary', content: data.summary || '' }
        ]
      },
      creative: {
        theme: { color: 'FF6B6B', bgColor: '2C3E50' },
        font: 'Georgia',
        slides: [
          { type: 'title', title: data.title || 'Creative Presentation', subtitle: data.subtitle || '' },
          { type: 'portfolio', items: data.portfolio || [] },
          { type: 'showcase', content: data.showcase || [] },
          { type: 'contact', info: data.contact || {} }
        ]
      }
    };
    
    const selectedTemplate = templates[template] || templates.business;
    
    // Apply customizations
    Object.assign(selectedTemplate, customizations);
    
    // Generate slides based on template
    selectedTemplate.slides.forEach(slideConfig => {
      const slide = pptx.addSlide();
      
      switch (slideConfig.type) {
        case 'title':
          slide.addText(slideConfig.title, {
            x: 1, y: 2, w: 8, h: 2,
            fontSize: 32, bold: true,
            color: selectedTemplate.theme.color,
            align: 'center'
          });
          if (slideConfig.subtitle) {
            slide.addText(slideConfig.subtitle, {
              x: 1, y: 4, w: 8, h: 1,
              fontSize: 18, color: '666666',
              align: 'center'
            });
          }
          break;
          
        case 'agenda':
        case 'objectives':
          slide.addText(slideConfig.type === 'agenda' ? 'Agenda' : 'Learning Objectives', {
            x: 0.5, y: 0.5, w: 9, h: 1,
            fontSize: 24, bold: true,
            color: selectedTemplate.theme.color
          });
          slideConfig.items.forEach((item, index) => {
            slide.addText(`${index + 1}. ${item}`, {
              x: 1, y: 1.5 + (index * 0.5), w: 8, h: 0.5,
              fontSize: 16, color: '333333'
            });
          });
          break;
          
        case 'content':
          slide.addText(slideConfig.title, {
            x: 0.5, y: 0.5, w: 9, h: 1,
            fontSize: 24, bold: true,
            color: selectedTemplate.theme.color
          });
          if (Array.isArray(slideConfig.content)) {
            slideConfig.content.forEach((item, index) => {
              slide.addText(`• ${item}`, {
                x: 1, y: 1.5 + (index * 0.5), w: 8, h: 0.5,
                fontSize: 16, color: '333333'
              });
            });
          } else {
            slide.addText(slideConfig.content, {
              x: 1, y: 1.5, w: 8, h: 4,
              fontSize: 16, color: '333333'
            });
          }
          break;
      }
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${template}-presentation.pptx"`);
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating template presentation:', error);
    res.status(500).json({ error: 'Failed to create template presentation', details: error.message });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Unhandled error:', error);
  res.status(500).json({ 
    error: 'Internal server error',
    details: process.env.NODE_ENV === 'development' ? error.message : 'Something went wrong'
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: 'Endpoint not found' });
});

app.listen(PORT, () => {
  console.log(`PptxGenJS API Server running on port ${PORT}`);
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`Health check: http://localhost:${PORT}/health`);
});

module.exports = app;
