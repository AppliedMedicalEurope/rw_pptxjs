const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const rateLimit = require('express-rate-limit');
const multer = require('multer');
const sharp = require('sharp');
const axios = require('axios');
const { v4: uuidv4 } = require('uuid');
const PptxGenJS = require('pptxgenjs');

const app = express();
const PORT = process.env.PORT || 3000;

// Handle graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully');
  process.exit(0);
});

// Middleware with Railway-specific configurations
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'"],
      scriptSrc: ["'self'"],
      imgSrc: ["'self'", "data:", "https:"],
    },
  },
  crossOriginEmbedderPolicy: false
}));

app.use(compression());
app.use(cors({
  origin: process.env.NODE_ENV === 'production' ? 
    [/\.railway\.app$/, /localhost/] : 
    true,
  credentials: true
}));

app.use(express.json({ 
  limit: '10mb',
  verify: (req, res, buf) => {
    try {
      JSON.parse(buf);
    } catch (e) {
      res.status(400).json({ error: 'Invalid JSON' });
      return;
    }
  }
}));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Rate limiting - more conservative for Railway
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 50, // reduced limit for Railway
  message: { error: 'Too many requests from this IP, please try again later' },
  standardHeaders: true,
  legacyHeaders: false
});
app.use('/api/', limiter);

// Multer for file uploads - reduced limits for Railway
const storage = multer.memoryStorage();
const upload = multer({ 
  storage: storage,
  limits: { 
    fileSize: 5 * 1024 * 1024, // 5MB limit
    files: 1
  },
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only JPEG, PNG, GIF, and WebP are allowed.'));
    }
  }
});

// Helper function to process images with better error handling
async function processImage(buffer, options = {}) {
  try {
    if (!buffer || buffer.length === 0) {
      throw new Error('Invalid image buffer');
    }

    let processedBuffer = buffer;
    
    if (options.resize && options.resize.width && options.resize.height) {
      const { width, height } = options.resize;
      processedBuffer = await sharp(buffer)
        .resize(Math.min(width, 1920), Math.min(height, 1080), { 
          fit: 'inside',
          withoutEnlargement: true
        })
        .jpeg({ quality: 85, progressive: true })
        .toBuffer();
    } else {
      // Convert to JPEG if not already
      processedBuffer = await sharp(buffer)
        .jpeg({ quality: 85, progressive: true })
        .toBuffer();
    }
    
    return `data:image/jpeg;base64,${processedBuffer.toString('base64')}`;
  } catch (error) {
    console.error('Image processing error:', error);
    throw new Error('Failed to process image: ' + error.message);
  }
}

// Helper function to fetch images from URL with timeout
async function fetchImageFromUrl(url, timeout = 10000) {
  try {
    const response = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout: timeout,
      maxContentLength: 5 * 1024 * 1024, // 5MB limit
      headers: {
        'User-Agent': 'PptxGenJS-API/1.0'
      }
    });
    
    return Buffer.from(response.data);
  } catch (error) {
    console.error('Error fetching image from URL:', error);
    throw new Error('Failed to fetch image from URL: ' + error.message);
  }
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.status(200).json({ 
    status: 'healthy', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    memory: process.memoryUsage(),
    env: process.env.NODE_ENV || 'development'
  });
});

// Root endpoint
app.get('/', (req, res) => {
  res.status(200).json({
    message: 'PptxGenJS API Server - Railway Deployment',
    version: '1.0.0',
    status: 'running',
    endpoints: {
      'GET /health': 'Health check',
      'POST /api/presentation/simple': 'Create a simple presentation (recommended)',
      'POST /api/presentation/create': 'Create a basic presentation',
      'POST /api/slide/text': 'Add text slide',
      'POST /api/slide/image': 'Add image slide'
    },
    documentation: 'Send POST requests with JSON data to create presentations'
  });
});


// Add this new endpoint to your Express.js server to handle the JSON structure properly

app.post('/api/presentation/from-json', async (req, res) => {
  try {
    const { slides = [] } = req.body;
    
    console.log('Processing presentation with', slides.length, 'slides');
    
    const pptx = new PptxGenJS();
    pptx.author = 'JSON Converter';
    pptx.title = 'Converted Presentation';
    pptx.layout = 'LAYOUT_16x9';
    
    // Process each slide from the JSON structure
    slides.forEach((slideData, slideIndex) => {
      try {
        const slide = pptx.addSlide();
        console.log(`Processing slide ${slideIndex + 1} with ${slideData.objects?.length || 0} objects`);
        
        // Process each object on the slide
        if (slideData.objects && Array.isArray(slideData.objects)) {
          slideData.objects.forEach((obj, objIndex) => {
            try {
              // Handle background rectangles
              if (obj.rect) {
                const rect = obj.rect;
                const x = typeof rect.x === 'string' && rect.x.includes('%') ? 0 : (parseFloat(rect.x) || 0);
                const y = typeof rect.y === 'string' && rect.y.includes('%') ? 0 : (parseFloat(rect.y) || 0);
                const w = typeof rect.w === 'string' && rect.w.includes('%') ? 10 : (parseFloat(rect.w) || 10);
                const h = typeof rect.h === 'string' && rect.h.includes('%') ? 7.5 : (parseFloat(rect.h) || 7.5);
                
                // Set slide background if this is a full-width background
                if (x === 0 && y === 0 && (w >= 10 || rect.w === '100%')) {
                  slide.background = { color: rect.fill?.color || 'FFFFFF' };
                } else {
                  // Add as a shape
                  slide.addShape('RECTANGLE', {
                    x: x,
                    y: y,
                    w: w,
                    h: h,
                    fill: { color: rect.fill?.color || 'FFFFFF' },
                    line: { width: 0 }
                  });
                }
              }
              
              // Handle text elements
              if (obj.text) {
                const textObj = obj.text;
                const options = textObj.options || {};
                
                slide.addText(textObj.text || '', {
                  x: parseFloat(options.x) || 0.5,
                  y: parseFloat(options.y) || 0.5,
                  w: parseFloat(options.w) || 9,
                  h: parseFloat(options.h) || 1,
                  fontSize: parseInt(options.fontSize) || 16,
                  fontFace: options.fontFace || 'Arial',
                  color: options.color || '000000',
                  bold: options.bold || false,
                  italic: options.italic || false,
                  underline: options.underline || false,
                  align: options.align || 'left',
                  valign: options.valign || 'top',
                  rotate: parseInt(options.rotate) || 0
                });
              }
              
              // Handle shape elements
              if (obj.shape) {
                const shape = obj.shape;
                const options = shape.options || {};
                
                slide.addShape(shape.type?.toUpperCase() || 'RECTANGLE', {
                  x: parseFloat(options.x) || 1,
                  y: parseFloat(options.y) || 1,
                  w: parseFloat(options.w) || 2,
                  h: parseFloat(options.h) || 2,
                  fill: options.fill || { color: '0066CC' },
                  line: options.line || { color: '000000', width: 1 },
                  rotate: parseInt(options.rotate) || 0
                });
              }
              
              // Handle line elements
              if (obj.line) {
                const line = obj.line;
                const options = line.options || {};
                
                slide.addShape('LINE', {
                  x: parseFloat(options.x1) || parseFloat(options.x) || 1,
                  y: parseFloat(options.y1) || parseFloat(options.y) || 1,
                  w: Math.abs(parseFloat(options.x2) - parseFloat(options.x1)) || 2,
                  h: Math.abs(parseFloat(options.y2) - parseFloat(options.y1)) || 0.1,
                  line: options.line || { color: '000000', width: 2 }
                });
              }
              
            } catch (objError) {
              console.warn(`Error processing object ${objIndex} on slide ${slideIndex}:`, objError.message);
            }
          });
        }
        
      } catch (slideError) {
        console.error(`Error processing slide ${slideIndex}:`, slideError.message);
        // Add a simple error slide
        const errorSlide = pptx.addSlide();
        errorSlide.addText(`Error on Slide ${slideIndex + 1}`, {
          x: 1, y: 1, w: 8, h: 1,
          fontSize: 20, color: 'FF0000', bold: true
        });
        errorSlide.addText(`${slideError.message}`, {
          x: 1, y: 2, w: 8, h: 4,
          fontSize: 14, color: '666666'
        });
      }
    });
    
    console.log('Generating PPTX buffer...');
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    console.log('PPTX generated successfully, size:', buffer.length);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="converted-presentation.pptx"');
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating presentation from JSON:', error);
    res.status(500).json({
      error: 'Failed to create presentation',
      message: error.message,
      details: process.env.NODE_ENV === 'development' ? error.stack : undefined,
      timestamp: new Date().toISOString()
    });
  }
});

// Helper endpoint to validate JSON structure
app.post('/api/presentation/validate-json', (req, res) => {
  try {
    const { slides } = req.body;
    
    if (!slides || !Array.isArray(slides)) {
      return res.status(400).json({
        valid: false,
        error: 'Missing or invalid slides array'
      });
    }
    
    const validation = {
      valid: true,
      slideCount: slides.length,
      slides: []
    };
    
    slides.forEach((slide, index) => {
      const slideValidation = {
        slideNumber: slide.slideNumber || index + 1,
        objectCount: slide.objects?.length || 0,
        objects: []
      };
      
      if (slide.objects && Array.isArray(slide.objects)) {
        slide.objects.forEach((obj, objIndex) => {
          const objValidation = {
            index: objIndex,
            type: obj.rect ? 'rectangle' : obj.text ? 'text' : obj.shape ? 'shape' : obj.line ? 'line' : 'unknown',
            valid: true,
            issues: []
          };
          
          // Validate text objects
          if (obj.text) {
            if (!obj.text.text) {
              objValidation.issues.push('Missing text content');
            }
            if (!obj.text.options) {
              objValidation.issues.push('Missing text options');
            }
          }
          
          // Validate rect objects
          if (obj.rect) {
            if (obj.rect.x === undefined || obj.rect.y === undefined) {
              objValidation.issues.push('Missing position (x, y)');
            }
            if (obj.rect.w === undefined || obj.rect.h === undefined) {
              objValidation.issues.push('Missing dimensions (w, h)');
            }
          }
          
          objValidation.valid = objValidation.issues.length === 0;
          slideValidation.objects.push(objValidation);
        });
      }
      
      validation.slides.push(slideValidation);
    });
    
    res.json(validation);
    
  } catch (error) {
    res.status(500).json({
      valid: false,
      error: error.message
    });
  }
});
// Simple presentation endpoint - Railway optimized
app.post('/api/presentation/simple', async (req, res) => {
  try {
    const { 
      title = 'My Presentation',
      slides = [],
      author = 'API User'
    } = req.body;

    console.log('Creating simple presentation:', { title, slideCount: slides.length });
    
    const pptx = new PptxGenJS();
    
    // Basic settings
    pptx.author = author;
    pptx.title = title;
    pptx.layout = 'LAYOUT_16x9';
    
    // Title slide
    const titleSlide = pptx.addSlide();
    titleSlide.addText(title, {
      x: 1,
      y: 2.5,
      w: 8,
      h: 1.5,
      fontSize: 28,
      bold: true,
      color: '363636',
      align: 'center'
    });
    
    // Content slides
    slides.forEach((slideData, index) => {
      try {
        const slide = pptx.addSlide();
        
        if (slideData.title) {
          slide.addText(slideData.title, {
            x: 0.5,
            y: 0.5,
            w: 9,
            h: 1,
            fontSize: 20,
            bold: true,
            color: '1f4e79'
          });
        }
        
        if (slideData.content) {
          slide.addText(slideData.content, {
            x: 0.5,
            y: slideData.title ? 1.5 : 0.5,
            w: 9,
            h: slideData.title ? 4.5 : 5.5,
            fontSize: 14,
            color: '444444'
          });
        }
      } catch (slideError) {
        console.warn(`Error processing slide ${index}:`, slideError);
      }
    });
    
    console.log('Generating PPTX buffer...');
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    console.log('PPTX generated successfully, size:', buffer.length);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(title)}.pptx"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating simple presentation:', error);
    res.status(500).json({ 
      error: 'Failed to create presentation', 
      message: error.message,
      timestamp: new Date().toISOString()
    });
  }
});
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

// Add simple text slide
app.post('/api/slide/text', async (req, res) => {
  try {
    const { 
      title = 'Text Slide',
      text = 'Sample text content',
      options = {}
    } = req.body;
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    // Add title
    slide.addText(title, {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontSize: 20,
      bold: true,
      color: '1f4e79'
    });
    
    // Add content
    slide.addText(text, {
      x: 0.5,
      y: 1.5,
      w: 9,
      h: 4,
      fontSize: options.fontSize || 14,
      color: options.color || '444444',
      align: options.align || 'left',
      bold: options.bold || false,
      italic: options.italic || false
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="text-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating text slide:', error);
    res.status(500).json({ error: 'Failed to create text slide', message: error.message });
  }
});

// Add simple image slide
app.post('/api/slide/image', upload.single('image'), async (req, res) => {
  try {
    const { title = 'Image Slide', options = {} } = req.body;
    let imageData;
    
    if (req.file) {
      imageData = await processImage(req.file.buffer, options);
    } else if (req.body.imageUrl) {
      const response = await fetch(req.body.imageUrl);
      const buffer = await response.buffer();
      imageData = await processImage(buffer, options);
    } else {
      return res.status(400).json({ error: 'No image provided. Send file or imageUrl.' });
    }
    
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    
    // Add title
    slide.addText(title, {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 0.8,
      fontSize: 20,
      bold: true,
      color: '1f4e79'
    });
    
    // Add image
    slide.addImage({
      data: imageData,
      x: 1.5,
      y: 1.5,
      w: 7,
      h: 4,
      sizing: { type: 'contain', w: 7, h: 4 }
    });
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="image-slide.pptx"');
    res.send(buffer);
    
  } catch (error) {
    console.error('Error creating image slide:', error);
    res.status(500).json({ error: 'Failed to create image slide', message: error.message });
  }
});
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

// Global error handler
app.use((error, req, res, next) => {
  console.error('Unhandled error:', {
    message: error.message,
    stack: error.stack,
    url: req.url,
    method: req.method,
    timestamp: new Date().toISOString()
  });
  
  // Handle specific error types
  if (error.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({ error: 'File too large', maxSize: '5MB' });
  }
  
  if (error.code === 'LIMIT_UNEXPECTED_FILE') {
    return res.status(400).json({ error: 'Unexpected file field' });
  }
  
  if (error.message.includes('Invalid file type')) {
    return res.status(400).json({ error: error.message });
  }
  
  res.status(500).json({ 
    error: 'Internal server error',
    message: process.env.NODE_ENV === 'development' ? error.message : 'Something went wrong',
    timestamp: new Date().toISOString()
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ 
    error: 'Endpoint not found',
    path: req.path,
    method: req.method,
    availableEndpoints: ['/', '/health', '/api/presentation/simple', '/api/slide/text', '/api/slide/image']
  });
});

// Start server
const server = app.listen(PORT, '0.0.0.0', () => {
  console.log(`🚀 PptxGenJS API Server running on port ${PORT}`);
  console.log(`📊 Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`🔗 Health check: http://localhost:${PORT}/health`);
  console.log(`📋 API docs: http://localhost:${PORT}/`);
});

// Handle server errors
server.on('error', (error) => {
  if (error.code === 'EADDRINUSE') {
    console.error(`❌ Port ${PORT} is already in use`);
  } else {
    console.error('❌ Server error:', error);
  }
  process.exit(1);
});

module.exports = app;
