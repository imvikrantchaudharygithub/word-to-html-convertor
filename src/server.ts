import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { convertDocxToHtml, isValidWordFile } from './converter';

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// Configure multer for file uploads
// Store files temporarily in memory with size limit of 10MB
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
    files: 1 // Only allow 1 file at a time
  },
  fileFilter: (req, file, cb) => {
    // Validate file type before processing
    if (isValidWordFile(file.originalname, file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Invalid file type. Only .docx and .doc files are allowed.'));
    }
  }
});

// Serve the main HTML page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../public/index.html'));
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', message: 'Word to HTML Converter API is running' });
});

// Debug conversion endpoint - shows detailed conversion steps
app.post('/api/debug-convert', upload.single('wordFile'), async (req, res): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ 
        error: 'No file uploaded. Please select a Word document (.docx or .doc).' 
      });
      return;
    }

    const { originalname, buffer, mimetype, size } = req.file;
    
    console.log(`DEBUG: Processing file: ${originalname} (${size} bytes, ${mimetype})`);

    if (!isValidWordFile(originalname, mimetype)) {
      res.status(400).json({ 
        error: 'Invalid file type. Only .docx and .doc files are allowed.' 
      });
      return;
    }

    // Import mammoth for debugging
    const mammoth = require('mammoth');
    
    // Step 1: Extract raw text
    const rawResult = await mammoth.extractRawText({ buffer });
    
    // Step 2: Convert with basic conversion (no custom styles)
    const basicResult = await mammoth.convertToHtml({ buffer });
    
    // Step 3: Convert with our enhanced conversion
    const { convertDocxToHtml } = await import('./converter');
    const enhancedHtml = await convertDocxToHtml(buffer);
    
    // Return debugging information
    res.json({
      filename: originalname,
      size,
      debugging: {
        rawText: rawResult.value.substring(0, 1000) + (rawResult.value.length > 1000 ? '...' : ''),
        rawTextLength: rawResult.value.length,
        basicHtml: basicResult.value,
        basicMessages: basicResult.messages,
        enhancedHtml: enhancedHtml,
        enhancedLength: enhancedHtml.length
      }
    });

  } catch (error) {
    console.error('Debug conversion error:', error);
    res.status(500).json({ 
      error: 'Debug conversion failed: ' + (error instanceof Error ? error.message : 'Unknown error')
    });
  }
});

// Main conversion endpoint
app.post('/api/convert', upload.single('wordFile'), async (req, res): Promise<void> => {
  try {
    // Check if file was uploaded
    if (!req.file) {
      res.status(400).json({ 
        error: 'No file uploaded. Please select a Word document (.docx or .doc).' 
      });
      return;
    }

    const { originalname, buffer, mimetype, size } = req.file;
    
    // Log upload details for debugging
    console.log(`Processing file: ${originalname} (${size} bytes, ${mimetype})`);

    // Double-check file validation (multer filter should have caught this)
    if (!isValidWordFile(originalname, mimetype)) {
      res.status(400).json({ 
        error: 'Invalid file type. Only .docx and .doc files are allowed.' 
      });
      return;
    }

    // Check file size (multer should have caught this too, but double-check)
    if (size > 10 * 1024 * 1024) {
      res.status(400).json({ 
        error: 'File too large. Maximum size is 10MB.' 
      });
      return;
    }

    // Convert the Word document to HTML
    console.log('Starting conversion...');
    const html = await convertDocxToHtml(buffer);
    console.log(`Conversion completed. HTML length: ${html.length} characters`);

    // Return the clean HTML
    res.json({ 
      html,
      filename: originalname,
      size: html.length
    });

  } catch (error) {
    console.error('Conversion error:', error);
    
    // Handle specific multer errors
    if (error instanceof multer.MulterError) {
      if (error.code === 'LIMIT_FILE_SIZE') {
        res.status(400).json({ 
          error: 'File too large. Maximum size is 10MB.' 
        });
        return;
      }
      if (error.code === 'LIMIT_FILE_COUNT') {
        res.status(400).json({ 
          error: 'Too many files. Please upload only one file at a time.' 
        });
        return;
      }
    }

    // Handle file type validation errors
    if (error instanceof Error && error.message && error.message.includes('Invalid file type')) {
      res.status(400).json({ 
        error: error.message 
      });
      return;
    }

    // Handle conversion errors
    if (error instanceof Error && error.message && error.message.includes('Failed to convert')) {
      res.status(500).json({ 
        error: 'Failed to convert the Word document. The file may be corrupted or in an unsupported format.' 
      });
      return;
    }

    // Generic error response
    res.status(500).json({ 
      error: 'An unexpected error occurred during file processing. Please try again.' 
    });
  }
});

// Error handling middleware
app.use((error: Error, req: express.Request, res: express.Response, next: express.NextFunction): void => {
  console.error('Unhandled error:', error);
  
  // Handle multer errors
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      res.status(400).json({ 
        error: 'File too large. Maximum size is 10MB.' 
      });
      return;
    }
    res.status(400).json({ 
      error: 'File upload error: ' + error.message 
    });
    return;
  }

  res.status(500).json({ 
    error: 'Internal server error' 
  });
});

// 404 handler
app.use('*', (req, res) => {
  res.status(404).json({ 
    error: 'Endpoint not found' 
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`🚀 Word to HTML Converter server is running on http://localhost:${PORT}`);
  console.log(`📁 Serving static files from: ${path.join(__dirname, '../public')}`);
  console.log(`🔄 API endpoint: POST http://localhost:${PORT}/api/convert`);
});
