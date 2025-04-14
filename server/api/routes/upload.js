const express = require('express');
const router = express.Router();
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const mammoth = require('mammoth');
const { generateDocxFile } = require('../utils/generateDocxFile');

// Load Gemini API key
const API_KEY = process.env.GEMINI_API_KEY;
if (!API_KEY) {
  console.error('[ERROR] Gemini API key not found in environment variables!');
} else {
  console.log('[INFO] Gemini API key loaded.');
}
const genAI = new GoogleGenerativeAI(API_KEY);

// Multer storage setup
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const uploadDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadDir)) {
      console.log('[INFO] Upload directory not found. Creating...');
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    const filename = uniqueSuffix + '-' + file.originalname;
    console.log(`[INFO] Saving file as: ${filename}`);
    cb(null, filename);
  }
});

const fileFilter = (req, file, cb) => {
  const allowedTypes = [
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/pdf'
  ];
  if (allowedTypes.includes(file.mimetype)) {
    console.log(`[INFO] File accepted: ${file.originalname} (${file.mimetype})`);
    cb(null, true);
  } else {
    console.warn(`[WARN] Rejected file type: ${file.mimetype}`);
    cb(new Error('Only DOCX and PDF files are allowed'), false);
  }
};

const upload = multer({
  storage: storage,
  fileFilter: fileFilter,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB
});

// Convert file to plain text for Gemini
async function fileToGenerativeText(filePath, mimeType) {
  console.log(`[DEBUG] Reading file for Gemini: ${filePath} (${mimeType})`);

  if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    const result = await mammoth.extractRawText({ path: filePath });
    return result.value;
  }

  if (mimeType === 'application/pdf') {
    throw new Error("PDF support not implemented yet");
  }

  throw new Error("Unsupported file type for Gemini");
}

// Generate structured data using Gemini
async function extractDataWithGemini(filePath, mimeType) {
  try {
    console.log('[STEP 3] Initializing Gemini model...');
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const fileText = await fileToGenerativeText(filePath, mimeType);

    const prompt = `
    The following is a document containing student scores:
    ----
    ${fileText}
    ----
    Extract all data from tables containing student records. 
    For each row, identify the Serial Number (sn), Registration Number (regNo), 
    Continuous Assessment (ca), Exam score (exam), and Total score (total). 

    If the Serial Number is missing or empty, automatically assign it based on the row index (starting from 1). 
    Return the data as a JSON array of objects, where each object represents one student record with these fields: sn, regNo, ca, exam, and total.

    Do not include blank or empty rows. Only return the JSON data, nothing else.
    `;


    console.log('[STEP 3] Sending prompt to Gemini...');
    const result = await model.generateContent(prompt);

    const response = await result.response;
    const text = response.text();

    console.log('[DEBUG] Gemini raw response:');
    console.log(text);

    const jsonMatch = text.match(/\[.*\]/s);
    if (!jsonMatch) {
      throw new Error("Could not extract JSON data from Gemini response");
    }

    const extractedData = JSON.parse(jsonMatch[0]);
    console.log('[SUCCESS] Extracted data from Gemini.');
    return extractedData;

  } catch (error) {
    console.error('[ERROR] Failed to extract data with Gemini:', error.message);
    throw error;
  }
}

// Upload route
router.post('/upload', upload.single('document'), async (req, res) => {
  try {
    console.log('[STEP 4] /upload hit. Processing file...');

    if (!req.file) {
      console.warn('[WARN] No file received in request.');
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    const mimeType = req.file.mimetype;

    console.log(`[INFO] File uploaded: ${filePath}`);

    // Step 1: Extract data using Gemini
    const extractedData = await extractDataWithGemini(filePath, mimeType);

    // Step 2: Generate DOCX file with extracted data
    const outputPath = path.join(__dirname, 'outputs', `results-${Date.now()}.docx`);
    await generateDocxFile(extractedData, outputPath);

    // Step 3: Send the processed .docx file to the client
    res.download(outputPath, 'processed-results.docx', (err) => {
      if (err) {
        console.error('[ERROR] Sending DOCX failed:', err.message);
        return res.status(500).send('Failed to download file.');
      }

      // Optional cleanup after successful download
      fs.unlinkSync(filePath);        // Delete uploaded file
      fs.unlinkSync(outputPath);      // Delete generated DOCX
    });

  } catch (error) {
    console.error('[ERROR] Upload processing failed:', error);
    res.status(500).json({
      success: false,
      message: 'Error processing file with Gemini',
      error: error.message
    });
  }
});

// Status check route
router.get('/status', (req, res) => {
  console.log('[INFO] /status route hit.');
  res.json({
    status: 'ok',
    message: 'Document upload API with Gemini integration is running'
  });
});

module.exports = router;
