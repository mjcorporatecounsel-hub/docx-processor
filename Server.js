const express = require('express');
const multer = require('multer');
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.post('/process-docx', upload.single('template'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No template file provided' });
    }

    if (!req.body.data) {
      return res.status(400).json({ error: 'No data provided' });
    }

    const templateBuffer = req.file.buffer;
    const formData = JSON.parse(req.body.data);

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.render(formData);

    const outputBuffer = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="processed.docx"');
    res.send(outputBuffer);

  } catch (error) {
    console.error('Error processing DOCX:', error);
    res.status(500).json({ error: 'Failed to process document', details: error.message });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`DOCX processing service running on port ${PORT}`);
});