const express = require('express');
const multer = require('multer');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// Servir arquivos estáticos (fotos processadas)
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = './uploads';
    if (!fs.existsSync(dir)) fs.mkdirSync(dir);
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname));
  }
});

const upload = multer({ storage });

app.post('/api/process-receipt', upload.single('receipt'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo enviado.' });

  const filePath = path.resolve(req.file.path);
  
  // Chamar o script Python
  exec(`python3 process_receipt.py "${filePath}"`, { cwd: __dirname }, (error, stdout, stderr) => {
    if (error) {
      console.error(`Erro: ${error.message}`);
      return res.status(500).json({ error: 'Erro ao processar imagem.' });
    }
    
    try {
      const result = JSON.parse(stdout);
      
      // Mover a imagem processada para a pasta uploads
      if (result.processed_image) {
        const oldPath = path.join(__dirname, result.processed_image);
        const newFileName = 'processed_' + req.file.filename;
        const newPath = path.join(__dirname, 'uploads', newFileName);
        
        fs.renameSync(oldPath, newPath);
        result.processed_image = `/uploads/${newFileName}`;
      }
      
      res.json(result);
    } catch (e) {
      res.status(500).json({ error: 'Erro ao ler resultado do processamento.' });
    }
  });
});

const PORT = 3001;
app.listen(PORT, () => {
  console.log(`Backend rodando na porta ${PORT}`);
});
