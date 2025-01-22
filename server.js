const express = require('express');
const multer = require('multer');
const transformExcel = require('./excel_transform');

// Multer belleğe kaydetme konfigürasyonu
const upload = multer({
    storage: multer.memoryStorage(),
    fileFilter: function (req, file, cb) {
        if (!file.originalname.match(/\.(xlsx|xls)$/)) {
            return cb(new Error('Sadece Excel dosyaları yüklenebilir!'));
        }
        cb(null, true);
    }
});

const app = express();

app.set('view engine', 'ejs');
app.use(express.static('public'));

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/api/upload', upload.single('excelFile'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({
            success: false,
            message: 'Lütfen bir dosya seçin.'
        });
    }

    try {
        const excelBuffer = await transformExcel(req.file.buffer);
        
        // Buffer'ı doğrudan response olarak gönder
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=output.xlsx');
        res.send(excelBuffer);

    } catch (error) {
        console.error('Hata:', error);
        res.status(500).json({
            success: false,
            message: 'Dosya dönüştürme sırasında bir hata oluştu: ' + error.message
        });
    }
});

// Vercel için export
module.exports = app;