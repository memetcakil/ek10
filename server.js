const express = require('express');
const multer = require('multer');
const transformExcel = require('./excel_transform');

// Multer belleğe kaydetme konfigürasyonu
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 5 * 1024 * 1024 // 5MB limit
    },
    fileFilter: function (req, file, cb) {
        if (!file.originalname.match(/\.(xlsx|xls)$/)) {
            return cb(new Error('Sadece Excel dosyaları yüklenebilir!'));
        }
        cb(null, true);
    }
}).single('excelFile');

const app = express();

app.set('view engine', 'ejs');
app.use(express.static('public'));

// Hata yakalama middleware'i
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({
        success: false,
        message: 'Bir hata oluştu: ' + err.message
    });
});

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/api/upload', (req, res) => {
    upload(req, res, async function(err) {
        if (err instanceof multer.MulterError) {
            return res.status(400).json({
                success: false,
                message: 'Dosya yükleme hatası: ' + err.message
            });
        } else if (err) {
            return res.status(400).json({
                success: false,
                message: err.message
            });
        }

        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: 'Lütfen bir dosya seçin.'
            });
        }

        try {
            const excelBuffer = await transformExcel(req.file.buffer);
            
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename=output.xlsx');
            return res.send(excelBuffer);

        } catch (error) {
            console.error('Dönüştürme hatası:', error);
            return res.status(500).json({
                success: false,
                message: 'Dosya dönüştürme sırasında bir hata oluştu: ' + error.message
            });
        }
    });
});

// Vercel için export
module.exports = app;