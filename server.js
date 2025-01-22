const express = require('express');
const multer = require('multer');
const path = require('path');
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

// View engine ayarları
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// Static dosyalar için
app.use(express.static(path.join(__dirname, 'public')));

// Ana sayfa
app.get('/', (req, res) => {
    try {
        res.render('index');
    } catch (error) {
        console.error('Render hatası:', error);
        res.status(500).send('Bir hata oluştu: ' + error.message);
    }
});

// Upload endpoint'i
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

        // Dosya tipi kontrolü
        if (!req.file.mimetype.includes('spreadsheet') && 
            !req.file.mimetype.includes('excel')) {
            return res.status(400).json({
                success: false,
                message: 'Lütfen geçerli bir Excel dosyası seçin.'
            });
        }

        try {
            const excelBuffer = await transformExcel(req.file.buffer);
            
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename=output.xlsx');
            return res.send(excelBuffer);

        } catch (error) {
            console.error('Dönüştürme hatası:', error);
            return res.status(400).json({
                success: false,
                message: error.message
            });
        }
    });
});

// Hata yakalama middleware'i
app.use((err, req, res, next) => {
    console.error('Uygulama hatası:', err.stack);
    res.status(500).json({
        success: false,
        message: 'Bir hata oluştu: ' + err.message
    });
});

// 404 handler
app.use((req, res) => {
    res.status(404).send('Sayfa bulunamadı');
});

// Express uygulamasını export et
module.exports = app;

// Eğer doğrudan çalıştırılıyorsa (development)
if (require.main === module) {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log(`Server ${PORT} portunda çalışıyor`);
    });
}