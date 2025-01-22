const Excel = require('exceljs');

async function transformExcel(fileBuffer) {
    try {
        // Buffer kontrolü
        if (!fileBuffer || fileBuffer.length === 0) {
            throw new Error('Geçersiz dosya: Boş dosya');
        }

        // Excel dosyası kontrolü
        try {
            // Gelen Excel dosyasını oku
            const sourceWorkbook = new Excel.Workbook();
            await sourceWorkbook.xlsx.load(fileBuffer);
            const sourceSheet = sourceWorkbook.getWorksheet(1);

            if (!sourceSheet) {
                throw new Error('Excel dosyası boş veya geçersiz');
            }
        } catch (error) {
            throw new Error('Geçersiz Excel dosyası: ' + error.message);
        }

        // Yeni workbook oluştur
        const newWorkbook = new Excel.Workbook();
        const newWorksheet = newWorkbook.addWorksheet('Transformed Data');

        // Sütun genişliklerini ayarla
        newWorksheet.columns = [
            { header: '', width: 5 },  // A sütunu (sıra no için)
            { header: '', width: 30 },   // B sütunu (isimler için)
            ...Array(21).fill({ width: 4 }), // C'den W'ye kadar modül sütunları
            { width: 4 }, // X sütunu (Başarı Puanı)
            { width: 4 }  // Y sütunu (Başarı Durumu)
        ];

        // A1'den Y1'e kadar hücreleri birleştir ve T.C. yaz
        newWorksheet.mergeCells('A1:Y1');
        const tcCell = newWorksheet.getCell('A1');
        tcCell.value = 'T.C.';
        tcCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        tcCell.font = {
            bold: true,
            size: 12
        };

        // A2'den Y2'ye kadar birleştir ve MİLLİ EĞİTİM BAKANLIĞI yaz
        newWorksheet.mergeCells('A2:Y2');
        const mebCell = newWorksheet.getCell('A2');
        mebCell.value = 'MİLLİ EĞİTİM BAKANLIĞI';
        mebCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        mebCell.font = {
            bold: true,
            size: 12
        };

        // A3'ten Y3'e kadar birleştir ve KAHTA HALK EĞİTİMİ MERKEZİ yaz
        newWorksheet.mergeCells('A3:Y3');
        const hemCell = newWorksheet.getCell('A3');
        hemCell.value = 'KAHTA HALK EĞİTİMİ MERKEZİ';
        hemCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        hemCell.font = {
            bold: true,
            size: 12
        };

        // A4'ten X4'e kadar birleştir ve MODÜL DEĞERLENDİRME ÇİZELGESİ yaz
        newWorksheet.mergeCells('A4:X4');
        const modulCell = newWorksheet.getCell('A4');
        modulCell.value = 'MODÜL DEĞERLENDİRME ÇİZELGESİ';
        modulCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        modulCell.font = {
            bold: true,
            size: 12
        };

        // Y4'e EK 10 yaz
        const ekCell = newWorksheet.getCell('Y4');
        ekCell.value = 'EK 10';
        ekCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        ekCell.font = {
            bold: true,
            size: 11
        };

        // Diğer başlıkları ve alanları ekle...
        // (Devam edecek...)

        // Buffer olarak döndür
        try {
            return await newWorkbook.xlsx.writeBuffer();
        } catch (error) {
            throw new Error('Excel dosyası oluşturulurken hata: ' + error.message);
        }

    } catch (error) {
        console.error('Detaylı hata:', error);
        throw new Error(error.message);
    }
}

module.exports = transformExcel; 