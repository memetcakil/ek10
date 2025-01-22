const XLSX = require('xlsx');
const Excel = require('exceljs');

async function transformExcel(fileBuffer) {
    try {
        // XLSX ile Excel dosyasını oku (buffer'dan)
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Veriyi işle
        const data = XLSX.utils.sheet_to_json(sheet, { header: ['name', 'module', 'status', 'grade'] });

        // Benzersiz isimleri ve modülleri al
        const uniqueNames = [...new Set(data.map(row => row.name))];
        const uniqueModules = [...new Set(data.map(row => row.module))];

        // ExcelJS ile yeni workbook oluştur
        const newWorkbook = new Excel.Workbook();
        const newWorksheet = newWorkbook.addWorksheet('Transformed Data');

        // Sütun genişliklerini ayarla
        newWorksheet.columns = [
            { header: '', width: 5 },  // A sütunu (sıra no için)
            { header: '', width: 30 },   // B sütunu (isimler için)
            ...uniqueModules.map(() => ({ width: 4 })), // Modül sütunları
            { width: 4 }, // Kalan sütunlar için de 4 birim genişlik
            { width: 4 },
            { width: 4 },
            { width: 4 }
        ];

        // Header row - ilk satıra boş hücre ve modül isimlerini ekle
        const headerRow = [''];
        uniqueModules.forEach(module => {
            headerRow.push(module);
        });
        transformedData.push(headerRow);

        // Her isim için bir satır oluştur
        uniqueNames.forEach(name => {
            const row = [name];
            // Her modül için notu bul
            uniqueModules.forEach(module => {
                const matchingRow = data.find(d => d.name === name && d.module === module);
                row.push(matchingRow ? matchingRow.grade : '');
            });
            transformedData.push(row);
        });

        // Yeni bir worksheet oluştur
        const ws = XLSX.utils.aoa_to_sheet(transformedData);

        // Yeni bir workbook oluştur ve worksheet'i ekle
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Transformed Data');

        // Buffer olarak döndür
        return await newWorkbook.xlsx.writeBuffer();

    } catch (error) {
        console.error('Detaylı hata:', error);
        throw new Error('Excel dönüştürme hatası: ' + error.message);
    }
}

module.exports = transformExcel; 