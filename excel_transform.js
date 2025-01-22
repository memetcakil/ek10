const Excel = require('exceljs');
const XLSX = require('xlsx');

async function transformExcel(fileBuffer) {
    try {
        // Buffer kontrolü
        if (!fileBuffer || fileBuffer.length === 0) {
            throw new Error('Geçersiz dosya: Boş dosya');
        }

        // Önce XLSX ile dosyayı oku
        console.log('XLSX ile dosya okunuyor...');
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const sourceData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        console.log('XLSX ile veri okundu:', sourceData.length, 'satır');

        // Yeni workbook oluştur
        const newWorkbook = new Excel.Workbook();
        const newWorksheet = newWorkbook.addWorksheet('Transformed Data');

        // Sütun genişliklerini ayarla
        newWorksheet.columns = [
            { header: '', width: 5 },  // A sütunu (sıra no için)
            { header: '', width: 30 },   // B sütunu (isimler için)
            ...Array(21).fill({ width: 4 }), // C'den W'ye kadar modül sütunları
            { width: 8 }, // X sütunu (Başarı Puanı)
            { width: 8 }  // Y sütunu (Başarı Durumu)
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

        // A5 ve B5'i birleştir ve KURSUN ADI yaz
        newWorksheet.mergeCells('A5:B5');
        const kursAdiCell = newWorksheet.getCell('A5');
        kursAdiCell.value = 'KURSUN ADI';
        kursAdiCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        kursAdiCell.font = {
            bold: true,
            size: 11
        };

        // A6 ve B6'yı birleştir ve KURSU NO yaz
        newWorksheet.mergeCells('A6:B6');
        const kursNoCell = newWorksheet.getCell('A6');
        kursNoCell.value = 'KURSU NO';
        kursNoCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        kursNoCell.font = {
            bold: true,
            size: 11
        };

        // A7 ve B7'yi birleştir ve DÜZENLENDİĞİ YER yaz
        newWorksheet.mergeCells('A7:B7');
        const yerCell = newWorksheet.getCell('A7');
        yerCell.value = 'DÜZENLENDİĞİ YER';
        yerCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        yerCell.font = {
            bold: true,
            size: 11
        };

        // B8'den W8'e kadar birleştir ve KURS MODÜL DEĞERLENDİRME NOTU yaz
        newWorksheet.mergeCells('B8:W8');
        const modulDegerlendirmeCell = newWorksheet.getCell('B8');
        modulDegerlendirmeCell.value = 'KURS MODÜL DEĞERLENDİRME NOTU';
        modulDegerlendirmeCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        modulDegerlendirmeCell.font = {
            bold: true,
            size: 14,
            underline: true
        };
        modulDegerlendirmeCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEEEEEE' }
        };

        // X8'den Y9'a kadar birleştir ve KURSUN BAŞARI PUANI VE DURUMU yaz
        newWorksheet.mergeCells('X8:Y9');
        const basariPuaniCell = newWorksheet.getCell('X8');
        basariPuaniCell.value = 'KURSUN BAŞARI PUANI VE DURUMU';
        basariPuaniCell.alignment = {
            vertical: 'middle',
            horizontal: 'center',
            wrapText: true
        };
        basariPuaniCell.font = {
            bold: true,
            size: 11
        };

        // C9'dan W9'a kadar sayıları yaz
        for (let i = 0; i < 21; i++) {
            const col = String.fromCharCode(67 + i); // 67 = 'C'
            const cell = newWorksheet.getCell(`${col}9`);
            cell.value = i + 1;
            cell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            cell.font = {
                bold: true,
                size: 11
            };
        }

        // A10 ve B10'daki yazıları alt çizgiye hizala
        const noCell = newWorksheet.getCell('A10');
        noCell.value = 'No';
        noCell.alignment = {
            vertical: 'bottom',
            horizontal: 'center'
        };
        noCell.font = {
            bold: true,
            size: 11
        };

        const adSoyadCell = newWorksheet.getCell('B10');
        adSoyadCell.value = 'AD SOYAD';
        adSoyadCell.alignment = {
            vertical: 'bottom',
            horizontal: 'center'
        };
        adSoyadCell.font = {
            bold: true,
            size: 11
        };

        // X10'a Başarı Puanı yaz (dikey)
        const basariPuaniBaslikCell = newWorksheet.getCell('X10');
        basariPuaniBaslikCell.value = 'Başarı Puanı';
        basariPuaniBaslikCell.alignment = {
            textRotation: 90,
            vertical: 'middle',
            horizontal: 'center'
        };
        basariPuaniBaslikCell.font = {
            bold: true,
            size: 11
        };

        // Y10'a Başarı Durumu yaz (dikey)
        const basariDurumuBaslikCell = newWorksheet.getCell('Y10');
        basariDurumuBaslikCell.value = 'Başarı Durumu';
        basariDurumuBaslikCell.alignment = {
            textRotation: 90,
            vertical: 'middle',
            horizontal: 'center'
        };
        basariDurumuBaslikCell.font = {
            bold: true,
            size: 11
        };

        // X11'den X50'ye kadar formülleri ekle
        for (let row = 11; row <= 50; row++) {
            const formulaCell = newWorksheet.getCell(`X${row}`);
            formulaCell.value = {
                formula: `=IF(COUNTIF(C${row}:W${row},"<>")=0,"Not Gir",ROUND(AVERAGE(C${row}:W${row}),1))`,
                date1904: false
            };
            formulaCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            formulaCell.numFmt = '0.0';

            // Y sütununa başarı durumu formülünü ekle
            const basariDurumuCell = newWorksheet.getCell(`Y${row}`);
            basariDurumuCell.value = {
                formula: `=IF(COUNTIF(C${row}:W${row},"<>")=0,"Not Gir",IF(COUNTIF(C${row}:W${row},"<50")<>0,"Başarısız","Başarılı"))`,
                date1904: false
            };
            basariDurumuCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
        }

        // Verileri işle
        try {
            let currentStudent = null;
            let studentGrades = new Map();
            let rowIndex = 11;

            // Her bir satırı işle (ilk satırı atla)
            for (let i = 1; i < sourceData.length; i++) {
                const row = sourceData[i];
                console.log(`İşlenen satır ${i}:`, row);

                if (!row || row.length < 2) {
                    console.log('Geçersiz satır atlandı');
                    continue;
                }

                // Yeni öğrenci başlangıcı
                if (row[0]) { // İlk sütunda isim varsa
                    console.log('Yeni öğrenci bulundu:', row[0]);
                    
                    // Önceki öğrencinin verilerini yaz
                    if (currentStudent) {
                        console.log('Önceki öğrenci yazılıyor:', currentStudent);
                        const wsRow = newWorksheet.getRow(rowIndex);
                        wsRow.getCell(1).value = rowIndex - 10;
                        wsRow.getCell(2).value = currentStudent;

                        // Notları yaz
                        studentGrades.forEach((grade, moduleIndex) => {
                            const col = String.fromCharCode(67 + moduleIndex);
                            console.log(`Not yazılıyor: ${col}${rowIndex} = ${grade}`);
                            wsRow.getCell(col).value = grade;
                        });

                        rowIndex++;
                    }

                    // Yeni öğrenci için hazırlık
                    currentStudent = row[0];
                    studentGrades = new Map();
                }

                // Not bilgisini ekle
                if (row[1] && row[3]) { // Modül adı ve not
                    const moduleIndex = Array.from(studentGrades.keys()).length;
                    if (moduleIndex < 21) {
                        console.log(`Not ekleniyor: ${row[1]} = ${row[3]}`);
                        studentGrades.set(moduleIndex, row[3]);
                    }
                }
            }

            // Son öğrencinin verilerini yaz
            if (currentStudent) {
                console.log('Son öğrenci yazılıyor:', currentStudent);
                const wsRow = newWorksheet.getRow(rowIndex);
                wsRow.getCell(1).value = rowIndex - 10;
                wsRow.getCell(2).value = currentStudent;

                studentGrades.forEach((grade, moduleIndex) => {
                    const col = String.fromCharCode(67 + moduleIndex);
                    console.log(`Not yazılıyor: ${col}${rowIndex} = ${grade}`);
                    wsRow.getCell(col).value = grade;
                });
            }

            console.log('Veri işleme tamamlandı');

        } catch (error) {
            console.error('Veri işleme hatası:', error);
            throw new Error('Veri işleme hatası: ' + error.message);
        }

        // Buffer olarak döndür
        return await newWorkbook.xlsx.writeBuffer();

    } catch (error) {
        console.error('Detaylı hata:', error);
        throw new Error(error.message);
    }
}

module.exports = transformExcel; 