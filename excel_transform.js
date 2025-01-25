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

        // Modül adlarını topla (ilk okuma sırasında)
        const moduleNames = new Set();
        for (let i = 1; i < sourceData.length; i++) {
            const row = sourceData[i];
            if (row && row[1]) { // Modül adı varsa
                moduleNames.add(row[1]);
            }
        }

        // Yeni workbook oluştur
        const newWorkbook = new Excel.Workbook();
        const newWorksheet = newWorkbook.addWorksheet('EK 10 Sayfa 1');

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

        // C5'ten L5'e kadar birleştir
        newWorksheet.mergeCells('C5:L5');
        const kursAdiValueCell = newWorksheet.getCell('C5');
        kursAdiValueCell.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            indent: 1
        };
        kursAdiValueCell.font = {
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

        // C6'dan L6'ya kadar birleştir
        newWorksheet.mergeCells('C6:L6');
        const kursNoValueCell = newWorksheet.getCell('C6');
        kursNoValueCell.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            indent: 1
        };
        kursNoValueCell.font = {
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

        // C7'den L7'ye kadar birleştir
        newWorksheet.mergeCells('C7:L7');
        const yerValueCell = newWorksheet.getCell('C7');
        yerValueCell.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            indent: 1
        };
        yerValueCell.font = {
            size: 11
        };

        // M7'den Y7'ye kadar birleştir
        newWorksheet.mergeCells('M7:Y7');
        const m7Cell = newWorksheet.getCell('M7');
        m7Cell.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            indent: 1
        };
        m7Cell.font = {
            size: 11
        };

        // M5'ten T5'e kadar birleştir ve BAŞLAMA TARİHİ yaz
        newWorksheet.mergeCells('M5:T5');
        const baslamaTarihiCell = newWorksheet.getCell('M5');
        baslamaTarihiCell.value = 'BAŞLAMA TARİHİ';
        baslamaTarihiCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        baslamaTarihiCell.font = {
            bold: true,
            size: 11
        };

        // U5'ten Y5'e kadar birleştir
        newWorksheet.mergeCells('U5:Y5');
        const u5Cell = newWorksheet.getCell('U5');
        u5Cell.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            indent: 1
        };
        u5Cell.font = {
            size: 11
        };

        // M6'dan T6'ya kadar birleştir ve BİTİŞ TARİHİ yaz
        newWorksheet.mergeCells('M6:T6');
        const bitisTarihiCell = newWorksheet.getCell('M6');
        bitisTarihiCell.value = 'BİTİŞ TARİHİ';
        bitisTarihiCell.alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        bitisTarihiCell.font = {
            bold: true,
            size: 11
        };

        // U6'dan Y6'ya kadar birleştir
        newWorksheet.mergeCells('U6:Y6');
        const u6Cell = newWorksheet.getCell('U6');
        u6Cell.alignment = {
            vertical: 'middle',
            horizontal: 'left',
            indent: 1
        };
        u6Cell.font = {
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

        // Modül adlarını topla
        const moduleNames = new Set();
        for (let i = 1; i < sourceData.length; i++) {
            const row = sourceData[i];
            if (row && row[1]) { // Modül adı varsa
                moduleNames.add(row[1]);
            }
        }

        // C10'dan başlayarak modül adlarını dikey yaz
        let moduleIndex = 0;
        for (const moduleName of moduleNames) {
            if (moduleIndex >= 21) break; // Maksimum 21 modül

            const col = String.fromCharCode(67 + moduleIndex); // C'den başla
            const cell = newWorksheet.getCell(`${col}10`);
            cell.value = moduleName;
            cell.alignment = {
                textRotation: 90, // Dikey yazı
                vertical: 'middle',
                horizontal: 'center'
            };
            cell.font = {
                bold: true,
                size: 11
            };
            moduleIndex++;
        }

        // Başlık satırı yüksekliğini ayarla
        newWorksheet.getRow(10).height = 150;

        // A1'den Y50'ye kadar olan tüm hücrelere kenar çizgisi ekle
        for (let row = 1; row <= 50; row++) {
            for (let col = 'A'; col <= 'Y'; col = String.fromCharCode(col.charCodeAt(0) + 1)) {
                const cell = newWorksheet.getCell(`${col}${row}`);
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
        }

        // Verileri işle
        try {
            let currentStudent = null;
            let studentGrades = new Map();
            let rowIndex = 11;
            let pageNumber = 1;
            let currentWorksheet = newWorksheet;

            // Kaynak dosyadan verileri oku
            for (let i = 1; i < sourceData.length; i++) {
                const row = sourceData[i];
                if (!row || row.length < 2) continue;

                // Yeni öğrenci başlangıcı
                if (row[0]) {
                    // Eğer 40 öğrenci olduysa yeni sayfa oluştur
                    if (rowIndex > 50) {
                        pageNumber++;
                        rowIndex = 11; // Yeni sayfada tekrar 11'den başla
                        
                        // Yeni worksheet oluştur (sayfa ismini güncelle)
                        currentWorksheet = newWorkbook.addWorksheet(`EK 10 Sayfa ${pageNumber}`);
                        
                        // Yeni sayfanın formatlamasını yap
                        setupNewWorksheet(currentWorksheet, moduleNames);
                    }

                    // Önceki öğrencinin verilerini yaz
                    if (currentStudent) {
                        const wsRow = currentWorksheet.getRow(rowIndex);
                        wsRow.getCell(1).value = rowIndex - 10;
                        wsRow.getCell(2).value = currentStudent;

                        // Notları yaz
                        studentGrades.forEach((grade, moduleIndex) => {
                            const col = String.fromCharCode(67 + moduleIndex);
                            wsRow.getCell(col).value = grade === -1 ? "M" : grade;
                        });

                        rowIndex++;
                    }

                    // Yeni öğrenci için hazırlık
                    currentStudent = row[0];
                    studentGrades = new Map();
                }

                // Not bilgisini ekle
                if (row[1] && row[3]) {
                    const moduleIndex = Array.from(studentGrades.keys()).length;
                    if (moduleIndex < 21) {
                        const grade = row[3];
                        studentGrades.set(moduleIndex, grade);
                    }
                }
            }

            // Son öğrencinin verilerini yaz
            if (currentStudent) {
                const wsRow = currentWorksheet.getRow(rowIndex);
                wsRow.getCell(1).value = rowIndex - 10;
                wsRow.getCell(2).value = currentStudent;

                studentGrades.forEach((grade, moduleIndex) => {
                    const col = String.fromCharCode(67 + moduleIndex);
                    wsRow.getCell(col).value = grade === -1 ? "M" : grade;
                });
            }

        } catch (error) {
            console.error('Veri işleme hatası:', error);
            throw new Error('Veri işleme hatası: ' + error.message);
        }

        // Yeni worksheet kurulum fonksiyonu
        function setupNewWorksheet(worksheet, moduleNames) {
            // Sütun genişliklerini ayarla
            worksheet.columns = [
                { header: '', width: 5 },  // A sütunu
                { header: '', width: 30 },   // B sütunu
                ...Array(21).fill({ width: 4 }), // C'den W'ye
                { width: 8 }, // X sütunu
                { width: 8 }  // Y sütunu
            ];

            // A1'den Y1'e kadar hücreleri birleştir ve T.C. yaz
            worksheet.mergeCells('A1:Y1');
            const tcCell = worksheet.getCell('A1');
            tcCell.value = 'T.C.';
            tcCell.alignment = { vertical: 'middle', horizontal: 'center' };
            tcCell.font = { bold: true, size: 12 };

            // A2'den Y2'ye kadar birleştir ve MİLLİ EĞİTİM BAKANLIĞI yaz
            worksheet.mergeCells('A2:Y2');
            const mebCell = worksheet.getCell('A2');
            mebCell.value = 'MİLLİ EĞİTİM BAKANLIĞI';
            mebCell.alignment = { vertical: 'middle', horizontal: 'center' };
            mebCell.font = { bold: true, size: 12 };

            // A3'ten Y3'e kadar birleştir ve KAHTA HALK EĞİTİMİ MERKEZİ yaz
            worksheet.mergeCells('A3:Y3');
            const hemCell = worksheet.getCell('A3');
            hemCell.value = 'KAHTA HALK EĞİTİMİ MERKEZİ';
            hemCell.alignment = { vertical: 'middle', horizontal: 'center' };
            hemCell.font = { bold: true, size: 12 };

            // A4'ten X4'e kadar birleştir ve MODÜL DEĞERLENDİRME ÇİZELGESİ yaz
            worksheet.mergeCells('A4:X4');
            const modulCell = worksheet.getCell('A4');
            modulCell.value = 'MODÜL DEĞERLENDİRME ÇİZELGESİ';
            modulCell.alignment = { vertical: 'middle', horizontal: 'center' };
            modulCell.font = { bold: true, size: 12 };

            // Y4'e EK 10 yaz
            const ekCell = worksheet.getCell('Y4');
            ekCell.value = 'EK 10';
            ekCell.alignment = { vertical: 'middle', horizontal: 'center' };
            ekCell.font = { bold: true, size: 11 };

            // A5 ve B5'i birleştir ve KURSUN ADI yaz
            worksheet.mergeCells('A5:B5');
            const kursAdiCell = worksheet.getCell('A5');
            kursAdiCell.value = 'KURSUN ADI';
            kursAdiCell.alignment = { vertical: 'middle', horizontal: 'center' };
            kursAdiCell.font = { bold: true, size: 11 };

            // C5'ten L5'e kadar birleştir
            worksheet.mergeCells('C5:L5');
            const kursAdiValueCell = worksheet.getCell('C5');
            kursAdiValueCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            kursAdiValueCell.font = { size: 11 };

            // A6 ve B6'yı birleştir ve KURSU NO yaz
            worksheet.mergeCells('A6:B6');
            const kursNoCell = worksheet.getCell('A6');
            kursNoCell.value = 'KURSU NO';
            kursNoCell.alignment = { vertical: 'middle', horizontal: 'center' };
            kursNoCell.font = { bold: true, size: 11 };

            // C6'dan L6'ya kadar birleştir
            worksheet.mergeCells('C6:L6');
            const kursNoValueCell = worksheet.getCell('C6');
            kursNoValueCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            kursNoValueCell.font = { size: 11 };

            // A7 ve B7'yi birleştir ve DÜZENLENDİĞİ YER yaz
            worksheet.mergeCells('A7:B7');
            const yerCell = worksheet.getCell('A7');
            yerCell.value = 'DÜZENLENDİĞİ YER';
            yerCell.alignment = { vertical: 'middle', horizontal: 'center' };
            yerCell.font = { bold: true, size: 11 };

            // C7'den L7'ye kadar birleştir
            worksheet.mergeCells('C7:L7');
            const yerValueCell = worksheet.getCell('C7');
            yerValueCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            yerValueCell.font = { size: 11 };

            // M7'den Y7'ye kadar birleştir
            worksheet.mergeCells('M7:Y7');
            const m7Cell = worksheet.getCell('M7');
            m7Cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            m7Cell.font = { size: 11 };

            // M5'ten T5'e kadar birleştir ve BAŞLAMA TARİHİ yaz
            worksheet.mergeCells('M5:T5');
            const baslamaTarihiCell = worksheet.getCell('M5');
            baslamaTarihiCell.value = 'BAŞLAMA TARİHİ';
            baslamaTarihiCell.alignment = { vertical: 'middle', horizontal: 'center' };
            baslamaTarihiCell.font = { bold: true, size: 11 };

            // U5'ten Y5'e kadar birleştir
            worksheet.mergeCells('U5:Y5');
            const u5Cell = worksheet.getCell('U5');
            u5Cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            u5Cell.font = { size: 11 };

            // M6'dan T6'ya kadar birleştir ve BİTİŞ TARİHİ yaz
            worksheet.mergeCells('M6:T6');
            const bitisTarihiCell = worksheet.getCell('M6');
            bitisTarihiCell.value = 'BİTİŞ TARİHİ';
            bitisTarihiCell.alignment = { vertical: 'middle', horizontal: 'center' };
            bitisTarihiCell.font = { bold: true, size: 11 };

            // U6'dan Y6'ya kadar birleştir
            worksheet.mergeCells('U6:Y6');
            const u6Cell = worksheet.getCell('U6');
            u6Cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            u6Cell.font = { size: 11 };

            // B8'den W8'e kadar birleştir ve KURS MODÜL DEĞERLENDİRME NOTU yaz
            worksheet.mergeCells('B8:W8');
            const modulDegerlendirmeCell = worksheet.getCell('B8');
            modulDegerlendirmeCell.value = 'KURS MODÜL DEĞERLENDİRME NOTU';
            modulDegerlendirmeCell.alignment = { vertical: 'middle', horizontal: 'center' };
            modulDegerlendirmeCell.font = { bold: true, size: 14, underline: true };
            modulDegerlendirmeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };

            // X8'den Y9'a kadar birleştir ve KURSUN BAŞARI PUANI VE DURUMU yaz
            worksheet.mergeCells('X8:Y9');
            const basariPuaniCell = worksheet.getCell('X8');
            basariPuaniCell.value = 'KURSUN BAŞARI PUANI VE DURUMU';
            basariPuaniCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            basariPuaniCell.font = { bold: true, size: 11 };

            // C9'dan W9'a kadar sayıları yaz
            for (let i = 0; i < 21; i++) {
                const col = String.fromCharCode(67 + i); // 67 = 'C'
                const cell = worksheet.getCell(`${col}9`);
                cell.value = i + 1;
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.font = { bold: true, size: 11 };
            }

            // A10 ve B10'daki yazıları alt çizgiye hizala
            const noCell = worksheet.getCell('A10');
            noCell.value = 'No';
            noCell.alignment = { vertical: 'bottom', horizontal: 'center' };
            noCell.font = { bold: true, size: 11 };

            const adSoyadCell = worksheet.getCell('B10');
            adSoyadCell.value = 'AD SOYAD';
            adSoyadCell.alignment = { vertical: 'bottom', horizontal: 'center' };
            adSoyadCell.font = { bold: true, size: 11 };

            // X10'a Başarı Puanı yaz (dikey)
            const basariPuaniBaslikCell = worksheet.getCell('X10');
            basariPuaniBaslikCell.value = 'Başarı Puanı';
            basariPuaniBaslikCell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
            basariPuaniBaslikCell.font = { bold: true, size: 11 };

            // Y10'a Başarı Durumu yaz (dikey)
            const basariDurumuBaslikCell = worksheet.getCell('Y10');
            basariDurumuBaslikCell.value = 'Başarı Durumu';
            basariDurumuBaslikCell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
            basariDurumuBaslikCell.font = { bold: true, size: 11 };

            // X11'den X50'ye kadar formülleri ekle
            for (let row = 11; row <= 50; row++) {
                const formulaCell = worksheet.getCell(`X${row}`);
                formulaCell.value = {
                    formula: `=IF(COUNTIF(C${row}:W${row},"<>")=0,"Not Gir",ROUND(AVERAGE(C${row}:W${row}),1))`,
                    date1904: false
                };
                formulaCell.alignment = { vertical: 'middle', horizontal: 'center' };
                formulaCell.numFmt = '0.0';

                // Y sütununa başarı durumu formülünü ekle
                const basariDurumuCell = worksheet.getCell(`Y${row}`);
                basariDurumuCell.value = {
                    formula: `=IF(COUNTIF(C${row}:W${row},"<>")=0,"Not Gir",IF(COUNTIF(C${row}:W${row},"<50")<>0,"Başarısız","Başarılı"))`,
                    date1904: false
                };
                basariDurumuCell.alignment = { vertical: 'middle', horizontal: 'center' };
            }

            // Kenar çizgilerini ekle
            for (let row = 1; row <= 50; row++) {
                for (let col = 'A'; col <= 'Y'; col = String.fromCharCode(col.charCodeAt(0) + 1)) {
                    const cell = worksheet.getCell(`${col}${row}`);
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                }
            }

            // İmza alanlarını ekle
            worksheet.mergeCells('A55:Y55');
            const aciklamaCell = worksheet.getCell('A55');
            aciklamaCell.value = 'İş bu modül değerlendirme çizelgesi kayıtlarımıza uygun olarak düzenlenmiş, bilgilerin doğru ve eksiksiz olduğu imza altına alınmıştır.';
            aciklamaCell.alignment = { vertical: 'middle', horizontal: 'center' };
            aciklamaCell.font = { size: 11 };

            // İmza alanları için hücre birleştirmeleri
            // Kurs Öğretmeni
            worksheet.mergeCells('A57:D57');
            const ogretmenCell = worksheet.getCell('A57');
            ogretmenCell.value = 'Murat BİLGİN';
            ogretmenCell.alignment = { horizontal: 'center' };

            worksheet.mergeCells('A58:D58');
            const ogretmenUnvanCell = worksheet.getCell('A58');
            ogretmenUnvanCell.value = 'Kurs Öğretmeni';
            ogretmenUnvanCell.alignment = { horizontal: 'center' };

            // Müdür Yardımcısı
            worksheet.mergeCells('H57:P57');
            const mudurYardCell = worksheet.getCell('H57');
            mudurYardCell.value = 'Murat BİLGİN';
            mudurYardCell.alignment = { horizontal: 'center' };

            worksheet.mergeCells('H58:P58');
            const mudurYardUnvanCell = worksheet.getCell('H58');
            mudurYardUnvanCell.value = 'Müdür Yardımcısı';
            mudurYardUnvanCell.alignment = { horizontal: 'center' };

            // Halk Eğitimi Merkezi Müdürü
            worksheet.mergeCells('R57:Y57');
            const mudurCell = worksheet.getCell('R57');
            mudurCell.value = 'Fahri DOĞAN';
            mudurCell.alignment = { horizontal: 'center' };

            worksheet.mergeCells('R58:Y58');
            const mudurUnvanCell = worksheet.getCell('R58');
            mudurUnvanCell.value = 'Halk Eğitimi Merkezi Müdürü';
            mudurUnvanCell.alignment = { horizontal: 'center' };

            // İmza alanlarının fontunu ayarla
            ['A57:Y58'].forEach(range => {
                worksheet.getCell(range).font = { bold: true, size: 11 };
            });

            // Modül adlarını C10'dan başlayarak dikey yaz
            let moduleIndex = 0;
            for (const moduleName of moduleNames) {
                if (moduleIndex >= 21) break; // Maksimum 21 modül

                const col = String.fromCharCode(67 + moduleIndex); // C'den başla
                const cell = worksheet.getCell(`${col}10`);
                cell.value = moduleName;
                cell.alignment = {
                    textRotation: 90, // Dikey yazı
                    vertical: 'middle',
                    horizontal: 'center'
                };
                cell.font = {
                    bold: true,
                    size: 11
                };
                moduleIndex++;
            }

            // Başlık satırı yüksekliğini ayarla
            worksheet.getRow(10).height = 150;
        }

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