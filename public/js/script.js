document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const formData = new FormData();
    const fileInput = document.getElementById('excelFile');
    
    if (!fileInput.files[0]) {
        alert('Lütfen bir dosya seçin');
        return;
    }

    // Excel dosyası kontrolü
    const fileName = fileInput.files[0].name;
    if (!fileName.match(/\.(xlsx|xls)$/)) {
        alert('Lütfen geçerli bir Excel dosyası seçin (.xlsx veya .xls)');
        return;
    }

    formData.append('excelFile', fileInput.files[0]);

    // Loading göster
    document.getElementById('loading').style.display = 'block';
    document.getElementById('result').style.display = 'none';

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.message || 'Dosya yükleme hatası');
        }
        
        if (data.success) {
            document.getElementById('downloadLink').href = data.downloadUrl;
            document.getElementById('result').style.display = 'block';
        } else {
            throw new Error(data.message);
        }
    } catch (error) {
        alert('Bir hata oluştu: ' + error.message);
    } finally {
        document.getElementById('loading').style.display = 'none';
    }
}); 