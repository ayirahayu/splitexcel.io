document.getElementById('processButton').addEventListener('click', function () {
    const fileInput = document.getElementById('fileInput');
    const status = document.getElementById('status');
  
    if (!fileInput.files.length) {
      status.textContent = 'Please upload an Excel file!';
      return;
    }
  
    const file = fileInput.files[0];
    const reader = new FileReader();
  
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
  
      // Process each sheet
      workbook.SheetNames.forEach((sheetName) => {
        const sheetData = workbook.Sheets[sheetName];
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, sheetData, sheetName);
  
        // Generate a file name
        const fileName = `${sheetName}.xlsx`;
  
        // Write and trigger download
        const blob = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'blob' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        link.click();
      });
  
      status.textContent = 'Sheets split successfully and downloaded!';
    };
  
    reader.readAsArrayBuffer(file);
  });
  