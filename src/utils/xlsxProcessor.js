function readXLSXFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            resolve(workbook);
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

function processWorkbook(workbook) {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = jsonData[0] || [];
    const presenzeIndex = headers.findIndex(h => String(h).trim().toLowerCase() === 'presenze');
    const processedData = [];

    // Find the last row that contains "TOTALE" (case-insensitive) and mark it to skip
    let lastTotaleIndex = -1;
    for (let i = jsonData.length - 1; i >= 1; i--) { // skip header at index 0
        const row = jsonData[i] || [];
        const hasTotale = row.some(cell => {
            if (cell == null) return false;
            const text = String(cell).trim().toUpperCase();
            return text.includes('TOTALE');
        });
        if (hasTotale) {
            lastTotaleIndex = i;
            break;
        }
    }

    jsonData.forEach((row, rowIndex) => {
        if (rowIndex === 0) {
            processedData.push([...row, "Percentuale"]);
            return;
        }

        // Skip the last totals row
        if (rowIndex === lastTotaleIndex) {
            return;
        }

        const isAllZeros = row.every((cell, index) => index !== presenzeIndex && (cell === 0 || cell === null || cell === undefined || (typeof cell === 'string' && cell.trim() === '')));
        if (!isAllZeros) {
            const presenzeValue = row[presenzeIndex] || 0;
            const totalAttendance = row.reduce((sum, cell, index) => index !== presenzeIndex ? sum + (cell || 0) : sum, 0);
            const percentuale = totalAttendance > 0 ? (presenzeValue / totalAttendance) * 100 : 0;
            processedData.push([...row, percentuale.toFixed(2)]);
        }
    });

    return processedData;
}

function exportToCSV(data) {
    const csvContent = data.map(row => row.join(",")).join("\n");
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "processed_attendance.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

export { readXLSXFile, processWorkbook, exportToCSV };