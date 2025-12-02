// Browser-only implementation using global XLSX from CDN.
// Handles upload, removes all-zero lecture columns, preserves Presenze as a formula,
// and adds Percentuale after Presenze.

(() => {
    const form = document.getElementById('upload-form');
    const fileInput = document.getElementById('file-input');
    const output = document.getElementById('output');

    if (!window.XLSX) {
        if (output) output.textContent = 'Error: XLSX library not loaded.';
        return;
    }

    form.addEventListener('submit', (e) => {
        e.preventDefault();
        const file = fileInput.files && fileInput.files[0];
        if (!file) {
            setOutput('Please select a .xlsx file.');
            return;
        }

        const reader = new FileReader();
        reader.onload = () => processArrayBuffer(reader.result, file.name).catch(err => {
            console.error(err);
            setOutput('Processing failed: ' + err.message);
        });
        reader.onerror = () => setOutput('Could not read the file.');
        reader.readAsArrayBuffer(file);
    });

    function setOutput(msgHtml) {
        if (!output) return;
        // Clear previous content
        output.innerHTML = msgHtml;
    }

    function toColLetter(n) { // 1-based -> A1 column letters
        let s = '';
        while (n > 0) {
            const m = (n - 1) % 26;
            s = String.fromCharCode(65 + m) + s;
            n = Math.floor((n - 1) / 26);
        }
        return s;
    }

    function isZeroLike(v) {
        if (v === null || v === undefined) return true;
        if (typeof v === 'number') return v === 0;
        if (typeof v === 'boolean') return v === false;
        if (typeof v === 'string') {
            const t = v.trim();
            if (!t) return true;
            const num = Number(t.replace(',', '.'));
            if (!Number.isNaN(num)) return num === 0;
            return false; // non-numeric text counts as non-zero
        }
        return false;
    }

    async function processArrayBuffer(buf, originalName) {
        const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true });
        const firstSheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[firstSheetName];

        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
        if (!aoa || aoa.length === 0) throw new Error('Empty sheet.');
        const header = aoa[0].map(h => (h == null ? '' : String(h)));
        if (header.length < 3) throw new Error('Unexpected format. Need at least: Name, Matricola, dates..., Presenze.');

        // Find "Presenze" column (case-insensitive)
        let presIdx = header.findIndex(h => h.trim().toLowerCase() === 'presenze');
        if (presIdx === -1) throw new Error('Column "Presenze" not found in header.');

        // Lecture columns are between index 2 and presIdx-1
        const lectureIdxs = [];
        for (let c = 2; c < presIdx; c++) lectureIdxs.push(c);

        // Determine which lecture columns are all zeros/blanks
        const keepLectureIdxs = lectureIdxs.filter(c => {
            for (let r = 1; r < aoa.length; r++) {
                const v = aoa[r]?.[c];
                if (!isZeroLike(v)) return true; // keep if any non-zero
            }
            return false; // drop if all zero-like
        });

        const lectureCount = keepLectureIdxs.length;

        // Build new header: Name, Matricola, kept dates..., Presenze, Percentuale
        const newHeader = [
            header[0] || 'Nome',
            header[1] || 'Matricola',
            ...keepLectureIdxs.map(c => header[c]),
            'Presenze',
            'Percentuale'
        ];

        const newAoa = [newHeader];

        // New positions in the rebuilt sheet
        const newLectureStart = 2; // 0-based
        const newLectureEnd = newLectureStart + Math.max(0, lectureCount - 1);
        const newPresIdx = newLectureStart + lectureCount; // presenze after lectures
        const newPercIdx = newPresIdx + 1;

        for (let r = 1; r < aoa.length; r++) {
            const srcRow = aoa[r] || [];
            const row = [];

            // Name, Matricola
            row[0] = srcRow[0] ?? '';
            row[1] = srcRow[1] ?? '';

            // Kept lecture values
            for (let i = 0; i < keepLectureIdxs.length; i++) {
                row[newLectureStart + i] = srcRow[keepLectureIdxs[i]] ?? 0;
            }

            const excelRow = r + 1; // A1 row number (header is row 1)

            // Presenze formula (kept as a formula on the reduced lecture range)
            if (lectureCount > 0) {
                const startColLetter = toColLetter(newLectureStart + 1); // 1-based
                const endColLetter = toColLetter(newLectureEnd + 1);
                row[newPresIdx] = { f: `COUNTIF(${startColLetter}${excelRow}:${endColLetter}${excelRow},1)` };
            } else {
                row[newPresIdx] = 0;
            }

            // Percentuale = Presenze / COUNTA(lecture columns)
            if (lectureCount > 0) {
                const startColLetter = toColLetter(newLectureStart + 1); // 1-based
                const endColLetter = toColLetter(newLectureEnd + 1);
                const presCellRef = `${toColLetter(newPresIdx + 1)}${excelRow}`;
                row[newPercIdx] = {
                    f: `${presCellRef}/COUNTA($${startColLetter}$1:$${endColLetter}$1)`,
                    z: '0.00%'
                };
            } else {
                row[newPercIdx] = '';
            }

            newAoa.push(row);
        }

        // Create new workbook and sheet
        const outWb = XLSX.utils.book_new();
        const outWs = XLSX.utils.aoa_to_sheet(newAoa);
        
        // Make first row bold
        const range = XLSX.utils.decode_range(outWs['!ref']);
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_col(C) + "1"; // Row 1
            if (!outWs[address]) continue;
            outWs[address].s = {
                font: { bold: true }
            };
        }
        
        XLSX.utils.book_append_sheet(outWb, outWs, 'Processed');

        const outName = (originalName?.replace(/\.xlsx$/i, '') || 'attendance') + '_processed.xlsx';
        XLSX.writeFile(outWb, outName);

        const removed = lectureIdxs.length - lectureCount;
        setOutput(`Done. Removed ${removed} lecture column(s). Downloaded: <strong>${outName}</strong>.`);
    }
})();