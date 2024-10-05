function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

const sheetName = getQueryParam('sheetName');
const fileUrl = getQueryParam('fileUrl');

(async () => {
    if (!fileUrl || !sheetName) {
        alert("Invalid sheet data.");
        return;
    }

    document.getElementById('download-sheet').addEventListener('click', async () => {
        try {
            const response = await fetch(fileUrl);
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
            const sheet = workbook.Sheets[sheetName];

            if (!sheet) {
                alert("Sheet not found.");
                return;
            }

            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, sheet, sheetName);
            const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });

            const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `${sheetName}.xlsx`;
            link.click();
        } catch (error) {
            console.error("Error downloading the Excel sheet:", error);
            alert("Failed to download the Excel sheet. Please try again.");
        }
    });

    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
})();
