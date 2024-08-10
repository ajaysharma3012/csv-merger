document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const fileList = document.getElementById('fileList');
    const message = document.getElementById('message');
    const progressBar = document.getElementById('progressBar').firstElementChild;
    const downloadButton = document.getElementById('downloadButton');

    form.addEventListener('submit', function(e) {
        e.preventDefault();
        const masterFile = document.getElementById('masterFile').files[0];
        const secondaryFiles = document.getElementById('secondaryFiles').files;
        const rowsToDelete = parseInt(document.getElementById('rowsToDelete').value) || 0;

        if (!masterFile || secondaryFiles.length === 0) {
            message.textContent = 'Please select all required files.';
            return;
        }

        progressBar.style.width = '0%';
        message.textContent = 'Merging files...';
        downloadButton.style.display = 'none';

        mergeFiles(masterFile, secondaryFiles, rowsToDelete)
            .then(mergedWorkbook => {
                progressBar.style.width = '100%';
                message.textContent = 'Files merged successfully!';
                downloadButton.style.display = 'block';
                downloadButton.onclick = function() {
                    XLSX.writeFile(mergedWorkbook, 'merged.xlsx');
                };
            })
            .catch(error => {
                console.error('Error:', error);
                message.textContent = 'An error occurred while merging files.';
                progressBar.style.width = '0%';
            });
    });

    document.getElementById('secondaryFiles').addEventListener('change', function(e) {
        fileList.innerHTML = '';
        for (let i = 0; i < this.files.length; i++) {
            fileList.innerHTML += `<p>${this.files[i].name}</p>`;
        }
    });

    async function mergeFiles(masterFile, secondaryFiles, rowsToDelete) {
        const masterData = await readFile(masterFile);
        const secondaryDataPromises = Array.from(secondaryFiles).map(file => readFile(file, rowsToDelete));
        const secondaryDataArrays = await Promise.all(secondaryDataPromises);

        let mergedData = [masterData.headers, ...masterData.data];
        secondaryDataArrays.forEach(secData => {
            mergedData.push(...secData.data);
        });

        // Apply special formatting to merged data
        const installationIdIndex = masterData.headers.findIndex(h => h === 'Installation ID');
        const l1ApprovalDateIndex = masterData.headers.findIndex(h => h === 'L1 Approval Date');
        const l2ApprovalDateIndex = masterData.headers.findIndex(h => h === 'L2 Approval Date');
        const installationDateIndex = masterData.headers.findIndex(h => h === 'Installation Date');

        mergedData = mergedData.map((row, index) => {
            if (index === 0) return row; // Skip header row
            return row.map((cell, cellIndex) => {
                if (cellIndex === installationIdIndex) {
                    // Remove leading zeros for Installation ID
                    return cell ? cell.toString().replace(/^0+/, '') : cell;
                } else if (cellIndex === l1ApprovalDateIndex || cellIndex === l2ApprovalDateIndex) {
                    // Keep only date part for L1 and L2 Approval Dates
                    return cell ? formatDateOnly(cell) : cell;
                } else if (cellIndex === installationDateIndex) {
                    // Format Installation Date as DD-MM-YYYY HH:MM:SS
                    return cell ? formatInstallationDate(cell) : cell;
                }
                return cell;
            });
        });

        const mergedWorkbook = XLSX.utils.book_new();
        const mergedSheet = XLSX.utils.aoa_to_sheet(mergedData);

        // Auto-size columns
        const colWidths = mergedData.reduce((widths, row) => {
            row.forEach((cell, i) => {
                const cellWidth = cell ? cell.toString().length : 0;
                widths[i] = Math.max(widths[i] || 0, cellWidth);
            });
            return widths;
        }, {});

        mergedSheet['!cols'] = Object.keys(colWidths).map(i => ({ wch: colWidths[i] }));

        XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, "Merged Data");
        return mergedWorkbook;
    }

    function readFile(file, rowsToDelete = 0) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    const range = XLSX.utils.decode_range(worksheet['!ref']);

                    let headers = [];
                    let jsonData = [];

                    // Read headers
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const headerCell = worksheet[XLSX.utils.encode_cell({r: range.s.r, c: C})];
                        headers.push(headerCell ? headerCell.v : '');
                    }

                    // Read data
                    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
                        let row = [];
                        for (let C = range.s.c; C <= range.e.c; ++C) {
                            const cell = worksheet[XLSX.utils.encode_cell({r: R, c: C})];
                            if (cell) {
                                row.push(cell.w || cell.v);
                            } else {
                                row.push(null);
                            }
                        }
                        jsonData.push(row);
                    }

                    jsonData = jsonData.slice(rowsToDelete);
                    resolve({ headers, data: jsonData });
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function formatDateOnly(dateString) {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString; // Return original if not a valid date
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }

    function formatInstallationDate(dateString) {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString; // Return original if not a valid date
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        const hours = date.getHours().toString().padStart(2, '0');
        const minutes = date.getMinutes().toString().padStart(2, '0');
        const seconds = date.getSeconds().toString().padStart(2, '0');
        return `${day}-${month}-${year}  ${hours}:${minutes}:${seconds}`;
    }
});