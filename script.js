// Event Listeners
document.getElementById('processButton').addEventListener('click', processFile);
document.getElementById('fileInput').addEventListener('change', handleFileSelect);

const dropZone = document.getElementById('dropZone');
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
});

['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, unhighlight, false);
});

dropZone.addEventListener('drop', handleDrop, false);

// Utility functions
function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e) {
    dropZone.classList.add('drag-over');
}

function unhighlight(e) {
    dropZone.classList.remove('drag-over');
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    document.getElementById('fileInput').files = files;
    handleFile(files[0]);
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    handleFile(file);
}

function handleFile(file) {
    if (file) {
        if (!file.name.endsWith('.xlsx')) {
            showStatus('Please select XLSX files only.', 'error');
            return;
        }

        updateFileList();
    }
}

function showStatus(message, type = '') {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
}

// Ordered source → output column mappings (output column order is preserved)
const columnMappings = [
    { source: 'DR code', output: 'DR Code' },
    { source: 'Cnt No', output: 'Contract No.' },
    { source: 'Action', output: 'Action' },
    { source: 'Name', output: 'Client Name' },
    { source: 'Code', output: 'Client Code' },
    { source: 'Counter', output: 'Counter' },
    { source: 'Qty', output: 'Quantity' },
    { source: 'Buy Price', output: 'Price' },
    { source: 'CCY', output: 'Traded Currency' },
    { source: 'Amount (traded)', output: 'Contract Value' },
    { source: 'CCY (Sett)', output: 'Settlement Currency' }
];

const ACTION_STYLES = {
    'due today': {
        fill: 'FFEB9C', // Excel Neutral fill
        font: '9C6500'  // Excel Neutral font
    },
    'force sell': {
        fill: 'FFC7CE', // Excel Bad fill
        font: '9C0006'  // Excel Bad font
    }
};

const ACTION_SORT_ORDER = {
    'force sell': 0,
    'due today': 1,
    'undue': 2
};

const QUANTITY_FORMAT = '#,##0';
const CONTRACT_VALUE_FORMAT = '#,##0.00';

function normalizeAction(actionValue) {
    if (actionValue == null) return '';
    return actionValue.toString().trim().toLowerCase();
}

function getActionStyle(actionValue) {
    return ACTION_STYLES[normalizeAction(actionValue)] || null;
}

function getActionSortRank(actionValue) {
    const rank = ACTION_SORT_ORDER[normalizeAction(actionValue)];
    return rank === undefined ? 99 : rank;
}

function getOutputFileName(fileIndex, total) {
    const now = new Date();
    const dd = String(now.getDate()).padStart(2, '0');
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const yyyy = String(now.getFullYear());
    const base = `Cross-Border Trade ${dd}${mm}${yyyy}`;
    return total > 1 ? `${base}_${fileIndex + 1}.xlsx` : `${base}.xlsx`;
}

const allowedDealerCodes = [
    'CC', 'CC1', 'CC2', 'CC6', 'CCD', 'CCD1', 'CD3', 'CD5', 'CD6', 'CD9', 'CEA', 'CEZ',
    'CE1', 'CE0', 'CE11', 'CPI', 'CSK', 'CSY', 'CSW', 'CSW5', 'CD2', 'CSM', 'CD12', 'DMU',
    'DMW', 'DMJ', 'DMC', 'DME', 'CSR1', 'CSW3', 'DMF', 'DMG', 'DMH', 'DMK', 'CE12', 'CT12', 'CT3', 'DMS'
];

let currentFileIndex = 0;
let totalFiles = 0;
let dealerCodeCounts = {};

function processFile() {
    const fileInput = document.getElementById('fileInput');
    const processButton = document.getElementById('processButton');
    const files = fileInput.files;

    if (!files.length) {
        showStatus('Please select at least one file.', 'error');
        return;
    }

    currentFileIndex = 0;
    totalFiles = files.length;
    processButton.disabled = true;
    processButton.classList.add('loading');

    processNextFile();
}

function processNextFile() {
    const fileInput = document.getElementById('fileInput');
    const processButton = document.getElementById('processButton');
    const files = fileInput.files;

    if (currentFileIndex >= totalFiles) {
        // All files processed
        processButton.disabled = false;
        processButton.classList.remove('loading');
        document.querySelector('.progress-container').style.display = 'none';

        // Generate and display the summary
        let summary = 'All files processed successfully!\n\nDealer Code Summary:';
        const sortedDealerCodes = Object.keys(dealerCodeCounts).sort();

        let hasCounts = false;
        sortedDealerCodes.forEach(code => {
            if (dealerCodeCounts[code] > 0) {
                hasCounts = true;
                summary += `\n${code}: ${dealerCodeCounts[code]} contract${dealerCodeCounts[code] > 1 ? 's' : ''}`;
            }
        });

        if (!hasCounts) {
            summary = 'All files processed successfully!\nNo matching contracts found.';
        }

        showStatus(summary, 'success');
        return;
    }

    updateProgress(currentFileIndex + 1, totalFiles);

    // Update file list item status
    const fileItems = document.querySelectorAll('.file-item');
    if (fileItems[currentFileIndex]) {
        fileItems[currentFileIndex].classList.add('processed');
    }

    const file = files[currentFileIndex];
    showStatus(`Processing file ${currentFileIndex + 1} of ${totalFiles}: ${file.name}...`);

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON
            let jsonData = XLSX.utils.sheet_to_json(worksheet);

            // Reset counts for first file
            if (currentFileIndex === 0) {
                dealerCodeCounts = {};
            }

            // Filter for CNPG contract numbers, specific dealer codes and rename columns
            const filteredData = jsonData
                .filter(row => {
                    const contractNum = row['Cnt No'];
                    const drCode = row['DR code'];

                    if (contractNum &&
                        contractNum.toString().startsWith('CNPG') &&
                        drCode &&
                        allowedDealerCodes.includes(drCode.toString())) {
                        // Count the dealer codes
                        dealerCodeCounts[drCode.toString()] = (dealerCodeCounts[drCode.toString()] || 0) + 1;
                        return true;
                    }
                    return false;
                })
                .map(row => {
                    const filteredRow = {};
                    columnMappings.forEach(({ source, output }) => {
                        if (!row.hasOwnProperty(source)) return;

                        let value = row[source];
                        if (output === 'Quantity' && value != null && value !== '') {
                            value = Math.round(Number(value));
                        } else if (output === 'Contract Value' && value != null && value !== '') {
                            value = Number(value);
                        }
                        filteredRow[output] = value;
                    });
                    return filteredRow;
                });

            // Sort: Force sell → Due today → Undue, then by Contract No.
            filteredData.sort((a, b) => {
                const actionDiff = getActionSortRank(a['Action']) - getActionSortRank(b['Action']);
                if (actionDiff !== 0) return actionDiff;
                if (a['Contract No.'] < b['Contract No.']) return -1;
                if (a['Contract No.'] > b['Contract No.']) return 1;
                return 0;
            });

            // Create new workbook with filtered data (header order follows columnMappings)
            const newWorkbook = XLSX.utils.book_new();
            const outputHeaders = columnMappings.map(({ output }) => output);
            const newWorksheet = XLSX.utils.json_to_sheet(filteredData, { header: outputHeaders });

            const quantityCol = outputHeaders.indexOf('Quantity');
            const contractValueCol = outputHeaders.indexOf('Contract Value');
            const actionCol = outputHeaders.indexOf('Action');

            // Number formats, Action highlights, and bold headers (no borders)
            const range = XLSX.utils.decode_range(newWorksheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!newWorksheet[cell_address]) continue;

                    const cellStyle = {};

                    if (R === 0) {
                        cellStyle.font = { bold: true };
                    } else {
                        if (C === quantityCol) {
                            newWorksheet[cell_address].z = QUANTITY_FORMAT;
                            cellStyle.numFmt = QUANTITY_FORMAT;
                        } else if (C === contractValueCol) {
                            newWorksheet[cell_address].z = CONTRACT_VALUE_FORMAT;
                            cellStyle.numFmt = CONTRACT_VALUE_FORMAT;
                        } else if (C === actionCol) {
                            const actionStyle = getActionStyle(newWorksheet[cell_address].v);
                            if (actionStyle) {
                                cellStyle.fill = {
                                    patternType: 'solid',
                                    fgColor: { rgb: actionStyle.fill }
                                };
                                if (actionStyle.font) {
                                    cellStyle.font = { color: { rgb: actionStyle.font } };
                                }
                            }
                        }
                    }

                    if (Object.keys(cellStyle).length) {
                        newWorksheet[cell_address].s = cellStyle;
                    }
                }
            }

            // Auto-fit columns
            const columnsWidth = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                let maxLength = 0;
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                    if (newWorksheet[cell_address]) {
                        const cellLength = String(newWorksheet[cell_address].v).length;
                        maxLength = Math.max(maxLength, cellLength);
                    }
                }
                columnsWidth[C] = maxLength + 2; // Add some padding
            }

            newWorksheet['!cols'] = columnsWidth.map(width => ({ width }));

            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Filtered Data');

            const outputFileName = getOutputFileName(currentFileIndex, totalFiles);

            XLSX.writeFile(newWorkbook, outputFileName);

            // Process next file
            currentFileIndex++;
            setTimeout(processNextFile, 100); // Small delay between files

        } catch (error) {
            showStatus(`Error processing file ${currentFileIndex + 1} (${file.name}): ${error.message}`, 'error');
            processButton.disabled = false;
            processButton.classList.remove('loading');
        }
    };

    reader.onerror = function () {
        showStatus(`Error reading file ${currentFileIndex + 1} (${file.name})`, 'error');
        processButton.disabled = false;
        processButton.classList.remove('loading');
    };

    reader.readAsArrayBuffer(file);
}

// Update the file input to allow multiple files
document.getElementById('fileInput').setAttribute('multiple', 'true');

function updateProgress(current, total) {
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const progressContainer = document.querySelector('.progress-container');

    if (progressContainer) {
        progressContainer.style.display = 'block';
    }

    const percentage = (current / total) * 100;
    progressBar.style.width = `${percentage}%`;
    progressText.textContent = `Processing file ${current} of ${total}`;
}

function updateFileList() {
    const files = document.getElementById('fileInput').files;
    const filePreview = document.getElementById('filePreview');
    const fileList = document.createElement('div');
    fileList.className = 'file-list';

    Array.from(files).forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <i class="fas fa-file-excel"></i>
            <span>${file.name}</span>
        `;
        fileList.appendChild(fileItem);
    });

    // Remove old file list if exists
    const oldFileList = filePreview.querySelector('.file-list');
    if (oldFileList) {
        oldFileList.remove();
    }

    filePreview.style.display = 'block';
    filePreview.appendChild(fileList);
    document.getElementById('processButton').disabled = files.length === 0;
}

// Initialize event listeners when document loads
document.addEventListener('DOMContentLoaded', () => {
    // Reset progress when new files are selected
    document.getElementById('fileInput').addEventListener('change', () => {
        document.querySelector('.progress-container').style.display = 'none';
        document.getElementById('progressBar').style.width = '0%';
    });
});