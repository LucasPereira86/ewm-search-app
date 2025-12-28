/**
 * EWM Search App - Excel File Search Application
 * Allows loading, searching, and exporting Excel file data
 */

// ===== Global State =====
const state = {
    data: [],           // All rows from Excel
    columns: [],        // Column headers
    filteredData: [],   // Search results
    fileName: ''        // Loaded file name
};

// ===== DOM Elements =====
const elements = {
    // Sections
    uploadSection: document.getElementById('uploadSection'),
    searchSection: document.getElementById('searchSection'),

    // Upload
    fileInput: document.getElementById('fileInput'),
    selectFileBtn: document.getElementById('selectFileBtn'),
    uploadCard: document.querySelector('.upload-card'),
    fileInfo: document.getElementById('fileInfo'),

    // Search
    searchInput: document.getElementById('searchInput'),
    clearSearchBtn: document.getElementById('clearSearchBtn'),
    columnFilter: document.getElementById('columnFilter'),

    // Stats
    resultCount: document.getElementById('resultCount'),
    totalCount: document.getElementById('totalCount'),

    // Table
    tableHead: document.getElementById('tableHead'),
    tableBody: document.getElementById('tableBody'),
    noResults: document.getElementById('noResults'),

    // Actions
    exportBtn: document.getElementById('exportBtn'),
    newFileBtn: document.getElementById('newFileBtn'),

    // UI
    toast: document.getElementById('toast'),
    loadingOverlay: document.getElementById('loadingOverlay')
};

// ===== Initialization =====
document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadFromLocalStorage();
});

function initEventListeners() {
    // File upload
    elements.selectFileBtn.addEventListener('click', () => elements.fileInput.click());
    elements.uploadCard.addEventListener('click', (e) => {
        if (e.target !== elements.selectFileBtn) {
            elements.fileInput.click();
        }
    });
    elements.fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop
    elements.uploadCard.addEventListener('dragover', handleDragOver);
    elements.uploadCard.addEventListener('dragleave', handleDragLeave);
    elements.uploadCard.addEventListener('drop', handleDrop);

    // Search
    elements.searchInput.addEventListener('input', debounce(performSearch, 200));
    elements.clearSearchBtn.addEventListener('click', clearSearch);
    elements.columnFilter.addEventListener('change', performSearch);

    // Actions
    elements.exportBtn.addEventListener('click', exportResults);
    elements.newFileBtn.addEventListener('click', resetApp);
}

// ===== File Handling =====
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) processFile(file);
}

function handleDragOver(event) {
    event.preventDefault();
    elements.uploadCard.classList.add('dragover');
}

function handleDragLeave(event) {
    event.preventDefault();
    elements.uploadCard.classList.remove('dragover');
}

function handleDrop(event) {
    event.preventDefault();
    elements.uploadCard.classList.remove('dragover');

    const file = event.dataTransfer.files[0];
    if (file) processFile(file);
}

async function processFile(file) {
    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv'
    ];
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!validTypes.includes(file.type) && !validExtensions.includes(fileExtension)) {
        showToast('Por favor, selecione um arquivo Excel (.xlsx, .xls) ou CSV', 'error');
        return;
    }

    showLoading(true);

    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Get first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 2) {
            showToast('O arquivo parece estar vazio ou não tem dados suficientes', 'error');
            showLoading(false);
            return;
        }

        // Extract headers and data
        state.columns = jsonData[0].map(col => String(col || '').trim());
        state.data = jsonData.slice(1).filter(row => row.some(cell => cell !== undefined && cell !== ''));
        state.filteredData = [...state.data];
        state.fileName = file.name;

        // Save to localStorage for offline use
        saveToLocalStorage();

        // Update UI
        setupSearchUI();
        showSearchSection();

        showToast(`Arquivo carregado: ${state.data.length} itens encontrados`, 'success');

    } catch (error) {
        console.error('Error processing file:', error);
        showToast('Erro ao processar o arquivo. Tente novamente.', 'error');
    }

    showLoading(false);
}

// ===== Search Functionality =====
function setupSearchUI() {
    // Populate column filter
    elements.columnFilter.innerHTML = '<option value="all">Todas as colunas</option>';
    state.columns.forEach((col, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = col;
        elements.columnFilter.appendChild(option);
    });

    // Setup table headers
    elements.tableHead.innerHTML = '';
    const headerRow = document.createElement('tr');
    state.columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        headerRow.appendChild(th);
    });
    elements.tableHead.appendChild(headerRow);

    // Update stats
    elements.totalCount.textContent = `Total: ${state.data.length} itens`;

    // Render initial data
    renderTable();
}

function performSearch() {
    const query = elements.searchInput.value.toLowerCase().trim();
    const columnFilter = elements.columnFilter.value;

    if (!query) {
        state.filteredData = [...state.data];
    } else {
        state.filteredData = state.data.filter(row => {
            if (columnFilter === 'all') {
                return row.some(cell =>
                    String(cell || '').toLowerCase().includes(query)
                );
            } else {
                const colIndex = parseInt(columnFilter);
                return String(row[colIndex] || '').toLowerCase().includes(query);
            }
        });
    }

    renderTable(query);
}

function clearSearch() {
    elements.searchInput.value = '';
    elements.columnFilter.value = 'all';
    state.filteredData = [...state.data];
    renderTable();
    elements.searchInput.focus();
}

function renderTable(highlightQuery = '') {
    elements.tableBody.innerHTML = '';

    if (state.filteredData.length === 0) {
        elements.noResults.classList.remove('hidden');
        document.getElementById('resultsTable').classList.add('hidden');
        elements.resultCount.textContent = '0 resultados';
        return;
    }

    elements.noResults.classList.add('hidden');
    document.getElementById('resultsTable').classList.remove('hidden');

    // Limit displayed rows for performance
    const maxRows = 500;
    const displayData = state.filteredData.slice(0, maxRows);

    displayData.forEach(row => {
        const tr = document.createElement('tr');

        state.columns.forEach((_, colIndex) => {
            const td = document.createElement('td');
            let cellValue = String(row[colIndex] ?? '');

            // Highlight matching text
            if (highlightQuery && cellValue.toLowerCase().includes(highlightQuery.toLowerCase())) {
                const regex = new RegExp(`(${escapeRegex(highlightQuery)})`, 'gi');
                cellValue = cellValue.replace(regex, '<span class="highlight">$1</span>');
                td.innerHTML = cellValue;
            } else {
                td.textContent = cellValue;
            }

            tr.appendChild(td);
        });

        elements.tableBody.appendChild(tr);
    });

    // Update result count
    const countText = state.filteredData.length === 1
        ? '1 resultado'
        : `${state.filteredData.length} resultados`;

    const extraText = state.filteredData.length > maxRows
        ? ` (mostrando ${maxRows})`
        : '';

    elements.resultCount.textContent = countText + extraText;
}

// ===== Export Functionality =====
function exportResults() {
    if (state.filteredData.length === 0) {
        showToast('Não há dados para exportar', 'warning');
        return;
    }

    showLoading(true);

    try {
        // Create workbook with filtered data
        const exportData = [state.columns, ...state.filteredData];
        const worksheet = XLSX.utils.aoa_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Resultados');

        // Generate filename
        const timestamp = new Date().toISOString().slice(0, 10);
        const baseName = state.fileName.replace(/\.[^/.]+$/, '');
        const exportName = `${baseName}_export_${timestamp}.xlsx`;

        // Download file
        XLSX.writeFile(workbook, exportName);

        showToast(`Exportado: ${state.filteredData.length} itens`, 'success');

    } catch (error) {
        console.error('Export error:', error);
        showToast('Erro ao exportar. Tente novamente.', 'error');
    }

    showLoading(false);
}

// ===== Local Storage =====
function saveToLocalStorage() {
    try {
        const dataToSave = {
            columns: state.columns,
            data: state.data,
            fileName: state.fileName
        };
        localStorage.setItem('ewmSearchData', JSON.stringify(dataToSave));
    } catch (error) {
        console.warn('Could not save to localStorage:', error);
    }
}

function loadFromLocalStorage() {
    try {
        const savedData = localStorage.getItem('ewmSearchData');
        if (savedData) {
            const parsed = JSON.parse(savedData);
            state.columns = parsed.columns;
            state.data = parsed.data;
            state.filteredData = [...state.data];
            state.fileName = parsed.fileName;

            setupSearchUI();
            showSearchSection();

            showToast(`Dados restaurados: ${state.fileName}`, 'success');
        }
    } catch (error) {
        console.warn('Could not load from localStorage:', error);
    }
}

function clearLocalStorage() {
    try {
        localStorage.removeItem('ewmSearchData');
    } catch (error) {
        console.warn('Could not clear localStorage:', error);
    }
}

// ===== UI Helpers =====
function showSearchSection() {
    elements.uploadSection.classList.add('hidden');
    elements.searchSection.classList.remove('hidden');
    elements.searchInput.focus();
}

function showUploadSection() {
    elements.searchSection.classList.add('hidden');
    elements.uploadSection.classList.remove('hidden');
}

function resetApp() {
    state.data = [];
    state.columns = [];
    state.filteredData = [];
    state.fileName = '';

    clearLocalStorage();

    elements.fileInput.value = '';
    elements.searchInput.value = '';
    elements.fileInfo.textContent = '';

    showUploadSection();
}

function showLoading(show) {
    elements.loadingOverlay.classList.toggle('hidden', !show);
}

function showToast(message, type = 'success') {
    const toast = elements.toast;
    toast.querySelector('.toast-message').textContent = message;
    toast.className = `toast ${type}`;
    toast.classList.remove('hidden');

    setTimeout(() => {
        toast.classList.add('hidden');
    }, 3000);
}

// ===== Utility Functions =====
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ===== Service Worker Registration =====
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('sw.js')
            .then(reg => console.log('Service Worker registered'))
            .catch(err => console.warn('Service Worker registration failed:', err));
    });
}
