// Dados do arquivo Excel
let globalData = [];

// ==========================================
// CONFIGURAÇÃO DA PLANILHA DO GOOGLE SHEETS
// ==========================================
// 1. No Google Sheets, vá em: Arquivo > Compartilhar > Publicar na Web
// 2. Em "Vincular", mude de "Página da Web" para "Microsoft Excel (.xlsx)"
// 3. Clique em Publicar, copie o link e cole-o entre as aspas abaixo:
const GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1g-eShoOU4zCh7IzN3eDr7hQjKBkBvawakXqh4UCHhBk/gviz/tq?tqx=out:csv&gid=1405378764";

// Colunas exigidas pelo usuário
const REQUIRED_COLUMNS = [
    "Local de Inst.", "Modelo", "N° Série", "Cidade",
    "UF", "Endereço de Inst.", "N° Contrato", "Término da Garantia"
];

// Elementos do DOM
const searchInput = document.getElementById('searchInput');
const clearBtn = document.getElementById('clearBtn');
const resultsList = document.getElementById('resultsList');
const emptyState = document.getElementById('emptyState');
const statusText = document.getElementById('statusText');
const loader = document.getElementById('loader');
const fallbackContainer = document.getElementById('fallbackContainer');
const fileUpload = document.getElementById('fileUpload');

// Utilitário de Debounce para busca
function debounce(func, wait) {
    let timeout;
    return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// Formatação de data do Excel para JS (se for número serial do Excel ou Date Object)
function formatExcelDate(dateVal) {
    if (!dateVal) return '-';

    // Se for um objeto Date nativo (SheetJS com cellDates: true)
    if (dateVal instanceof Date) {
        const day = String(dateVal.getUTCDate()).padStart(2, '0');
        const month = String(dateVal.getUTCMonth() + 1).padStart(2, '0');
        const year = dateVal.getUTCFullYear();
        return `${day}/${month}/${year}`;
    }

    // Se já for string (ex: '15/05/2025'), retorna ela mesma
    if (typeof dateVal === 'string') {
        // Tenta garantir formato DD/MM/AAAA se a string vier com traços (ex: 2024-01-24)
        if (dateVal.includes('-')) {
            const parts = dateVal.split('T')[0].split('-');
            if (parts.length === 3 && parts[0].length === 4) {
                return `${parts[2]}/${parts[1]}/${parts[0]}`;
            }
        }
        return dateVal;
    }

    // Se for um número (data serial do Excel)
    if (typeof dateVal === 'number') {
        // Excel usa 1 de janeiro de 1900 como base, mas tem um bug no ano bissexto de 1900.
        // O offset em dias entre 01/01/1900 e 01/01/1970 é 25569.
        const excelEpochOffset = 25569;
        const msPerDay = 86400 * 1000;
        const date = new Date((dateVal - excelEpochOffset) * msPerDay);

        // Ajuste de fuso horário pra evitar que caia um dia antes (ex: 23:00)
        date.setMinutes(date.getMinutes() + date.getTimezoneOffset());

        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
    }

    return String(dateVal);
}

// Analisa a data para determinar o status da garantia
function getWarrantyStatus(dateStr) {
    try {
        if (!dateStr || dateStr === '-') return { class: '', text: 'Sem Data' };
        if (dateStr === 'Sem contrato') return { class: 'status-expired', text: 'Sem Contrato' };

        // Assume formato DD/MM/YYYY ou similar
        const parts = dateStr.includes('/') ? dateStr.split('/') : null;
        if (!parts || parts.length !== 3) return { class: 'status-active', text: 'Ativa' }; // Fallback

        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);

        const warrantyDate = new Date(year, month, day);
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const diffTime = warrantyDate - today;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        if (diffDays < 0) return { class: 'status-expired', text: 'Expirada' };
        if (diffDays <= 30) return { class: 'status-warning', text: 'Expira em breve' };
        return { class: 'status-active', text: 'Ativa' };
    } catch (e) {
        return { class: '', text: 'Data Inválida' };
    }
}

// Realçar (Highlight) o texto buscado
function highlightText(text, query) {
    if (!query || !text) return text || '-';
    const str = String(text);
    const regex = new RegExp(`(${query})`, 'gi');
    return str.replace(regex, '<mark>$1</mark>');
}

// Processa o workbook do SheetJS
function processWorkbook(workbook) {
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // A planilha oficial começa o cabeçalho na linha 3 ou possui colunas em branco.
    // Tentaremos achar a linha de cabeçalho correta buscando palavras chave.
    const rawDataFromSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    // Encontra The header row index
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(10, rawDataFromSheet.length); i++) {
        const rowStr = rawDataFromSheet[i].join(' ').toLowerCase();
        if (rowStr.includes('série') || rowStr.includes('serie') || rowStr.includes('modelo')) {
            headerRowIndex = i;
            break;
        }
    }

    // Converte com o header offset correto:
    const rawData = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex, defval: "" });

    // Normalizar chaves para facilitar busca (ignorar acentos nos nomes das colunas)
    globalData = rawData.map(row => {
        // Mapeia colunas encontradas para o que vamos usar, sendo flexivel com nomes exatos
        const getVal = (possibleKeys) => {
            for (let key of possibleKeys) {
                // Procura match parcial de key se existir
                const foundKey = Object.keys(row).find(k => k.toLowerCase().trim() === key.toLowerCase().trim() || k.replace(/[^a-zA-Z0-9]/g, '') === key.replace(/[^a-zA-Z0-9]/g, ''));
                if (foundKey) return row[foundKey];
            }
            return '-';
        };

            let valGaran = getVal(['Término da Garantia', 'Termino Garantia', 'Garantia', 'Fim Garantia']);
            let valManu = getVal(['Término da Manutenção', 'Termino da Manutencao', 'Termino Manutencao', 'Manutenção', 'Manutencao']);
            let valLoca = getVal(['Término da Locação', 'Término de Locação', 'Termino da Locacao', 'Termino de Locacao', 'Termino Locacao', 'Fim Locacao']);
            
            let dtVal = valGaran;
            let dtLabel = 'Término Garantia';
            
            if (dtVal === '-' || !dtVal) {
                dtVal = valManu;
                dtLabel = 'Término Manutenção';
            }
            if (dtVal === '-' || !dtVal) {
                dtVal = valLoca;
                dtLabel = 'Término Locação';
            }
            
            let finalDateStr = '';
            if (dtVal === '-' || !dtVal) {
                finalDateStr = 'Sem contrato';
                dtLabel = 'Status Contrato';
            } else {
                finalDateStr = formatExcelDate(dtVal);
            }

            // Verifica os status
            const isAtivo = (val) => {
                if (val === '-' || !val) return false;
                const status = getWarrantyStatus(formatExcelDate(val));
                return status.class === 'status-active' || status.class === 'status-warning';
            };
            
            let isGarantiaActive = isAtivo(valGaran);
            let isLocacaoActive = isAtivo(valLoca);
            let isManutencaoActive = isAtivo(valManu);
            
            let originalTroca = String(getVal(['Contempla troca de peças', 'Contempla Troca de Peças', 'Troca de peças', 'Troca de pecas', 'Peças', 'Pecas', 'Troca Peças'])).trim().toLowerCase();
            let trocaPecasValor;
            
            if (originalTroca === 'sim' || originalTroca === 'não' || originalTroca === 'nao') {
                trocaPecasValor = originalTroca.charAt(0).toUpperCase() + originalTroca.slice(1);
                if (trocaPecasValor === 'Nao') trocaPecasValor = 'Não';
            } else {
                if (isGarantiaActive || isLocacaoActive || isManutencaoActive) {
                    trocaPecasValor = 'Sim';
                } else {
                    trocaPecasValor = 'Não';
                }
            }

            return {
                local: getVal(['Local de Inst.', 'Local de Instalacao', 'Local Inst', 'Local']),
                modelo: getVal(['Modelo', 'Equipamento']),
                serie: getVal(['N° Série', 'N Série', 'Numero Serie', 'Num Serie', 'Série', 'Serie']),
                cidade: getVal(['Cidade']),
                uf: getVal(['UF', 'Estado']),
                endereco: getVal(['Endereço de Inst.', 'Endereço', 'Endereco de Inst', 'Endereco']),
                contrato: getVal(['N° Contrato', 'N Contrato', 'Contrato']),
                garantiaLabel: dtLabel,
                garantia: finalDateStr,
                trocaPecas: trocaPecasValor
            };
        });
}

// Processar Arquivo Excel
async function loadExcelData() {
    try {
        loader.classList.remove('hidden');
        statusText.textContent = "Carregando base de dados...";
        fallbackContainer.classList.add('hidden');

        // Determina a URL: usa a do Google Sheets se configurada, senão tenta o arquivo local
        const urlToFetch = GOOGLE_SHEET_URL.includes("COLE_AQUI_O_LINK_GERADO_PELO_GOOGLE_SHEETS")
            ? './CONTRATOS 2025 - REV03 - Online.xlsx'
            : GOOGLE_SHEET_URL;

        // Faz o download do arquivo
        const response = await fetch(urlToFetch);
        if (!response.ok) throw new Error('Falha ao carregar a planilha');

        const arrayBuffer = await response.arrayBuffer();

        // Leitura usando SheetJS
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

        processWorkbook(workbook);

        loader.classList.add('hidden');
        statusText.textContent = `${globalData.length} registros carregados.`;

        // Remove text after 3 seconds
        setTimeout(() => {
            statusText.textContent = "";
        }, 3000);

        // Habilita a busca
        searchInput.disabled = false;
        searchInput.focus();

    } catch (error) {
        console.error("Erro ao carregar Excel:", error);
        loader.classList.add('hidden');
        statusText.textContent = "Erro: Restrição de navegador (abra via link ou use o botão abaixo para selecionar o arquivo localmente).";
        statusText.style.color = "var(--warning-color)";
        fallbackContainer.classList.remove('hidden');
    }
}

// Handlers de Upload Manual de Arquivo
fileUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    loader.classList.remove('hidden');
    statusText.textContent = "Processando arquivo...";
    statusText.style.color = "var(--text-secondary)";
    fallbackContainer.classList.add('hidden');

    const reader = new FileReader();
    reader.onload = function (evt) {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });

            processWorkbook(workbook);

            loader.classList.add('hidden');
            statusText.textContent = `${globalData.length} registros carregados com sucesso!`;
            statusText.style.color = "var(--success-color)";

            setTimeout(() => {
                statusText.textContent = "";
            }, 3500);

            searchInput.disabled = false;
            searchInput.focus();

        } catch (err) {
            console.error(err);
            loader.classList.add('hidden');
            statusText.textContent = "Erro ao processar o arquivo selecionado.";
            statusText.style.color = "var(--danger-color)";
            fallbackContainer.classList.remove('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
});

// Renderizar os cards
function renderResults(results, query) {
    if (results.length === 0) {
        emptyState.classList.add('hidden');
        resultsList.innerHTML = `
            <div class="no-results">
                Nenhum equipamento encontrado com a série "<strong>${query}</strong>".
            </div>
        `;
        return;
    }

    emptyState.classList.add('hidden');

    // Limitando a 50 resultados para performance do DOM
    const dataToRender = results.slice(0, 50);

    const htmlCards = dataToRender.map(item => {
        const warrantyStatus = getWarrantyStatus(item.garantia);

        return `
        <div class="result-card">
            <div class="card-header">
                <div class="serial-number">
                    <svg class="serial-icon" xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path>
                        <line x1="9" y1="9" x2="15" y2="9"></line>
                        <line x1="9" y1="13" x2="15" y2="13"></line>
                        <path d="M12 17h.01"></path>
                    </svg>
                    ${highlightText(item.serie, query)}
                </div>
                <div class="model-badge">${item.modelo || 'N/A'}</div>
            </div>
            
            <div class="card-body">
                <div class="data-group">
                    <span class="data-label">Local de Inst.</span>
                    <span class="data-value">${item.local}</span>
                </div>
                
                <div class="data-group">
                    <span class="data-label">N° Contrato</span>
                    <span class="data-value">${item.contrato}</span>
                </div>
                
                <div class="data-group address-group">
                    <span class="data-label">Endereço</span>
                    <span class="data-value">${item.endereco} - ${item.cidade}/${item.uf}</span>
                </div>
                
                <div class="data-group">
                    <span class="data-label">${item.garantiaLabel || 'Status de Contrato'}</span>
                    <span class="data-value warranty-status ${item.garantia === 'Sem contrato' ? '' : warrantyStatus.class}">
                        ${item.garantia !== 'Sem contrato' ? `<span class="status-dot"></span>` : ''}
                        ${item.garantia !== 'Sem contrato' ? `<span style="opacity: 0.8; font-size: 0.8em; margin-right: 4px;">(${warrantyStatus.text})</span>` : ''}
                        ${item.garantia}
                    </span>
                </div>
                
                <div class="data-group">
                    <span class="data-label">Contempla Peças</span>
                    <span class="data-value" style="font-weight: 600;">${item.trocaPecas !== '-' ? item.trocaPecas : 'N/A'}</span>
                </div>
            </div>
        </div>
        `;
    }).join('');

    resultsList.innerHTML = htmlCards;

    // Indicador caso haja mais resultados não exibidos
    if (results.length > 50) {
        resultsList.innerHTML += `
            <div style="text-align: center; font-size: 0.8rem; color: var(--text-secondary); margin-top: 1rem;">
                Exibindo 50 de ${results.length} resultados. Refine sua busca.
            </div>
        `;
    }
}

// Executar Busca
function handleSearch() {
    const query = searchInput.value.trim().toLowerCase();

    if (query.length > 0) {
        clearBtn.classList.remove('hidden');
    } else {
        clearBtn.classList.add('hidden');
        resultsList.innerHTML = '';
        emptyState.classList.remove('hidden');
        return;
    }

    // Busca apenas se tiver pelomenos 2 caracteres
    if (query.length >= 2) {
        const results = globalData.filter(item => {
            const serieStr = String(item.serie).toLowerCase();
            return serieStr.includes(query);
        });
        renderResults(results, query);
    }
}

// Event Listeners
searchInput.addEventListener('input', debounce(handleSearch, 300));

clearBtn.addEventListener('click', () => {
    searchInput.value = '';
    clearBtn.classList.add('hidden');
    resultsList.innerHTML = '';
    emptyState.classList.remove('hidden');
    searchInput.focus();
});

// Tema Escuro/Claro
const themeToggleBtn = document.getElementById('themeToggleBtn');
const sunIcon = document.querySelector('.sun-icon');
const moonIcon = document.querySelector('.moon-icon');

function initTheme() {
    const savedTheme = localStorage.getItem('theme');
    
    // Default system is light now. Set to dark only if explicitly requested.
    const isDark = savedTheme === 'dark';
    
    if (isDark) {
        document.documentElement.removeAttribute('data-theme');
        sunIcon.style.display = 'block';
        moonIcon.style.display = 'none';
    } else {
        document.documentElement.setAttribute('data-theme', 'light');
        sunIcon.style.display = 'none';
        moonIcon.style.display = 'block';
    }
}

themeToggleBtn.addEventListener('click', () => {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    
    if (currentTheme === 'light') {
        document.documentElement.removeAttribute('data-theme');
        localStorage.setItem('theme', 'dark');
        sunIcon.style.display = 'block';
        moonIcon.style.display = 'none';
    } else {
        document.documentElement.setAttribute('data-theme', 'light');
        localStorage.setItem('theme', 'light');
        sunIcon.style.display = 'none';
        moonIcon.style.display = 'block';
    }
});

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    // Desabilitar input até carregar os dados
    searchInput.disabled = true;
    loadExcelData();
});
