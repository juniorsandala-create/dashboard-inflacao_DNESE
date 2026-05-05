// ================= FIREBASE CONFIGURATION =================
const firebaseConfig = {
  apiKey: "AIzaSyAB8Qj3cWDQh1nqCYGSWPATktAORX2UpBc",
  authDomain: "dashboard-ipcn.firebaseapp.com",
  projectId: "dashboard-ipcn",
  storageBucket: "dashboard-ipcn.firebasestorage.app",
  messagingSenderId: "523269078636",
  appId: "1:523269078636:web:a9562dbda7e7c93c252249",
  measurementId: "G-SLNVMX9LWS"
};

// Global State
let globalData = { variacaoIpcn: [], varProdutos: [], provincias: [], sadc: [] };
let charts = {};
let currentUserRole = 'viewer';

// Initialize Firebase
let auth, db;
if (window.FB && firebaseConfig.apiKey !== "SUA_API_KEY_AQUI") {
    const app = FB.initializeApp(firebaseConfig);
    auth = FB.getAuth(app);
    db = FB.getFirestore(app);
    setupAuthListeners();
} else {
    console.warn("Firebase não configurado. O sistema funcionará em modo demonstração local.");
    // Fallback for local demo
    document.getElementById('login-form').addEventListener('submit', (e) => {
        e.preventDefault();
        enterDashboard('admin');
    });
}

// ================= AUTHENTICATION LOGIC =================
function setupAuthListeners() {
    FB.onAuthStateChanged(auth, async (user) => {
        if (user) {
            // Get user role from Firestore
            const docRef = FB.doc(db, "users", user.uid);
            const docSnap = await FB.getDoc(docRef);
            
            if (docSnap.exists()) {
                currentUserRole = docSnap.data().role || 'viewer';
            } else {
                // Se for o primeiro utilizador do sistema, torna-o Admin automaticamente
                const usersSnap = await FB.getDocs(FB.collection(db, "users"));
                currentUserRole = usersSnap.empty ? 'admin' : 'viewer';
                
                await FB.setDoc(docRef, {
                    email: user.email,
                    role: currentUserRole,
                    createdAt: new Date().toISOString()
                });
            }
            enterDashboard(currentUserRole, user.email);
        } else {
            exitDashboard();
        }
    });

    // Auth Form Toggle
    let isSignUp = false;
    document.getElementById('toggle-signup').addEventListener('click', (e) => {
        e.preventDefault();
        isSignUp = !isSignUp;
        document.getElementById('auth-title').textContent = isSignUp ? 'Pedir Acesso' : 'Acesso ao Sistema';
        document.getElementById('btn-auth-submit').textContent = isSignUp ? 'Criar Conta' : 'Entrar no Sistema';
        e.target.textContent = isSignUp ? 'Já tenho conta' : 'Pedir Acesso';
    });

    // Login/Sign Up Submission
    document.getElementById('login-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('email').value;
        const pass = document.getElementById('password').value;
        const errorEl = document.getElementById('auth-error');
        errorEl.textContent = '';

        try {
            if (isSignUp) {
                await FB.createUserWithEmailAndPassword(auth, email, pass);
            } else {
                await FB.signInWithEmailAndPassword(auth, email, pass);
            }
        } catch (error) {
            errorEl.textContent = "Erro: " + error.message;
        }
    });

    document.getElementById('btn-logout').addEventListener('click', () => FB.signOut(auth));
}

function enterDashboard(role, email) {
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('dashboard-screen').classList.remove('hidden');
    
    document.querySelector('.header-user span').textContent = email || 'Administrador';
    
    // Apply role-based visibility
    const adminElements = document.querySelectorAll('.admin-only');
    adminElements.forEach(el => {
        if (role === 'admin') el.classList.remove('hidden');
        else el.classList.add('hidden');
    });
}

function exitDashboard() {
    document.getElementById('login-screen').classList.remove('hidden');
    document.getElementById('dashboard-screen').classList.add('hidden');
}

// ================= USER MANAGEMENT =================
document.getElementById('btn-users-mgmt')?.addEventListener('click', async () => {
    document.getElementById('section-users').classList.remove('hidden');
    loadUsersList();
});

document.getElementById('btn-close-users')?.addEventListener('click', () => {
    document.getElementById('section-users').classList.add('hidden');
});

async function loadUsersList() {
    const listEl = document.getElementById('users-list');
    listEl.innerHTML = '<tr><td colspan="4">Carregando utilizadores...</td></tr>';
    
    try {
        const querySnapshot = await FB.getDocs(FB.collection(db, "users"));
        listEl.innerHTML = '';
        
        querySnapshot.forEach((doc) => {
            const u = doc.data();
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${u.email}</td>
                <td><span class="role-badge role-${u.role}">${u.role}</span></td>
                <td>${u.createdAt ? new Date(u.createdAt).toLocaleDateString() : 'N/A'}</td>
                <td>
                    ${u.role !== 'admin' ? `<button class="btn-action btn-promote" onclick="promoteUser('${doc.id}')">Promover</button>` : '---'}
                </td>
            `;
            listEl.appendChild(tr);
        });
    } catch (error) {
        listEl.innerHTML = '<tr><td colspan="4">Erro ao carregar lista.</td></tr>';
    }
}

window.promoteUser = async (uid) => {
    if (!confirm("Deseja promover este utilizador a Administrador?")) return;
    try {
        const userRef = FB.doc(db, "users", uid);
        await FB.updateDoc(userRef, { role: 'admin' });
        loadUsersList();
    } catch (error) {
        alert("Erro ao promover utilizador.");
    }
};

// ================= ORIGINAL DASHBOARD LOGIC (WRAPPED) =================

// Chart registration
Chart.register(ChartDataLabels);
Chart.defaults.plugins.datalabels.display = false;

const chartColors = ['#001F3F', '#D32F2F', '#800020', '#FF9800', '#4CAF50', '#9C27B0', '#00BCD4', '#795548', '#607D8B', '#E91E63', '#9E9E9E'];

document.getElementById('excel-upload').addEventListener('change', handleFileUpload);

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const statusMsg = document.getElementById('upload-status');
    statusMsg.textContent = 'Processando...';
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            parseExcelData(workbook);
            updateFilters();
            updateDashboard();
            statusMsg.textContent = '✓ Sucesso';
            statusMsg.className = 'status-msg status-success';
        } catch (error) {
            statusMsg.textContent = '✕ Erro: ' + error.message;
            statusMsg.className = 'status-msg status-error';
        }
    };
    reader.readAsArrayBuffer(file);
}

// ... rest of your parsing and rendering functions (parseExcelData, formatExcelDate, renderChartX, etc.)
// (Mantenha as funções originais abaixo deste ponto)

function parseExcelData(workbook) {
    const getRawData = (sheetName) => {
        const targetUpper = sheetName.toUpperCase();
        const key = Object.keys(workbook.Sheets).find(k => k.toUpperCase().includes(targetUpper));
        const sheet = workbook.Sheets[key || sheetName];
        if (!sheet) return [];
        return XLSX.utils.sheet_to_json(sheet, {header: 1, defval: null})
            .filter(r => r && r.length > 0 && r.some(c => c !== null && c !== ''));
    };

    // 1. VARIAÇÃO IPCN (Vertical: DATAS | INDICADOR | DADOS)
    let ipcnRaw = getRawData('VARIAÇÃO_IPCN');
    let colData = -1, colInd = -1, colVal = -1;
    
    // Find headers dynamically
    for(let r of ipcnRaw) {
        for(let i=0; i<r.length; i++) {
            if(r[i] && typeof r[i] === 'string') {
                let up = r[i].toUpperCase().trim();
                if(up.includes('DATA')) colData = i;
                if(up.includes('INDICADOR')) colInd = i;
                if(up.includes('DADOS')) colVal = i;
            }
        }
        if(colData !== -1 && colInd !== -1 && colVal !== -1) break;
    }
    
    // Fallback if headers not found by name
    if (colData === -1) colData = 1;
    if (colInd === -1) colInd = 2;
    if (colVal === -1) colVal = 3;

    let pivotMap = {};
    for(let r of ipcnRaw) {
        let dateVal = r[colData];
        let indVal = r[colInd];
        let valVal = r[colVal];
        
        if(dateVal === null || dateVal === undefined) continue;
        if(typeof dateVal === 'string' && dateVal.toUpperCase().includes('DATA')) continue; 
        
        let d = formatExcelDate(dateVal);
        if(!d) continue;
        
        if(!pivotMap[d]) pivotMap[d] = { DATA: d };
        if(indVal && typeof indVal === 'string') {
            pivotMap[d][indVal.trim()] = valVal;
        }
    }
    globalData.variacaoIpcn = Object.values(pivotMap);

    // 2. PROVINCIAS (Find columns: Name, Date, Value)
    let provRaw = getRawData('PROVINCIAS');
    let pColName = 1, pColDate = 2, pColVal = 3;
    
    // Find name column (first that is text and not a date)
    for(let r of provRaw.slice(0, 20)) {
        for(let i=0; i<r.length; i++) {
            if (r[i] && typeof r[i] === 'string' && !formatExcelDate(r[i]) && isNaN(parseExcelNumber(r[i]))) {
                pColName = i; break;
            }
        }
    }

    globalData.provincias = provRaw.map(r => {
        return {
            prov: String(r[pColName] || '').trim(),
            data: formatExcelDate(r[pColName + 1] || r[pColName - 1] || r[1]),
            valor: parseExcelNumber(r[pColName + 2] || r[pColName + 1] || r[3])
        };
    }).filter(x => x.prov && x.data && !isNaN(x.valor) && !x.prov.toUpperCase().includes('PROV') && isNaN(parseExcelNumber(x.prov)));

    // 3. SADC (Find columns: Name, Date, Value)
    let sadcRaw = getRawData('SADC');
    let sColName = 1;
    for(let r of sadcRaw.slice(0, 20)) {
        for(let i=0; i<r.length; i++) {
            if (r[i] && typeof r[i] === 'string' && !formatExcelDate(r[i]) && isNaN(parseExcelNumber(r[i]))) {
                sColName = i; break;
            }
        }
    }

    globalData.sadc = sadcRaw.map(r => {
        return {
            pais: String(r[sColName] || '').trim(),
            data: formatExcelDate(r[sColName + 1] || r[sColName - 1] || r[1]),
            valor: parseExcelNumber(r[sColName + 2] || r[sColName + 1] || r[3])
        };
    }).filter(x => x.pais && x.data && !isNaN(x.valor) && !x.pais.toUpperCase().includes('PAÍS') && !x.pais.toUpperCase().includes('PAIS') && isNaN(parseExcelNumber(x.pais)));

    // 4. VAR_DOS PRODUTOS (Can be horizontal or vertical)
    globalData.varProdutos = getRawData('VAR_DOS PRODUTOS');

    if (globalData.variacaoIpcn.length === 0) {
        throw new Error("Não foi possível ler os dados da folha VARIAÇÃO_IPCN.");
    }
}

// Helpers
function parseExcelNumber(val) {
    if (val === null || val === undefined || val === '') return NaN;
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
        let cleanVal = val.replace('%', '').replace(/\s/g, '');
        cleanVal = cleanVal.replace(',', '.'); 
        return parseFloat(cleanVal);
    }
    return NaN;
}

function formatExcelDate(val) {
    if (val === null || val === undefined || val === '') return null;
    if (typeof val === 'string') {
        let s = val.trim();
        if (s.toUpperCase().includes('DATA')) return null;
        if (/^[a-z]{3}\/\d{2}$/i.test(s)) return s.toLowerCase();
        if (!isNaN(Number(s)) && s.length > 3) {
            // Fall through to serial conversion
        } else {
            return null; 
        }
    }
    const num = Number(val);
    if (isNaN(num) || num < 20000) return null; // Likely not an Excel date
    const jsDate = new Date((num - 25569) * 86400 * 1000);
    const months = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
    return `${months[jsDate.getMonth()]}/${String(jsDate.getFullYear()).slice(-2)}`;
}

function formatPercentUI(val) {
    if (isNaN(val)) return '--%';
    return val.toFixed(2).replace('.', ',') + '%';
}

function populateSelect(selectId, items, selectedByDefault = true) {
    const sel = document.getElementById(selectId);
    if (!sel) return;
    sel.innerHTML = '';
    items.forEach((item, index) => {
        let opt = document.createElement('option');
        opt.value = item;
        opt.textContent = item;
        // For indicators, don't select all by default, just top 2
        if (selectId === 'filter-indicador') {
            opt.selected = item.toUpperCase().includes('MENSAL') || item.toUpperCase().includes('ALIMENTAÇÃO') || item.toUpperCase().includes('ALIMENTACAO') || index < 2;
        } else {
            opt.selected = selectedByDefault;
        }
        sel.appendChild(opt);
    });
}

function getSelectedOptions(selectId) {
    const select = document.getElementById(selectId);
    if (!select) return [];
    return Array.from(select.selectedOptions).map(opt => opt.value);
}

function updateFilters() {
    if (globalData.variacaoIpcn.length === 0) return;

    // 1. Periods (from variacaoIpcn)
    const periods = [...new Set(globalData.variacaoIpcn.map(row => row.DATA).filter(Boolean))];
    populateSelect('filter-period', periods);

    // 2. Classes (from variacaoIpcn keys)
    const excludeKeys = ['DATA', 'VAR', 'TAXA', 'IPCN', 'MENSAL', 'HOMÓLOGA', 'HOMOLOGA'];
    const classes = Object.keys(globalData.variacaoIpcn[0]).filter(k => 
        !excludeKeys.some(ex => k.toUpperCase().includes(ex))
    );
    populateSelect('filter-class', classes);

    // 3. Indicadores (for Evolução chart)
    const indicadores = Object.keys(globalData.variacaoIpcn[0]).filter(k => k !== 'DATA');
    populateSelect('filter-indicador', indicadores);

    // 4. Provincias
    const provincias = [...new Set(globalData.provincias.map(r => r.prov))];
    populateSelect('filter-province', provincias);

    // 5. SADC
    const sadc = [...new Set(globalData.sadc.map(r => r.pais))];
    populateSelect('filter-sadc', sadc);

    // Add events
    ['filter-period', 'filter-class', 'filter-province', 'filter-sadc', 'filter-indicador'].forEach(id => {
        const el = document.getElementById(id);
        if(el) {
            // Remove old listener to avoid duplicates
            el.replaceWith(el.cloneNode(true));
            document.getElementById(id).addEventListener('change', updateDashboard);
        }
    });
}

function updateDashboard() {
    if (globalData.variacaoIpcn.length === 0) return;

    const selectedPeriods = getSelectedOptions('filter-period');
    const selectedClasses = getSelectedOptions('filter-class');
    const selectedProvinces = getSelectedOptions('filter-province');
    const selectedSadc = getSelectedOptions('filter-sadc');
    const selectedIndicadores = getSelectedOptions('filter-indicador');

    if (selectedPeriods.length === 0) return;

    const filteredIpcn = globalData.variacaoIpcn.filter(row => selectedPeriods.includes(row.DATA));
    if (filteredIpcn.length === 0) return;

    updateKPIs(filteredIpcn);
    renderChartEvolucaoMensal(filteredIpcn, selectedIndicadores);
    renderChartHomologa(filteredIpcn);
    renderChartClasses(filteredIpcn, selectedClasses);
    renderChartProvincias(selectedProvinces, selectedPeriods);
    renderChartSadc(selectedSadc, selectedPeriods);
    renderChartProdutos(selectedPeriods);
    renderChartRegressao();
}

function updateKPIs(filteredIpcn) {
    const latestData = filteredIpcn[filteredIpcn.length - 1];
    
    const kHomologa = Object.keys(latestData).find(k => k.toUpperCase().includes('HOMÓLOGA') || k.toUpperCase().includes('HOMOLOGA'));
    const kMensal = Object.keys(latestData).find(k => k.toUpperCase().includes('MENSAL'));
    const kAlim = Object.keys(latestData).find(k => k.toUpperCase().includes('ALIMENTAÇÃO') || k.toUpperCase().includes('ALIMENTACAO'));
    const kSaude = Object.keys(latestData).find(k => k.toUpperCase().includes('SAÚDE') || k.toUpperCase().includes('SAUDE'));

    const valHomologa = parseExcelNumber(latestData[kHomologa]);
    const valMensal = parseExcelNumber(latestData[kMensal]);
    const valAlim = parseExcelNumber(latestData[kAlim]);
    const valSaude = parseExcelNumber(latestData[kSaude]);

    document.getElementById('kpi-homologa').textContent = formatPercentUI(valHomologa);
    document.getElementById('kpi-mensal').textContent = formatPercentUI(valMensal);
    document.getElementById('kpi-alimentacao').textContent = formatPercentUI(valAlim);
    document.getElementById('kpi-saude').textContent = formatPercentUI(valSaude);

    // SADC KPIs
    if (globalData.sadc.length > 0) {
        const latestPeriod = filteredIpcn[filteredIpcn.length - 1].DATA;
        const sadcData = globalData.sadc.filter(r => r.data === latestPeriod);
        
        if (sadcData.length > 0) {
            const angola = sadcData.find(r => r.pais.toUpperCase().includes('ANGOLA'));
            if (angola) {
                document.getElementById('kpi-sadc-angola').textContent = `🇦🇴 Angola: ${formatPercentUI(angola.valor)}`;
            }

            let maxVal = -Infinity, maxCountry = '';
            let minVal = Infinity, minCountry = '';

            sadcData.forEach(row => {
                let val = row.valor;
                if (!isNaN(val)) {
                    if (val > maxVal) { maxVal = val; maxCountry = row.pais; }
                    if (val < minVal) { minVal = val; minCountry = row.pais; }
                }
            });

            document.getElementById('kpi-sadc-high').textContent = `⬆️ ${maxCountry}: ${formatPercentUI(maxVal)}`;
            document.getElementById('kpi-sadc-low').textContent = `⬇️ ${minCountry}: ${formatPercentUI(minVal)}`;
        }
    }
}

function getChartInstance(chartId) {
    if (charts[chartId]) charts[chartId].destroy();
    return document.getElementById(chartId).getContext('2d');
}

function renderChartEvolucaoMensal(data, selectedIndicadores) {
    if(selectedIndicadores.length === 0) return;
    
    const labels = data.map(row => row.DATA);
    
    const datasets = selectedIndicadores.map((ind, index) => {
        const datasetData = data.map(row => {
            return parseExcelNumber(row[ind]);
        });
        
        return { 
            label: ind, 
            data: datasetData, 
            borderColor: chartColors[index % chartColors.length], 
            backgroundColor: chartColors[index % chartColors.length], 
            borderWidth: 2, 
            tension: 0.1, // Lower tension for more precise data representation
            pointRadius: 2,
            pointHoverRadius: 6
        };
    });

    charts['chartEvolucao'] = new Chart(getChartInstance('chartEvolucao'), {
        type: 'line',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false,
            layout: { padding: { left: 20, right: 20, top: 20, bottom: 5 } },
            plugins: { 
                legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8, font: { size: 11 } } },
                datalabels: {
                    display: function(context) {
                        return context.dataIndex === 0 || context.dataIndex === context.dataset.data.length - 1;
                    },
                    align: function(context) {
                        return context.dataIndex === 0 ? 'left' : 'right';
                    },
                    anchor: 'center',
                    offset: 6,
                    color: function(context) {
                        return context.dataset.borderColor;
                    },
                    font: { weight: 'bold', size: 11 },
                    formatter: function(value) {
                        return value.toFixed(2).replace('.', ',');
                    }
                }
            },
            scales: { 
                x: { grid: { display: false } },
                y: { display: false } 
            }
        }
    });
}

function renderChartHomologa(data) {
    const labels = data.map(row => row.DATA);
    const kHomologa = Object.keys(data[0]).find(k => k.toUpperCase().includes('HOMÓLOGA') || k.toUpperCase().includes('HOMOLOGA'));
    const dataHomologa = data.map(row => {
        return parseExcelNumber(row[kHomologa]);
    });

    charts['chartHomologa'] = new Chart(getChartInstance('chartHomologa'), {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{ 
                label: 'Taxa de Inflação Homóloga', 
                data: dataHomologa, 
                borderColor: chartColors[1], // red
                backgroundColor: 'transparent',
                borderWidth: 2, 
                tension: 0.2,
                pointRadius: 0,
                pointHoverRadius: 5
            }]
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false,
            layout: { padding: { top: 30, right: 20, left: 10 } },
            plugins: {
                legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8 } },
                datalabels: {
                    display: true,
                    align: 'top',
                    offset: 8,
                    color: chartColors[1],
                    font: { weight: 'bold', size: 11 },
                    formatter: function(value) {
                        return value.toFixed(2).replace('.', ',');
                    }
                }
            },
            scales: {
                x: { grid: { display: false } },
                y: { display: false, min: Math.min(...dataHomologa) - 2, max: Math.max(...dataHomologa) + 5 }
            }
        }
    });
}

function renderChartClasses(data, selectedClasses) {
    const latest = data[data.length - 1];
    const values = [], labels = [];
    
    selectedClasses.forEach(cls => {
        if (latest[cls] !== undefined) {
            let v = parseExcelNumber(latest[cls]);
            if (!isNaN(v)) {
                labels.push(cls.substring(0, 15) + (cls.length > 15 ? '...' : ''));
                values.push(v);
            }
        }
    });

    if (values.length === 0) return;

    charts['chartClasses'] = new Chart(getChartInstance('chartClasses'), {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{ data: values, backgroundColor: chartColors }]
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false, 
            plugins: { 
                legend: { position: 'right' },
                datalabels: {
                    display: true,
                    color: '#fff',
                    font: { weight: 'bold' },
                    formatter: (value) => value.toFixed(1).replace('.', ',') + '%'
                }
            } 
        }
    });
}

function renderChartProvincias(selectedProvinces, selectedPeriods) {
    if (globalData.provincias.length === 0 || selectedPeriods.length === 0) return;
    
    const latestPeriod = selectedPeriods[selectedPeriods.length - 1];
    
    let provData = globalData.provincias.filter(r => r.data === latestPeriod);
    if (selectedProvinces && selectedProvinces.length > 0) {
        provData = provData.filter(row => selectedProvinces.includes(row.prov));
    }

    const labels = provData.map(r => r.prov);
    const values = provData.map(r => r.valor);

    charts['chartProvincias'] = new Chart(getChartInstance('chartProvincias'), {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{ label: 'IPCN por Província', data: values, backgroundColor: chartColors[2], borderRadius: 4 }] // wine
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false,
            plugins: {
                datalabels: {
                    display: true,
                    anchor: 'end',
                    align: 'top',
                    formatter: (value) => value.toFixed(2).replace('.', ',')
                }
            }
        }
    });
}

function renderChartSadc(selectedSadc, selectedPeriods) {
    if (globalData.sadc.length === 0 || selectedPeriods.length === 0) return;
    
    const latestPeriod = selectedPeriods[selectedPeriods.length - 1];
    
    let sadcData = globalData.sadc.filter(r => r.data === latestPeriod);
    if (selectedSadc && selectedSadc.length > 0) {
        sadcData = sadcData.filter(row => selectedSadc.includes(row.pais));
    }

    const labels = sadcData.map(r => r.pais);
    const values = sadcData.map(r => r.valor);

    charts['chartSADC'] = new Chart(getChartInstance('chartSADC'), {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{ label: 'SADC', data: values, backgroundColor: chartColors[0], borderRadius: 4 }] // navy
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false,
            plugins: {
                datalabels: {
                    display: true,
                    anchor: 'end',
                    align: 'top',
                    formatter: (value) => value.toFixed(2).replace('.', ',')
                }
            }
        }
    });
}

function renderChartProdutos(selectedPeriods) {
    if (globalData.varProdutos.length === 0 || selectedPeriods.length === 0) return;
    
    const latestPeriod = selectedPeriods[selectedPeriods.length - 1];
    const rows = globalData.varProdutos;
    
    // 1. Find the header row to locate the column for the latest period
    let headerRowIndex = rows.findIndex(r => r.some(c => c && typeof c === 'string' && c.toUpperCase().includes('PRODUTO')));
    if (headerRowIndex === -1) headerRowIndex = 1;
    const headerRow = rows[headerRowIndex];

    // 2. Find the column index for the selected period
    let valColIndex = -1;
    for (let i = 0; i < headerRow.length; i++) {
        if (formatExcelDate(headerRow[i]) === latestPeriod) {
            valColIndex = i;
            break;
        }
    }

    // 3. Fallback: If not found in header (horizontal), check if it's a vertical layout (Column 3 has dates)
    let validData = [];
    if (valColIndex === -1) {
        // Vertical layout: [Classe, Produto, Data, Valor]
        validData = rows.map(r => ({
            produto: String(r[1] || '').trim(),
            data: formatExcelDate(r[2]),
            valor: parseExcelNumber(r[3])
        })).filter(x => x.data === latestPeriod && !isNaN(x.valor) && x.produto.length > 1);
    } else {
        // Horizontal layout: values are in valColIndex
        validData = rows.slice(headerRowIndex + 1).map(r => ({
            produto: String(r[1] || '').trim(),
            valor: parseExcelNumber(r[valColIndex])
        })).filter(x => x.produto && !isNaN(x.valor) && x.produto.length > 1);
    }

    if (validData.length === 0) return;
    
    validData.sort((a, b) => b.valor - a.valor);
    const top11 = validData.slice(0, 11);
    
    const labels = [], values = [];
    top11.forEach(r => {
        labels.push(r.produto.substring(0, 25));
        values.push(r.valor);
    });

    charts['chartProdutos'] = new Chart(getChartInstance('chartProdutos'), {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{ label: 'Maiores Altas', data: values, backgroundColor: chartColors[0] }]
        },
        options: { 
            indexAxis: 'y',
            responsive: true, 
            maintainAspectRatio: false,
            plugins: {
                datalabels: {
                    display: true,
                    anchor: 'end',
                    align: 'right',
                    formatter: (value) => value.toFixed(2).replace('.', ',')
                }
            }
        }
    });
}

function renderChartRegressao() {
    const scatterData = [];
    const kMensal = Object.keys(globalData.variacaoIpcn[0] || {}).find(k => k.toUpperCase().includes('MENSAL'));
    
    if (!kMensal) return;

    globalData.variacaoIpcn.forEach(row => {
        const date = row.DATA;
        const valNacional = parseExcelNumber(row[kMensal]);
        
        // Find Luanda for same date
        const provRow = globalData.provincias.find(p => p.data === date && p.prov.toUpperCase().includes('LUANDA'));
        if (provRow && !isNaN(valNacional)) {
            let valLuanda = provRow.valor;
            scatterData.push({ x: valNacional, y: valLuanda });
        }
    });

    if (scatterData.length === 0) return;

    charts['chartRegressao'] = new Chart(getChartInstance('chartRegressao'), {
        type: 'scatter',
        data: { 
            datasets: [{ 
                label: 'Nacional vs Luanda', 
                data: scatterData, 
                backgroundColor: chartColors[0],
                pointRadius: 4
            }] 
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false,
            scales: {
                x: { title: { display: true, text: 'Nacional (%)' } },
                y: { title: { display: true, text: 'Luanda (%)' } }
            }
        }
    });
}
