// State Management
let currentClass = 'SSC';
let seatNumbers = [];
let processedData = [];
let ratioChart, gradeChart;

// Private Azure Proxy URL (Relative path for Azure Static Web Apps)
const PROXY_URL = "/api/fetchResult";

// Constants for GSEB
const ENDPOINTS = {
    SSC: {
        index: 'https://www.gseb.org/',
        view: 'https://www.gseb.org/Result/SSCResultView'
    },
    HSC: {
        index: 'https://www.gseb.org/Result/Index',
        view: 'https://www.gseb.org/Result/ResultView'
    }
};

// UI Handlers
function selectClass(className) {
    currentClass = className;
    document.getElementById('btnSSC').classList.toggle('active', className === 'SSC');
    document.getElementById('btnHSC').classList.toggle('active', className === 'HSC');
    document.getElementById('hscOptions').classList.toggle('hidden', className === 'SSC');
}

async function startProcessing() {
    seatNumbers = [];
    processedData = [];
    
    // 1. Get seat numbers from Excel
    const fileInput = document.getElementById('fileUpload');
    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        const data = await readExcel(file);
        seatNumbers = seatNumbers.concat(data);
    }

    // 2. Get seat numbers from Manual Input
    const manualVal = document.getElementById('manualInput').value;
    if (manualVal) {
        const manualSeats = manualVal.split(/[,\s\n]+/).filter(s => s.trim().length > 0);
        seatNumbers = seatNumbers.concat(manualSeats);
    }

    if (seatNumbers.length === 0) {
        alert("Please provide at least one seat number.");
        return;
    }

    // Filter Commerce for HSC if selected
    if (currentClass === 'HSC' && document.getElementById('chkCommerceOnly').checked) {
        seatNumbers = seatNumbers.filter(s => s.toUpperCase().startsWith('C'));
    }

    showStatus(true);
    updateProgress(0, seatNumbers.length);

    // 3. Process each seat number
    for (let i = 0; i < seatNumbers.length; i++) {
        const seat = seatNumbers[i].trim();
        updateStatus(`Processing ${seat}...`);
        
        try {
            const result = await fetchResult(seat);
            if (result) {
                processedData.push(result);
            }
        } catch (error) {
            console.error(`Error processing ${seat}:`, error);
        }
        
        updateProgress(i + 1, seatNumbers.length);
    }

    showResults();
}

// Data Fetching Logic (The core "Scraper")
async function fetchResult(seatFull) {
    // The Azure Function now handles the entire 3-step process in one go
    try {
        const fetchUrl = `${PROXY_URL}?seat=${seatFull}&class=${currentClass}`;
        
        const response = await fetch(fetchUrl);
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(errorText || `Server Error: ${response.status}`);
        }
        
        const resultHtml = await response.text();
        
        if (resultHtml.includes("Proxy Error:")) throw new Error(resultHtml);

        console.log(`Received HTML for ${seatFull}, length: ${resultHtml.length}`);

        const data = parseResultHtml(resultHtml, seatFull);
        
        if (!data) {
            if (resultHtml.includes("SeatNo is required") || resultHtml.includes("Invalid Captcha")) {
                throw new Error("GSEB rejected the request (Invalid Captcha/Session).");
            }
            throw new Error("Data parsing failed. Result might not be available.");
        }
        return data;

    } catch (e) {
        showBackendError(`Error processing ${seatFull}: ${e.message}`);
        return null;
    }
}

function showBackendError(msg) {
    const statusText = document.getElementById('statusText');
    statusText.innerHTML += `<br><span style="color: #e74c3c; font-weight: bold;">❌ ${msg}</span>`;
    console.error(msg);
}

function parseResultHtml(html, seat) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    
    // Check if result found
    const nameLabel = Array.from(doc.querySelectorAll('td, span, div')).find(el => el.textContent.includes('Candidate Name'));
    if (!nameLabel) return null;

    const name = nameLabel.parentElement.nextElementSibling?.textContent.trim() || "Unknown";
    
    // Extracting Subjects and Marks (Dynamic depending on structure)
    const subjects = [];
    const rows = Array.from(doc.querySelectorAll('tr'));
    
    let totalMarks = 0;
    let resultStatus = "Pass";
    let isGrace = false;

    // Example parsing logic (adjust based on actual GSEB HTML structure provided in user context)
    // This is a simplified version for the demonstration
    const marksData = {
        seat: seat,
        name: name,
        subjects: [],
        total: 0,
        percentage: 0,
        pr: 0,
        status: "Pass",
        grade: "A1"
    };

    // Logic to find PR, Percentage, Grade, etc.
    const prElement = Array.from(doc.querySelectorAll('td')).find(el => el.textContent.includes('Percentile Rank'));
    marksData.pr = prElement?.nextElementSibling?.textContent.trim() || "0";

    const percElement = Array.from(doc.querySelectorAll('td')).find(el => el.textContent.includes('Percentage'));
    marksData.percentage = percElement?.nextElementSibling?.textContent.trim() || "0";

    const resultElement = Array.from(doc.querySelectorAll('td')).find(el => el.textContent.includes('Result'));
    marksData.status = resultElement?.nextElementSibling?.textContent.trim() || "Pass";

    // Grace logic: look for '*' in marks
    if (html.includes('*')) {
        isGrace = true;
    }
    marksData.isGrace = isGrace;

    // Simple Grade calculation if not found
    const p = parseFloat(marksData.percentage);
    if (p >= 91) marksData.grade = "A1";
    else if (p >= 81) marksData.grade = "A2";
    else if (p >= 71) marksData.grade = "B1";
    else if (p >= 61) marksData.grade = "B2";
    else if (p >= 51) marksData.grade = "C1";
    else if (p >= 41) marksData.grade = "C2";
    else if (p >= 33) marksData.grade = "D";
    else marksData.grade = "E1/F";

    return marksData;
}

// Utility: Read Excel
function readExcel(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // Flatten and find strings that look like seat numbers (e.g. B1234567)
            const seats = rows.flat().filter(cell => cell && typeof cell === 'string' && /^[A-Z]\d{6,7}$/i.test(cell.trim()));
            resolve(seats);
        };
        reader.readAsArrayBuffer(file);
    });
}

// UI: Progress & Results
function showStatus(show) {
    document.getElementById('statusSection').classList.toggle('hidden', !show);
    document.getElementById('resultsSection').classList.add('hidden');
}

function updateProgress(current, total) {
    const percent = Math.round((current / total) * 100);
    document.getElementById('progressBar').style.width = percent + '%';
    document.getElementById('statusText').textContent = `Processed ${current} of ${total} students...`;
}

function updateStatus(text) {
    document.getElementById('statusTitle').textContent = text;
}

function showResults() {
    document.getElementById('statusSection').classList.add('hidden');
    document.getElementById('resultsSection').classList.remove('hidden');

    updateSummary();
    renderTable();
    renderCharts();
}

function updateSummary() {
    const total = processedData.length;
    const pass = processedData.filter(d => d.status.toLowerCase().includes('pass')).length;
    const grace = processedData.filter(d => d.isGrace).length;
    const fail = total - pass;
    const ratio = total > 0 ? Math.round((pass / total) * 100) : 0;

    document.getElementById('statTotal').textContent = total;
    document.getElementById('statPass').textContent = pass;
    document.getElementById('statFail').textContent = fail;
    document.getElementById('statGrace').textContent = grace;
    document.getElementById('statRatio').textContent = ratio + '%';
}

function renderTable() {
    const head = document.getElementById('tableHead');
    const body = document.getElementById('tableBody');
    
    head.innerHTML = `
        <th>Seat No</th>
        <th>Name</th>
        <th>Percentage</th>
        <th>PR</th>
        <th>Grade</th>
        <th>Status</th>
    `;

    body.innerHTML = processedData.map(d => `
        <tr>
            <td>${d.seat}</td>
            <td>${d.name}</td>
            <td>${d.percentage}%</td>
            <td>${d.pr}</td>
            <td>${d.grade}</td>
            <td class="${d.status.toLowerCase().includes('pass') ? 'text-success' : 'text-danger'}">
                ${d.status} ${d.isGrace ? '(Grace)' : ''}
            </td>
        </tr>
    `).join('');
}

function renderCharts() {
    const ratioCtx = document.getElementById('ratioChart').getContext('2d');
    const gradeCtx = document.getElementById('gradeChart').getContext('2d');

    // Destroy existing charts if any
    if (ratioChart) ratioChart.destroy();
    if (gradeChart) gradeChart.destroy();

    const pass = processedData.filter(d => d.status.toLowerCase().includes('pass') && !d.isGrace).length;
    const grace = processedData.filter(d => d.isGrace).length;
    const fail = processedData.length - (pass + grace);

    ratioChart = new Chart(ratioCtx, {
        type: 'doughnut',
        data: {
            labels: ['Pass', 'Grace Pass', 'Fail'],
            datasets: [{
                data: [pass, grace, fail],
                backgroundColor: ['#27ae60', '#f39c12', '#e74c3c']
            }]
        }
    });

    const grades = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2', 'D', 'E1/F'];
    const gradeCounts = grades.map(g => processedData.filter(d => d.grade === g).length);

    gradeChart = new Chart(gradeCtx, {
        type: 'bar',
        data: {
            labels: grades,
            datasets: [{
                label: 'Students',
                data: gradeCounts,
                backgroundColor: '#3498db'
            }]
        },
        options: {
            scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
        }
    });
}

function exportToExcel() {
    const ws = XLSX.utils.json_to_sheet(processedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    XLSX.writeFile(wb, `GSEB_Results_${currentClass}_${new Date().getTime()}.xlsx`);
}