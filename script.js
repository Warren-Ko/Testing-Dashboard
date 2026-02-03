const SHEET_ID = '1wKFnVUTAZT5vMSf7e0Vl4nec99RaLOTWbal_6WtMVHA';
const API_KEY = 'AIzaSyCBFge35ER_x3Kf5C377e1O5dp82uVj6-U';
const SHEET_RANGE_MAIN = 'A1:Z1000'; // Main Data (First Sheet)
const SHEET_RANGE_ORDER = 'status order!A:C'; // Status Order Tab
const API_URL_MAIN = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_RANGE_MAIN}?key=${API_KEY}`;
const API_URL_ORDER = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_RANGE_ORDER}?key=${API_KEY}`;

let STAGE_ORDER = ['New', 'Recruiter Review', 'Manager Review', 'Interview', 'Offer']; // Default Fallback

// Color Palette for Charts
const CHART_COLORS = [
    'rgba(59, 130, 246, 0.8)',   // Blue
    'rgba(16, 185, 129, 0.8)',  // Green
    'rgba(245, 158, 11, 0.8)',   // Orange
    'rgba(139, 92, 246, 0.8)',   // Purple
    'rgba(239, 68, 68, 0.8)',    // Red
    'rgba(14, 165, 233, 0.8)',   // Sky
    'rgba(236, 72, 153, 0.8)',   // Pink
    'rgba(20, 184, 166, 0.8)'    // Teal
];

const CHART_COLORS_BORDER = [
    'rgba(59, 130, 246, 1)',
    'rgba(16, 185, 129, 1)',
    'rgba(245, 158, 11, 1)',
    'rgba(139, 92, 246, 1)',
    'rgba(239, 68, 68, 1)',
    'rgba(14, 165, 233, 1)',
    'rgba(236, 72, 153, 1)',
    'rgba(20, 184, 166, 1)'
];

let roleChartInstance = null;
let funnelChartInstance = null;
let distributionChartInstance = null;

// Global Data Store
let allData = [];
let uniqueRolesList = [];
let availableDates = [];

document.addEventListener('DOMContentLoaded', () => {
    fetchData();
    // Auto refresh every 60 seconds
    setInterval(fetchData, 60000);
});

async function fetchData() {
    const btn = document.querySelector('.refresh-btn');
    const icon = btn.querySelector('i');

    // UI Loading State
    icon.classList.add('fa-spin');
    btn.disabled = true;
    document.body.style.cursor = 'wait';

    try {
        // Fetch BOTH Main Data and Status Order
        const [mainResponse, orderResponse] = await Promise.all([
            fetch(API_URL_MAIN),
            fetch(API_URL_ORDER)
        ]);

        const mainData = await mainResponse.json();
        const orderData = await orderResponse.json();

        if (mainData.error) throw new Error(`Main Data: ${mainData.error.message}`);
        // Status Order is optional but preferred, don't crash if missing, just log
        if (orderData.error) console.warn(`Status Order API Error: ${orderData.error.message}`);

        // 1. Process Order (if available)
        if (orderData.values && orderData.values.length > 1) { // >1 to skip header
            processStatusOrder(orderData.values);
        }

        // 2. Process Main Data
        if (mainData.values && mainData.values.length > 0) {
            processData(mainData.values);
        } else {
            console.error('No data found in sheet');
            alert('Connected to Sheet, but found no data.');
        }

    } catch (error) {
        console.error('Error fetching data:', error);
        alert(`Failed to fetch data: ${error.message}\n\nCheck console for details.`);
    } finally {
        // UI Reset State
        icon.classList.remove('fa-spin');
        btn.disabled = false;
        document.body.style.cursor = 'default';
    }
}

function processStatusOrder(rows) {
    // Determine Column Mapping
    // We expect one column to be the Name (String) and one to be the Index (Number)
    if (!rows || rows.length < 2) return;

    const firstRow = rows[1]; // Use first data row (skip header 0)
    let nameCol = 0;
    let indexCol = 1;

    // Heuristic: If Col 0 is a number and Col 1 is NOT, then Col 1 is likely the Name
    // (using flexible parsing for potentially string-encoded numbers)
    const col0IsNum = !isNaN(parseFloat(firstRow[0])) && isFinite(firstRow[0]);
    const col1IsNum = !isNaN(parseFloat(firstRow[1])) && isFinite(firstRow[1]);

    if (col0IsNum && !col1IsNum) {
        nameCol = 1;
        indexCol = 0;
    }

    // Process
    const validRows = rows.slice(1).filter(r => r[nameCol] && r[indexCol]);

    // Sort by Index
    validRows.sort((a, b) => parseInt(a[indexCol]) - parseInt(b[indexCol]));

    // Update Global STAGE_ORDER
    const newOrder = validRows.map(r => r[nameCol]);

    if (newOrder.length > 0) {
        STAGE_ORDER = newOrder;
        console.log('Updated Stage Order:', STAGE_ORDER);
    }
}

function processData(rows) {
    // Row 0 is headers: Index, Date, Role, Status, Candidates, Candidate Delta
    const rawData = rows.slice(1);

    // Parse and Store Data
    allData = rawData.map(row => {
        const dateStr = row[1]; // Date column
        if (!dateStr) return null;

        const [day, month, year] = dateStr.split('/');
        return {
            date: new Date(`${year}-${month}-${day}`),
            originalDate: dateStr,
            role: row[2],
            status: row[3],
            count: parseInt(row[4]) || 0,
            delta: parseInt(row[5]) || 0,
            row: row
        };
    }).filter(item => item !== null && !isNaN(item.date));

    if (allData.length === 0) return;

    // Sort descending by date for processing
    allData.sort((a, b) => b.date - a.date);

    // Extract Unique Roles and Dates
    uniqueRolesList = [...new Set(allData.map(d => d.role))].sort();
    // Extract unique dates as strings, sorted new to old
    availableDates = [...new Set(allData.map(d => d.originalDate))];

    // Sort Date Strings properly (Oldest to Newest for Chart)
    availableDates.sort((a, b) => {
        const [d1, m1, y1] = a.split('/');
        const [d2, m2, y2] = b.split('/');
        return new Date(`${y1}-${m1}-${d1}`) - new Date(`${y2}-${m2}-${d2}`);
    });

    // Initialize All Controls
    initControls();

    // Initial Dashboard Render
    updateDashboard();
}

function initControls() {
    // 1. Global Role Select
    const roleSelect = document.getElementById('roleSelect');
    const currentGlobalRole = roleSelect.value;
    roleSelect.innerHTML = '<option value="All">All Roles</option>';
    uniqueRolesList.forEach(role => {
        const option = document.createElement('option');
        option.value = role;
        option.textContent = role;
        roleSelect.appendChild(option);
    });
    // Restore selection if valid
    if (uniqueRolesList.includes(currentGlobalRole)) roleSelect.value = currentGlobalRole;

    // 2. Role Chart Date Filter (NEW)
    // Same logic as Funnel: Newest dates first
    const datesDesc = [...availableDates].reverse();
    const roleDateSelect = document.getElementById('roleChartDateFilter');
    const currentRoleDate = roleDateSelect.value;

    roleDateSelect.innerHTML = '';
    // Add "Latest" option
    const latestOptRole = document.createElement('option');
    latestOptRole.value = 'Latest';
    latestOptRole.textContent = `Latest (${datesDesc[0]})`;
    roleDateSelect.appendChild(latestOptRole);

    datesDesc.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        roleDateSelect.appendChild(option);
    });
    if (datesDesc.includes(currentRoleDate) || currentRoleDate === 'Latest') {
        roleDateSelect.value = currentRoleDate;
    }

    // 3. Funnel Date Filter
    // ... (reuse datesDesc)
    const funnelDateSelect = document.getElementById('funnelDateFilter');
    const currentFunnelDate = funnelDateSelect.value;

    funnelDateSelect.innerHTML = '';
    const latestOptFunnel = document.createElement('option');
    latestOptFunnel.value = 'Latest';
    latestOptFunnel.textContent = `Latest (${datesDesc[0]})`;
    funnelDateSelect.appendChild(latestOptFunnel);

    datesDesc.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        funnelDateSelect.appendChild(option);
    });
    if (datesDesc.includes(currentFunnelDate) || currentFunnelDate === 'Latest') {
        funnelDateSelect.value = currentFunnelDate;
    }

    // 4. Status Toggles (Main Chart)
    initStatusToggles();

    // Attach Listeners
    roleSelect.onchange = () => updateDashboard();
    roleDateSelect.onchange = () => updateDashboard(); // New Listener
    funnelDateSelect.onchange = () => updateDashboard();
}

function initStatusToggles() {
    const container = document.getElementById('statusToggles');
    container.innerHTML = ''; // Clear existing

    STAGE_ORDER.forEach(status => {
        const label = document.createElement('label');
        label.className = 'toggle-label';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = status;
        checkbox.checked = true; // Default all checked

        checkbox.addEventListener('change', () => {
            updateDashboard();
        });

        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(status));
        container.appendChild(label);
    });
}

function getSelectedStatuses() {
    const checkboxes = document.querySelectorAll('#statusToggles input[type="checkbox"]');
    const selected = [];
    checkboxes.forEach(cb => {
        if (cb.checked) {
            selected.push(cb.value);
        }
    });
    return selected;
}

function updateDashboard() {
    // 0. Get Current Control Values
    const globalRole = document.getElementById('roleSelect').value;
    // const localRoleFilter = document.getElementById('roleChartSpecificFilter').value; // REMOVED
    let roleChartDate = document.getElementById('roleChartDateFilter').value; // NEW
    let funnelDate = document.getElementById('funnelDateFilter').value;

    const latestDateStr = availableDates[availableDates.length - 1]; // Last in sorted array

    // Resolve "Latest" dates
    if (funnelDate === 'Latest') funnelDate = latestDateStr;
    if (roleChartDate === 'Latest') roleChartDate = latestDateStr;


    // --- KPI SECTION ---
    // Source: Latest Date, Filtered by Global Role
    const kpiData = allData.filter(d =>
        d.originalDate === latestDateStr &&
        (globalRole === 'All' || d.role === globalRole)
    );

    let totalCandidates = 0;
    const uniqueRoles = new Set();
    let interviewsActiveCount = 0; // "Total Interviews" / "Interview Set"
    let interviewedCount = 0;      // "Total Interviewed"

    kpiData.forEach(item => {
        totalCandidates += item.count;
        uniqueRoles.add(item.role);

        // Split Logic:
        if (item.status.includes('Interviewed')) {
            interviewedCount += item.count;
        } else if (item.status.includes('Interview')) {
            interviewsActiveCount += item.count;
        }
    });

    animateValue(document.getElementById('kpi-total'), totalCandidates);
    animateValue(document.getElementById('kpi-roles'), uniqueRoles.size);
    animateValue(document.getElementById('kpi-interviews'), interviewsActiveCount);
    animateValue(document.getElementById('kpi-total-interviewed'), interviewedCount);

    document.getElementById('last-updated').innerHTML = `<i class="fa-regular fa-clock"></i> Data from: ${latestDateStr}`;


    // --- ROLE CHART SECTION (STACKED) ---
    // Source: Specific Role Date, Global Role Filter applied? 

    const roleChartRaw = allData.filter(d =>
        d.originalDate === roleChartDate &&
        (globalRole === 'All' || d.role === globalRole)
    );

    // Identify "Interview" related statuses dynamically
    // Look for any status string containing "Interview"
    const interviewStatuses = STAGE_ORDER.filter(s => s.includes('Interview'));

    // Structure: { [RoleName]: { [Status1]: count, [Status2]: count, ... } }
    const roleStackedData = {};

    // Initialize roles to ensure 0 counts are handled if needed, or build dynamic
    roleChartRaw.forEach(item => {
        // Only care about Interview statuses
        if (interviewStatuses.includes(item.status)) {
            if (!roleStackedData[item.role]) roleStackedData[item.role] = {};
            roleStackedData[item.role][item.status] = (roleStackedData[item.role][item.status] || 0) + item.count;
        }
    });

    const roleChartCard = document.getElementById('roleChartCard');
    roleChartCard.classList.remove('hidden');
    renderRoleChart(roleStackedData, interviewStatuses);


    // --- FUNNEL CHART SECTION ---
    // Source: Specific Funnel Date, Filtered by Global Role
    const funnelDataRaw = allData.filter(d =>
        d.originalDate === funnelDate &&
        (globalRole === 'All' || d.role === globalRole)
    );

    const funnelCounts = {};
    funnelDataRaw.forEach(item => {
        funnelCounts[item.status] = (funnelCounts[item.status] || 0) + item.count;
    });
    renderFunnelChart(funnelCounts);


    // --- DISTRIBUTION CHART SECTION ---
    const selectedStatuses = getSelectedStatuses();
    renderDistributionChart(globalRole, selectedStatuses);

    // --- INSIGHTS SECTION ---
    generateInsights(kpiData, globalRole);
}

function generateInsights(data, globalRole) {
    if (!data || data.length === 0) return;

    // 1. Top Role (Volume)
    const roleCounts = {};
    data.forEach(d => {
        if (d.status.includes('Interview')) {
            roleCounts[d.role] = (roleCounts[d.role] || 0) + d.count;
        }
    });
    const sortedRoles = Object.entries(roleCounts).sort((a, b) => b[1] - a[1]);
    const topRole = sortedRoles.length > 0 ? sortedRoles[0] : ['None', 0];

    document.getElementById('insight-top-role').textContent = topRole[0];
    document.getElementById('insight-top-role-desc').textContent = `${topRole[1]} candidates in interview`;


    // 2. Best Conversion Rates (NEW & REVIEW separately)
    const roleNewCounts = {};
    const roleReviewCounts = {};
    const roleInterviewedCounts = {};

    // Use latest company-wide data for Best Role insights
    const latestDateStr = availableDates[availableDates.length - 1];
    const latestDataAll = allData.filter(d => d.originalDate === latestDateStr);

    latestDataAll.forEach(d => {
        if (d.status === 'New') roleNewCounts[d.role] = (roleNewCounts[d.role] || 0) + d.count;
        if (d.status === 'Recruiter Review') roleReviewCounts[d.role] = (roleReviewCounts[d.role] || 0) + d.count;
        if (d.status.includes('Interviewed')) {
            roleInterviewedCounts[d.role] = (roleInterviewedCounts[d.role] || 0) + d.count;
        }
    });

    // Best from NEW
    let bestRoleNew = 'None';
    let bestRateNew = 0;
    // Best from REVIEW
    let bestRoleReview = 'None';
    let bestRateReview = 0;

    Object.keys(roleInterviewedCounts).forEach(role => {
        const interviewed = roleInterviewedCounts[role] || 0;

        // From NEW baseline to Interviewed
        const newCount = roleNewCounts[role] || 0;
        if (newCount > 0) {
            const rate = Math.min(interviewed / newCount, 1.0);
            if (rate > bestRateNew) {
                bestRateNew = rate;
                bestRoleNew = role;
            }
        }

        // From REVIEW baseline to Interviewed
        const reviewCount = roleReviewCounts[role] || 0;
        if (reviewCount > 0) {
            const rate = Math.min(interviewed / reviewCount, 1.0);
            if (rate > bestRateReview) {
                bestRateReview = rate;
                bestRoleReview = role;
            }
        }
    });

    // Update Card 2: Best Conv (New)
    document.getElementById('insight-conv-new').textContent = `${(bestRateNew * 100).toFixed(0)}%`;
    document.getElementById('insight-conv-new-desc').textContent = bestRoleNew;

    // Update Card 3: Best Conv (Review)
    document.getElementById('insight-conv-review').textContent = `${(bestRateReview * 100).toFixed(0)}%`;
    document.getElementById('insight-conv-review-desc').textContent = bestRoleReview;


    // 3. Weekly Pipeline Growth
    const previousDateStr = availableDates.length >= 2 ? availableDates[availableDates.length - 2] : null;

    if (previousDateStr) {
        const prevData = allData.filter(d => d.originalDate === previousDateStr && (globalRole === 'All' || d.role === globalRole));
        const prevTotal = prevData.reduce((sum, d) => sum + d.count, 0);
        const currentTotal = data.reduce((sum, d) => sum + d.count, 0);

        const growth = currentTotal - prevTotal;
        const growthPercent = prevTotal > 0 ? ((growth / prevTotal) * 100).toFixed(1) : 0;
        const arrow = growth >= 0 ? '↑' : '↓';

        document.getElementById('insight-velocity').textContent = `${arrow} ${Math.abs(growth)}`;
        document.getElementById('insight-velocity-desc').textContent = `${growthPercent}% ${growth >= 0 ? 'increase' : 'decrease'} vs last week`;

        const card4Title = document.querySelector('#insight-velocity').parentElement.querySelector('h3');
        if (card4Title) card4Title.textContent = 'Weekly Growth';
    } else {
        document.getElementById('insight-velocity').textContent = '--';
        document.getElementById('insight-velocity-desc').textContent = 'Initial data snapshot';
    }
}

function renderFunnelChart(dataObj) {
    const ctx = document.getElementById('funnelChart').getContext('2d');

    const rawCounts = STAGE_ORDER.map(stage => dataObj[stage] || 0);
    const cumulativeCounts = [];

    let runningTotal = 0;
    for (let i = rawCounts.length - 1; i >= 0; i--) {
        runningTotal += rawCounts[i];
        cumulativeCounts[i] = runningTotal;
    }

    const conversionRates = cumulativeCounts.map((count, index) => {
        if (index === 0) return '100% (Baseline)';
        const prev = cumulativeCounts[index - 1];
        if (prev === 0) return '0%';
        return `${((count / prev) * 100).toFixed(1)}% of previous stage`;
    });

    if (funnelChartInstance) {
        funnelChartInstance.destroy();
    }

    funnelChartInstance = new Chart(ctx, {
        type: 'funnel',
        data: {
            labels: STAGE_ORDER,
            datasets: [{
                label: 'Cumulative Candidates',
                data: cumulativeCounts,
                backgroundColor: [
                    'rgba(59, 130, 246, 0.8)',   // New (Blue)
                    'rgba(245, 158, 11, 0.8)',   // Recruiter (Orange)
                    'rgba(16, 185, 129, 0.8)',   // Manager (Green)
                    'rgba(139, 92, 246, 0.8)',   // Interview (Purple)
                    'rgba(236, 72, 153, 0.8)'    // Offer (Pink)
                ],
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        afterLabel: function (context) {
                            return `Conversion: ${conversionRates[context.dataIndex]}`;
                        }
                    },
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleColor: '#f8fafc',
                    bodyColor: '#cbd5e1',
                    padding: 10,
                    cornerRadius: 8
                }
            }
        }
    });
}

function renderDistributionChart(roleFilter, selectedStatuses) {
    const ctx = document.getElementById('distributionChart').getContext('2d');

    // 1. Get Selected Statuses from Checkboxes
    const selectedStatusesSet = new Set(selectedStatuses);

    // 2. Filter data by Role
    const trendData = roleFilter === 'All'
        ? allData
        : allData.filter(d => d.role === roleFilter);

    // 3. Determine unique statuses to show (Order enforced, Selection respected)
    // Only show statuses that are both in STAGE_ORDER and Checked
    const uniqueStatuses = STAGE_ORDER.filter(stage => selectedStatusesSet.has(stage));

    // 4. Build Datasets
    const datasets = uniqueStatuses.map((status, index) => {
        const dataPoints = availableDates.map(date => {
            return trendData
                .filter(d => d.originalDate === date && d.status === status)
                .reduce((sum, current) => sum + current.count, 0);
        });

        const stageIndex = STAGE_ORDER.indexOf(status);
        const colorIdx = stageIndex >= 0 ? stageIndex % CHART_COLORS.length : index % CHART_COLORS.length;

        return {
            label: status,
            data: dataPoints,
            backgroundColor: CHART_COLORS[colorIdx],
            borderColor: CHART_COLORS_BORDER[colorIdx],
            borderWidth: 1,
            fill: true
        };
    });

    if (distributionChartInstance) {
        distributionChartInstance.destroy();
    }

    distributionChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: availableDates,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false,
            },
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8', padding: 20, usePointStyle: true }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleColor: '#f8fafc',
                    bodyColor: '#cbd5e1',
                    padding: 10,
                    cornerRadius: 8,
                    displayColors: true
                }
            },
            scales: {
                y: {
                    stacked: true, // STACKED AREA
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#94a3b8' }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#94a3b8' }
                }
            }
        }
    });
}

function renderRoleChart(roleStackedData, statuses) {
    const ctx = document.getElementById('roleChart').getContext('2d');

    // 1. Sort Roles by Total Count (Most -> Least)
    const roleTotals = Object.entries(roleStackedData).map(([role, counts]) => {
        const total = Object.values(counts).reduce((a, b) => a + b, 0);
        return { role, total, counts };
    });

    roleTotals.sort((a, b) => b.total - a.total);

    const labels = roleTotals.map(r => r.role);

    // 2. Build Datasets (one per Status)
    const datasets = statuses.map((status, index) => {
        // Find color for this status from STAGE_ORDER index
        let color = CHART_COLORS[index % CHART_COLORS.length];

        // Try to sync color with Global STAGE_ORDER if possible
        const stageIdx = STAGE_ORDER.indexOf(status);
        if (stageIdx >= 0) {
            color = CHART_COLORS[stageIdx % CHART_COLORS.length];
        }

        return {
            label: status,
            data: roleTotals.map(r => r.counts[status] || 0),
            backgroundColor: color,
            borderRadius: 4,
            borderWidth: 0,
            barPercentage: 0.6,
        };
    });

    if (roleChartInstance) {
        roleChartInstance.destroy();
    }

    roleChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: { color: '#94a3b8' }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleColor: '#f8fafc',
                    bodyColor: '#cbd5e1',
                    padding: 10,
                    cornerRadius: 8,
                    displayColors: true,
                    callbacks: {
                        footer: (tooltipItems) => {
                            let total = 0;
                            tooltipItems.forEach((t) => total += t.parsed.y);
                            return 'Total: ' + total;
                        }
                    }
                }
            },
            scales: {
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#94a3b8', stepSize: 1 }
                },
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { color: '#94a3b8' }
                }
            }
        }
    });
}

function animateValue(obj, end, duration = 1000) {
    let startTimestamp = null;
    const start = parseInt(obj.innerHTML) || 0;
    if (isNaN(start)) { obj.innerHTML = end; return; }

    const step = (timestamp) => {
        if (!startTimestamp) startTimestamp = timestamp;
        const progress = Math.min((timestamp - startTimestamp) / duration, 1);
        obj.innerHTML = Math.floor(progress * (end - start) + start);
        if (progress < 1) {
            window.requestAnimationFrame(step);
        } else {
            obj.innerHTML = end;
        }
    };
    window.requestAnimationFrame(step);
}
