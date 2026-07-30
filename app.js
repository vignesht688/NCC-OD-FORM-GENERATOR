// Application Controller for NCC Attendance & OD Management System

// Destructure docx components from global docx object (provided by UMD build)
let docxLib = window.docx;

// State management
let cadetDatabase = []; // Full list of cadets
let attendanceRecords = {}; // Key: reg_no or name.toLowerCase(), Value: 'present' | 'absent' | 'od'
let selectedDate = new Date().toISOString().split('T')[0];
let attendanceParticulars = "";
let attendanceODReason = "";

let activeTab = 'tab-attendance';
let activeSubTab = 'subtab-sessions';
let isServerConnected = false;
let editingCadetKey = null;

// Search and filter states
let searchAttendanceQuery = "";
let filterAttendanceStatus = "ALL";
let searchCadetsQuery = "";
let filterShift = "";

// Document Config (Header/Signatures)
const docConfig = {
    college: "SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)",
    unit: "NATIONAL CADET CROPS",
    formType: "OD FORM",
    sigLeft: "NCC Officer",
    sigRight: "Principal"
};

// Helper to sort cadets by year (III first, then II, then I) and name
function getYearWeight(department) {
    const dept = (department || '').trim().toUpperCase();
    if (dept.startsWith('III')) return 3;
    if (dept.startsWith('II')) return 2;
    if (dept.startsWith('I')) return 1;
    return 0;
}

function sortCadetsForOD(a, b) {
    // 1. Sort by Year (III first, then II, then I)
    const weightA = getYearWeight(a.department);
    const weightB = getYearWeight(b.department);
    if (weightB !== weightA) {
        return weightB - weightA;
    }

    // 2. Sort by Name (alphabetically)
    const nameA = (a.name || '').trim().toUpperCase();
    const nameB = (b.name || '').trim().toUpperCase();
    if (nameA !== nameB) {
        return nameA.localeCompare(nameB);
    }

    // 3. Fallback to Register Number
    const regA = (a.register_no || '').trim();
    const regB = (b.register_no || '').trim();
    return regA.localeCompare(regB);
}

// On Page Load
document.addEventListener('DOMContentLoaded', () => {
    initApp();
});

// Initialize the application
async function initApp() {
    setupEventListeners();
    setTodayDate();
    loadConfig();
    await checkServerConnection();
    await loadCadetDatabase();
    await loadAttendanceForDate(selectedDate);
    updateUI();
}

// Load Document Config from LocalStorage
function loadConfig() {
    const savedConfig = localStorage.getItem('ncc_od_config');
    if (savedConfig) {
        try {
            Object.assign(docConfig, JSON.parse(savedConfig));
            // Pre-fill inputs
            document.getElementById('config-college').value = docConfig.college;
            document.getElementById('config-unit').value = docConfig.unit;
            document.getElementById('config-form-type').value = docConfig.formType;
            document.getElementById('config-sig-left').value = docConfig.sigLeft;
            document.getElementById('config-sig-right').value = docConfig.sigRight;
        } catch (e) {
            console.error("Error parsing saved config", e);
        }
    }
}

// Set initial date picker to today
function setTodayDate() {
    const dateInput = document.getElementById('attendance-date');
    if (dateInput) {
        dateInput.value = selectedDate;
    }
}

// Check if Express backend is running
async function checkServerConnection() {
    try {
        const res = await fetch('/api/cadets', { method: 'GET' });
        if (res.ok) {
            isServerConnected = true;
            updateSyncStatus(true, "Connected to Database Server");
        } else {
            isServerConnected = false;
            updateSyncStatus(false, "Offline Mode - Browser storage only");
        }
    } catch (err) {
        isServerConnected = false;
        updateSyncStatus(false, "Offline Mode - Server unreachable");
    }
}

// Update Server Connection Status Indicator UI
function updateSyncStatus(connected, message) {
    const badge = document.getElementById('badge-sync-status');
    const text = document.getElementById('text-sync-status');
    const serverBadge = document.getElementById('db-sync-info-box');
    const serverText = document.getElementById('db-sync-details');

    if (badge && text) {
        if (connected) {
            badge.className = "badge badge-sync status-connected";
            text.innerText = "Sync Active";
            badge.title = "Connected to local database: " + message;
        } else {
            badge.className = "badge badge-sync status-offline";
            text.innerText = "Offline Mode";
            badge.title = "Database offline: " + message;
        }
    }

    if (serverBadge && serverText) {
        if (connected) {
            serverBadge.className = "db-sync-info-box status-connected";
            serverBadge.style.background = "rgba(16, 185, 129, 0.15)";
            serverBadge.style.color = "#10b981";
            serverText.innerText = "Synced: " + message;
        } else {
            serverBadge.className = "db-sync-info-box status-offline";
            serverBadge.style.background = "rgba(245, 158, 11, 0.08)";
            serverBadge.style.color = "#f59e0b";
            serverText.innerText = "Local fallback: " + message;
        }
    }
}

// Load Cadets database from server (or local fallback)
async function loadCadetDatabase() {
    if (isServerConnected) {
        try {
            const res = await fetch('/api/cadets');
            if (res.ok) {
                cadetDatabase = await res.json();
                cadetDatabase.sort(sortCadetsForOD);
            }
        } catch (err) {
            console.error("Error fetching cadets:", err);
            loadCadetsFallback();
        }
    } else {
        loadCadetsFallback();
    }
    document.getElementById('stat-total-cadets').innerText = cadetDatabase.length;
}

// Local Storage Fallback for Cadets
function loadCadetsFallback() {
    const local = localStorage.getItem('ncc_cadets');
    if (local) {
        try {
            cadetDatabase = JSON.parse(local);
            if (Array.isArray(cadetDatabase)) {
                cadetDatabase.sort(sortCadetsForOD);
            }
        } catch (e) {
            cadetDatabase = [];
        }
    } else {
        cadetDatabase = [];
    }
}

// Save Cadet database (to server or localFallback)
async function saveCadetDatabase() {
    if (isServerConnected) {
        try {
            await fetch('/api/cadets', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(cadetDatabase)
            });
        } catch (err) {
            console.error("Error saving cadets to server:", err);
        }
    }
    localStorage.setItem('ncc_cadets', JSON.stringify(cadetDatabase));
    document.getElementById('stat-total-cadets').innerText = cadetDatabase.length;
}

// Load Attendance Records for a specific Date
async function loadAttendanceForDate(date) {
    if (isServerConnected) {
        try {
            const res = await fetch(`/api/attendance?date=${date}`);
            if (res.ok) {
                const data = await res.json();
                attendanceParticulars = data.event_name || "";
                attendanceODReason = data.od_reason || "";
                attendanceRecords = data.records || {};
                
                // Set form values
                document.getElementById('attendance-particulars').value = attendanceParticulars;
                document.getElementById('attendance-od-reason').value = attendanceODReason;
            }
        } catch (err) {
            console.error("Error fetching attendance:", err);
            loadAttendanceFallback(date);
        }
    } else {
        loadAttendanceFallback(date);
    }

    // Default unmarked cadets to 'present'
    cadetDatabase.forEach(cadet => {
        const key = getCadetKey(cadet);
        if (!attendanceRecords[key]) {
            attendanceRecords[key] = 'present';
        }
    });

    renderAttendanceChecklist();
    updateAttendanceSummaryCounts();
    generateLivePreview();
}

// Local Storage fallback for Attendance
function loadAttendanceFallback(date) {
    const local = localStorage.getItem(`ncc_att_${date}`);
    if (local) {
        try {
            const data = JSON.parse(local);
            attendanceParticulars = data.event_name || "";
            attendanceODReason = data.od_reason || "";
            attendanceRecords = data.records || {};
            document.getElementById('attendance-particulars').value = attendanceParticulars;
            document.getElementById('attendance-od-reason').value = attendanceODReason;
        } catch (e) {
            attendanceRecords = {};
        }
    } else {
        attendanceRecords = {};
        document.getElementById('attendance-particulars').value = "";
        document.getElementById('attendance-od-reason').value = "";
    }
}

// Save Attendance to Server (or local fallback)
async function saveAttendance() {
    // Collect event particulars
    attendanceParticulars = document.getElementById('attendance-particulars').value.trim();
    attendanceODReason = document.getElementById('attendance-od-reason').value.trim();

    const payload = {
        date: selectedDate,
        event_name: attendanceParticulars,
        od_reason: attendanceODReason,
        records: attendanceRecords
    };

    if (isServerConnected) {
        try {
            const res = await fetch('/api/attendance', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (res.ok) {
                showToast("Attendance saved successfully!");
            } else {
                showToast("Failed to save attendance on server.", "danger");
            }
        } catch (err) {
            console.error("Server save failed, using local storage", err);
            saveAttendanceLocal(payload);
        }
    } else {
        saveAttendanceLocal(payload);
    }

    updateAttendanceSummaryCounts();
    await generateLivePreview();
}

function saveAttendanceLocal(payload) {
    localStorage.setItem(`ncc_att_${selectedDate}`, JSON.stringify(payload));
    
    // Add date to index of local attendance dates
    let dates = [];
    const localDates = localStorage.getItem('ncc_att_dates');
    if (localDates) {
        try { dates = JSON.parse(localDates); } catch (e) {}
    }
    if (!dates.includes(selectedDate)) {
        dates.push(selectedDate);
        dates.sort();
        localStorage.setItem('ncc_att_dates', JSON.stringify(dates));
    }
    showToast("Attendance saved to browser storage!");
}

// Get unique cadet key
function getCadetKey(cadet) {
    return cadet.register_no ? cadet.register_no.trim() : cadet.name.trim().toLowerCase();
}

// Setup Event Listeners
function setupEventListeners() {
    // Attendance date picker change
    document.getElementById('attendance-date').addEventListener('change', (e) => {
        selectedDate = e.target.value;
        loadAttendanceForDate(selectedDate);
    });

    // Attendance search input
    document.getElementById('search-attendance').addEventListener('input', (e) => {
        searchAttendanceQuery = e.target.value.toLowerCase().trim();
        renderAttendanceChecklist();
    });

    // Attendance status filter
    document.getElementById('filter-attendance-status').addEventListener('change', (e) => {
        filterAttendanceStatus = e.target.value;
        renderAttendanceChecklist();
    });

    // Bulk actions
    document.getElementById('btn-mark-all-present').addEventListener('click', () => {
        bulkMarkStatus('present');
    });
    document.getElementById('btn-mark-all-absent').addEventListener('click', () => {
        bulkMarkStatus('absent');
    });

    // Save attendance button
    document.getElementById('btn-save-attendance').addEventListener('click', saveAttendance);

    // Event Particulars inputs for preview updating
    document.getElementById('attendance-particulars').addEventListener('input', () => {
        // Auto-show/hide OD clause based on particulars
        const particulars = document.getElementById('attendance-particulars').value.trim();
        const odGroup = document.getElementById('attendance-od-reason-group');
        if (particulars) {
            odGroup.style.display = 'block';
            // Auto generate OD Clause date
            const dateStr = formatDateString(selectedDate);
            document.getElementById('attendance-od-reason').value = `EXEMPTION FROM ALL ACADEMIC WORKS AND REGULAR CLASSES ON ${dateStr.toUpperCase()}`;
        } else {
            odGroup.style.display = 'none';
        }
        generateLivePreview();
    });

    document.getElementById('attendance-od-reason').addEventListener('input', generateLivePreview);

    // Doc type selector (OD Form / Attendance Sheet)
    const docTypeSel = document.getElementById('preview-doc-type');
    if (docTypeSel) {
        docTypeSel.addEventListener('change', generateLivePreview);
    }

    // Directory Search & Filter
    document.getElementById('search-cadets').addEventListener('input', (e) => {
        searchCadetsQuery = e.target.value.toLowerCase().trim();
        renderDirectoryChecklist();
    });

    document.getElementById('filter-shift').addEventListener('change', (e) => {
        filterShift = e.target.value;
        renderDirectoryChecklist();
    });

    // Directory Show Add Form Button
    document.getElementById('btn-show-add-form').addEventListener('click', () => {
        const form = document.getElementById('directory-add-form-container');
        const isHidden = form.style.display === 'none';
        form.style.display = isHidden ? 'block' : 'none';
        document.getElementById('add-form-title').innerText = "Add New Cadet";
        clearAddCadetForm();
    });

    // Cancel Add/Edit Cadet
    document.getElementById('btn-cancel-edit-cadet').addEventListener('click', () => {
        document.getElementById('directory-add-form-container').style.display = 'none';
        clearAddCadetForm();
    });

    // Save/Add Cadet button
    document.getElementById('btn-add-cadet').addEventListener('click', saveCadetAction);

    // Config inputs change
    const configs = ['config-college', 'config-unit', 'config-form-type', 'config-sig-left', 'config-sig-right'];
    configs.forEach(id => {
        document.getElementById(id).addEventListener('input', (e) => {
            const key = id.replace('config-', '').replace(/-([a-z])/g, (g) => g[1].toUpperCase());
            docConfig[key] = e.target.value;
            localStorage.setItem('ncc_od_config', JSON.stringify(docConfig));
            generateLivePreview();
        });
    });

    // Word doc generate button
    document.getElementById('btn-generate-docx').addEventListener('click', generateWordDocument);

    // Database manual backup exports
    document.getElementById('btn-export-db').addEventListener('click', () => {
        downloadJSON(cadetDatabase, "cadets.json");
    });
    document.getElementById('btn-export-attendance-db').addEventListener('click', async () => {
        let attData = {};
        if (isServerConnected) {
            try {
                // Fetch full logs
                const res = await fetch('/api/attendance/history');
                if (res.ok) {
                    const history = await res.json();
                    attData = history;
                }
            } catch (e) {}
        } else {
            // Read all local att logs
            const dates = getLocalAttendanceDates();
            dates.forEach(d => {
                const log = localStorage.getItem(`ncc_att_${d}`);
                if (log) attData[d] = JSON.parse(log);
            });
        }
        downloadJSON(attData, "attendance.json");
    });

    // Database JSON imports
    document.getElementById('import-db-input').addEventListener('change', (e) => {
        handleJSONImport(e, async (imported) => {
            if (Array.isArray(imported)) {
                cadetDatabase = imported;
                await saveCadetDatabase();
                renderDirectoryChecklist();
                renderAttendanceChecklist();
                showToast("Cadets directory updated!");
            } else {
                showToast("Invalid format. Expected an array.", "danger");
            }
        });
    });

    document.getElementById('import-attendance-db-input').addEventListener('change', (e) => {
        handleJSONImport(e, async (imported) => {
            if (imported && typeof imported === 'object' && !Array.isArray(imported)) {
                if (isServerConnected) {
                    for (const date in imported) {
                        await fetch('/api/attendance', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ date, ...imported[date] })
                        });
                    }
                } else {
                    for (const date in imported) {
                        localStorage.setItem(`ncc_att_${date}`, JSON.stringify(imported[date]));
                    }
                    localStorage.setItem('ncc_att_dates', JSON.stringify(Object.keys(imported).sort()));
                }
                showToast("Attendance history updated!");
                if (activeTab === 'tab-history') {
                    loadAttendanceHistory();
                }
            } else {
                showToast("Invalid format. Expected a JSON dictionary.", "danger");
            }
        });
    });
}

// Bulk mark status
function bulkMarkStatus(status) {
    cadetDatabase.forEach(cadet => {
        // Apply search query filters
        const matchesSearch = !searchAttendanceQuery || 
            cadet.name.toLowerCase().includes(searchAttendanceQuery) || 
            (cadet.register_no && cadet.register_no.toLowerCase().includes(searchAttendanceQuery)) ||
            (cadet.department && cadet.department.toLowerCase().includes(searchAttendanceQuery));
        
        if (matchesSearch) {
            const key = getCadetKey(cadet);
            attendanceRecords[key] = status;
        }
    });
    renderAttendanceChecklist();
    updateAttendanceSummaryCounts();
    generateLivePreview();
}

// Render Attendance Checklist tab
function renderAttendanceChecklist() {
    const container = document.getElementById('attendance-checklist-container');
    if (!container) return;

    container.innerHTML = "";

    const matched = cadetDatabase.filter(cadet => {
        const key = getCadetKey(cadet);
        const status = attendanceRecords[key] || 'present';
        
        const matchesSearch = !searchAttendanceQuery || 
            cadet.name.toLowerCase().includes(searchAttendanceQuery) || 
            (cadet.register_no && cadet.register_no.toLowerCase().includes(searchAttendanceQuery)) ||
            (cadet.department && cadet.department.toLowerCase().includes(searchAttendanceQuery));

        const matchesStatus = filterAttendanceStatus === 'ALL' || 
            (filterAttendanceStatus === 'PRESENT' && status === 'present') ||
            (filterAttendanceStatus === 'ABSENT' && status === 'absent') ||
            (filterAttendanceStatus === 'OD' && status === 'od') ||
            (filterAttendanceStatus === 'UNMARKED' && !status);

        return matchesSearch && matchesStatus;
    });

    if (matched.length === 0) {
        container.innerHTML = `
            <div class="preview-empty-state" style="padding: 2rem 0;">
                <i class="fa-solid fa-users-slash"></i>
                <p>No cadets match filters.</p>
            </div>`;
        return;
    }

    matched.forEach(cadet => {
        const key = getCadetKey(cadet);
        const status = attendanceRecords[key] || 'present';

        const item = document.createElement('div');
        item.className = `cadet-item ${status === 'od' ? 'selected' : ''}`;
        item.style.display = 'flex';
        item.style.justifyContent = 'space-between';
        item.style.alignItems = 'center';

        const info = document.createElement('div');
        info.className = 'cadet-info';
        
        // Custom color shift badges
        const shiftBadge = cadet.shift ? `
            <span class="cadet-badge-shift ${cadet.shift === 'SHIFT-I' ? 'shift-1' : 'shift-2'}">
                ${cadet.shift}
            </span>` : '';

        info.innerHTML = `
            <div class="cadet-text">
                <span class="cadet-name">${cadet.name} ${shiftBadge}</span>
                <span class="cadet-sub">${cadet.register_no || "No Reg"} | ${cadet.department || "No Dept"}</span>
            </div>`;

        const actions = document.createElement('div');
        actions.className = 'attendance-actions';
        
        const group = document.createElement('div');
        group.className = 'attendance-toggle-group';

        const statuses = [
            { code: 'present', label: 'P', activeClass: 'active-present' },
            { code: 'absent', label: 'A', activeClass: 'active-absent' },
            { code: 'od', label: 'OD', activeClass: 'active-od' }
        ];

        statuses.forEach(s => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = `attendance-toggle-btn ${status === s.code ? s.activeClass : ''}`;
            btn.innerText = s.code === 'od' ? 'OD' : s.label;
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                // Toggle status
                attendanceRecords[key] = s.code;
                renderAttendanceChecklist();
                updateAttendanceSummaryCounts();
                generateLivePreview();
            });
            group.appendChild(btn);
        });

        actions.appendChild(group);
        item.appendChild(info);
        item.appendChild(actions);
        container.appendChild(item);
    });
}

// Update Counts Summary badges
function updateAttendanceSummaryCounts() {
    let total = cadetDatabase.length;
    let present = 0;
    let absent = 0;
    let od = 0;

    cadetDatabase.forEach(cadet => {
        const key = getCadetKey(cadet);
        const status = attendanceRecords[key];
        if (status === 'present') present++;
        else if (status === 'absent') absent++;
        else if (status === 'od') od++;
    });

    document.getElementById('count-total').innerText = total;
    document.getElementById('count-present').innerText = present;
    document.getElementById('count-absent').innerText = absent;
    document.getElementById('count-od').innerText = od;
    document.getElementById('stat-selected-count').innerText = od;
}

// Render Directory Tab
function renderDirectoryChecklist() {
    const container = document.getElementById('cadets-checklist-container');
    if (!container) return;

    container.innerHTML = "";

    const matched = cadetDatabase.filter(cadet => {
        const matchesSearch = !searchCadetsQuery || 
            cadet.name.toLowerCase().includes(searchCadetsQuery) || 
            (cadet.register_no && cadet.register_no.toLowerCase().includes(searchCadetsQuery)) ||
            (cadet.department && cadet.department.toLowerCase().includes(searchCadetsQuery));

        const matchesShift = !filterShift || cadet.shift === filterShift;

        return matchesSearch && matchesShift;
    });

    if (matched.length === 0) {
        container.innerHTML = `
            <div class="preview-empty-state" style="padding: 2rem 0;">
                <i class="fa-solid fa-users-slash"></i>
                <p>No cadets found in directory.</p>
            </div>`;
        return;
    }

    matched.forEach(cadet => {
        const key = getCadetKey(cadet);

        const item = document.createElement('div');
        item.className = 'cadet-item';
        item.style.display = 'flex';
        item.style.justifyContent = 'space-between';
        item.style.alignItems = 'center';

        const info = document.createElement('div');
        info.className = 'cadet-info';
        
        const shiftBadge = cadet.shift ? `
            <span class="cadet-badge-shift ${cadet.shift === 'SHIFT-I' ? 'shift-1' : 'shift-2'}">
                ${cadet.shift}
            </span>` : '';

        info.innerHTML = `
            <div class="cadet-text">
                <span class="cadet-name">${cadet.name} ${shiftBadge}</span>
                <span class="cadet-sub">${cadet.register_no || "No Reg"} | ${cadet.department || "No Dept"}</span>
            </div>`;

        const actions = document.createElement('div');
        actions.className = 'cadet-actions';

        // Edit Button
        const editBtn = document.createElement('button');
        editBtn.type = 'button';
        editBtn.className = 'btn-edit-cadet';
        editBtn.title = "Edit Cadet details";
        editBtn.innerHTML = '<i class="fa-solid fa-pen-to-square"></i>';
        editBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            editCadetDetails(cadet);
        });

        // Delete Button
        const delBtn = document.createElement('button');
        delBtn.type = 'button';
        delBtn.className = 'btn-delete-permanent';
        delBtn.title = "Delete Cadet from Directory";
        delBtn.innerHTML = '<i class="fa-solid fa-trash"></i>';
        delBtn.addEventListener('click', async (e) => {
            e.stopPropagation();
            if (confirm(`Are you sure you want to permanently delete ${cadet.name} from the directory?`)) {
                await deleteCadetAction(cadet);
            }
        });

        actions.appendChild(editBtn);
        actions.appendChild(delBtn);
        item.appendChild(info);
        item.appendChild(actions);
        container.appendChild(item);
    });
}

// Edit Cadet Details Pre-fill
function editCadetDetails(cadet) {
    document.getElementById('directory-add-form-container').style.display = 'block';
    document.getElementById('add-form-title').innerText = "Edit Cadet Details";
    
    document.getElementById('new-reg-no').value = cadet.register_no || "";
    document.getElementById('new-name').value = cadet.name || "";
    document.getElementById('new-dept').value = cadet.department || "";
    document.getElementById('new-shift').value = cadet.shift || "";
    
    editingCadetKey = getCadetKey(cadet);
    document.getElementById('btn-cancel-edit-cadet').style.display = 'inline-flex';
}

// Save or Update Cadet
async function saveCadetAction() {
    const regNo = document.getElementById('new-reg-no').value.trim();
    const name = document.getElementById('new-name').value.trim();
    const dept = document.getElementById('new-dept').value.trim();
    const shift = document.getElementById('new-shift').value;

    if (!name) {
        showToast("Cadet Name is required!", "danger");
        return;
    }

    const cadet = { register_no: regNo, name: name.toUpperCase(), department: dept.toUpperCase(), shift };
    const key = getCadetKey(cadet);

    if (editingCadetKey) {
        // Update existing
        const idx = cadetDatabase.findIndex(c => getCadetKey(c) === editingCadetKey);
        if (idx !== -1) {
            cadetDatabase[idx] = cadet;
            showToast("Cadet updated successfully!");
        }
        editingCadetKey = null;
        document.getElementById('btn-cancel-edit-cadet').style.display = 'none';
    } else {
        // Add new
        // Check for duplicates
        if (cadetDatabase.some(c => getCadetKey(c) === key)) {
            showToast("Cadet already exists in directory!", "warning");
            return;
        }
        cadetDatabase.push(cadet);
        showToast("Cadet added to directory!");
    }

    await saveCadetDatabase();
    clearAddCadetForm();
    document.getElementById('directory-add-form-container').style.display = 'none';
    
    // Sort directory
    cadetDatabase.sort(sortCadetsForOD);
    
    renderDirectoryChecklist();
    renderAttendanceChecklist();
}

// Delete Cadet Action
async function deleteCadetAction(cadet) {
    const key = getCadetKey(cadet);
    cadetDatabase = cadetDatabase.filter(c => getCadetKey(c) !== key);
    await saveCadetDatabase();
    
    // Delete from current attendance
    if (attendanceRecords[key]) {
        delete attendanceRecords[key];
    }
    
    renderDirectoryChecklist();
    renderAttendanceChecklist();
    updateAttendanceSummaryCounts();
    showToast("Cadet deleted from directory.");
}

function clearAddCadetForm() {
    document.getElementById('new-reg-no').value = "";
    document.getElementById('new-name').value = "";
    document.getElementById('new-dept').value = "";
    document.getElementById('new-shift').value = "";
    editingCadetKey = null;
}

// Load Attendance History Summary
async function loadAttendanceHistory() {
    const container = document.getElementById('history-sessions-container');
    if (!container) return;

    container.innerHTML = `
        <div class="preview-empty-state" style="padding: 2rem 0;">
            <i class="fa-solid fa-circle-notch fa-spin"></i>
            <p>Loading attendance log records...</p>
        </div>`;

    let history = [];
    if (isServerConnected) {
        try {
            const res = await fetch('/api/attendance/history');
            if (res.ok) history = await res.json();
        } catch (e) {
            console.error(e);
            history = getLocalAttendanceHistory();
        }
    } else {
        history = getLocalAttendanceHistory();
    }

    if (history.length === 0) {
        container.innerHTML = `
            <div class="preview-empty-state" style="padding: 2rem 0;">
                <i class="fa-solid fa-history"></i>
                <p>No attendance logs recorded yet.</p>
            </div>`;
        return;
    }

    container.innerHTML = "";
    history.forEach(log => {
        const card = document.createElement('div');
        card.className = 'history-session-card';

        const header = document.createElement('div');
        header.className = 'history-card-header';
        
        const dateStr = formatDateString(log.date);
        header.innerHTML = `
            <span class="history-card-date">${dateStr}</span>
            <span class="history-card-particulars">${log.event_name || 'Regular parade'}</span>`;

        const stats = document.createElement('div');
        stats.className = 'history-card-stats';
        stats.innerHTML = `
            <span class="history-stat-badge present">P: ${log.counts.present}</span>
            <span class="history-stat-badge absent">A: ${log.counts.absent}</span>
            <span class="history-stat-badge od">OD: ${log.counts.od}</span>`;

        const actions = document.createElement('div');
        actions.className = 'history-card-actions';

        // Load Button
        const editBtn = document.createElement('button');
        editBtn.className = 'btn btn-secondary btn-sm';
        editBtn.innerHTML = '<i class="fa-solid fa-folder-open"></i> Load';
        editBtn.addEventListener('click', () => {
            selectedDate = log.date;
            document.getElementById('attendance-date').value = selectedDate;
            loadAttendanceForDate(selectedDate);
            switchTab('tab-attendance');
            showToast(`Loaded attendance log for ${dateStr}`);
        });

        // Delete Button
        const delBtn = document.createElement('button');
        delBtn.className = 'btn btn-danger-outline btn-sm';
        delBtn.innerHTML = '<i class="fa-solid fa-trash"></i>';
        delBtn.addEventListener('click', async () => {
            if (confirm(`Delete attendance records for ${dateStr}?`)) {
                await deleteAttendanceAction(log.date);
            }
        });

        // Generate Sheet
        const sheetBtn = document.createElement('button');
        sheetBtn.className = 'btn btn-secondary btn-sm';
        sheetBtn.innerHTML = '<i class="fa-solid fa-file-invoice"></i> Sheet';
        sheetBtn.addEventListener('click', () => {
            generateSessionReport(log.date, 'attendance');
        });

        // Generate OD Form
        const odBtn = document.createElement('button');
        odBtn.className = 'btn btn-primary btn-sm';
        odBtn.innerHTML = '<i class="fa-solid fa-file-word"></i> OD Form';
        odBtn.disabled = log.counts.od === 0;
        odBtn.addEventListener('click', () => {
            generateSessionReport(log.date, 'od');
        });

        actions.appendChild(editBtn);
        actions.appendChild(sheetBtn);
        actions.appendChild(odBtn);
        actions.appendChild(delBtn);

        card.appendChild(header);
        card.appendChild(stats);
        card.appendChild(actions);
        container.appendChild(card);
    });

    // Populate cadet-wise stats list in Analytics sub-tab
    renderCadetAnalytics(history);
}

// Delete Attendance Record
async function deleteAttendanceAction(date) {
    if (isServerConnected) {
        try {
            const res = await fetch(`/api/attendance?date=${date}`, { method: 'DELETE' });
            if (res.ok) {
                showToast("Record deleted.");
            }
        } catch (e) {
            deleteAttendanceLocal(date);
        }
    } else {
        deleteAttendanceLocal(date);
    }
    loadAttendanceHistory();
    if (selectedDate === date) {
        loadAttendanceForDate(selectedDate);
    }
}

function deleteAttendanceLocal(date) {
    localStorage.removeItem(`ncc_att_${date}`);
    let dates = [];
    const localDates = localStorage.getItem('ncc_att_dates');
    if (localDates) {
        try { dates = JSON.parse(localDates); } catch (e) {}
    }
    dates = dates.filter(d => d !== date);
    localStorage.setItem('ncc_att_dates', JSON.stringify(dates));
    showToast("Record deleted from browser storage.");
}

// Generate reports directly from History logs
async function generateSessionReport(date, type) {
    showToast("Compiling document...");
    
    // Fetch details for that date first
    let record = null;
    if (isServerConnected) {
        try {
            const res = await fetch(`/api/attendance?date=${date}`);
            if (res.ok) record = await res.json();
        } catch (e) {}
    }
    
    if (!record) {
        const local = localStorage.getItem(`ncc_att_${date}`);
        if (local) record = JSON.parse(local);
    }

    if (!record || !record.records) {
        showToast("Failed to compile session records.", "danger");
        return;
    }

    if (type === 'od') {
        // Compile OD Form
        const selectedCadetsList = [];
        cadetDatabase.forEach(cadet => {
            const key = getCadetKey(cadet);
            if (record.records[key] === 'od') {
                selectedCadetsList.push(cadet);
            }
        });

        selectedCadetsList.sort(sortCadetsForOD);

        if (selectedCadetsList.length === 0) {
            showToast("No cadets marked OD on that date.", "warning");
            return;
        }

        const payload = {
            selectedCadets: selectedCadetsList,
            particularsVal: record.event_name,
            odReasonVal: record.od_reason,
            selectedDate: date,
            docConfig
        };
        
        await triggerDocDownload('/api/generate-docx', payload, `od-${date}`);
    } else {
        // Compile Attendance Sheet
        const recordsList = [];
        cadetDatabase.forEach(cadet => {
            const key = getCadetKey(cadet);
            recordsList.push({
                register_no: cadet.register_no,
                name: cadet.name,
                department: cadet.department,
                shift: cadet.shift,
                status: record.records[key] || 'present'
            });
        });

        const payload = {
            date,
            records: recordsList,
            docConfig
        };

        await triggerDocDownload('/api/generate-attendance-sheet-docx', payload, `attendance-${date}`);
    }
}

// Render Cadet-wise Analytics
function renderCadetAnalytics(history) {
    const container = document.getElementById('stats-cadets-container');
    if (!container) return;

    container.innerHTML = "";

    // Count sessions per cadet
    const cadetStats = {};
    cadetDatabase.forEach(cadet => {
        const key = getCadetKey(cadet);
        cadetStats[key] = {
            cadet,
            present: 0,
            absent: 0,
            od: 0,
            total: 0
        };
    });

    history.forEach(log => {
        if (log.records) {
            for (const key in log.records) {
                if (cadetStats[key]) {
                    const status = log.records[key];
                    cadetStats[key].total++;
                    if (status === 'present') cadetStats[key].present++;
                    else if (status === 'absent') cadetStats[key].absent++;
                    else if (status === 'od') cadetStats[key].od++;
                }
            }
        }
    });

    const list = Object.values(cadetStats);
    // Sort by name
    list.sort((a,b) => a.cadet.name.localeCompare(b.cadet.name));

    list.forEach(item => {
        const row = document.createElement('div');
        row.className = 'cadet-stats-row';

        const info = document.createElement('div');
        info.className = 'cadet-stats-info';
        info.innerHTML = `
            <span class="cadet-stats-name">${item.cadet.name}</span>
            <span class="cadet-stats-sub">
                P: ${item.present} | A: ${item.absent} | OD: ${item.od} (Total: ${item.total})
            </span>`;

        // Calculate attendance %: (Present + OD) / Total
        const pct = item.total > 0 ? Math.round(((item.present + item.od) / item.total) * 100) : 100;
        
        let pctClass = 'good';
        if (pct < 75) pctClass = 'danger';
        else if (pct < 85) pctClass = 'warning';

        const percentage = document.createElement('div');
        percentage.className = 'cadet-stats-percentage';
        percentage.innerHTML = `
            <span class="percentage-badge ${pctClass}">${pct}%</span>`;

        row.appendChild(info);
        row.appendChild(percentage);
        container.appendChild(row);
    });
}

// Get Local storage history summary
function getLocalAttendanceHistory() {
    const dates = getLocalAttendanceDates();
    const history = [];

    dates.forEach(d => {
        const log = localStorage.getItem(`ncc_att_${d}`);
        if (log) {
            try {
                const parsed = JSON.parse(log);
                const counts = { present: 0, absent: 0, od: 0, total: 0 };
                if (parsed.records) {
                    for (const key in parsed.records) {
                        counts.total++;
                        const s = parsed.records[key];
                        if (s === 'present') counts.present++;
                        else if (s === 'absent') counts.absent++;
                        else if (s === 'od') counts.od++;
                    }
                }
                history.push({
                    date: d,
                    event_name: parsed.event_name,
                    od_reason: parsed.od_reason,
                    records: parsed.records,
                    counts
                });
            } catch (e) {}
        }
    });

    // Sort descending
    history.sort((a, b) => b.date.localeCompare(a.date));
    return history;
}

function getLocalAttendanceDates() {
    const localDates = localStorage.getItem('ncc_att_dates');
    if (localDates) {
        try { return JSON.parse(localDates); } catch (e) {}
    }
    return [];
}

// Re-generate live document preview on right side panel
async function generateLivePreview() {
    const emptyPlaceholder = document.getElementById('preview-empty-placeholder');
    const contentArea = document.getElementById('preview-content-area');
    const previewTableBody = document.getElementById('preview-table-body');
    const docTypeSelector = document.getElementById('preview-doc-type');
    const generateBtn = document.getElementById('btn-generate-docx');

    const particulars = document.getElementById('attendance-particulars').value.trim();
    const odReason = document.getElementById('attendance-od-reason').value.trim();

    // Determine document type to preview
    const docType = docTypeSelector ? docTypeSelector.value : 'attendance';

    // Count selected/OD cadets
    const odCadets = [];
    cadetDatabase.forEach(cadet => {
        const key = getCadetKey(cadet);
        if (attendanceRecords[key] === 'od') {
            odCadets.push(cadet);
        }
    });

    odCadets.sort(sortCadetsForOD);

    if (docType === 'od') {
        // OD Form Preview requires particulars and at least one OD cadet
        if (!particulars || odCadets.length === 0) {
            emptyPlaceholder.style.display = 'flex';
            contentArea.style.display = 'none';
            generateBtn.disabled = true;
            return;
        }

        emptyPlaceholder.style.display = 'none';
        contentArea.style.display = 'flex';
        generateBtn.disabled = false;

        document.getElementById('preview-form-title').innerText = docConfig.formType;
        
        // Render preview table body
        previewTableBody.innerHTML = "";
        odCadets.forEach((cadet, index) => {
            const tr = document.createElement('tr');
            
            // Build cell contents
            let particularsCell = "";
            let odCell = "";
            
            if (index === 0) {
                // Vertical rowspan mockup for first cell
                particularsCell = `<td rowspan="${odCadets.length}"><div class="merged-cell-content">${particulars}</div></td>`;
                odCell = `<td rowspan="${odCadets.length}"><div class="merged-cell-content">${odReason}</div></td>`;
            }
            
            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${cadet.register_no || "-"}</td>
                <td style="text-align: left; padding-left: 0.25rem;">${cadet.name}</td>
                <td style="text-align: left; padding-left: 0.25rem;">${cadet.department}</td>
                <td>${cadet.shift || "-"}</td>
                ${particularsCell}
                ${odCell}`;
                
            previewTableBody.appendChild(tr);
        });
    } else {
        // Attendance Sheet Preview (always visible if database loaded)
        if (cadetDatabase.length === 0) {
            emptyPlaceholder.style.display = 'flex';
            contentArea.style.display = 'none';
            generateBtn.disabled = true;
            return;
        }

        emptyPlaceholder.style.display = 'none';
        contentArea.style.display = 'flex';
        generateBtn.disabled = false;

        const dateStr = formatDateString(selectedDate);
        document.getElementById('preview-form-title').innerText = `ATTENDANCE SHEET - ${dateStr.toUpperCase()}`;

        // Render preview table body
        previewTableBody.innerHTML = "";
        
        // S.No | Register No | Name | Department | Shift | Status
        cadetDatabase.forEach((cadet, index) => {
            const key = getCadetKey(cadet);
            const status = (attendanceRecords[key] || 'present').toUpperCase();

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${cadet.register_no || "-"}</td>
                <td style="text-align: left; padding-left: 0.25rem;">${cadet.name}</td>
                <td style="text-align: left; padding-left: 0.25rem;">${cadet.department || "-"}</td>
                <td>${cadet.shift || "-"}</td>
                <td colspan="2" style="font-weight: bold; color: ${status === 'ABSENT' ? 'var(--accent-danger)' : status === 'OD' ? 'var(--accent-blue)' : 'var(--accent-success)'}">
                    ${status}
                </td>`;
            previewTableBody.appendChild(tr);
        });
    }
}

// Download dynamic Word Document (Attendance sheet or OD form)
async function generateWordDocument() {
    const docTypeSelector = document.getElementById('preview-doc-type');
    const docType = docTypeSelector ? docTypeSelector.value : 'attendance';

    if (docType === 'od') {
        const odCadets = [];
        cadetDatabase.forEach(cadet => {
            const key = getCadetKey(cadet);
            if (attendanceRecords[key] === 'od') {
                odCadets.push(cadet);
            }
        });

        odCadets.sort(sortCadetsForOD);

        if (odCadets.length === 0) {
            showToast("No cadets marked OD to export.", "warning");
            return;
        }

        const particulars = document.getElementById('attendance-particulars').value.trim();
        const odReason = document.getElementById('attendance-od-reason').value.trim();

        const payload = {
            selectedCadets: odCadets,
            particularsVal: particulars,
            odReasonVal: odReason,
            selectedDate,
            docConfig
        };

        showToast("Generating OD Form Word Document...");
        await triggerDocDownload('/api/generate-docx', payload, `od-${selectedDate}`);
    } else {
        const recordsList = [];
        cadetDatabase.forEach(cadet => {
            const key = getCadetKey(cadet);
            recordsList.push({
                register_no: cadet.register_no,
                name: cadet.name,
                department: cadet.department,
                shift: cadet.shift,
                status: attendanceRecords[key] || 'present'
            });
        });

        const payload = {
            date: selectedDate,
            records: recordsList,
            docConfig
        };

        showToast("Generating Attendance Sheet Word Document...");
        await triggerDocDownload('/api/generate-attendance-sheet-docx', payload, `attendance-${selectedDate}`);
    }
}

// AJAX Download Trigger
async function triggerDocDownload(url, payload, filePrefix) {
    if (isServerConnected) {
        try {
            const res = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (res.ok) {
                const blob = await res.blob();
                const downloadUrl = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = downloadUrl;
                a.download = `${filePrefix}.docx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(downloadUrl);
                showToast("Document downloaded successfully!");
                return;
            }
        } catch (e) {
            console.error("Server download failed, falling back to browser-only download", e);
        }
    }
    
    // Browser fallback
    showToast("Offline mode: Client-side docx rendering fallback.", "warning");
    generateDocxClientFallback(payload, url.includes('sheet') ? 'attendance' : 'od');
}

// Client Side Docx Generation fallback using UMD docx
function generateDocxClientFallback(payload, type) {
    if (!docxLib) {
        showToast("Error: docx library not loaded. Check internet connection.", "danger");
        return;
    }
    showToast("Generating document locally...", "info");
    try {
        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, VerticalAlign, BorderStyle } = docxLib;
        
        let docRows = [];
        
        if (type === 'od') {
            const { selectedCadets, particularsVal, odReasonVal } = payload;
            
            // Header Row
            docRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", bold: true, size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", bold: true, size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", bold: true, size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", bold: true, size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", bold: true, size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PARTICULARS", bold: true, size: 24 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ON DUTY", bold: true, size: 24 })] })] })
                ]
            }));

            const N = selectedCadets.length;
            let P = 1;
            if (N > 27) {
                P = Math.ceil((N - 27) / 32) + 1;
            }
            const midPage = Math.ceil(P / 2);
            const startRow = midPage === 1 ? 0 : 27 + (midPage - 2) * 32;

            selectedCadets.forEach((cadet, index) => {
                let isPartRestart = false;
                let partVal = "";
                let isOdRestart = false;
                let odVal = "";
                
                if (startRow > 0) {
                    if (index === 0) {
                        isPartRestart = true;
                        partVal = "";
                        isOdRestart = true;
                        odVal = "";
                    } else if (index === startRow) {
                        isPartRestart = true;
                        partVal = particularsVal;
                        isOdRestart = true;
                        odVal = odReasonVal;
                    }
                } else {
                    if (index === 0) {
                        isPartRestart = true;
                        partVal = particularsVal;
                        isOdRestart = true;
                        odVal = odReasonVal;
                    }
                }

                const partMerge = startRow > 0
                    ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
                    : (index === 0 ? 'restart' : 'continue');

                const odMerge = startRow > 0
                    ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
                    : (index === 0 ? 'restart' : 'continue');

                const partCell = new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isPartRestart ? partVal : "", size: 24 })] })],
                    verticalMerge: partMerge
                });

                const odCell = new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isOdRestart ? odVal : "", size: 24 })] })],
                    verticalMerge: odMerge
                });

                docRows.push(new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), size: 24 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "-", size: 24 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, size: 24 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department || "-", size: 24 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "-", size: 24 })] })] }),
                        partCell,
                        odCell
                    ]
                }));
            });
        } else {
            const { records } = payload;
            
            docRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", bold: true, size: 22 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", bold: true, size: 22 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", bold: true, size: 22 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", bold: true, size: 22 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", bold: true, size: 22 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", bold: true, size: 22 })] })] })
                ]
            }));

            records.forEach((cadet, index) => {
                docRows.push(new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), size: 22 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "-", size: 22 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, size: 22 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department || "-", size: 22 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "-", size: 22 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.status.toUpperCase(), bold: true, size: 22 })] })] })
                    ]
                }));
            });
        }

        const dateStr = formatDateString(selectedDate);
        
        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: [
                    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: docConfig.college, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: docConfig.unit, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: type === 'od' ? docConfig.formType : `ATTENDANCE SHEET - ${dateStr.toUpperCase()}`, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: docRows }),
                    new Paragraph({ spacing: { before: 1440 }, children: [new TextRun("")] }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0 },
                            bottom: { style: BorderStyle.NONE, size: 0 },
                            left: { style: BorderStyle.NONE, size: 0 },
                            right: { style: BorderStyle.NONE, size: 0 },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0 },
                            insideVertical: { style: BorderStyle.NONE, size: 0 }
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: docConfig.sigLeft, font: "Bookman Old Style", size: 24 })] })] }),
                                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: docConfig.sigRight, font: "Bookman Old Style", size: 24 })] })] })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        Packer.toBlob(doc).then(blob => {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${type === 'od' ? 'od' : 'attendance'}-${selectedDate}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast("Document downloaded locally!");
        });

    } catch (err) {
        console.error("Local docx build failed", err);
        showToast("Local docx build failed. See console.", "danger");
    }
}

// Switch Tabs
window.switchTab = function(tabId) {
    logActionToServer('Switch Tab', { tabId });
    // Hide all tab content
    const contents = document.querySelectorAll('.tab-content');
    contents.forEach(c => c.classList.remove('active'));

    // Remove active class from all buttons
    const btns = document.querySelectorAll('.tab-btn');
    btns.forEach(b => b.classList.remove('active'));

    // Show selected content and set button active
    const activeContent = document.getElementById(tabId);
    if (activeContent) {
        activeContent.classList.add('active');
    }

    // Set button active
    const activeBtn = Array.from(btns).find(b => b.getAttribute('onclick') && b.getAttribute('onclick').includes(tabId));
    if (activeBtn) {
        activeBtn.classList.add('active');
    }

    activeTab = tabId;

    // Trigger tab-specific loaders
    if (tabId === 'tab-attendance') {
        renderAttendanceChecklist();
        updateAttendanceSummaryCounts();
        generateLivePreview();
    } else if (tabId === 'tab-history') {
        loadAttendanceHistory();
    } else if (tabId === 'tab-select') {
        renderDirectoryChecklist();
    }
    
    // Add dropdown event handler to live preview container if present
    const docTypeSelector = document.getElementById('preview-doc-type');
    if (docTypeSelector && !docTypeSelector.dataset.listenerBound) {
        docTypeSelector.addEventListener('change', () => {
            generateLivePreview();
        });
        docTypeSelector.dataset.listenerBound = 'true';
    }
};

// Switch SubTabs in History & Stats
window.switchSubTab = function(subTabId) {
    logActionToServer('Switch Sub-Tab', { subTabId });
    const contents = document.querySelectorAll('.subtab-content');
    contents.forEach(c => c.style.display = 'none');

    const subBtns = document.getElementById('tab-history').querySelectorAll('.tab-btn');
    subBtns.forEach(b => b.classList.remove('active'));

    const activeContent = document.getElementById(subTabId);
    if (activeContent) {
        activeContent.style.display = 'block';
    }

    const activeBtn = document.getElementById(subTabId === 'subtab-sessions' ? 'btn-subtab-sessions' : 'btn-subtab-stats');
    if (activeBtn) {
        activeBtn.classList.add('active');
    }

    activeSubTab = subTabId;
};

// Helper: Format YYYY-MM-DD to DD MMM YYYY (e.g. 08 JUL 2026)
function formatDateString(dateStr) {
    if (!dateStr) return "";
    const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    const day = String(d.getDate()).padStart(2, '0');
    const month = months[d.getMonth()];
    const year = d.getFullYear();
    return `${day} ${month} ${year}`;
}

// Toast notification
function showToast(message, type = "success") {
    logActionToServer('Toast Shown', { message, type });
    const toast = document.getElementById('toast-notification');
    const text = document.getElementById('toast-message');
    if (!toast || !text) return;

    text.innerText = message;
    
    // Reset colors
    toast.className = "toast show";
    if (type === 'danger') {
        toast.classList.add('toast-danger');
    } else if (type === 'warning') {
        toast.classList.add('toast-warning');
    }

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Helper: Trigger Download of JSON data
function downloadJSON(data, filename) {
    logActionToServer('Download JSON', { filename });
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Helper: Read files imported via upload input
function handleJSONImport(event, callback) {
    logActionToServer('Import JSON File');
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const parsed = JSON.parse(e.target.result);
            callback(parsed);
        } catch (err) {
            showToast("Failed to parse JSON file.", "danger");
        }
    };
    reader.readAsText(file);
    event.target.value = "";
}

// Standalone updateUI dummy to maintain initialization calls
function updateUI() {
    generateLivePreview();
}

// Helper function to log frontend user actions to the terminal
function logActionToServer(action, details = {}) {
    if (isServerConnected) {
        fetch('/api/log-action', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action, details })
        }).catch(err => {});
    }
}
