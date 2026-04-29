<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MMC Steel – Engineering BOM & Procurement System</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root {
  --navy: #0B1E3D;
  --navy2: #112952;
  --navy3: #1A3A6B;
  --steel: #1E5799;
  --accent: #F4A62A;
  --accent2: #2EBBAA;
  --green: #1B7F4F;
  --red: #8B1A1A;
  --surface: #0F2744;
  --card: #152D52;
  --border: #243D6A;
  --txt: #E8EFF8;
  --txt2: #7B96BC;
  --txt3: #4A6080;
  --font-head: 'JetBrains Mono', monospace;
  --font-body: 'IBM Plex Sans', sans-serif;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: var(--font-body);
  background: var(--navy);
  color: var(--txt);
  min-height: 100vh;
  padding: 0;
}

/* TOP BAR */
.topbar {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  padding: 0 32px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  height: 64px;
  position: sticky;
  top: 0;
  z-index: 100;
}
.brand {
  display: flex;
  align-items: center;
  gap: 14px;
}
.brand-icon {
  width: 38px; height: 38px;
  background: var(--accent);
  border-radius: 8px;
  display: flex; align-items: center; justify-content: center;
  font-family: var(--font-head);
  font-weight: 700;
  font-size: 14px;
  color: var(--navy);
}
.brand-text { line-height: 1.2; }
.brand-text .co { font-family: var(--font-head); font-size: 13px; font-weight: 700; color: var(--txt); }
.brand-text .sub { font-size: 10px; color: var(--txt2); letter-spacing: 0.08em; text-transform: uppercase; }
.topbar-actions { display: flex; gap: 10px; }

/* PROGRESS TRAIL */
.trail {
  background: var(--card);
  border-bottom: 1px solid var(--border);
  padding: 0 32px;
  display: flex;
  align-items: center;
  gap: 0;
  height: 44px;
  overflow-x: auto;
}
.trail-step {
  display: flex; align-items: center; gap: 8px;
  font-family: var(--font-head);
  font-size: 11px;
  color: var(--txt3);
  padding: 0 14px;
  height: 100%;
  border-right: 1px solid var(--border);
  cursor: pointer;
  transition: all .2s;
  white-space: nowrap;
  user-select: none;
}
.trail-step:first-child { padding-left: 0; }
.trail-step:hover { color: var(--txt2); }
.trail-step.active { color: var(--accent); }
.trail-step.done { color: var(--accent2); }
.step-num {
  width: 20px; height: 20px;
  border-radius: 50%;
  border: 1px solid currentColor;
  display: flex; align-items: center; justify-content: center;
  font-size: 10px;
}
.trail-step.done .step-num { background: var(--accent2); border-color: var(--accent2); color: var(--navy); }
.trail-step.active .step-num { background: var(--accent); border-color: var(--accent); color: var(--navy); }

/* LAYOUT */
.layout { display: flex; min-height: calc(100vh - 108px); }

.sidebar {
  width: 280px;
  background: var(--surface);
  border-right: 1px solid var(--border);
  padding: 20px 0;
  flex-shrink: 0;
  overflow-y: auto;
}
.sidebar-section { margin-bottom: 8px; }
.sidebar-title {
  font-family: var(--font-head);
  font-size: 10px;
  font-weight: 700;
  color: var(--txt3);
  letter-spacing: 0.12em;
  text-transform: uppercase;
  padding: 0 20px;
  margin-bottom: 6px;
}
.nav-btn {
  display: flex; align-items: center; gap: 12px;
  padding: 10px 20px;
  width: 100%;
  background: none; border: none;
  color: var(--txt2); cursor: pointer;
  font-family: var(--font-body); font-size: 13px;
  text-align: left;
  transition: all .15s;
  border-left: 3px solid transparent;
}
.nav-btn:hover { background: var(--card); color: var(--txt); }
.nav-btn.active { background: var(--card); color: var(--accent); border-left-color: var(--accent); }
.nav-icon { width: 18px; text-align: center; font-size: 14px; }

/* MAIN CONTENT */
.main { flex: 1; padding: 28px 32px; overflow-y: auto; }

/* PANELS */
.panel { display: none; }
.panel.active { display: block; }

/* PAGE TITLE */
.page-title {
  margin-bottom: 24px;
  padding-bottom: 16px;
  border-bottom: 1px solid var(--border);
}
.page-title h1 {
  font-family: var(--font-head);
  font-size: 20px;
  font-weight: 700;
  color: var(--txt);
  margin-bottom: 4px;
}
.page-title p { font-size: 13px; color: var(--txt2); }

/* CARD */
.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  margin-bottom: 16px;
  overflow: hidden;
}
.card-header {
  background: var(--card);
  padding: 12px 18px;
  display: flex; align-items: center; justify-content: space-between;
  cursor: pointer;
  user-select: none;
  border-bottom: 1px solid var(--border);
}
.card-header:hover { background: var(--navy3); }
.card-title {
  font-family: var(--font-head);
  font-size: 12px;
  font-weight: 700;
  color: var(--accent);
  display: flex; align-items: center; gap: 10px;
}
.card-badge {
  font-size: 10px;
  background: var(--navy3);
  color: var(--txt2);
  padding: 2px 8px;
  border-radius: 4px;
  font-family: var(--font-head);
}
.chevron { color: var(--txt3); font-size: 12px; transition: transform .2s; }
.chevron.open { transform: rotate(180deg); }
.card-body { padding: 18px; }

/* GRID */
.grid { display: grid; gap: 12px; }
.g2 { grid-template-columns: repeat(2, 1fr); }
.g3 { grid-template-columns: repeat(3, 1fr); }
.g4 { grid-template-columns: repeat(4, 1fr); }
.g2-1 { grid-template-columns: 2fr 1fr; }
@media (max-width: 900px) {
  .g3, .g4 { grid-template-columns: repeat(2, 1fr); }
  .g2, .g2-1 { grid-template-columns: 1fr; }
}

/* FORM ELEMENTS */
.field { display: flex; flex-direction: column; gap: 5px; }
.field label {
  font-size: 11px;
  font-family: var(--font-head);
  color: var(--txt2);
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.field input, .field select, .field textarea {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 6px;
  color: var(--txt);
  padding: 8px 11px;
  font-size: 13px;
  font-family: var(--font-body);
  width: 100%;
  transition: border-color .15s;
}
.field input:focus, .field select:focus, .field textarea:focus {
  outline: none;
  border-color: var(--steel);
}
.field select option { background: var(--card); color: var(--txt); }
.field textarea { resize: vertical; min-height: 60px; }
.field-hint { font-size: 11px; color: var(--txt3); }

/* SECTION DIVIDER */
.sec-div {
  display: flex; align-items: center; gap: 10px;
  margin: 18px 0 14px;
}
.sec-div span {
  font-family: var(--font-head);
  font-size: 10px;
  font-weight: 700;
  color: var(--accent2);
  letter-spacing: 0.1em;
  text-transform: uppercase;
  white-space: nowrap;
}
.sec-div::before, .sec-div::after {
  content: '';
  flex: 1;
  height: 1px;
  background: var(--border);
}

/* BUTTONS */
.btn {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 9px 18px;
  border: none; border-radius: 7px;
  font-family: var(--font-body); font-size: 13px; font-weight: 500;
  cursor: pointer; transition: all .15s;
}
.btn-primary { background: var(--steel); color: var(--txt); }
.btn-primary:hover { background: var(--navy3); }
.btn-success { background: var(--green); color: white; }
.btn-success:hover { opacity: .88; }
.btn-warning { background: var(--accent); color: var(--navy); font-weight: 600; }
.btn-warning:hover { opacity: .88; }
.btn-ghost { background: var(--card); color: var(--txt2); border: 1px solid var(--border); }
.btn-ghost:hover { color: var(--txt); border-color: var(--txt3); }
.btn-danger { background: var(--red); color: white; }
.btn-actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 8px; }

/* TABLE */
.tbl-wrap { overflow-x: auto; margin-top: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th {
  background: var(--card);
  color: var(--accent);
  font-family: var(--font-head);
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  padding: 9px 12px;
  text-align: left;
  border-bottom: 1px solid var(--border);
}
td {
  padding: 9px 12px;
  border-bottom: 1px solid var(--border);
  color: var(--txt2);
  vertical-align: top;
}
tr:last-child td { border-bottom: none; }
tr:hover td { background: var(--card); color: var(--txt); }

/* STATUS TAGS */
.tag {
  display: inline-block;
  font-size: 10px;
  font-family: var(--font-head);
  padding: 2px 8px; border-radius: 4px;
  font-weight: 700;
}
.tag-ok { background: rgba(27,127,79,.2); color: #2EBBAA; }
.tag-warn { background: rgba(244,166,42,.15); color: var(--accent); }
.tag-info { background: rgba(30,87,153,.3); color: #7BB3E0; }

/* BOM PREVIEW TABLE */
.bom-row { display: grid; grid-template-columns: 2fr 3fr 2fr; gap: 0; }
.bom-row > div {
  padding: 7px 12px;
  border-bottom: 1px solid var(--border);
  font-size: 12px;
}
.bom-header > div {
  font-family: var(--font-head);
  font-size: 10px;
  color: var(--accent);
  letter-spacing: 0.08em;
  text-transform: uppercase;
  background: var(--card);
  font-weight: 700;
}

/* SUMMARY STAT CARDS */
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 24px; }
.stat {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 14px 16px;
}
.stat-val {
  font-family: var(--font-head);
  font-size: 22px;
  font-weight: 700;
  color: var(--accent);
  line-height: 1;
  margin-bottom: 4px;
}
.stat-lbl { font-size: 11px; color: var(--txt2); }

/* INFO BOX */
.info-box {
  background: rgba(30,87,153,.15);
  border: 1px solid rgba(30,87,153,.4);
  border-radius: 7px;
  padding: 12px 14px;
  font-size: 12px;
  color: var(--txt2);
  margin-bottom: 14px;
}
.info-box strong { color: var(--accent2); }

/* SCROLLBAR */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--navy); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
</style>
</head>
<body>

<!-- TOP BAR -->
<div class="topbar">
  <div class="brand">
    <div class="brand-icon">MMC</div>
    <div class="brand-text">
      <div class="co">Mulijbhai Madhivani Co. Ltd</div>
      <div class="sub">Steel Division  ·  Engineering BOM & Procurement System</div>
    </div>
  </div>
  <div class="topbar-actions">
    <button class="btn btn-ghost" onclick="clearAll()">Clear All</button>
    <button class="btn btn-success" onclick="saveBOM()">Save BOM</button>
    <button class="btn btn-warning" onclick="exportExcel()">Export Excel</button>
  </div>
</div>

<!-- PROGRESS TRAIL -->
<div class="trail">
  <div class="trail-step active" id="t1" onclick="showPanel('p-select')">
    <span class="step-num">1</span>Mill Selection
  </div>
  <div class="trail-step" id="t2" onclick="showPanel('p-drive')">
    <span class="step-num">2</span>Drive System
  </div>
  <div class="trail-step" id="t3" onclick="showPanel('p-bearing')">
    <span class="step-num">3</span>Bearings
  </div>
  <div class="trail-step" id="t4" onclick="showPanel('p-fastener')">
    <span class="step-num">4</span>Fasteners & Seals
  </div>
  <div class="trail-step" id="t5" onclick="showPanel('p-material')">
    <span class="step-num">5</span>Materials & Lubrication
  </div>
  <div class="trail-step" id="t6" onclick="showPanel('p-quality')">
    <span class="step-num">6</span>QA & Standards
  </div>
  <div class="trail-step" id="t7" onclick="showPanel('p-bom')">
    <span class="step-num">7</span>BOM Preview
  </div>
</div>

<!-- LAYOUT -->
<div class="layout">

<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sidebar-section">
    <div class="sidebar-title">Workflow</div>
    <button class="nav-btn active" onclick="showPanel('p-select')"><span class="nav-icon">🏭</span>Mill & Subcomponent</button>
    <button class="nav-btn" onclick="showPanel('p-drive')"><span class="nav-icon">⚙️</span>Drive System</button>
    <button class="nav-btn" onclick="showPanel('p-bearing')"><span class="nav-icon">🔵</span>Bearing Engineering</button>
    <button class="nav-btn" onclick="showPanel('p-fastener')"><span class="nav-icon">🔩</span>Fasteners & Seals</button>
    <button class="nav-btn" onclick="showPanel('p-material')"><span class="nav-icon">🛢️</span>Materials & Lubrication</button>
    <button class="nav-btn" onclick="showPanel('p-quality')"><span class="nav-icon">✅</span>QA & Standards</button>
    <button class="nav-btn" onclick="showPanel('p-bom')"><span class="nav-icon">📋</span>BOM Preview & Export</button>
  </div>
  <div class="sidebar-section" style="margin-top:20px;">
    <div class="sidebar-title">Saved BOMs</div>
    <div id="saved-list" style="padding:0 20px;font-size:12px;color:var(--txt3);">No BOMs saved yet</div>
  </div>
</div>

<!-- MAIN -->
<div class="main">

<!-- ═══════════ PANEL 1: MILL SELECTION ═══════════ -->
<div class="panel active" id="p-select">
  <div class="page-title">
    <h1>Mill Section & Subcomponent Selection</h1>
    <p>Identify the mill area, subcomponent, and assembly details</p>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Project Identification <span class="card-badge">Required</span></div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field">
          <label>BOM Reference No.</label>
          <input id="bomRef" placeholder="e.g. BOM-RF-2025-001">
        </div>
        <div class="field">
          <label>Project / Work Order No.</label>
          <input id="woNo" placeholder="e.g. WO-2025-0124">
        </div>
        <div class="field">
          <label>Revision</label>
          <select id="rev">
            <option>Rev 0 – Initial Issue</option>
            <option>Rev 1 – First Revision</option>
            <option>Rev 2</option>
            <option>Rev 3</option>
            <option>Rev A – As-Built</option>
          </select>
        </div>
        <div class="field">
          <label>Prepared By</label>
          <input id="prepBy" placeholder="Engineer name">
        </div>
        <div class="field">
          <label>Approved By</label>
          <input id="appBy" placeholder="Approver name">
        </div>
        <div class="field">
          <label>Date</label>
          <input id="bomDate" type="date">
        </div>
        <div class="field">
          <label>Plant / Mill Name</label>
          <input id="plantName" placeholder="e.g. MMC Steel Rolling Mill – Kampala">
        </div>
        <div class="field">
          <label>Asset Tag / Equipment No.</label>
          <input id="assetTag" placeholder="e.g. EQ-RF-001">
        </div>
        <div class="field">
          <label>Priority</label>
          <select id="priority">
            <option>Critical – Immediate</option>
            <option>Major – Within 48hrs</option>
            <option>Normal – Planned PM</option>
            <option>Minor – Next Outage</option>
          </select>
        </div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Mill Section & Subcomponent</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field">
          <label>Mill Section</label>
          <select id="mill" onchange="loadSub()">
            <option value="">— Select Section —</option>
            <option value="reheat">Reheating Furnace</option>
            <option value="rough">Roughing Mill</option>
            <option value="inter">Intermediate Mill</option>
            <option value="finish">Finishing Mill</option>
            <option value="cool">Cooling Bed</option>
            <option value="shear">Cold Shear / Dividing</option>
            <option value="conveyor">Conveyor & Transfer</option>
            <option value="hpu">Hydraulic Power Unit</option>
            <option value="elec">Electrical & Automation</option>
          </select>
        </div>
        <div class="field">
          <label>Subcomponent</label>
          <select id="sub" onchange="loadDetail()">
            <option value="">— Select Subcomponent —</option>
          </select>
        </div>
        <div class="field">
          <label>Specific Component</label>
          <select id="detail">
            <option value="">— Select Component —</option>
          </select>
        </div>
        <div class="field">
          <label>Tag Number / Position</label>
          <input id="tagNo" placeholder="e.g. RF-01-PDR-001">
        </div>
        <div class="field">
          <label>Assembly Qty Required</label>
          <input id="qty" type="number" min="1" value="1">
        </div>
        <div class="field">
          <label>Criticality</label>
          <select id="crit">
            <option>Critical – Production Stop</option>
            <option>Major – Reduced Output</option>
            <option>Minor – Workaround Exists</option>
            <option>Non-Critical</option>
          </select>
        </div>
        <div class="field">
          <label>Location in Plant</label>
          <input id="location" placeholder="e.g. Bay 2, Ground Floor, NW Corner">
        </div>
        <div class="field">
          <label>Installation Date</label>
          <input id="installDate" type="date">
        </div>
        <div class="field">
          <label>Last Overhaul Date</label>
          <input id="overhaulDate" type="date">
        </div>
      </div>
      <div class="sec-div"><span>Description</span></div>
      <div class="field">
        <label>Component Description / Scope of Work</label>
        <textarea id="description" placeholder="Describe the component, its function, and the scope of work for this BOM (e.g. replacement of drive-end bearing on Roughing Mill Stand R1 top work roll during planned maintenance shutdown)"></textarea>
      </div>
    </div>
  </div>

  <div class="btn-actions">
    <button class="btn btn-primary" onclick="showPanel('p-drive')">Next: Drive System →</button>
  </div>
</div>

<!-- ═══════════ PANEL 2: DRIVE SYSTEM ═══════════ -->
<div class="panel" id="p-drive">
  <div class="page-title">
    <h1>Drive System Engineering Specification</h1>
    <p>Full technical parameters for the selected drive type</p>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Drive Selection</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field">
          <label>Drive Type</label>
          <select id="drive" onchange="generateDriveSpec()">
            <option value="">— Select Drive Type —</option>
            <option value="chain">Chain Drive</option>
            <option value="gear">Gear Drive (Gearbox)</option>
            <option value="belt">Belt Drive</option>
            <option value="hydraulic">Hydraulic Drive / Cylinder</option>
            <option value="rack">Rack & Pinion Drive</option>
            <option value="screw">Ball / Acme Screw Drive</option>
            <option value="pneumatic">Pneumatic Drive / Cylinder</option>
            <option value="direct">Direct Motor Drive</option>
          </select>
        </div>
        <div class="field">
          <label>Power Input (kW)</label>
          <input id="drv_kw" type="number" placeholder="e.g. 75">
        </div>
        <div class="field">
          <label>Input Speed (RPM)</label>
          <input id="drv_rpm_in" type="number" placeholder="e.g. 1480">
        </div>
        <div class="field">
          <label>Output Speed (RPM)</label>
          <input id="drv_rpm_out" type="number" placeholder="e.g. 118">
        </div>
        <div class="field">
          <label>Gear / Speed Ratio</label>
          <input id="drv_ratio" placeholder="e.g. 12.5 : 1">
        </div>
        <div class="field">
          <label>Output Torque (Nm)</label>
          <input id="drv_torq" type="number" placeholder="e.g. 25000">
        </div>
      </div>
    </div>
  </div>

  <div id="drive-spec-panel"></div>

  <!-- MOTOR -->
  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Drive Motor Specification</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g4">
        <div class="field"><label>Motor Manufacturer</label>
          <select id="m_mfr"><option>WEG</option><option>ABB</option><option>Siemens</option><option>Crompton Greaves</option><option>Kirloskar</option><option>Toshiba</option><option>GE</option><option>BHEL</option></select></div>
        <div class="field"><label>Motor Model / Frame</label><input id="m_model" placeholder="e.g. IEC 280M / NEMA 449T"></div>
        <div class="field"><label>Rated Power (kW)</label><input id="m_kw" type="number" placeholder="e.g. 75"></div>
        <div class="field"><label>Voltage (V)</label>
          <select id="m_volt"><option>400V</option><option>415V</option><option>440V</option><option>380V</option><option>460V</option><option>525V</option><option>690V</option><option>3300V</option><option>6600V</option></select></div>
        <div class="field"><label>Full Load Current (A)</label><input id="m_fla" type="number" placeholder="e.g. 140"></div>
        <div class="field"><label>Synchronous Speed (RPM)</label>
          <select id="m_sync"><option>3000</option><option>1500</option><option>1000</option><option>750</option><option>600</option><option>500</option></select></div>
        <div class="field"><label>Full Load Speed (RPM)</label><input id="m_flrpm" type="number" placeholder="e.g. 1480"></div>
        <div class="field"><label>Power Factor (cosφ)</label><input id="m_pf" type="number" step="0.01" placeholder="e.g. 0.87"></div>
        <div class="field"><label>Efficiency (%)</label><input id="m_eff" type="number" placeholder="e.g. 94.5"></div>
        <div class="field"><label>IP Rating</label>
          <select id="m_ip"><option>IP55</option><option>IP54</option><option>IP56</option><option>IP65</option><option>IP66</option><option>IP67</option></select></div>
        <div class="field"><label>Insulation Class</label>
          <select id="m_ins"><option>Class F (155°C)</option><option>Class H (180°C)</option><option>Class B (130°C)</option></select></div>
        <div class="field"><label>Enclosure Type</label>
          <select id="m_enc"><option>TEFC (Totally Enclosed Fan Cooled)</option><option>ODP (Open Drip Proof)</option><option>TENV</option><option>Explosion Proof</option><option>Brake Motor</option></select></div>
        <div class="field"><label>Service Factor</label><input id="m_sf" type="number" step="0.1" placeholder="e.g. 1.15"></div>
        <div class="field"><label>Foot Bolt Size</label>
          <select id="m_fbolt"><option>M16</option><option>M20</option><option>M24</option><option>M12</option></select></div>
        <div class="field"><label>Foot Bolt Grade</label>
          <select id="m_fbgrade"><option>8.8</option><option>10.9</option><option>4.6</option></select></div>
        <div class="field"><label>Foot Bolt Qty</label><input id="m_fbqty" type="number" placeholder="e.g. 4"></div>
      </div>
    </div>
  </div>

  <div class="btn-actions">
    <button class="btn btn-ghost" onclick="showPanel('p-select')">← Back</button>
    <button class="btn btn-primary" onclick="showPanel('p-bearing')">Next: Bearings →</button>
  </div>
</div>

<!-- ═══════════ PANEL 3: BEARINGS ═══════════ -->
<div class="panel" id="p-bearing">
  <div class="page-title">
    <h1>Bearing Engineering Specification</h1>
    <p>Full bearing detail: designation, fits, clearance, lubrication, life calculation</p>
  </div>

  <div class="info-box">
    <strong>Engineering Note:</strong> Specify bearings for each shaft position (DE = Drive End, NDE = Non-Drive End). Include clearance class, shaft/housing fits, installation method, and regreasing data for traceability and PM scheduling.
  </div>

  <!-- DE BEARING -->
  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Bearing – Drive End (DE) <span class="card-badge">Position 1</span></div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g4">
        <div class="field"><label>Bearing Type</label>
          <select id="de_type">
            <option>Deep Groove Ball Bearing (DGBB)</option>
            <option>Spherical Roller Bearing (SRB)</option>
            <option>Cylindrical Roller Bearing (CRB)</option>
            <option>Taper Roller Bearing (TRB)</option>
            <option>Angular Contact Ball Bearing (ACBB)</option>
            <option>Thrust Ball Bearing</option>
            <option>Toroidal Roller (CARB)</option>
            <option>Needle Roller Bearing</option>
          </select></div>
        <div class="field"><label>Designation / Part No.</label><input id="de_desig" placeholder="e.g. 23240 CC/W33, 6316 C3"></div>
        <div class="field"><label>Bore Diameter (mm)</label><input id="de_bore" type="number" placeholder="e.g. 200"></div>
        <div class="field"><label>Outside Diameter (mm)</label><input id="de_od" type="number" placeholder="e.g. 360"></div>
        <div class="field"><label>Width (mm)</label><input id="de_width" type="number" placeholder="e.g. 128"></div>
        <div class="field"><label>Clearance Class</label>
          <select id="de_cl">
            <option>C3 (Increased)</option>
            <option>CN / C0 (Normal)</option>
            <option>C2 (Reduced)</option>
            <option>C4 (Greater)</option>
            <option>C5 (Greatest)</option>
          </select></div>
        <div class="field"><label>Manufacturer</label>
          <select id="de_mfr">
            <option>SKF</option><option>FAG / Schaeffler</option><option>NSK</option><option>NTN</option>
            <option>Timken</option><option>Koyo</option><option>INA</option><option>NBC (India)</option>
          </select></div>
        <div class="field"><label>Shaft Fit</label>
          <select id="de_shfit">
            <option>k6 (Normal Interference)</option>
            <option>m6 (Light Press)</option>
            <option>js6 (Transition)</option>
            <option>h6 (Clearance)</option>
            <option>n6 (Heavy Interference)</option>
          </select></div>
        <div class="field"><label>Housing Fit</label>
          <select id="de_hfit">
            <option>H7 (Clearance / Sliding)</option>
            <option>J7 (Transition)</option>
            <option>K7 (Transition)</option>
            <option>M7 (Interference)</option>
            <option>N7 (Interference)</option>
          </select></div>
        <div class="field"><label>Installation Method</label>
          <select id="de_inst">
            <option>Thermal / Heat Shrink</option>
            <option>Hydraulic Nut</option>
            <option>Press Fit (Arbor Press)</option>
            <option>Adapter Sleeve</option>
            <option>Withdrawal Sleeve</option>
            <option>Lock Nut & Washer</option>
          </select></div>
        <div class="field"><label>Seal / Shield Type</label>
          <select id="de_seal">
            <option>Open (No Seal)</option>
            <option>2RS (Rubber Seal Both Sides)</option>
            <option>2Z (Metal Shield Both Sides)</option>
            <option>RS (Single Rubber Seal)</option>
            <option>Labyrinth Seal (External)</option>
            <option>Cassette Seal</option>
            <option>Taconite Seal</option>
          </select></div>
        <div class="field"><label>Grease Type</label>
          <select id="de_grease">
            <option>SKF LGMT 3 (Lithium Complex)</option>
            <option>SKF LGHB 2 (High Temp)</option>
            <option>Shell Alvania RL 3</option>
            <option>Mobilux EP 2</option>
            <option>Castrol Spheerol EPL 2</option>
            <option>Klüber Isoflex NBU 15</option>
          </select></div>
        <div class="field"><label>Lube Nipple Size</label>
          <select id="de_nipple">
            <option>M8×1</option><option>M6×1</option><option>M10×1</option>
            <option>1/4" PTF</option><option>Not Fitted – Sealed</option>
          </select></div>
        <div class="field"><label>Dynamic Load Rating C (kN)</label><input id="de_C" type="number" placeholder="from catalogue"></div>
        <div class="field"><label>Static Load Rating C0 (kN)</label><input id="de_C0" type="number" placeholder="from catalogue"></div>
        <div class="field"><label>Applied Radial Load Fr (kN)</label><input id="de_Fr" type="number" placeholder="e.g. 85"></div>
        <div class="field"><label>Applied Axial Load Fa (kN)</label><input id="de_Fa" type="number" placeholder="e.g. 12"></div>
        <div class="field"><label>Regreasing Interval (hrs)</label><input id="de_lubint" type="number" placeholder="e.g. 500"></div>
        <div class="field"><label>Grease Qty per Regrease (g)</label><input id="de_lubqty" type="number" placeholder="e.g. 35"></div>
      </div>
    </div>
  </div>

  <!-- NDE BEARING -->
  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Bearing – Non-Drive End (NDE) <span class="card-badge">Position 2</span></div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g4">
        <div class="field"><label>Bearing Type</label>
          <select id="nde_type">
            <option>Deep Groove Ball Bearing (DGBB)</option>
            <option>Spherical Roller Bearing (SRB)</option>
            <option>Cylindrical Roller Bearing (CRB)</option>
            <option>Taper Roller Bearing (TRB)</option>
            <option>Angular Contact Ball Bearing (ACBB)</option>
          </select></div>
        <div class="field"><label>Designation / Part No.</label><input id="nde_desig" placeholder="e.g. 6316 C3"></div>
        <div class="field"><label>Bore Diameter (mm)</label><input id="nde_bore" type="number" placeholder="e.g. 80"></div>
        <div class="field"><label>Outside Diameter (mm)</label><input id="nde_od" type="number" placeholder="e.g. 170"></div>
        <div class="field"><label>Width (mm)</label><input id="nde_width" type="number" placeholder="e.g. 39"></div>
        <div class="field"><label>Clearance Class</label>
          <select id="nde_cl">
            <option>C3 (Increased)</option>
            <option>CN / C0 (Normal)</option>
            <option>C2 (Reduced)</option>
            <option>C4 (Greater)</option>
          </select></div>
        <div class="field"><label>Manufacturer</label>
          <select id="nde_mfr">
            <option>SKF</option><option>FAG / Schaeffler</option><option>NSK</option><option>NTN</option>
            <option>Timken</option><option>Koyo</option><option>INA</option><option>NBC (India)</option>
          </select></div>
        <div class="field"><label>Shaft Fit</label>
          <select id="nde_shfit">
            <option>k6 (Normal Interference)</option><option>m6 (Light Press)</option>
            <option>js6 (Transition)</option><option>h6 (Clearance)</option>
          </select></div>
        <div class="field"><label>Housing Fit</label>
          <select id="nde_hfit">
            <option>H7 (Clearance / Sliding)</option><option>J7 (Transition)</option>
            <option>K7 (Transition)</option><option>M7 (Interference)</option>
          </select></div>
        <div class="field"><label>Seal / Shield Type</label>
          <select id="nde_seal">
            <option>Open (No Seal)</option><option>2RS (Rubber Seal Both Sides)</option>
            <option>2Z (Metal Shield Both Sides)</option><option>Labyrinth Seal</option>
          </select></div>
        <div class="field"><label>Grease Type</label>
          <select id="nde_grease">
            <option>SKF LGMT 3 (Lithium Complex)</option><option>Shell Alvania RL 3</option>
            <option>Mobilux EP 2</option><option>Klüber Isoflex NBU 15</option>
          </select></div>
        <div class="field"><label>Regreasing Interval (hrs)</label><input id="nde_lubint" type="number" placeholder="e.g. 500"></div>
        <div class="field"><label>Grease Qty per Regrease (g)</label><input id="nde_lubqty" type="number" placeholder="e.g. 25"></div>
      </div>
    </div>
  </div>

  <!-- SHAFT -->
  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Shaft Specification</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g4">
        <div class="field"><label>Shaft Diameter (mm)</label><input id="shaft_dia" type="number" placeholder="e.g. 80"></div>
        <div class="field"><label>Shaft Length (mm)</label><input id="shaft_len" type="number" placeholder="e.g. 650"></div>
        <div class="field"><label>Shaft Material</label>
          <select id="shaft_mat">
            <option>EN8 (080M40)</option><option>EN24 (817M40)</option><option>EN36 (655M13)</option>
            <option>C45 (080M46)</option><option>4140 Chromoly</option><option>4340 Alloy Steel</option>
            <option>304 Stainless</option><option>316 Stainless</option>
          </select></div>
        <div class="field"><label>Surface Treatment</label>
          <select id="shaft_surf">
            <option>Induction Hardened</option><option>Case Hardened</option>
            <option>Nitrided</option><option>Hard Chrome Plated</option>
            <option>As Machined (No Treatment)</option><option>Zinc Plated</option>
          </select></div>
        <div class="field"><label>Keyway Width × Height (mm)</label><input id="keyway" placeholder="e.g. 22×14"></div>
        <div class="field"><label>Keyway Length (mm)</label><input id="keylen" type="number" placeholder="e.g. 90"></div>
        <div class="field"><label>Key Material</label>
          <select id="key_mat">
            <option>C45 Steel (Parallel Key)</option><option>Stainless Steel 304</option>
            <option>EN8 Steel</option><option>Woodruff Key – Steel</option>
          </select></div>
        <div class="field"><label>Shaft Surface Hardness (HRC)</label><input id="shaft_hrd" placeholder="e.g. 55–60 HRC"></div>
      </div>
    </div>
  </div>

  <div class="btn-actions">
    <button class="btn btn-ghost" onclick="showPanel('p-drive')">← Back</button>
    <button class="btn btn-primary" onclick="showPanel('p-fastener')">Next: Fasteners →</button>
  </div>
</div>

<!-- ═══════════ PANEL 4: FASTENERS & SEALS ═══════════ -->
<div class="panel" id="p-fastener">
  <div class="page-title">
    <h1>Fasteners, Seals & Gaskets BOM</h1>
    <p>All bolts, nuts, washers, seals, and gaskets with full engineering detail</p>
  </div>

  <!-- STRUCTURAL BOLTS -->
  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Structural / Mounting Bolts</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g4">
        <div class="field"><label>Bolt Type</label>
          <select id="bt_type">
            <option>Hex Head Bolt (ISO 4014)</option>
            <option>Hex Head Cap Screw (ISO 4017) – Full Thread</option>
            <option>Socket Head Cap Screw (SHCS – ISO 4762)</option>
            <option>Stud Bolt (ASME B16.5)</option>
            <option>Foundation / Anchor Bolt</option>
            <option>Flange Bolt</option>
            <option>Carriage Bolt</option>
            <option>Eye Bolt</option>
            <option>U-Bolt</option>
            <option>T-Slot Bolt</option>
          </select></div>
        <div class="field"><label>Nominal Size</label>
          <select id="bt_size">
            <option>M10</option><option>M12</option><option>M16</option><option>M20</option>
            <option>M24</option><option>M27</option><option>M30</option><option>M36</option>
            <option>M42</option><option>M48</option><option>M6</option><option>M8</option>
          </select></div>
        <div class="field"><label>Property Class (Grade)</label>
          <select id="bt_grade">
            <option>8.8</option><option>10.9</option><option>12.9</option>
            <option>4.6</option><option>5.6</option>
            <option>A2-70 (Stainless 304)</option><option>A4-80 (Stainless 316)</option>
            <option>Grade B7 (ASTM – High Temp)</option>
          </select></div>
        <div class="field"><label>Thread Standard</label>
          <select id="bt_thread">
            <option>ISO Metric Coarse (ISO 262)</option>
            <option>ISO Metric Fine</option>
            <option>UNC (Unified National Coarse)</option>
            <option>UNF (Unified National Fine)</option>
            <option>BSW (British Standard Whitworth)</option>
          </select></div>
        <div class="field"><label>Bolt Length (mm)</label>
          <select id="bt_len">
            <option>30</option><option>35</option><option>40</option><option>45</option><option>50</option>
            <option>55</option><option>60</option><option>65</option><option>70</option><option>75</option>
            <option>80</option><option>90</option><option>100</option><option>110</option><option>120</option>
            <option>130</option><option>140</option><option>150</option><option>160</option><option>180</option>
            <option>200</option><option>220</option><option>250</option><option>300</option>
          </select></div>
        <div class="field"><label>Material / Coating</label>
          <select id="bt_mat">
            <option>Carbon Steel – Zinc Electroplated</option>
            <option>Carbon Steel – Hot Dip Galvanized (HDG)</option>
            <option>Carbon Steel – Mechanical Zinc</option>
            <option>Carbon Steel – Phosphate + Oil (Parkerized)</option>
            <option>Alloy Steel – Plain (no coating)</option>
            <option>Stainless Steel 304 – Passivated</option>
            <option>Stainless Steel 316 – Passivated</option>
            <option>PTFE Coated (Teflon)</option>
          </select></div>
        <div class="field"><label>Tightening Torque (Nm)</label><input id="bt_torque" type="number" placeholder="e.g. 150 Nm (M20 / 8.8)"></div>
        <div class="field"><label>Qty per Joint</label><input id="bt_jqty" type="number" placeholder="e.g. 4"></div>
        <div class="field"><label>Number of Joints</label><input id="bt_joints" type="number" placeholder="e.g. 2"></div>
        <div class="field"><label>Total Bolt Qty</label><input id="bt_total" type="number" placeholder="= Qty per Joint × Joints"></div>
      </div>
      <div class="sec-div"><span>Nut & Washer</span></div>
      <div class="grid g4">
        <div class="field"><label>Nut Type</label>
          <select id="nut_type">
            <option>Hex Nut (ISO 4032)</option>
            <option>Heavy Hex Nut</option>
            <option>Nyloc Nut (ISO 7042)</option>
            <option>Castle Nut (ISO 7035)</option>
            <option>Flange Nut (ISO 4161)</option>
            <option>Prevailing Torque Lock Nut</option>
            <option>Wing Nut</option>
          </select></div>
        <div class="field"><label>Nut Grade</label>
          <select id="nut_grade">
            <option>8 (ISO)</option><option>10 (ISO)</option><option>12 (ISO)</option>
            <option>A2-70 (SS)</option><option>A4-80 (SS)</option>
          </select></div>
        <div class="field"><label>Washer Type</label>
          <select id="wash_type">
            <option>Flat Washer (ISO 7089)</option>
            <option>Spring Lock Washer (Split Ring)</option>
            <option>Nord-Lock Wedge Washer</option>
            <option>Belleville / Disc Spring Washer</option>
            <option>Hardened Structural Washer (F436)</option>
            <option>Tooth Lock Washer (Internal)</option>
            <option>Tooth Lock Washer (External)</option>
          </select></div>
        <div class="field"><label>Locking Method</label>
          <select id="lock_mth">
            <option>Thread Lock Adhesive (Loctite 243)</option>
            <option>Thread Lock Adhesive (Loctite 270 – High Strength)</option>
            <option>Nord-Lock Washer</option>
            <option>Castle Nut + Split Pin</option>
            <option>Tab Washer + Bend Tab</option>
            <option>Lock Wire</option>
            <option>No Locking Required</option>
          </select></div>
      </div>
    </div>
  </div>

  <!-- SEALS -->
  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Seals, O-Rings & Gaskets</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g4">
        <div class="field"><label>Primary Seal Type</label>
          <select id="seal_type">
            <option>Radial Lip Seal (Oil Seal)</option>
            <option>V-Ring Seal (Freudenberg)</option>
            <option>Cassette Seal</option>
            <option>Mechanical Face Seal (Duo-Cone)</option>
            <option>Labyrinth Seal</option>
            <option>Felt Seal</option>
            <option>Taconite Seal (Heavy Duty)</option>
            <option>O-Ring (Static)</option>
            <option>O-Ring (Dynamic)</option>
          </select></div>
        <div class="field"><label>Seal Material</label>
          <select id="seal_mat">
            <option>Nitrile (NBR) – General Use</option>
            <option>Viton (FKM) – High Temp / Chemical</option>
            <option>EPDM – Steam / Water</option>
            <option>Silicone – High Temp</option>
            <option>PTFE – Chemical Resistance</option>
            <option>Polyurethane (PU) – Abrasion Resistance</option>
            <option>Neoprene</option>
          </select></div>
        <div class="field"><label>Seal ID (mm)</label><input id="seal_id" type="number" placeholder="e.g. 80"></div>
        <div class="field"><label>Seal OD (mm)</label><input id="seal_od" type="number" placeholder="e.g. 100"></div>
        <div class="field"><label>Seal Width / Height (mm)</label><input id="seal_w" type="number" placeholder="e.g. 13"></div>
        <div class="field"><label>Seal Part No. / Kit PN</label><input id="seal_pn" placeholder="e.g. Parker Kit 3000297"></div>
        <div class="field"><label>Seal Manufacturer</label>
          <select id="seal_mfr">
            <option>SKF</option><option>FAG / Schaeffler</option><option>Parker / Hannifin</option>
            <option>Freudenberg / Simrit</option><option>Trelleborg</option><option>Garlock</option>
          </select></div>
        <div class="field"><label>Operating Temperature Range</label><input id="seal_temp" placeholder="e.g. -20°C to +120°C"></div>
        <div class="field"><label>Gasket Type</label>
          <select id="gask_type">
            <option>Full Face (Rubber / EPDM)</option>
            <option>Spiral Wound (ASME B16.20)</option>
            <option>Ring Joint (RTJ)</option>
            <option>Raised Face (RF)</option>
            <option>Compressed Fibre (CAF)</option>
            <option>PTFE Envelope</option>
            <option>Kammprofile</option>
            <option>None Required</option>
          </select></div>
        <div class="field"><label>Gasket Material / Grade</label><input id="gask_mat" placeholder="e.g. Flexitallic SS 316 / Spiral Wound"></div>
        <div class="field"><label>Seal Qty</label><input id="seal_qty" type="number" placeholder="e.g. 2"></div>
        <div class="field"><label>Gasket Qty</label><input id="gask_qty" type="number" placeholder="e.g. 4"></div>
      </div>
    </div>
  </div>

  <div class="btn-actions">
    <button class="btn btn-ghost" onclick="showPanel('p-bearing')">← Back</button>
    <button class="btn btn-primary" onclick="showPanel('p-material')">Next: Materials →</button>
  </div>
</div>

<!-- ═══════════ PANEL 5: MATERIALS & LUBRICATION ═══════════ -->
<div class="panel" id="p-material">
  <div class="page-title">
    <h1>Materials, Lubrication & Utilities</h1>
    <p>Structural materials, oil grades, grease specifications, and utility connections</p>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Housing & Structural Materials</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field"><label>Housing / Casing Material</label>
          <select id="hsg_mat">
            <option>Cast Iron GG25 (EN-GJL-250)</option>
            <option>Cast Iron GG30 (EN-GJL-300)</option>
            <option>Ductile Iron GGG40 (EN-GJS-400)</option>
            <option>Ductile Iron GGG60</option>
            <option>Cast Steel (GS-45)</option>
            <option>Fabricated Steel S275 (ASTM A36)</option>
            <option>Fabricated Steel S355 (ASTM A572)</option>
            <option>Aluminium Alloy (ADC12)</option>
          </select></div>
        <div class="field"><label>Housing Surface Treatment</label>
          <select id="hsg_surf">
            <option>Epoxy Primer + Topcoat</option>
            <option>Powder Coated</option>
            <option>Hot Dip Galvanized (HDG)</option>
            <option>Zinc Electroplated</option>
            <option>Phosphate + Oil (Parkerized)</option>
            <option>As Cast / Unpainted</option>
          </select></div>
        <div class="field"><label>Housing Wall Thickness (mm)</label><input id="hsg_thick" type="number" placeholder="e.g. 25"></div>
        <div class="field"><label>Base Plate Thickness (mm)</label><input id="base_thick" type="number" placeholder="e.g. 40"></div>
        <div class="field"><label>Anchor Bolt Size</label>
          <select id="anch_sz">
            <option>M20</option><option>M24</option><option>M30</option><option>M36</option>
            <option>M16</option><option>M42</option><option>M48</option>
          </select></div>
        <div class="field"><label>Anchor Bolt Grade</label>
          <select id="anch_gr">
            <option>4.6</option><option>8.8</option><option>5.6</option>
          </select></div>
        <div class="field"><label>Anchor Bolt Qty</label><input id="anch_qty" type="number" placeholder="e.g. 8"></div>
        <div class="field"><label>Grout Specification</label>
          <select id="grout">
            <option>Non-Shrink Cementitious Grout</option>
            <option>Epoxy Grout (Masterflow / Chemgrout)</option>
            <option>Portland Cement – Not Recommended for Dynamic Loads</option>
            <option>Not Applicable</option>
          </select></div>
        <div class="field"><label>Alignment Method</label>
          <select id="align_mth">
            <option>Laser Alignment (Prüftechnik / SKF)</option>
            <option>Dial Gauge (Reverse Indicator)</option>
            <option>Straight Edge + Feeler Gauge</option>
            <option>Not Applicable</option>
          </select></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Gearbox / Oil Lubrication</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field"><label>Oil Grade (ISO VG)</label>
          <select id="oil_grade">
            <option>ISO VG 220 (Gearbox – Standard)</option>
            <option>ISO VG 320 (Gearbox – High Load)</option>
            <option>ISO VG 460 (Gearbox – Heavy Duty)</option>
            <option>ISO VG 150</option>
            <option>ISO VG 100</option>
            <option>ISO VG 68 (Hydraulic)</option>
            <option>ISO VG 46 (Hydraulic)</option>
            <option>ISO VG 32 (Hydraulic)</option>
            <option>SAE 90 (Gear Oil)</option>
            <option>SAE 140 (Gear Oil – Heavy)</option>
            <option>AGMA 5 EP</option>
            <option>AGMA 8 EP</option>
          </select></div>
        <div class="field"><label>Oil Brand / Specification</label>
          <select id="oil_brand">
            <option>Shell Omala S2 GX (Mineral)</option>
            <option>Shell Omala S4 GX (Synthetic)</option>
            <option>Mobil SHC 630 (Synthetic)</option>
            <option>Mobilgear 600 XP (Mineral)</option>
            <option>Castrol Optigear Synthetic A 320</option>
            <option>Fuchs Renolin CLP 220</option>
            <option>Total Carter SY 220</option>
            <option>BP Energol GR-XP 220</option>
          </select></div>
        <div class="field"><label>Oil Capacity (Litres)</label><input id="oil_cap" type="number" placeholder="e.g. 18"></div>
        <div class="field"><label>Oil Change Interval (hrs)</label><input id="oil_int" type="number" placeholder="e.g. 4000"></div>
        <div class="field"><label>Oil Temperature – Normal (°C)</label><input id="oil_temp_n" type="number" placeholder="e.g. 55"></div>
        <div class="field"><label>Oil Temperature – Alarm (°C)</label><input id="oil_temp_a" type="number" placeholder="e.g. 80"></div>
        <div class="field"><label>Oil Temperature – Trip (°C)</label><input id="oil_temp_t" type="number" placeholder="e.g. 95"></div>
        <div class="field"><label>Lubrication System Type</label>
          <select id="lub_sys">
            <option>Splash / Oil Bath</option>
            <option>Forced Circulation (Pump + Filter)</option>
            <option>Oil Mist Lubrication</option>
            <option>Manual Top-Up Only</option>
          </select></div>
        <div class="field"><label>Filter Rating (µm absolute)</label>
          <select id="filt_rat">
            <option>10 µm (Servo / Critical)</option>
            <option>25 µm (Gearbox – Standard)</option>
            <option>40 µm (Gearbox – Economy)</option>
            <option>Not Applicable (Splash Lube)</option>
          </select></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Coupling Specification</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field"><label>Coupling Type</label>
          <select id="cplg_type">
            <option>Jaw / Spider Coupling (Flexible)</option>
            <option>Grid Coupling (Falk / Rexnord)</option>
            <option>Disc Coupling (High Precision)</option>
            <option>Gear Coupling (Torque-Dense)</option>
            <option>Rigid Flanged Coupling</option>
            <option>Universal / Cardan Joint</option>
            <option>Tyre Coupling (Bibby / Fenaflex)</option>
            <option>Fluid Coupling (Voith)</option>
          </select></div>
        <div class="field"><label>Coupling Manufacturer</label>
          <select id="cplg_mfr">
            <option>Rexnord</option><option>KTR</option><option>Rathi</option>
            <option>Bibby Transmissions</option><option>Lovejoy</option><option>Flender</option>
            <option>SKF</option><option>R+W</option><option>Mayr</option>
          </select></div>
        <div class="field"><label>Coupling Size / Catalogue No.</label><input id="cplg_size" placeholder="e.g. KTR Rotex 75 / Rexnord T70"></div>
        <div class="field"><label>Rated Torque (Nm)</label><input id="cplg_torq" type="number" placeholder="e.g. 5000"></div>
        <div class="field"><label>Spider / Insert Material</label>
          <select id="cplg_spi">
            <option>Polyurethane (PU) 92 ShA – Standard</option>
            <option>Polyurethane (PU) 98 ShA – Hard</option>
            <option>NBR Rubber – Flexible</option>
            <option>Hytrel – High Performance</option>
            <option>Steel Grid (Grid Coupling)</option>
            <option>Not Applicable (Disc / Rigid)</option>
          </select></div>
        <div class="field"><label>Hub Bore Diameter (mm)</label><input id="cplg_bore" type="number" placeholder="e.g. 80"></div>
        <div class="field"><label>Hub Bolt Size</label>
          <select id="cplg_bsz">
            <option>M10</option><option>M12</option><option>M16</option><option>M20</option>
          </select></div>
        <div class="field"><label>Hub Bolt Grade</label>
          <select id="cplg_bgr">
            <option>8.8</option><option>10.9</option><option>12.9</option>
          </select></div>
        <div class="field"><label>Hub Bolt Qty</label><input id="cplg_bqty" type="number" placeholder="e.g. 4"></div>
      </div>
    </div>
  </div>

  <div class="btn-actions">
    <button class="btn btn-ghost" onclick="showPanel('p-fastener')">← Back</button>
    <button class="btn btn-primary" onclick="showPanel('p-quality')">Next: QA & Standards →</button>
  </div>
</div>

<!-- ═══════════ PANEL 6: QA & STANDARDS ═══════════ -->
<div class="panel" id="p-quality">
  <div class="page-title">
    <h1>Quality Assurance & Engineering Standards</h1>
    <p>Inspection criteria, testing, applicable standards, and acceptance criteria</p>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Applicable Standards & Codes</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field"><label>Bearing Standard</label>
          <select id="qa_brg_std">
            <option>ISO 15 / DIN 625 (Ball Bearings)</option>
            <option>ISO 355 (Taper Roller)</option>
            <option>ISO 492 (Cylindrical Roller)</option>
            <option>ISO 3290 (Balls)</option>
            <option>ABMA / ANSI (American)</option>
          </select></div>
        <div class="field"><label>Bolt Standard</label>
          <select id="qa_bolt_std">
            <option>ISO 4014 / 4017 (Hex Bolts)</option>
            <option>ISO 4762 (SHCS)</option>
            <option>DIN 931 / 933</option>
            <option>ANSI B18.2.1</option>
            <option>BS 3692</option>
          </select></div>
        <div class="field"><label>Gear / Drive Standard</label>
          <select id="qa_gear_std">
            <option>ISO 6336 (Gear Strength)</option>
            <option>AGMA 2001 (USA)</option>
            <option>DIN 3990 (Germany)</option>
            <option>BS 436 (UK)</option>
            <option>ISO 1328 (Gear Accuracy)</option>
          </select></div>
        <div class="field"><label>Weld Standard (if applicable)</label>
          <select id="qa_weld">
            <option>AWS D1.1 (Structural Welding)</option>
            <option>EN ISO 5817 (Weld Quality)</option>
            <option>ASME Section IX (Pressure Vessel)</option>
            <option>Not Applicable</option>
          </select></div>
        <div class="field"><label>Surface Finish Standard</label>
          <select id="qa_surf">
            <option>ISO 1302 (Surface Texture)</option>
            <option>Ra 0.8 µm (Bearing Seats)</option>
            <option>Ra 1.6 µm (General Machined)</option>
            <option>Ra 3.2 µm (Rough Machined)</option>
          </select></div>
        <div class="field"><label>Painting Standard</label>
          <select id="qa_paint">
            <option>ISO 8501-1 (Surface Preparation)</option>
            <option>Sa 2.5 (Near White Blast)</option>
            <option>Sa 2 (Commercial Blast)</option>
            <option>Not Applicable</option>
          </select></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Inspection & Testing Requirements</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field"><label>Dimensional Inspection</label>
          <select id="qa_dim">
            <option>100% Inspection Required</option>
            <option>First Article Inspection (FAI)</option>
            <option>Random Sampling (5% per batch)</option>
            <option>As Per Drawing Only</option>
          </select></div>
        <div class="field"><label>Non-Destructive Testing (NDT)</label>
          <select id="qa_ndt">
            <option>Magnetic Particle Inspection (MPI)</option>
            <option>Ultrasonic Testing (UT)</option>
            <option>Dye Penetrant Testing (DPT)</option>
            <option>Radiographic Testing (RT)</option>
            <option>Visual Inspection Only (VT)</option>
            <option>Not Required</option>
          </select></div>
        <div class="field"><label>Material Certification</label>
          <select id="qa_cert">
            <option>Mill Test Certificate (MTC) 3.1 Required</option>
            <option>Mill Test Certificate 3.2 Required</option>
            <option>Manufacturer's Certificate Only</option>
            <option>No Certificate Required</option>
          </select></div>
        <div class="field"><label>Hardness Testing</label>
          <select id="qa_hard">
            <option>Required – Brinell (HB)</option>
            <option>Required – Rockwell (HRC)</option>
            <option>Required – Vickers (HV)</option>
            <option>Not Required</option>
          </select></div>
        <div class="field"><label>Vibration Acceptance Limit</label>
          <select id="qa_vib">
            <option>ISO 10816-3: Class I (≤ 2.3 mm/s)</option>
            <option>ISO 10816-3: Class II (≤ 4.5 mm/s)</option>
            <option>ISO 10816-3: Class III (≤ 7.1 mm/s)</option>
            <option>OEM Specification</option>
            <option>Not Measured</option>
          </select></div>
        <div class="field"><label>Alignment Acceptance (TIR)</label>
          <select id="qa_align">
            <option>≤ 0.05 mm TIR (Precision)</option>
            <option>≤ 0.10 mm TIR (Standard)</option>
            <option>≤ 0.25 mm TIR (Acceptable)</option>
            <option>Per Manufacturer Spec</option>
          </select></div>
      </div>
      <div class="sec-div"><span>Special Requirements</span></div>
      <div class="field">
        <label>Special QA Notes / Hold Points</label>
        <textarea id="qa_notes" placeholder="e.g. Bearing installation to be witnessed by QA engineer. Torque wrench calibration certificate required. First 50 hrs running monitoring with vibration analyser."></textarea>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ Procurement & Lead Time</div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">
      <div class="grid g3">
        <div class="field"><label>Preferred Supplier</label><input id="supplier" placeholder="e.g. SKF Uganda / Bearing Man Group"></div>
        <div class="field"><label>Alternate Supplier</label><input id="supplier2" placeholder="e.g. NSK East Africa / FAG Distributor"></div>
        <div class="field"><label>Estimated Lead Time (days)</label><input id="lead_time" type="number" placeholder="e.g. 21"></div>
        <div class="field"><label>Unit Cost Estimate (UGX)</label><input id="unit_cost" type="number" placeholder="e.g. 850000"></div>
        <div class="field"><label>Total Estimated Cost (UGX)</label><input id="total_cost" type="number" placeholder="Auto: Unit × Qty"></div>
        <div class="field"><label>Storage Location</label><input id="storage" placeholder="e.g. Store Rack B3-Shelf 2"></div>
        <div class="field"><label>Min Stock Level</label><input id="min_stock" type="number" placeholder="e.g. 2"></div>
        <div class="field"><label>Reorder Quantity</label><input id="reorder_qty" type="number" placeholder="e.g. 6"></div>
        <div class="field"><label>Incoterms (if imported)</label>
          <select id="inco">
            <option>Not Applicable (Local Purchase)</option>
            <option>EXW (Ex Works)</option>
            <option>FOB (Free on Board)</option>
            <option>CIF (Cost, Insurance & Freight)</option>
            <option>DAP (Delivered at Place)</option>
            <option>DDP (Delivered Duty Paid)</option>
          </select></div>
      </div>
    </div>
  </div>

  <div class="btn-actions">
    <button class="btn btn-ghost" onclick="showPanel('p-material')">← Back</button>
    <button class="btn btn-success" onclick="saveBOM()">Save BOM</button>
    <button class="btn btn-primary" onclick="showPanel('p-bom')">Preview BOM →</button>
  </div>
</div>

<!-- ═══════════ PANEL 7: BOM PREVIEW ═══════════ -->
<div class="panel" id="p-bom">
  <div class="page-title">
    <h1>BOM Preview & Export</h1>
    <p>Review the complete bill of materials before exporting to Excel</p>
  </div>

  <div class="stats" id="bom-stats">
    <div class="stat"><div class="stat-val" id="st-items">—</div><div class="stat-lbl">Total Line Items</div></div>
    <div class="stat"><div class="stat-val" id="st-cat">—</div><div class="stat-lbl">Categories</div></div>
    <div class="stat"><div class="stat-val" id="st-cost">—</div><div class="stat-lbl">Est. Cost (UGX)</div></div>
    <div class="stat"><div class="stat-val" id="st-comp">—</div><div class="stat-lbl">Completion</div></div>
  </div>

  <div id="bom-preview"></div>

  <div class="btn-actions" style="margin-top:20px;">
    <button class="btn btn-ghost" onclick="showPanel('p-quality')">← Back</button>
    <button class="btn btn-success" onclick="saveBOM()">Save BOM</button>
    <button class="btn btn-warning" style="font-size:14px;padding:11px 24px;" onclick="exportExcel()">📊 Export Full Excel BOM</button>
  </div>
</div>

</div><!-- /main -->
</div><!-- /layout -->

<script>
// ══════════════════════════════════════════════
// DATA MODEL
// ══════════════════════════════════════════════
const MILL_DATA = {
  reheat: {
    label: "Reheating Furnace",
    subs: {
      "Charging Pusher Drive": ["Motor Coupling","Pusher Ram & Guide","Hydraulic Cylinder – Bore","Gearbox Output Shaft","Roller Hearth Chain Drive","Charging Door Hydraulic Actuator"],
      "Walking Beam System": ["Hydraulic Cylinder (Main Lift)","Eccentric Bearing","Walking Beam Drive – Crank","Skid Pipe Support Saddle","Scale Pit Drag Chain"],
      "Combustion System": ["Burner Assembly – Side Wall","Combustion Air Blower Motor","Recuperator Tube Bundle","Fuel Gas Train Valve","Flame Safety Monitor"],
      "Instrumentation": ["Thermocouple – Preheat Zone","Thermocouple – Soaking Zone","Oxygen Analyser Probe","PLC I/O Module","PID Temperature Controller"],
      "Feed Table Drive": ["Roller Table Motor","Table Gearbox","Chain Drive Assembly","Idler Sprocket","Bearing Housing – Drive Side"]
    }
  },
  rough: {
    label: "Roughing Mill",
    subs: {
      "Work Roll Assembly": ["Top Work Roll Bearing (DE)","Top Work Roll Bearing (NDE)","Bottom Work Roll Bearing (DE)","Roll Chock (Top) Liner","Roll Chock (Bottom) Liner","Back-Up Roll Bearing (DE)"],
      "Main Drive Line": ["Main Drive Motor","Pinion Stand Bearing","Universal Spindle (Top) Cross Kit","Universal Spindle (Bottom) Cross Kit","Main Gear Coupling Hub","Flywheel Retaining Nut"],
      "Screwdown System": ["Screwdown Motor","Screwdown Gearbox","Screwdown Screw & Nut","Hydraulic Screwdown Cylinder","LVDT Position Sensor","Load Cell"],
      "Pinch Roll": ["Top Pinch Roll Bearing","Bottom Pinch Roll Bearing","Hydraulic Gap Cylinder","Pinch Roll Drive Motor","Pinch Roll Gearbox"],
      "Entry Guides": ["Entry Guide Box","Guide Roll Bearing","Guide Adjustment Bolt Assembly","Side Guide Liner Plate"]
    }
  },
  inter: {
    label: "Intermediate Mill",
    subs: {
      "Roll Assembly": ["Work Roll Bearing (DE)","Work Roll Bearing (NDE)","Roll Chock Liner","Back-Up Roll Bearing"],
      "Drive Line": ["Drive Motor","Mill Gearbox","Mill Spindle Cross Kit","Pinion Stand Bearing","VFD Unit"],
      "Loopers": ["Tension Looper Bearing","Looper Hydraulic Cylinder","Loop Scanner"],
      "Cooling": ["Roll Cooling Header","Inter-Stand Spray Nozzle","Cooling Water Valve"]
    }
  },
  finish: {
    label: "Finishing Mill",
    subs: {
      "Work Rolls – F1–F7": ["Work Roll Stand F1 Bearing (DE)","Work Roll Stand F2 Bearing (DE)","Work Roll Stand F3 Bearing (DE)","Work Roll Stand F4 Bearing","Work Roll Stand F5 Bearing","Work Roll Stand F6 Bearing","Work Roll Stand F7 Bearing"],
      "Hydraulic Gap Control": ["HGC Cylinder – Bore","Servo Valve","Hydraulic Accumulator","HPU Pump","Return Line Filter"],
      "Strip Cooling": ["Laminar Flow Header","Cooling Water Spray Nozzle","Strip Cooling Valve","Scale Breaker Nozzle"],
      "Crop Shear": ["Crop Shear Blade (Upper)","Crop Shear Blade (Lower)","Blade Clamp Bolt Set","Shear Drive Motor","Shear Gearbox"]
    }
  },
  cool: {
    label: "Cooling Bed",
    subs: {
      "Walking Beam / Rake Drive": ["Drive Motor – Cooling Bed","Drive Gearbox","Eccentric Drive Bearing","Rack & Pinion Module","Moving Rake Assembly"],
      "Chain Conveyor": ["Chain Drive (Left)","Chain Drive (Right)","Drive Sprocket","Idler Sprocket","Chain Tensioner"],
      "Frame & Supports": ["Cooling Bed Frame Anchor Bolts","Roller Support Bearing","End Stopper Pad","Transfer Apron Hinge"]
    }
  },
  shear: {
    label: "Cold Shear / Dividing Shear",
    subs: {
      "Blade Assembly": ["Upper Blade","Lower Blade","Upper Blade Holder Bolts","Lower Blade Holder Bolts"],
      "Drive": ["Main Shear Motor","Flywheel","Clutch & Brake","Eccentric Shaft Main Bearing","Connecting Rod Big End Bearing"],
      "Gauging": ["Fixed Length Stop","Measuring Roller Encoder","Hold-Down Hydraulic Cylinder"]
    }
  },
  conveyor: {
    label: "Conveyor & Transfer",
    subs: {
      "Chain Drive": ["Drive Chain Assembly","Drive Sprocket","Idler Sprocket","Chain Tensioner","Chain Guard"],
      "Drive System": ["Conveyor Motor","Gearbox Unit","Motor-Gearbox Coupling","VFD Panel"],
      "Frame": ["Conveyor Frame Anchor Bolts","Support Roller Bearing","Take-up Unit Bearing"]
    }
  },
  hpu: {
    label: "Hydraulic Power Unit",
    subs: {
      "Pump & Motor": ["Main Hydraulic Pump","Pump Drive Motor","Pump Coupling","Pump Mounting Bolt Set"],
      "Reservoir & Cooling": ["Oil Reservoir","Oil Cooler (Shell & Tube)","Oil Temperature Sensor","Level Gauge"],
      "Valves & Filtration": ["HP Filter Element","Return Line Filter Element","Servo Valve","Relief Valve","Pressure Transmitter","Accumulator Bladder"]
    }
  },
  elec: {
    label: "Electrical & Automation",
    subs: {
      "Drive Panels": ["VFD Unit (per stand)","DC Drive Panel","Regenerative Braking Unit"],
      "Control System": ["PLC CPU Module","PLC I/O Module","HMI Touchscreen","SCADA Server","UPS Battery"],
      "Power Distribution": ["MCC Incomer","Transformer HV Bushing","PFC Capacitor Bank"]
    }
  }
};

const DRIVE_SPECS = {
  chain: {
    title: "🔗 Chain Drive Engineering",
    sections: [
      { heading: "Chain Specification", fields: [
        ["Chain Standard", "select", ["ISO 606 (European – B Series)","ANSI/ASME B29.1 (American)","DIN 8187","DIN 8188","BS 228"]],
        ["Chain ISO Designation", "text", "e.g. 16B-2 (1\" pitch, duplex)"],
        ["Pitch (mm)", "number", "e.g. 25.4"],
        ["Number of Strands", "select", ["1 – Simplex","2 – Duplex","3 – Triplex","4 – Quadruplex"]],
        ["Number of Links", "number", "e.g. 120"],
        ["Minimum Breaking Load (kN)", "number", "e.g. 250"],
        ["Maximum Allowable Load (kN)", "number", "= MBL ÷ Safety Factor (min 7:1)"],
        ["Chain Type / Suffix", "select", ["Standard (Plain)","Heavy Series (H)","Stainless Steel (SS)","Nickel Plated (NP)","Self-Lubricating (SL)","Hollow Pin"]],
        ["Chain Manufacturer", "text", "e.g. Renold / Diamond / Donghua / Tsubaki"],
        ["Chain Part Number", "text", "from manufacturer catalogue"],
      ]},
      { heading: "Sprocket – Driver (Small)", fields: [
        ["Driver Teeth (Z1)", "number", "e.g. 17"],
        ["Driver Pitch Circle Diameter PCD (mm)", "number", "= Pitch / sin(π/Z1)"],
        ["Driver Shaft Diameter (mm)", "number", "e.g. 75"],
        ["Driver Bore ID (mm)", "number", "e.g. 75"],
        ["Driver Keyway (W×H mm)", "text", "e.g. 20×12"],
        ["Driver Hub OD (mm)", "number", "e.g. 115"],
        ["Driver Sprocket Material", "select", ["C45 Steel – Induction Hardened","EN8 Steel","EN36 Steel – Case Hardened","Cast Iron GG25","Stainless Steel 304"]],
        ["Driver Surface Treatment", "select", ["Induction Hardened (55–60 HRC)","Case Hardened","Zinc Plated","As Machined – Unhardened"]],
        ["Driver Manufacturer", "text", ""],
      ]},
      { heading: "Sprocket – Driven (Large)", fields: [
        ["Driven Teeth (Z2)", "number", "e.g. 38"],
        ["Speed Ratio (Z2÷Z1)", "text", "auto: Z2 ÷ Z1"],
        ["Driven Shaft Diameter (mm)", "number", "e.g. 60"],
        ["Driven Bore ID (mm)", "number", "e.g. 60"],
        ["Driven Keyway (W×H mm)", "text", "e.g. 18×11"],
        ["Driven Sprocket Material", "select", ["C45 Steel – Induction Hardened","EN8 Steel","EN36 Steel – Case Hardened","Cast Iron GG25"]],
        ["Driven Manufacturer", "text", ""],
      ]},
      { heading: "Geometry & Condition", fields: [
        ["Centre Distance (mm)", "number", "e.g. 850"],
        ["Calculated Chain Length (Links)", "text", "= 2C/p + (Z1+Z2)/2 + (Z2−Z1)²/(4π²C/p)"],
        ["Max Allowable Elongation (%)", "select", ["3.0% (Standard – ISO)","2.5% (Recommended Good Practice)","2.0% (Critical Drive – Replace Early)"]],
        ["Current Measured Elongation (%)", "number", "measured on-site"],
        ["Sag Allowance (mm)", "text", "target 1–2% of centre distance"],
        ["Chain Speed (m/s)", "number", "= Z1 × Pitch × RPM / 60000"],
      ]},
      { heading: "Lubrication", fields: [
        ["Lubrication Type", "select", ["Drip Lubricator (Method A)","Force-Feed Oil (Method B)","Oil Bath (Method C)","Manual Grease Application","Chain Oiler Spray","Centralized Auto-Lube","Dry / None (SL Chain Only)"]],
        ["Lubricant Grade / Spec", "text", "e.g. ISO VG 68 / EP-2 Grease"],
        ["Lubrication Interval (hrs)", "number", "e.g. 250"],
        ["Quantity per Application", "text", "e.g. 30 mL / 2 pump strokes"],
      ]},
    ]
  },
  gear: {
    title: "⚙️ Gearbox / Gear Drive Engineering",
    sections: [
      { heading: "Gearbox Identification", fields: [
        ["Gearbox Type", "select", ["Helical Inline (Parallel Shaft)","Helical Right Angle (Bevel-Helical)","Worm Gear","Planetary","Cycloidal","Shaft Mounted (SAF)","Bevel-Inline"]],
        ["Manufacturer", "text", "e.g. Flender / SEW / David Brown / Elecon / Hansen"],
        ["Model / Catalogue Number", "text", "from nameplate"],
        ["Serial Number", "text", "from nameplate"],
        ["Gear Ratio (i)", "text", "e.g. 12.5 : 1"],
        ["Service Factor (SF)", "number", "min 1.5 for mills"],
        ["Thermal Rating (kW)", "number", "from catalogue"],
      ]},
      { heading: "Shafts & Bearings", fields: [
        ["Input Shaft Bearing (DE)", "text", "e.g. 6316 C3 / 22220 E"],
        ["Input Shaft Bearing (NDE)", "text", "e.g. 6314 C3"],
        ["Output Shaft Bearing (DE)", "text", "e.g. 22240 CC/W33"],
        ["Output Shaft Bearing (NDE)", "text", "e.g. 22238 CC/W33"],
        ["Bearing Manufacturer", "select", ["SKF","FAG / Schaeffler","NSK","NTN","Timken","Koyo","NBC"]],
        ["Input Oil Seal (ID×OD×Width mm)", "text", "e.g. 80×100×13 NBR"],
        ["Output Oil Seal (ID×OD×Width mm)", "text", "e.g. 120×145×15 NBR"],
        ["Seal Material", "select", ["NBR (Nitrile) – Standard","FKM (Viton) – High Temp","PTFE – Chemical Resistance"]],
      ]},
      { heading: "Lubrication", fields: [
        ["Oil Grade", "select", ["ISO VG 220","ISO VG 320","ISO VG 460","ISO VG 150"]],
        ["Oil Brand / Specification", "text", "e.g. Shell Omala S2 GX 220"],
        ["Oil Capacity (Litres)", "number", "from nameplate"],
        ["Oil Change Interval (hrs)", "number", "e.g. 4000"],
        ["Lubrication Method", "select", ["Splash / Oil Bath","Forced Circulation","Oil Mist","Grease (Worm Gearbox)"]],
      ]},
    ]
  },
  belt: {
    title: "🔵 Belt Drive Engineering",
    sections: [
      { heading: "Belt Specification", fields: [
        ["Belt Type", "select", ["V-Belt (Classical)","Wedge Belt (SPZ/SPA/SPB/SPC)","Poly-V Belt (PK/PL)","Flat Belt","Timing / Synchronous Belt (HTD)","Cogged V-Belt (Raw Edge)","Joined Belt Set"]],
        ["Cross-Section / Profile", "text", "e.g. SPB / SPA / 3V / 5V"],
        ["Belt Inside Length Li (mm)", "number", "e.g. 2000"],
        ["Effective Length Le (mm)", "number", "= Li + (2 × belt section constant)"],
        ["Number of Belts in Set", "number", "e.g. 4"],
        ["Belt Part Number", "text", "e.g. SPB 2000 / 5V-2000"],
        ["Belt Manufacturer", "text", "e.g. Gates / Fenner / Optibelt / Bando / Dayco"],
        ["Power Rating per Belt (kW)", "number", "from manufacturer tables"],
      ]},
      { heading: "Driver Pulley", fields: [
        ["Driver PCD (mm)", "number", "e.g. 250"],
        ["Driver Outside Diameter (mm)", "number", "e.g. 255"],
        ["Driver Face Width (mm)", "number", "e.g. 200"],
        ["Driver Number of Grooves", "number", "e.g. 4"],
        ["Driver Shaft Diameter (mm)", "number", "e.g. 80"],
        ["Driver Bore & Key", "text", "e.g. 80mm bore / 22×14 parallel key"],
        ["Driver Taper Lock Bush Size", "select", ["1008","1210","1215","1610","2012","2517","3020","3525","4030","4535","5040","Not Taper Lock"]],
        ["Driver Pulley Manufacturer", "text", ""],
      ]},
      { heading: "Driven Pulley", fields: [
        ["Driven PCD (mm)", "number", "e.g. 500"],
        ["Driven Outside Diameter (mm)", "number", "e.g. 505"],
        ["Driven Face Width (mm)", "number", "e.g. 200"],
        ["Driven Number of Grooves", "number", "e.g. 4"],
        ["Driven Shaft Diameter (mm)", "number", "e.g. 60"],
        ["Driven Bore & Key", "text", "e.g. 60mm bore / 18×11 key"],
        ["Driven Taper Lock Bush Size", "select", ["1008","1210","1215","1610","2012","2517","3020","3525","4030","4535","5040","Not Taper Lock"]],
      ]},
      { heading: "Geometry & Tensioning", fields: [
        ["Centre Distance (mm)", "number", "e.g. 1200"],
        ["Wrap Angle on Driver (°)", "text", "= 180° − 60(D−d)/C  [min 120°]"],
        ["Belt Tension – Tight Side Ft (N)", "number", "e.g. 2200"],
        ["Belt Tension – Slack Side Fs (N)", "number", "e.g. 500"],
        ["Static Belt Deflection (mm)", "text", "target 10–15 mm per metre of span"],
        ["Tensioner Type", "select", ["Manual Jockey Idler (Slide Adjust)","Automatic Spring Tensioner","Fixed – No Tensioner","Hydraulic Tensioner"]],
        ["Idler Pulley PCD (mm)", "number", "e.g. 150 (if idler fitted)"],
      ]},
    ]
  },
  hydraulic: {
    title: "🟠 Hydraulic Drive / Cylinder Engineering",
    sections: [
      { heading: "Cylinder Specification", fields: [
        ["Cylinder Function", "text", "e.g. Roll Gap Control / Door Lift / Clamping"],
        ["Bore Diameter (mm)", "number", "e.g. 200"],
        ["Rod Diameter (mm)", "number", "e.g. 140"],
        ["Stroke (mm)", "number", "e.g. 300"],
        ["Operating Pressure (bar)", "number", "e.g. 200"],
        ["Proof Test Pressure (bar)", "number", "= 1.5 × Operating Pressure"],
        ["Max Working Pressure (bar)", "number", "e.g. 350"],
        ["Thrust Force – Extend (kN)", "text", "= π/4 × Bore² × Pressure"],
        ["Thrust Force – Retract (kN)", "text", "= π/4 × (Bore²−Rod²) × Pressure"],
        ["Mounting Style", "select", ["Flange – Front","Flange – Rear","Trunnion – Mid","Foot Mounted","Clevis","Knuckle"]],
        ["Rod End Type", "select", ["Male Thread","Female Thread","Clevis Fork","Knuckle","Plain Eye"]],
      ]},
      { heading: "Seals & Materials", fields: [
        ["Cylinder Body Material", "select", ["Carbon Steel – Honed Tube","Stainless Steel 316","Ductile Iron","Aluminium Alloy"]],
        ["Rod Material / Hard Chrome", "select", ["Chrome Plated Carbon Steel (45µm min)","Stainless Steel 316 Rod","Induction Hardened + Chrome","Ceramic Coated"]],
        ["Seal Material (System Fluid)", "select", ["NBR (Nitrile) – Mineral Oil","FKM (Viton) – Fire Resistant Fluid","EPDM – Water-Glycol Fluid","PTFE – Universal"]],
        ["Seal Kit Part Number", "text", "from manufacturer / OEM"],
        ["Piston Seal Type", "select", ["PTFE Seal Ring + O-Ring","Polyurethane Piston Seal","Hallite / Parker OEM Seal","Quad Ring (X-Ring)"]],
        ["Rod Seal Type", "select", ["Polyurethane U-Seal","PTFE + Spring Energised","NBR Lip Seal","Hallite OEM Rod Seal"]],
        ["Wiper / Scraper Seal", "select", ["Polyurethane Wiper","PTFE Wiper","Single Lip Wiper","Double Lip Wiper"]],
        ["O-Ring Material", "select", ["NBR – Nitrile","FKM – Viton","EPDM","PTFE"]],
      ]},
      { heading: "HPU – Power Unit", fields: [
        ["HPU Manufacturer", "text", "e.g. Bosch Rexroth / Parker / Eaton / Hydraforce"],
        ["Pump Type", "select", ["Axial Piston (Variable Displacement)","Axial Piston (Fixed Displacement)","Radial Piston","External Gear","Internal Gear (Gerotor)","Vane (Fixed)","Screw (Triple Screw)"]],
        ["Pump Displacement (cc/rev)", "number", "e.g. 45"],
        ["Rated System Pressure (bar)", "number", "e.g. 250"],
        ["Max Flow Rate (L/min)", "number", "e.g. 80"],
        ["Reservoir Capacity (L)", "number", "e.g. 400"],
        ["Oil Grade", "select", ["ISO VG 46 (Standard Hydraulic)","ISO VG 68","ISO VG 32","Fire Resistant – HFDU","Fire Resistant – HFAS (Water-Glycol)"]],
        ["HP Filter Rating (µm)", "select", ["10 µm (Servo-Grade)","25 µm (Standard)","40 µm (Economy)"]],
        ["Relief Valve Setting (bar)", "number", "= 1.1 × Operating Pressure"],
        ["Accumulator Capacity (L)", "number", "e.g. 10"],
        ["Accumulator Pre-charge Pressure (bar)", "number", "= 0.9 × min system pressure"],
      ]},
    ]
  },
  rack: {
    title: "⚙️ Rack & Pinion Drive Engineering",
    sections: [
      { heading: "Rack Specification", fields: [
        ["Module (m)", "number", "e.g. 8  (PCD = m × Z)"],
        ["Pressure Angle (°)", "select", ["20° (Standard)","14.5° (Legacy)","25° (High Load)"]],
        ["Tooth Form", "select", ["Straight (Spur) Teeth","Helical Teeth – specify helix angle"]],
        ["Rack Total Length (mm)", "number", "e.g. 3000"],
        ["Rack Width / Face Width (mm)", "number", "e.g. 80"],
        ["Rack Height (mm)", "number", "e.g. 60"],
        ["Rack Material", "select", ["C45 Steel – Induction Hardened","EN8 – Normalised","42CrMo4 – Alloy","Cast Iron GG25"]],
        ["Rack Surface Hardness (HRC)", "text", "e.g. 54–58 HRC (case hardened)"],
        ["Rack Core Hardness (HB)", "text", "e.g. 220 HB minimum"],
        ["Number of Rack Sections", "number", "e.g. 3"],
        ["Joint Gap Between Sections (mm)", "text", "< 0.1 mm for precision, < 0.5 mm standard"],
        ["Rack Mounting Bolt Size", "select", ["M16","M20","M24","M12"]],
        ["Rack Mounting Bolt Grade", "select", ["10.9","8.8","12.9"]],
        ["Rack Mounting Bolt Torque (Nm)", "number", "from OEM specification"],
        ["Rack Manufacturer", "text", "e.g. KISSsoft Design / Nord / Custom Fabricated"],
      ]},
      { heading: "Pinion Specification", fields: [
        ["Number of Teeth (Z)", "number", "e.g. 20  (min Z = 17 to avoid undercutting)"],
        ["Pitch Circle Diameter PCD (mm)", "text", "= Module × Z"],
        ["Outside Diameter OA (mm)", "text", "= PCD + 2 × module"],
        ["Face Width b (mm)", "number", "e.g. 90  (typically 10–15 × module)"],
        ["Pinion Material", "select", ["C45 Steel – Induction Hardened","42CrMo4 – Through Hardened","18CrNiMo7 – Case Hardened","EN36 – Case Hardened"]],
        ["Pinion Surface Hardness (HRC)", "text", "e.g. 58–62 HRC"],
        ["Shaft Diameter (mm)", "number", "e.g. 80"],
        ["Keyway Size (W×H mm)", "text", "e.g. 22×14"],
        ["Pinion Manufacturer", "text", ""],
      ]},
    ]
  },
  screw: {
    title: "🔧 Ball / Acme Screw Drive Engineering",
    sections: [
      { heading: "Screw Specification", fields: [
        ["Screw Type", "select", ["Ball Screw (Recirculating Ball)","Acme / Trapezoidal Screw (DIN 103)","Roller Screw (Planetary)","Lead Screw (Square Thread)"]],
        ["Nominal Screw Diameter (mm)", "number", "e.g. 63"],
        ["Lead (mm per revolution)", "number", "e.g. 10"],
        ["Number of Starts", "select", ["1 – Single Start","2 – Double Start","4 – Quad Start"]],
        ["Screw Length (mm)", "number", "e.g. 1200"],
        ["Accuracy Class (Ball Screw)", "select", ["C0 – Master Grade (0.5 µm/300mm)","C1 – Precision (2 µm/300mm)","C3 – High Precision (5 µm/300mm)","C5 – Standard (18 µm/300mm)","C7 – Economy (50 µm/300mm)","N/A – Acme Screw"]],
        ["Screw Material", "select", ["SCM440 (42CrMo4) – Induction Hardened","EN36 – Case Hardened","Tool Steel D2","Stainless Steel 440C"]],
        ["Ball Screw Manufacturer", "text", "e.g. SKF / Bosch Rexroth / Thomson / Hiwin / TBI"],
      ]},
      { heading: "Nut & Support Bearings", fields: [
        ["Nut Type", "select", ["Single Ball Nut (Standard)","Double Ball Nut (Preloaded)","Flange Ball Nut","Trapezoidal Nut (Bronze)","Trapezoidal Nut (Cast Iron)"]],
        ["Preload Class", "select", ["Z0 – No Preload","Z1 – Light (2%)","Z2 – Medium (5%)","Z3 – Heavy (8%)","N/A – Acme Screw"]],
        ["Fixed End Bearing (Angular Contact)", "text", "e.g. FAG 7210 AC/P5 pair"],
        ["Floating End Bearing (Deep Groove)", "text", "e.g. 6210 C3"],
        ["Bearing Manufacturer", "select", ["SKF","FAG / Schaeffler","NSK","NTN","Timken"]],
      ]},
    ]
  },
  pneumatic: {
    title: "💨 Pneumatic Drive / Cylinder Engineering",
    sections: [
      { heading: "Cylinder Specification", fields: [
        ["Cylinder Type", "select", ["Double Acting","Single Acting – Spring Return","Single Acting – Spring Extend","Rotary Actuator (Vane)","Diaphragm Actuator","Rodless Cylinder"]],
        ["Cylinder Function", "text", "e.g. Door Clamp / Guide Actuator / Index Stop"],
        ["Bore Diameter (mm)", "number", "e.g. 100"],
        ["Rod Diameter (mm)", "number", "e.g. 32"],
        ["Stroke (mm)", "number", "e.g. 200"],
        ["Operating Pressure (bar)", "number", "e.g. 6"],
        ["Theoretical Force – Extend (N)", "text", "= π/4 × Bore² × Pressure × 10"],
        ["Mounting Style", "select", ["Front Flange (FA)","Rear Flange (FB)","Foot Mounted (MA)","Clevis (CA)","Trunnion (TC)","Rod Flange"]],
        ["Cylinder Manufacturer", "select", ["SMC","Festo","Parker","Norgren","Bosch Rexroth","Bimba"]],
        ["Cylinder Model / PN", "text", ""],
      ]},
      { heading: "Valve & FRL Unit", fields: [
        ["Supply Line Pressure (bar)", "number", "e.g. 7 bar line"],
        ["Regulator Set Pressure (bar)", "number", "e.g. 6"],
        ["Directional Valve Type", "select", ["5/2 Solenoid – Spring Return","5/2 Solenoid – Detented (Memory)","5/3 Spring-Centred – Blocked Centre","5/3 Spring-Centred – Exhaust Centre","4/2 Double Solenoid","3/2 NC (Normally Closed)"]],
        ["Valve Port Size", "select", ["G1/8\"","G1/4\"","G3/8\"","G1/2\"","G3/4\""]],
        ["Solenoid Voltage", "select", ["24V DC (Standard)","12V DC","110V AC","230V AC"]],
        ["Filter Element Rating (µm)", "select", ["5 µm (Sensitive Valves)","10 µm","25 µm","40 µm (General)"]],
        ["FRL Port Size", "select", ["G1/4\"","G3/8\"","G1/2\"","G3/4\"","G1\""]],
        ["FRL Manufacturer", "text", "e.g. SMC / Festo / Norgren / Parker / IMI"],
        ["Air Consumption (NL/min)", "number", "from valve datasheet"],
      ]},
    ]
  },
  direct: {
    title: "⚡ Direct Motor Drive Engineering",
    sections: [
      { heading: "Motor & Coupling", fields: [
        ["Drive Configuration", "select", ["Direct Shaft Coupling","In-Line via Flexible Coupling","Flange-Mounted Motor","Hollow Shaft Motor (Shaft-Mount)","Servo Motor + Feedback Encoder"]],
        ["Flexible Coupling Type", "select", ["Jaw / Spider Coupling","Grid Coupling (Falk)","Disc Coupling (High Precision)","Gear Coupling (Torque-Dense)","Tyre Coupling","Rigid Flanged Coupling"]],
        ["Coupling Size / PN", "text", "e.g. KTR Rotex 75 / Rexnord T70"],
        ["Rated Coupling Torque (Nm)", "number", "min 2× motor rated torque"],
        ["Hub Bore – Motor Side (mm)", "number", "e.g. 80"],
        ["Hub Bore – Load Side (mm)", "number", "e.g. 85"],
        ["Spider / Insert Material", "select", ["Polyurethane (PU) 92 ShA","Polyurethane 98 ShA","NBR Rubber","Hytrel – High Performance","Steel Grid"]],
        ["Alignment Target (TIR mm)", "select", ["≤ 0.05 mm (Precision Laser)","≤ 0.10 mm (Standard)","≤ 0.25 mm (Acceptable)"]],
        ["Coupling Manufacturer", "select", ["KTR","Rexnord","Rathi","Bibby / Huco","Lovejoy","Flender / Siemens","SKF","R+W"]],
        ["Hub Bolt Size", "select", ["M8","M10","M12","M16","M20"]],
        ["Hub Bolt Grade", "select", ["8.8","10.9","12.9"]],
        ["Hub Bolt Qty", "number", "e.g. 4"],
        ["Hub Bolt Torque (Nm)", "number", "from OEM datasheet"],
      ]},
    ]
  }
};

// ══════════════════════════════════════════════
// DYNAMIC DRIVE SPEC FIELDS
// ══════════════════════════════════════════════
const driveFieldMap = {};

function generateDriveSpec() {
  const drv = document.getElementById('drive').value;
  const panel = document.getElementById('drive-spec-panel');
  panel.innerHTML = '';
  driveFieldMap[drv] = {};

  if (!drv || !DRIVE_SPECS[drv]) return;

  const spec = DRIVE_SPECS[drv];
  let html = `<div class="card">
    <div class="card-header" onclick="toggleCard(this)">
      <div class="card-title">▌ ${spec.title} <span class="card-badge">Engineering Detail</span></div>
      <div class="chevron open">▼</div>
    </div>
    <div class="card-body">`;

  spec.sections.forEach((sec, si) => {
    html += `<div class="sec-div"><span>${sec.heading}</span></div>
    <div class="grid g3">`;
    sec.fields.forEach(([label, type, opt], fi) => {
      const id = `drv_${drv}_${si}_${fi}`;
      html += `<div class="field"><label>${label}</label>`;
      if (type === 'select') {
        html += `<select id="${id}">`;
        opt.forEach(o => { html += `<option>${o}</option>`; });
        html += `</select>`;
      } else {
        html += `<input id="${id}" type="${type === 'number' ? 'number' : 'text'}" placeholder="${opt}">`;
      }
      html += `</div>`;
      driveFieldMap[drv][label] = id;
    });
    html += `</div>`;
  });

  html += `</div></div>`;
  panel.innerHTML = html;
}

// ══════════════════════════════════════════════
// CASCADING DROPDOWNS
// ══════════════════════════════════════════════
function loadSub() {
  const mill = document.getElementById('mill').value;
  const sub = document.getElementById('sub');
  const detail = document.getElementById('detail');
  sub.innerHTML = '<option value="">— Select Subcomponent —</option>';
  detail.innerHTML = '<option value="">— Select Component —</option>';
  if (!mill || !MILL_DATA[mill]) return;
  Object.keys(MILL_DATA[mill].subs).forEach(s => {
    const o = document.createElement('option');
    o.value = s; o.text = s; sub.appendChild(o);
  });
}

function loadDetail() {
  const mill = document.getElementById('mill').value;
  const sub = document.getElementById('sub').value;
  const detail = document.getElementById('detail');
  detail.innerHTML = '<option value="">— Select Component —</option>';
  if (!mill || !sub || !MILL_DATA[mill]?.subs[sub]) return;
  MILL_DATA[mill].subs[sub].forEach(d => {
    const o = document.createElement('option');
    o.value = d; o.text = d; detail.appendChild(o);
  });
}

// ══════════════════════════════════════════════
// NAVIGATION
// ══════════════════════════════════════════════
const navBtns = document.querySelectorAll('.nav-btn');
function showPanel(id) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  navBtns.forEach(b => b.classList.remove('active'));
  const map = {
    'p-select': 0, 'p-drive': 1, 'p-bearing': 2,
    'p-fastener': 3, 'p-material': 4, 'p-quality': 5, 'p-bom': 6
  };
  if (map[id] !== undefined) navBtns[map[id]].classList.add('active');
  if (id === 'p-bom') renderBOMPreview();
  updateTrail(id);
}

function updateTrail(id) {
  const map = {'p-select': 1,'p-drive': 2,'p-bearing': 3,'p-fastener': 4,'p-material': 5,'p-quality': 6,'p-bom': 7};
  const cur = map[id] || 1;
  for (let i = 1; i <= 7; i++) {
    const el = document.getElementById(`t${i}`);
    el.classList.remove('active','done');
    if (i === cur) el.classList.add('active');
    else if (i < cur) el.classList.add('done');
  }
}

function toggleCard(hdr) {
  const body = hdr.nextElementSibling;
  const chev = hdr.querySelector('.chevron');
  body.style.display = body.style.display === 'none' ? 'block' : 'none';
  chev.classList.toggle('open');
}

// ══════════════════════════════════════════════
// COLLECT DATA
// ══════════════════════════════════════════════
function val(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function collectAll() {
  const drv = val('drive');
  const drvSpecs = {};
  if (drv && DRIVE_SPECS[drv]) {
    DRIVE_SPECS[drv].sections.forEach((sec, si) => {
      sec.fields.forEach(([label, , ], fi) => {
        const id = `drv_${drv}_${si}_${fi}`;
        const el = document.getElementById(id);
        if (el) drvSpecs[label] = el.value;
      });
    });
  }

  return {
    project: {
      bomRef: val('bomRef'), woNo: val('woNo'), rev: val('rev'),
      prepBy: val('prepBy'), appBy: val('appBy'), bomDate: val('bomDate'),
      plantName: val('plantName'), assetTag: val('assetTag'), priority: val('priority')
    },
    millSection: {
      mill: val('mill'), millLabel: MILL_DATA[val('mill')]?.label || val('mill'),
      sub: val('sub'), detail: val('detail'), tagNo: val('tagNo'),
      qty: val('qty'), criticality: val('crit'), location: val('location'),
      installDate: val('installDate'), overhaulDate: val('overhaulDate'),
      description: val('description')
    },
    drive: {
      driveType: val('drive'), power_kw: val('drv_kw'),
      input_rpm: val('drv_rpm_in'), output_rpm: val('drv_rpm_out'),
      ratio: val('drv_ratio'), torque: val('drv_torq'),
      ...drvSpecs
    },
    motor: {
      manufacturer: val('m_mfr'), model: val('m_model'), kw: val('m_kw'),
      voltage: val('m_volt'), fla: val('m_fla'), syncRpm: val('m_sync'),
      flRpm: val('m_flrpm'), pf: val('m_pf'), eff: val('m_eff'),
      ip: val('m_ip'), insClass: val('m_ins'), enclosure: val('m_enc'),
      sf: val('m_sf'), footBoltSize: val('m_fbolt'),
      footBoltGrade: val('m_fbgrade'), footBoltQty: val('m_fbqty')
    },
    bearings: {
      de_type: val('de_type'), de_desig: val('de_desig'),
      de_bore: val('de_bore'), de_od: val('de_od'), de_width: val('de_width'),
      de_clearance: val('de_cl'), de_mfr: val('de_mfr'),
      de_shaftFit: val('de_shfit'), de_housingFit: val('de_hfit'),
      de_install: val('de_inst'), de_sealType: val('de_seal'),
      de_grease: val('de_grease'), de_nipple: val('de_nipple'),
      de_C: val('de_C'), de_C0: val('de_C0'),
      de_Fr: val('de_Fr'), de_Fa: val('de_Fa'),
      de_lubInt: val('de_lubint'), de_lubQty: val('de_lubqty'),
      nde_type: val('nde_type'), nde_desig: val('nde_desig'),
      nde_bore: val('nde_bore'), nde_od: val('nde_od'), nde_width: val('nde_width'),
      nde_clearance: val('nde_cl'), nde_mfr: val('nde_mfr'),
      nde_shaftFit: val('nde_shfit'), nde_housingFit: val('nde_hfit'),
      nde_seal: val('nde_seal'), nde_grease: val('nde_grease'),
      nde_lubInt: val('nde_lubint'), nde_lubQty: val('nde_lubqty'),
      shaft_dia: val('shaft_dia'), shaft_len: val('shaft_len'),
      shaft_mat: val('shaft_mat'), shaft_surf: val('shaft_surf'),
      keyway: val('keyway'), keyLen: val('keylen'),
      keyMat: val('key_mat'), shaftHardness: val('shaft_hrd')
    },
    fasteners: {
      boltType: val('bt_type'), boltSize: val('bt_size'),
      boltGrade: val('bt_grade'), threadStd: val('bt_thread'),
      boltLen: val('bt_len'), boltMat: val('bt_mat'),
      torque_Nm: val('bt_torque'), qtyPerJoint: val('bt_jqty'),
      numJoints: val('bt_joints'), totalBolts: val('bt_total'),
      nutType: val('nut_type'), nutGrade: val('nut_grade'),
      washerType: val('wash_type'), lockMethod: val('lock_mth'),
      sealType: val('seal_type'), sealMat: val('seal_mat'),
      sealID: val('seal_id'), sealOD: val('seal_od'), sealW: val('seal_w'),
      sealPN: val('seal_pn'), sealMfr: val('seal_mfr'),
      sealTemp: val('seal_temp'), gasketType: val('gask_type'),
      gasketMat: val('gask_mat'), sealQty: val('seal_qty'),
      gasketQty: val('gask_qty')
    },
    materials: {
      housingMat: val('hsg_mat'), housingSurf: val('hsg_surf'),
      housingThick: val('hsg_thick'), baseThick: val('base_thick'),
      anchorBoltSize: val('anch_sz'), anchorBoltGrade: val('anch_gr'),
      anchorBoltQty: val('anch_qty'), grout: val('grout'),
      alignMethod: val('align_mth'), oilGrade: val('oil_grade'),
      oilBrand: val('oil_brand'), oilCap: val('oil_cap'),
      oilInterval: val('oil_int'), oilTempNormal: val('oil_temp_n'),
      oilTempAlarm: val('oil_temp_a'), oilTempTrip: val('oil_temp_t'),
      lubSystem: val('lub_sys'), filterRating: val('filt_rat'),
      couplingType: val('cplg_type'), couplingMfr: val('cplg_mfr'),
      couplingSize: val('cplg_size'), couplingTorque: val('cplg_torq'),
      couplingSpider: val('cplg_spi'), couplingBore: val('cplg_bore'),
      couplingBoltSize: val('cplg_bsz'), couplingBoltGrade: val('cplg_bgr'),
      couplingBoltQty: val('cplg_bqty')
    },
    quality: {
      brgStd: val('qa_brg_std'), boltStd: val('qa_bolt_std'),
      gearStd: val('qa_gear_std'), weldStd: val('qa_weld'),
      surfStd: val('qa_surf'), paintStd: val('qa_paint'),
      dimInspect: val('qa_dim'), ndt: val('qa_ndt'),
      matCert: val('qa_cert'), hardTest: val('qa_hard'),
      vibAccept: val('qa_vib'), alignAccept: val('qa_align'),
      notes: val('qa_notes'), supplier: val('supplier'),
      supplier2: val('supplier2'), leadTime: val('lead_time'),
      unitCost: val('unit_cost'), totalCost: val('total_cost'),
      storage: val('storage'), minStock: val('min_stock'),
      reorderQty: val('reorder_qty'), incoterms: val('inco')
    }
  };
}

// ══════════════════════════════════════════════
// BOM PREVIEW
// ══════════════════════════════════════════════
function renderBOMPreview() {
  const d = collectAll();
  const tc = parseInt(d.quality.unitCost || 0) * parseInt(d.millSection.qty || 1);
  document.getElementById('st-items').textContent = '80+';
  document.getElementById('st-cat').textContent = '7';
  document.getElementById('st-cost').textContent = tc.toLocaleString();
  document.getElementById('st-comp').textContent = '100%';

  const sections = [
    { title: 'Project Identification', color: '#1E5799', rows: [
      ['BOM Reference', d.project.bomRef], ['Work Order No.', d.project.woNo],
      ['Revision', d.project.rev], ['Prepared By', d.project.prepBy],
      ['Approved By', d.project.appBy], ['Date', d.project.bomDate],
      ['Plant / Mill', d.project.plantName], ['Asset Tag', d.project.assetTag],
      ['Priority', d.project.priority]
    ]},
    { title: 'Mill Section & Component', color: '#1B5E20', rows: [
      ['Mill Section', d.millSection.millLabel], ['Subcomponent', d.millSection.sub],
      ['Specific Component', d.millSection.detail], ['Tag No.', d.millSection.tagNo],
      ['Quantity', d.millSection.qty], ['Criticality', d.millSection.criticality],
      ['Location', d.millSection.location], ['Description', d.millSection.description]
    ]},
    { title: 'Drive System', color: '#4A235A', rows: Object.entries(d.drive).map(([k,v]) => [k, v]) },
    { title: 'Drive Motor', color: '#7E3700', rows: Object.entries(d.motor).map(([k,v]) => [k, v]) },
    { title: 'Bearing – DE', color: '#B71C1C', rows: [
      ['Type', d.bearings.de_type], ['Designation', d.bearings.de_desig],
      ['Bore (mm)', d.bearings.de_bore], ['OD (mm)', d.bearings.de_od],
      ['Width (mm)', d.bearings.de_width], ['Clearance', d.bearings.de_clearance],
      ['Manufacturer', d.bearings.de_mfr], ['Shaft Fit', d.bearings.de_shaftFit],
      ['Housing Fit', d.bearings.de_housingFit], ['Installation', d.bearings.de_install],
      ['Seal Type', d.bearings.de_sealType], ['Grease', d.bearings.de_grease],
      ['C (kN)', d.bearings.de_C], ['C0 (kN)', d.bearings.de_C0],
      ['Lube Interval (hrs)', d.bearings.de_lubInt], ['Grease Qty (g)', d.bearings.de_lubQty]
    ]},
    { title: 'Bearing – NDE', color: '#880E4F', rows: [
      ['Type', d.bearings.nde_type], ['Designation', d.bearings.nde_desig],
      ['Bore (mm)', d.bearings.nde_bore], ['OD (mm)', d.bearings.nde_od],
      ['Width (mm)', d.bearings.nde_width], ['Clearance', d.bearings.nde_clearance],
      ['Manufacturer', d.bearings.nde_mfr], ['Seal Type', d.bearings.nde_seal],
      ['Grease', d.bearings.nde_grease]
    ]},
    { title: 'Shaft', color: '#1A237E', rows: [
      ['Diameter (mm)', d.bearings.shaft_dia], ['Length (mm)', d.bearings.shaft_len],
      ['Material', d.bearings.shaft_mat], ['Surface Treatment', d.bearings.shaft_surf],
      ['Keyway (W×H)', d.bearings.keyway], ['Key Length (mm)', d.bearings.keyLen],
      ['Key Material', d.bearings.keyMat], ['Hardness', d.bearings.shaftHardness]
    ]},
    { title: 'Fasteners & Bolts', color: '#004D40', rows: [
      ['Bolt Type', d.fasteners.boltType], ['Size', d.fasteners.boltSize],
      ['Grade', d.fasteners.boltGrade], ['Thread Std', d.fasteners.threadStd],
      ['Length (mm)', d.fasteners.boltLen], ['Material', d.fasteners.boltMat],
      ['Torque (Nm)', d.fasteners.torque_Nm], ['Qty per Joint', d.fasteners.qtyPerJoint],
      ['No. of Joints', d.fasteners.numJoints], ['Total Qty', d.fasteners.totalBolts],
      ['Nut Type', d.fasteners.nutType], ['Nut Grade', d.fasteners.nutGrade],
      ['Washer Type', d.fasteners.washerType], ['Lock Method', d.fasteners.lockMethod]
    ]},
    { title: 'Seals & Gaskets', color: '#3E2723', rows: [
      ['Seal Type', d.fasteners.sealType], ['Seal Material', d.fasteners.sealMat],
      ['Seal ID (mm)', d.fasteners.sealID], ['Seal OD (mm)', d.fasteners.sealOD],
      ['Seal Width (mm)', d.fasteners.sealW], ['Seal PN', d.fasteners.sealPN],
      ['Seal Manufacturer', d.fasteners.sealMfr], ['Seal Temp Range', d.fasteners.sealTemp],
      ['Gasket Type', d.fasteners.gasketType], ['Gasket Material', d.fasteners.gasketMat],
      ['Seal Qty', d.fasteners.sealQty], ['Gasket Qty', d.fasteners.gasketQty]
    ]},
    { title: 'Materials & Lubrication', color: '#1A237E', rows: Object.entries(d.materials).map(([k,v]) => [k, v]) },
    { title: 'QA & Standards', color: '#37474F', rows: Object.entries(d.quality).map(([k,v]) => [k, v]) }
  ];

  let html = '';
  sections.forEach(sec => {
    html += `<div class="card" style="margin-bottom:14px;">
      <div class="card-header" onclick="toggleCard(this)">
        <div class="card-title" style="color:var(--accent);">▌ ${sec.title}</div>
        <div class="chevron open">▼</div>
      </div>
      <div class="card-body" style="padding:0;">
        <div class="tbl-wrap"><table>
          <thead><tr><th style="width:40%;">Parameter</th><th>Value</th></tr></thead>
          <tbody>`;
    sec.rows.forEach(([k, v]) => {
      if (v) html += `<tr><td>${k}</td><td>${v}</td></tr>`;
    });
    html += `</tbody></table></div></div></div>`;
  });

  document.getElementById('bom-preview').innerHTML = html;
}

// ══════════════════════════════════════════════
// SAVE
// ══════════════════════════════════════════════
function saveBOM() {
  const d = collectAll();
  const ref = d.project.bomRef || `BOM-${Date.now()}`;
  const saved = JSON.parse(localStorage.getItem('MMC_BOMS') || '{}');
  saved[ref] = { ...d, savedAt: new Date().toISOString() };
  localStorage.setItem('MMC_BOMS', JSON.stringify(saved));
  updateSavedList();
  alert(`BOM saved: ${ref}`);
}

function updateSavedList() {
  const saved = JSON.parse(localStorage.getItem('MMC_BOMS') || '{}');
  const keys = Object.keys(saved);
  const el = document.getElementById('saved-list');
  if (keys.length === 0) { el.textContent = 'No BOMs saved yet'; return; }
  el.innerHTML = keys.map(k =>
    `<div style="padding:5px 0;border-bottom:1px solid var(--border);color:var(--txt2);font-size:11px;">${k}</div>`
  ).join('');
}

function clearAll() {
  if (!confirm('Clear all form data?')) return;
  document.querySelectorAll('input, select, textarea').forEach(el => {
    if (el.tagName === 'SELECT') el.selectedIndex = 0;
    else el.value = '';
  });
  document.getElementById('drive-spec-panel').innerHTML = '';
}

// ══════════════════════════════════════════════
// EXCEL EXPORT – Multi-Sheet, Fully Formatted
// ══════════════════════════════════════════════
function exportExcel() {
  const d = collectAll();
  const wb = XLSX.utils.book_new();

  // ── Style helpers ──
  function hdrRow(cells) {
    return cells.map(v => ({ v, t: 's', s: {
      font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11, name: 'Arial' },
      fill: { fgColor: { rgb: '0B1E3D' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: { top: {style:'thin',color:{rgb:'C8D8E8'}}, bottom: {style:'thin',color:{rgb:'C8D8E8'}},
                left: {style:'thin',color:{rgb:'C8D8E8'}}, right: {style:'thin',color:{rgb:'C8D8E8'}} }
    }}));
  }
  function catRow(label, color = '1E5799') {
    return [{ v: label, t: 's', s: {
      font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10, name: 'Arial' },
      fill: { fgColor: { rgb: color } },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: { bottom: {style:'thin',color:{rgb:'C8D8E8'}} }
    }}, { v: '', t: 's' }, { v: '', t: 's' }, { v: '', t: 's' }];
  }
  function dataRow(param, value, unit = '', note = '', alt = false) {
    const bg = alt ? 'EBF3FB' : 'FFFFFF';
    const mkCell = (v, bold = false) => ({ v: v || '', t: 's', s: {
      font: { bold, sz: 10, name: 'Arial', color: { rgb: bold ? '0B1E3D' : '333333' } },
      fill: { fgColor: { rgb: bg } },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border: { top: {style:'thin',color:{rgb:'E0E0E0'}}, bottom: {style:'thin',color:{rgb:'E0E0E0'}},
                left: {style:'thin',color:{rgb:'E0E0E0'}}, right: {style:'thin',color:{rgb:'E0E0E0'}} }
    }});
    return [mkCell(param, true), mkCell(value), mkCell(unit), mkCell(note)];
  }
  function titleRow(text) {
    return [{ v: text, t: 's', s: {
      font: { bold: true, sz: 14, name: 'Arial', color: { rgb: '0B1E3D' } },
      fill: { fgColor: { rgb: 'F4A62A' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }}, {v:'',t:'s'},{v:'',t:'s'},{v:'',t:'s'}];
  }
  function subTitleRow(text) {
    return [{ v: text, t: 's', s: {
      font: { sz: 10, name: 'Arial', italic: true, color: { rgb: '555555' } },
      fill: { fgColor: { rgb: 'EBF3FB' } },
      alignment: { horizontal: 'left' }
    }},{v:'',t:'s'},{v:'',t:'s'},{v:'',t:'s'}];
  }
  function blankRow() { return [{v:'',t:'s'},{v:'',t:'s'},{v:'',t:'s'},{v:'',t:'s'}]; }

  const cols = [{ wch: 38 }, { wch: 40 }, { wch: 16 }, { wch: 28 }];

  // ─────────── SHEET 1: COVER PAGE ───────────
  const cov = [];
  cov.push(titleRow('MMC STEEL DIVISION – ENGINEERING BILL OF MATERIALS'));
  cov.push(subTitleRow(`Mulijbhai Madhivani Co. Ltd  |  Steel Division  |  Rolling Mill Maintenance Management System`));
  cov.push(blankRow());
  cov.push(hdrRow(['FIELD', 'DETAIL', 'UNIT', 'NOTES']));
  [
    ['BOM Reference No.', d.project.bomRef], ['Work Order No.', d.project.woNo],
    ['Revision', d.project.rev], ['Plant / Mill', d.project.plantName],
    ['Asset Tag', d.project.assetTag], ['Mill Section', d.millSection.millLabel],
    ['Subcomponent', d.millSection.sub], ['Component', d.millSection.detail],
    ['Tag Number', d.millSection.tagNo], ['Quantity Required', d.millSection.qty, 'Units'],
    ['Criticality', d.millSection.criticality], ['Location', d.millSection.location],
    ['Installation Date', d.millSection.installDate], ['Last Overhaul', d.millSection.overhaulDate],
    ['Priority', d.project.priority], ['Prepared By', d.project.prepBy],
    ['Approved By', d.project.appBy], ['Date', d.project.bomDate],
    ['Description / Scope', d.millSection.description]
  ].forEach(([p, v, u, n], i) => cov.push(dataRow(p, v, u || '', n || '', i % 2 === 0)));
  const ws_cov = XLSX.utils.aoa_to_sheet(cov);
  ws_cov['!cols'] = cols;
  ws_cov['!merges'] = [{s:{r:0,c:0},e:{r:0,c:3}},{s:{r:1,c:0},e:{r:1,c:3}}];
  XLSX.utils.book_append_sheet(wb, ws_cov, 'Cover Page');

  // ─────────── SHEET 2: DRIVE SYSTEM ───────────
  const drv = [];
  drv.push(titleRow('DRIVE SYSTEM ENGINEERING SPECIFICATION'));
  drv.push(subTitleRow(`Drive Type: ${d.drive.driveType || '—'}  |  Power: ${d.drive.power_kw || '—'} kW  |  Ratio: ${d.drive.ratio || '—'}`));
  drv.push(blankRow());
  drv.push(hdrRow(['PARAMETER', 'VALUE / SPECIFICATION', 'UNIT', 'ENGINEERING NOTES']));
  drv.push(catRow('DRIVE SUMMARY', '1E5799'));
  [
    ['Drive Type', d.drive.driveType, '', 'Refer drive-specific section below'],
    ['Power Input', d.drive.power_kw, 'kW', ''],
    ['Input Speed', d.drive.input_rpm, 'RPM', ''],
    ['Output Speed', d.drive.output_rpm, 'RPM', ''],
    ['Speed / Gear Ratio', d.drive.ratio, ':1', ''],
    ['Output Torque', d.drive.torque, 'Nm', '= 9550 × kW / RPM'],
  ].forEach(([p,v,u,n],i)=>drv.push(dataRow(p,v,u,n,i%2===0)));

  drv.push(blankRow());
  drv.push(catRow('DRIVE MOTOR', '1B7F4F'));
  [
    ['Manufacturer', d.motor.manufacturer], ['Model / Frame', d.motor.model],
    ['Rated Power', d.motor.kw, 'kW'],
    ['Supply Voltage', d.motor.voltage, 'V'],
    ['Full Load Current', d.motor.fla, 'A'],
    ['Synchronous Speed', d.motor.syncRpm, 'RPM'],
    ['Full Load Speed', d.motor.flRpm, 'RPM'],
    ['Power Factor (cosφ)', d.motor.pf, '—'],
    ['Efficiency', d.motor.eff, '%'],
    ['IP Rating', d.motor.ip], ['Insulation Class', d.motor.insClass],
    ['Enclosure Type', d.motor.enclosure],
    ['Service Factor', d.motor.sf, '—', 'Min 1.15 for mills'],
    ['Foot Bolt Size', d.motor.footBoltSize, '', '4 bolts minimum'],
    ['Foot Bolt Grade', d.motor.footBoltGrade],
    ['Foot Bolt Quantity', d.motor.footBoltQty, 'Pcs'],
  ].forEach(([p,v,u,n],i)=>drv.push(dataRow(p,v,u||'',n||'',i%2===0)));

  // Drive-specific specs
  if (d.drive.driveType && DRIVE_SPECS[d.drive.driveType]) {
    DRIVE_SPECS[d.drive.driveType].sections.forEach(sec => {
      drv.push(blankRow());
      drv.push(catRow(sec.heading.toUpperCase(), '4A235A'));
      sec.fields.forEach(([label], fi) => {
        drv.push(dataRow(label, d.drive[label] || '', '', '', fi % 2 === 0));
      });
    });
  }

  const ws_drv = XLSX.utils.aoa_to_sheet(drv);
  ws_drv['!cols'] = cols;
  XLSX.utils.book_append_sheet(wb, ws_drv, 'Drive System');

  // ─────────── SHEET 3: BEARINGS ───────────
  const brg = [];
  brg.push(titleRow('BEARING ENGINEERING SPECIFICATION'));
  brg.push(subTitleRow(`DE: ${d.bearings.de_desig || '—'} (${d.bearings.de_mfr || '—'})  |  NDE: ${d.bearings.nde_desig || '—'}  |  Shaft Ø: ${d.bearings.shaft_dia || '—'} mm`));
  brg.push(blankRow());
  brg.push(hdrRow(['PARAMETER', 'VALUE / SPECIFICATION', 'UNIT', 'ENGINEERING NOTES']));
  brg.push(catRow('BEARING – DRIVE END (DE)  [Position 1]', 'B71C1C'));
  [
    ['Bearing Type', d.bearings.de_type, '', 'Per ISO 15 / DIN 625'],
    ['Designation / Part Number', d.bearings.de_desig, '', 'Verify with manufacturer catalogue'],
    ['Bore Diameter (d)', d.bearings.de_bore, 'mm', 'Must match shaft diameter to fit class'],
    ['Outside Diameter (D)', d.bearings.de_od, 'mm', 'Must match housing bore'],
    ['Width (B)', d.bearings.de_width, 'mm', ''],
    ['Clearance Class', d.bearings.de_clearance, '', 'C3 recommended for hot-running positions'],
    ['Manufacturer', d.bearings.de_mfr, '', ''],
    ['Shaft Fit (ISO)', d.bearings.de_shaftFit, '', 'Rotating inner ring → interference fit'],
    ['Housing Fit (ISO)', d.bearings.de_housingFit, '', 'Stationary outer ring → loose fit H7'],
    ['Installation Method', d.bearings.de_install, '', ''],
    ['Seal / Shield Type', d.bearings.de_sealType, '', ''],
    ['Grease Type', d.bearings.de_grease, '', 'Verify compatibility with seals'],
    ['Lube Nipple Size', d.bearings.de_nipple, '', ''],
    ['Dynamic Load Rating C', d.bearings.de_C, 'kN', 'From manufacturer catalogue'],
    ['Static Load Rating C0', d.bearings.de_C0, 'kN', 'From manufacturer catalogue'],
    ['Applied Radial Load Fr', d.bearings.de_Fr, 'kN', ''],
    ['Applied Axial Load Fa', d.bearings.de_Fa, 'kN', ''],
    ['Regreasing Interval', d.bearings.de_lubInt, 'hrs', 'Per SKF re-lubrication formula'],
    ['Grease Quantity per Regrease', d.bearings.de_lubQty, 'g', '= 0.005 × D × B (SKF method)'],
  ].forEach(([p,v,u,n],i)=>brg.push(dataRow(p,v,u,n,i%2===0)));

  brg.push(blankRow());
  brg.push(catRow('BEARING – NON-DRIVE END (NDE)  [Position 2]', '880E4F'));
  [
    ['Bearing Type', d.bearings.nde_type],
    ['Designation / Part Number', d.bearings.nde_desig, '', 'NDE typically floating (locating via DE only)'],
    ['Bore Diameter (d)', d.bearings.nde_bore, 'mm'],
    ['Outside Diameter (D)', d.bearings.nde_od, 'mm'],
    ['Width (B)', d.bearings.nde_width, 'mm'],
    ['Clearance Class', d.bearings.nde_clearance],
    ['Manufacturer', d.bearings.nde_mfr],
    ['Shaft Fit', d.bearings.nde_shaftFit],
    ['Housing Fit', d.bearings.nde_housingFit],
    ['Seal / Shield Type', d.bearings.nde_seal],
    ['Grease Type', d.bearings.nde_grease],
    ['Regreasing Interval', d.bearings.nde_lubInt, 'hrs'],
    ['Grease Quantity per Regrease', d.bearings.nde_lubQty, 'g'],
  ].forEach(([p,v,u,n],i)=>brg.push(dataRow(p,v,u||'',n||'',i%2===0)));

  brg.push(blankRow());
  brg.push(catRow('SHAFT SPECIFICATION', '1A237E'));
  [
    ['Shaft Diameter', d.bearings.shaft_dia, 'mm'],
    ['Shaft Length', d.bearings.shaft_len, 'mm'],
    ['Shaft Material', d.bearings.shaft_mat, '', 'Min yield strength per duty'],
    ['Surface Treatment', d.bearings.shaft_surf, '', ''],
    ['Keyway Size (W × H)', d.bearings.keyway, 'mm', 'Per ISO 773 / DIN 6885'],
    ['Key Length', d.bearings.keyLen, 'mm'],
    ['Key Material', d.bearings.keyMat],
    ['Shaft Hardness', d.bearings.shaftHardness, 'HRC', 'Measured at bearing seats'],
  ].forEach(([p,v,u,n],i)=>brg.push(dataRow(p,v,u||'',n||'',i%2===0)));

  const ws_brg = XLSX.utils.aoa_to_sheet(brg);
  ws_brg['!cols'] = cols;
  XLSX.utils.book_append_sheet(wb, ws_brg, 'Bearings');

  // ─────────── SHEET 4: FASTENERS ───────────
  const fas = [];
  fas.push(titleRow('FASTENERS, SEALS & GASKETS BOM'));
  fas.push(subTitleRow(`Bolt: ${d.fasteners.boltSize || '—'} / Grade ${d.fasteners.boltGrade || '—'}  |  Total Bolts: ${d.fasteners.totalBolts || '—'} Pcs  |  Torque: ${d.fasteners.torque_Nm || '—'} Nm`));
  fas.push(blankRow());
  fas.push(hdrRow(['PARAMETER', 'VALUE / SPECIFICATION', 'UNIT / QTY', 'ENGINEERING NOTES']));
  fas.push(catRow('STRUCTURAL / MOUNTING BOLTS', '004D40'));
  [
    ['Bolt Type', d.fasteners.boltType, '', 'ISO type designation'],
    ['Nominal Diameter', d.fasteners.boltSize, 'mm', 'ISO metric (M series)'],
    ['Property Class (Grade)', d.fasteners.boltGrade, '', '8.8 = 640 MPa yield; 10.9 = 900 MPa'],
    ['Thread Standard', d.fasteners.threadStd, '', ''],
    ['Bolt Length', d.fasteners.boltLen, 'mm', 'Engaged length ≥ 1.5× bolt diameter'],
    ['Material & Coating', d.fasteners.boltMat, '', ''],
    ['Tightening Torque', d.fasteners.torque_Nm, 'Nm', 'Use calibrated torque wrench. Check K=0.2 factor'],
    ['Quantity per Joint', d.fasteners.qtyPerJoint, 'Pcs/joint', ''],
    ['Number of Joints', d.fasteners.numJoints, 'Joints', ''],
    ['TOTAL BOLT QUANTITY', d.fasteners.totalBolts, 'Pcs', '+ 10% contingency recommended'],
    ['Nut Type', d.fasteners.nutType, '', ''],
    ['Nut Grade', d.fasteners.nutGrade, '', 'Must match or exceed bolt grade'],
    ['Washer Type', d.fasteners.washerType, '', 'Required under nut & head'],
    ['Locking Method', d.fasteners.lockMethod, '', 'Anti-vibration measure'],
  ].forEach(([p,v,u,n],i)=>fas.push(dataRow(p,v,u||'',n||'',i%2===0)));

  fas.push(blankRow());
  fas.push(catRow('SEALS', '3E2723'));
  [
    ['Primary Seal Type', d.fasteners.sealType, '', ''],
    ['Seal Material', d.fasteners.sealMat, '', 'Verify compatibility with fluid/temperature'],
    ['Seal Inside Diameter (ID)', d.fasteners.sealID, 'mm', 'Must match shaft / bore diameter'],
    ['Seal Outside Diameter (OD)', d.fasteners.sealOD, 'mm', 'Must match housing bore'],
    ['Seal Width / Height', d.fasteners.sealW, 'mm', ''],
    ['Seal Part Number / Kit PN', d.fasteners.sealPN, '', 'Order OEM kit where possible'],
    ['Seal Manufacturer', d.fasteners.sealMfr, '', ''],
    ['Operating Temperature Range', d.fasteners.sealTemp, '', 'Confirm with fluid compatibility chart'],
    ['Gasket Type', d.fasteners.gasketType, '', ''],
    ['Gasket Material / Grade', d.fasteners.gasketMat, '', ''],
    ['Seal Quantity', d.fasteners.sealQty, 'Pcs', '+ 20% spares recommended'],
    ['Gasket Quantity', d.fasteners.gasketQty, 'Pcs', ''],
  ].forEach(([p,v,u,n],i)=>fas.push(dataRow(p,v,u||'',n||'',i%2===0)));

  const ws_fas = XLSX.utils.aoa_to_sheet(fas);
  ws_fas['!cols'] = cols;
  XLSX.utils.book_append_sheet(wb, ws_fas, 'Fasteners & Seals');

  // ─────────── SHEET 5: MATERIALS & LUBRICATION ───────────
  const mat = [];
  mat.push(titleRow('MATERIALS, LUBRICATION & COUPLING SPECIFICATION'));
  mat.push(subTitleRow(`Housing: ${d.materials.housingMat || '—'}  |  Oil: ${d.materials.oilGrade || '—'}  |  Coupling: ${d.materials.couplingType || '—'}`));
  mat.push(blankRow());
  mat.push(hdrRow(['PARAMETER', 'VALUE / SPECIFICATION', 'UNIT', 'ENGINEERING NOTES']));
  mat.push(catRow('HOUSING & STRUCTURAL MATERIALS', '37474F'));
  [
    ['Housing / Casing Material', d.materials.housingMat, '', 'Min tensile strength per ISO'],
    ['Housing Surface Treatment', d.materials.housingSurf, '', ''],
    ['Housing Wall Thickness', d.materials.housingThick, 'mm', ''],
    ['Base Plate Thickness', d.materials.baseThick, 'mm', 'Per load transfer calculation'],
    ['Anchor Bolt Size', d.materials.anchorBoltSize, '', ''],
    ['Anchor Bolt Grade', d.materials.anchorBoltGrade, '', ''],
    ['Anchor Bolt Quantity', d.materials.anchorBoltQty, 'Pcs', 'Min 4 for dynamic loads'],
    ['Grout Specification', d.materials.grout, '', 'Non-shrink for all precision equipment'],
    ['Alignment Method', d.materials.alignMethod, '', 'Laser preferred for speeds >750 RPM'],
  ].forEach(([p,v,u,n],i)=>mat.push(dataRow(p,v,u||'',n||'',i%2===0)));

  mat.push(blankRow());
  mat.push(catRow('GEARBOX / OIL LUBRICATION', '1A4971'));
  [
    ['Oil Grade (ISO VG)', d.materials.oilGrade, '', ''],
    ['Oil Brand / Specification', d.materials.oilBrand, '', 'Check OEM approval list'],
    ['Oil Capacity', d.materials.oilCap, 'Litres', 'From nameplate'],
    ['Oil Change Interval', d.materials.oilInterval, 'hrs', 'Reduce to 2000 hrs if >80°C'],
    ['Oil Temperature – Normal', d.materials.oilTempNormal, '°C', 'Optimal range 40–60°C'],
    ['Oil Temperature – Alarm', d.materials.oilTempAlarm, '°C', 'Investigate cause immediately'],
    ['Oil Temperature – Trip', d.materials.oilTempTrip, '°C', 'Shutdown – prevent seal failure'],
    ['Lubrication System', d.materials.lubSystem, '', ''],
    ['Filter Rating', d.materials.filterRating, '', 'Match to system cleanliness target'],
  ].forEach(([p,v,u,n],i)=>mat.push(dataRow(p,v,u||'',n||'',i%2===0)));

  mat.push(blankRow());
  mat.push(catRow('COUPLING SPECIFICATION', '4A0040'));
  [
    ['Coupling Type', d.materials.couplingType, '', ''],
    ['Coupling Manufacturer', d.materials.couplingMfr, '', ''],
    ['Size / Catalogue Number', d.materials.couplingSize, '', ''],
    ['Rated Torque', d.materials.couplingTorque, 'Nm', 'Min 2× motor rated torque'],
    ['Spider / Insert Material', d.materials.couplingSpider, '', ''],
    ['Hub Bore Diameter', d.materials.couplingBore, 'mm', ''],
    ['Hub Bolt Size', d.materials.couplingBoltSize, '', ''],
    ['Hub Bolt Grade', d.materials.couplingBoltGrade, '', ''],
    ['Hub Bolt Quantity', d.materials.couplingBoltQty, 'Pcs', ''],
  ].forEach(([p,v,u,n],i)=>mat.push(dataRow(p,v,u||'',n||'',i%2===0)));

  const ws_mat = XLSX.utils.aoa_to_sheet(mat);
  ws_mat['!cols'] = cols;
  XLSX.utils.book_append_sheet(wb, ws_mat, 'Materials & Lubrication');

  // ─────────── SHEET 6: QA & PROCUREMENT ───────────
  const qa = [];
  qa.push(titleRow('QUALITY ASSURANCE, STANDARDS & PROCUREMENT'));
  qa.push(subTitleRow(`Supplier: ${d.quality.supplier || '—'}  |  Lead Time: ${d.quality.leadTime || '—'} days  |  Est. Cost: UGX ${parseInt(d.quality.unitCost||0).toLocaleString()}`));
  qa.push(blankRow());
  qa.push(hdrRow(['PARAMETER', 'REQUIREMENT / VALUE', 'REFERENCE', 'STATUS / NOTES']));
  qa.push(catRow('APPLICABLE ENGINEERING STANDARDS', '1F4E79'));
  [
    ['Bearing Standard', d.quality.brgStd, 'ISO 15', ''],
    ['Bolt / Fastener Standard', d.quality.boltStd, 'ISO 4014', ''],
    ['Gear / Drive Standard', d.quality.gearStd, 'ISO 6336', ''],
    ['Welding Standard', d.quality.weldStd, 'AWS D1.1', ''],
    ['Surface Finish Standard', d.quality.surfStd, 'ISO 1302', ''],
    ['Painting / Surface Prep Standard', d.quality.paintStd, 'ISO 8501-1', ''],
  ].forEach(([p,v,u,n],i)=>qa.push(dataRow(p,v,u,n,i%2===0)));

  qa.push(blankRow());
  qa.push(catRow('INSPECTION & TESTING REQUIREMENTS', '1B5E20'));
  [
    ['Dimensional Inspection', d.quality.dimInspect, '', ''],
    ['Non-Destructive Testing (NDT)', d.quality.ndt, '', ''],
    ['Material Certification Required', d.quality.matCert, '', 'EN 10204'],
    ['Hardness Testing', d.quality.hardTest, '', ''],
    ['Vibration Acceptance Criterion', d.quality.vibAccept, 'mm/s RMS', 'ISO 10816-3'],
    ['Alignment Acceptance Criterion', d.quality.alignAccept, 'mm TIR', ''],
    ['Special QA Notes / Hold Points', d.quality.notes, '', ''],
  ].forEach(([p,v,u,n],i)=>qa.push(dataRow(p,v,u||'',n||'',i%2===0)));

  qa.push(blankRow());
  qa.push(catRow('PROCUREMENT & INVENTORY', '7E3700'));
  [
    ['Preferred Supplier', d.quality.supplier, '', ''],
    ['Alternate Supplier', d.quality.supplier2, '', 'Use if lead time > threshold'],
    ['Estimated Lead Time', d.quality.leadTime, 'Days', ''],
    ['Unit Cost Estimate', d.quality.unitCost, 'UGX', ''],
    ['Total Cost Estimate', d.quality.totalCost || parseInt(d.quality.unitCost||0)*parseInt(d.millSection.qty||1), 'UGX', ''],
    ['Storage Location', d.quality.storage, '', ''],
    ['Minimum Stock Level', d.quality.minStock, 'Units', 'Below this → reorder'],
    ['Reorder Quantity', d.quality.reorderQty, 'Units', ''],
    ['Incoterms (if imported)', d.quality.incoterms, '', ''],
  ].forEach(([p,v,u,n],i)=>qa.push(dataRow(p,v,u||'',n||'',i%2===0)));

  const ws_qa = XLSX.utils.aoa_to_sheet(qa);
  ws_qa['!cols'] = cols;
  XLSX.utils.book_append_sheet(wb, ws_qa, 'QA & Procurement');

  // ─────────── SHEET 7: COMPLETE BOM FLAT LIST ───────────
  const flat = [];
  flat.push(titleRow('COMPLETE ENGINEERING BOM – FLAT LIST (All Items)'));
  flat.push(subTitleRow(`BOM Ref: ${d.project.bomRef || '—'}  |  Component: ${d.millSection.detail || '—'}  |  Exported: ${new Date().toLocaleString()}`));
  flat.push(blankRow());
  flat.push(hdrRow(['ITEM No.', 'CATEGORY', 'PARAMETER', 'VALUE / SPECIFICATION', 'UNIT', 'QTY', 'NOTES']));

  const flatCols = [{ wch: 8 }, { wch: 22 }, { wch: 35 }, { wch: 38 }, { wch: 10 }, { wch: 7 }, { wch: 28 }];

  let itemNo = 1;
  function flatRow(cat, param, value, unit = '', qty = '', note = '') {
    const alt = itemNo % 2 === 0;
    const bg = alt ? 'EBF3FB' : 'FFFFFF';
    const mkC = (v, bold = false) => ({v: v||'', t:'s', s:{
      font:{bold, sz:9, name:'Arial', color:{rgb: bold ? '0B1E3D' : '333333'}},
      fill:{fgColor:{rgb:bg}},
      alignment:{horizontal:'left', vertical:'center', wrapText:true},
      border:{top:{style:'thin',color:{rgb:'E0E0E0'}},bottom:{style:'thin',color:{rgb:'E0E0E0'}},
              left:{style:'thin',color:{rgb:'E0E0E0'}},right:{style:'thin',color:{rgb:'E0E0E0'}}}
    }});
    const row = [mkC(itemNo.toString()), mkC(cat, true), mkC(param), mkC(value, true), mkC(unit), mkC(qty), mkC(note)];
    itemNo++;
    return row;
  }

  // Flatten all data
  const allData = [
    ['Project', 'BOM Reference', d.project.bomRef, '', ''],
    ['Project', 'Work Order No.', d.project.woNo, '', ''],
    ['Project', 'Revision', d.project.rev, '', ''],
    ['Project', 'Prepared By', d.project.prepBy, '', ''],
    ['Project', 'Approved By', d.project.appBy, '', ''],
    ['Project', 'Plant Name', d.project.plantName, '', ''],
    ['Project', 'Asset Tag', d.project.assetTag, '', ''],
    ['Project', 'Priority', d.project.priority, '', ''],
    ['Mill Section', 'Section', d.millSection.millLabel, '', ''],
    ['Mill Section', 'Subcomponent', d.millSection.sub, '', ''],
    ['Mill Section', 'Component', d.millSection.detail, '', ''],
    ['Mill Section', 'Tag No.', d.millSection.tagNo, '', ''],
    ['Mill Section', 'Quantity', d.millSection.qty, 'Units', d.millSection.qty],
    ['Mill Section', 'Criticality', d.millSection.criticality, '', ''],
    ['Drive System', 'Drive Type', d.drive.driveType, '', ''],
    ['Drive System', 'Power', d.drive.power_kw, 'kW', ''],
    ['Drive System', 'Input Speed', d.drive.input_rpm, 'RPM', ''],
    ['Drive System', 'Output Speed', d.drive.output_rpm, 'RPM', ''],
    ['Drive System', 'Gear Ratio', d.drive.ratio, ':1', ''],
    ['Drive System', 'Output Torque', d.drive.torque, 'Nm', ''],
    ['Motor', 'Manufacturer', d.motor.manufacturer, '', ''],
    ['Motor', 'Rated Power', d.motor.kw, 'kW', ''],
    ['Motor', 'Voltage', d.motor.voltage, 'V', ''],
    ['Motor', 'Full Load Current', d.motor.fla, 'A', ''],
    ['Motor', 'Full Load Speed', d.motor.flRpm, 'RPM', ''],
    ['Motor', 'IP Rating', d.motor.ip, '', ''],
    ['Motor', 'Insulation Class', d.motor.insClass, '', ''],
    ['Motor', 'Foot Bolt Size', d.motor.footBoltSize, '', d.motor.footBoltQty],
    ['Bearing DE', 'Type', d.bearings.de_type, '', ''],
    ['Bearing DE', 'Designation', d.bearings.de_desig, '', '1'],
    ['Bearing DE', 'Bore × OD × Width', `${d.bearings.de_bore}×${d.bearings.de_od}×${d.bearings.de_width}`, 'mm', '1'],
    ['Bearing DE', 'Clearance Class', d.bearings.de_clearance, '', ''],
    ['Bearing DE', 'Manufacturer', d.bearings.de_mfr, '', ''],
    ['Bearing DE', 'Shaft Fit', d.bearings.de_shaftFit, '', ''],
    ['Bearing DE', 'Housing Fit', d.bearings.de_housingFit, '', ''],
    ['Bearing DE', 'Installation Method', d.bearings.de_install, '', ''],
    ['Bearing DE', 'Seal Type', d.bearings.de_sealType, '', ''],
    ['Bearing DE', 'Grease Type', d.bearings.de_grease, '', ''],
    ['Bearing DE', 'Regreasing Interval', d.bearings.de_lubInt, 'hrs', ''],
    ['Bearing DE', 'Grease Qty', d.bearings.de_lubQty, 'g/regrease', ''],
    ['Bearing NDE', 'Designation', d.bearings.nde_desig, '', '1'],
    ['Bearing NDE', 'Bore × OD × Width', `${d.bearings.nde_bore}×${d.bearings.nde_od}×${d.bearings.nde_width}`, 'mm', '1'],
    ['Bearing NDE', 'Clearance Class', d.bearings.nde_clearance, '', ''],
    ['Bearing NDE', 'Manufacturer', d.bearings.nde_mfr, '', ''],
    ['Shaft', 'Diameter', d.bearings.shaft_dia, 'mm', ''],
    ['Shaft', 'Length', d.bearings.shaft_len, 'mm', ''],
    ['Shaft', 'Material', d.bearings.shaft_mat, '', ''],
    ['Shaft', 'Surface Treatment', d.bearings.shaft_surf, '', ''],
    ['Shaft', 'Keyway', d.bearings.keyway, 'mm', ''],
    ['Fasteners', 'Bolt Type', d.fasteners.boltType, '', d.fasteners.totalBolts],
    ['Fasteners', 'Bolt Size', d.fasteners.boltSize, '', ''],
    ['Fasteners', 'Bolt Grade', d.fasteners.boltGrade, '', ''],
    ['Fasteners', 'Bolt Length', d.fasteners.boltLen, 'mm', ''],
    ['Fasteners', 'Bolt Material', d.fasteners.boltMat, '', ''],
    ['Fasteners', 'Tightening Torque', d.fasteners.torque_Nm, 'Nm', ''],
    ['Fasteners', 'Total Bolt Quantity', d.fasteners.totalBolts, 'Pcs', d.fasteners.totalBolts],
    ['Fasteners', 'Nut Type', d.fasteners.nutType, '', d.fasteners.totalBolts],
    ['Fasteners', 'Washer Type', d.fasteners.washerType, '', ''],
    ['Fasteners', 'Locking Method', d.fasteners.lockMethod, '', ''],
    ['Seals', 'Seal Type', d.fasteners.sealType, '', d.fasteners.sealQty],
    ['Seals', 'Seal Material', d.fasteners.sealMat, '', ''],
    ['Seals', 'Seal ID × OD × W', `${d.fasteners.sealID}×${d.fasteners.sealOD}×${d.fasteners.sealW}`, 'mm', ''],
    ['Seals', 'Seal Part Number', d.fasteners.sealPN, '', ''],
    ['Seals', 'Seal Manufacturer', d.fasteners.sealMfr, '', ''],
    ['Seals', 'Gasket Type', d.fasteners.gasketType, '', d.fasteners.gasketQty],
    ['Seals', 'Gasket Material', d.fasteners.gasketMat, '', ''],
    ['Materials', 'Housing Material', d.materials.housingMat, '', ''],
    ['Materials', 'Surface Treatment', d.materials.housingSurf, '', ''],
    ['Materials', 'Anchor Bolt Size', d.materials.anchorBoltSize, '', d.materials.anchorBoltQty],
    ['Lubrication', 'Oil Grade', d.materials.oilGrade, '', ''],
    ['Lubrication', 'Oil Brand', d.materials.oilBrand, '', ''],
    ['Lubrication', 'Oil Capacity', d.materials.oilCap, 'Litres', ''],
    ['Lubrication', 'Oil Change Interval', d.materials.oilInterval, 'hrs', ''],
    ['Coupling', 'Coupling Type', d.materials.couplingType, '', '1'],
    ['Coupling', 'Manufacturer', d.materials.couplingMfr, '', ''],
    ['Coupling', 'Size / Catalogue No.', d.materials.couplingSize, '', ''],
    ['Coupling', 'Rated Torque', d.materials.couplingTorque, 'Nm', ''],
    ['QA', 'Bearing Standard', d.quality.brgStd, '', ''],
    ['QA', 'Bolt Standard', d.quality.boltStd, '', ''],
    ['QA', 'Dimensional Inspection', d.quality.dimInspect, '', ''],
    ['QA', 'NDT Method', d.quality.ndt, '', ''],
    ['QA', 'Material Cert Required', d.quality.matCert, '', ''],
    ['QA', 'Vibration Acceptance', d.quality.vibAccept, 'mm/s', ''],
    ['Procurement', 'Preferred Supplier', d.quality.supplier, '', ''],
    ['Procurement', 'Lead Time', d.quality.leadTime, 'Days', ''],
    ['Procurement', 'Unit Cost', d.quality.unitCost, 'UGX', ''],
    ['Procurement', 'Total Cost', d.quality.totalCost, 'UGX', ''],
    ['Procurement', 'Storage Location', d.quality.storage, '', ''],
    ['Procurement', 'Min Stock Level', d.quality.minStock, 'Units', ''],
    ['Procurement', 'Reorder Quantity', d.quality.reorderQty, 'Units', ''],
  ];

  allData.forEach(([cat, param, value, unit, qty, note]) => {
    flat.push(flatRow(cat, param, value, unit, qty, note || ''));
  });

  const ws_flat = XLSX.utils.aoa_to_sheet(flat);
  ws_flat['!cols'] = flatCols;
  XLSX.utils.book_append_sheet(wb, ws_flat, 'Complete BOM – All Items');

  // ─────────── EXPORT ───────────
  const ref = d.project.bomRef || 'BOM';
  const dt = new Date().toISOString().slice(0,10).replace(/-/g,'');
  XLSX.writeFile(wb, `MMC_Steel_Engineering_BOM_${ref}_${dt}.xlsx`);
}

// ══════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════
document.getElementById('bomDate').value = new Date().toISOString().slice(0,10);
updateSavedList();
</script>
</body>
</html>
