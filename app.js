/* ============================================================
   MM Steel Nesting & Procurement Tool
   ------------------------------------------------------------
   SETUP REQUIRED:
   1. Create a Supabase project (or reuse the MM Platform one).
   2. Run this SQL once in the Supabase SQL editor:

      create table offcut_inventory (
        id uuid primary key default gen_random_uuid(),
        profile text not null,
        grade text not null,
        length_mm numeric not null,
        source_job text,
        note text,
        created_at timestamptz default now()
      );
      alter table offcut_inventory enable row level security;
      create policy "allow all" on offcut_inventory for all using (true) with check (true);

      (The "allow all" policy matches the open-access pattern used by the
      other MM Platform tables. Tighten with real auth later if needed.)

   3. Fill in SUPABASE_URL and SUPABASE_ANON_KEY below.
   ============================================================ */

const SUPABASE_URL = "https://github.com/MMAdmin327/MM-Steel-Nesting-";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVnY21sZXlxYnRqZHd1c3BnYnNpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkwOTQ3MDgsImV4cCI6MjA5NDY3MDcwOH0.Bc43J1OzmTKaVNCdKT1bXvIfak1jcxmCqVuyJKZINfw";

let supabaseClient = null;
try {
  if (SUPABASE_URL.startsWith("http")) {
    supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  }
} catch (e) {
  console.error("Supabase init failed", e);
}

// ---------- Global state ----------
let bomRows = [];          // [{profile, grade, length, qty, description}]
let groupSettings = {};    // key -> stockLength
let offcutInventory = [];  // live from Supabase: [{id, profile, grade, length_mm, source_job, note, created_at}]
let nestResult = null;     // output of runNesting()

// ---------- Helpers ----------
function groupKey(profile, grade) {
  return (profile || "").trim().toLowerCase() + "||" + (grade || "").trim().toLowerCase();
}
function fmt(n) {
  return Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 });
}
function money(n) {
  return Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function setMsg(elId, text, type) {
  const el = document.getElementById(elId);
  if (!text) { el.innerHTML = ""; return; }
  el.innerHTML = `<div class="msg ${type}">${text}</div>`;
}

// ---------- Tab navigation ----------
document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
  });
});

// ---------- BOM upload ----------
document.getElementById("bomFile").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      ingestRows(rows);
    } catch (err) {
      setMsg("uploadMsg", "Could not read file: " + err.message, "error");
    }
  };
  reader.readAsArrayBuffer(file);
});

function ingestRows(rows) {
  const cleaned = [];
  rows.forEach(r => {
    const profile = r.Profile || r.profile || r.PROFILE || "";
    const grade = r.Grade || r.grade || r.GRADE || "Mild Steel";
    const length = Number(r.Length_mm || r.length_mm || r.Length || r.length || 0);
    const qty = Number(r.Qty || r.qty || r.Quantity || r.quantity || 0);
    const description = r.Description || r.description || "";
    if (profile && length > 0 && qty > 0) {
      cleaned.push({ profile: String(profile).trim(), grade: String(grade).trim(), length, qty, description });
    }
  });
  if (cleaned.length === 0) {
    setMsg("uploadMsg", "No valid rows found. Check column headers: Profile, Grade, Length_mm, Qty.", "error");
    return;
  }
  bomRows = bomRows.concat(cleaned);
  setMsg("uploadMsg", `Loaded ${cleaned.length} BOM line(s).`, "ok");
  renderBomTable();
  renderGroupSettings();
}

document.getElementById("loadSampleBtn").addEventListener("click", () => {
  ingestRows([
    { Profile: "50x50x3 Equal Angle", Grade: "Mild Steel", Length_mm: 1850, Qty: 6, Description: "Frame upright" },
    { Profile: "50x50x3 Equal Angle", Grade: "Mild Steel", Length_mm: 900, Qty: 10, Description: "Cross brace" },
    { Profile: "50x50x3 Equal Angle", Grade: "Mild Steel", Length_mm: 2400, Qty: 4, Description: "Long rail" },
    { Profile: "76x76x6 Equal Angle", Grade: "Mild Steel", Length_mm: 3100, Qty: 3, Description: "Column" },
    { Profile: "76x76x6 Equal Angle", Grade: "Mild Steel", Length_mm: 1200, Qty: 8, Description: "Base plate stiffener" },
    { Profile: "IPE 200", Grade: "S355", Length_mm: 4200, Qty: 2, Description: "Beam" },
  ]);
});

document.getElementById("addRowBtn").addEventListener("click", () => {
  bomRows.push({ profile: "", grade: "Mild Steel", length: 0, qty: 1, description: "" });
  renderBomTable();
  renderGroupSettings();
});

function renderBomTable() {
  const card = document.getElementById("bomTableCard");
  const tbody = document.getElementById("bomTableBody");
  if (bomRows.length === 0) { card.style.display = "none"; return; }
  card.style.display = "block";
  tbody.innerHTML = "";
  bomRows.forEach((row, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input class="cell-input" data-i="${i}" data-f="profile" value="${escapeHtml(row.profile)}"></td>
      <td><input class="cell-input" data-i="${i}" data-f="grade" value="${escapeHtml(row.grade)}"></td>
      <td><input class="cell-input" type="number" data-i="${i}" data-f="length" value="${row.length}"></td>
      <td><input class="cell-input" type="number" data-i="${i}" data-f="qty" value="${row.qty}"></td>
      <td><input class="cell-input" data-i="${i}" data-f="description" value="${escapeHtml(row.description)}"></td>
      <td><button class="del-btn" data-i="${i}">✕</button></td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll(".cell-input").forEach(inp => {
    inp.addEventListener("change", (e) => {
      const i = Number(e.target.dataset.i), f = e.target.dataset.f;
      bomRows[i][f] = (f === "length" || f === "qty") ? Number(e.target.value) : e.target.value;
      renderGroupSettings();
    });
  });
  tbody.querySelectorAll(".del-btn").forEach(btn => {
    btn.addEventListener("click", (e) => {
      bomRows.splice(Number(e.target.dataset.i), 1);
      renderBomTable();
      renderGroupSettings();
    });
  });
}

function escapeHtml(s) {
  return String(s || "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// ---------- Nesting settings (per group stock length) ----------
function renderGroupSettings() {
  const card = document.getElementById("nestSettingsCard");
  const tbody = document.getElementById("groupTableBody");
  if (bomRows.length === 0) { card.style.display = "none"; return; }
  card.style.display = "block";

  const groups = {};
  bomRows.forEach(r => {
    if (!r.profile || !r.length || !r.qty) return;
    const key = groupKey(r.profile, r.grade);
    if (!groups[key]) groups[key] = { profile: r.profile, grade: r.grade, totalCuts: 0 };
    groups[key].totalCuts += r.qty;
  });

  const stockLengths = getStockLengthOptions();
  tbody.innerHTML = "";
  Object.entries(groups).forEach(([key, g]) => {
    if (!groupSettings[key]) groupSettings[key] = stockLengths[stockLengths.length - 1] || stockLengths[0];
    const options = stockLengths.map(len =>
      `<option value="${len}" ${Number(groupSettings[key]) === len ? "selected" : ""}>${fmt(len)} mm</option>`
    ).join("");
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(g.profile)}</td>
      <td>${escapeHtml(g.grade)}</td>
      <td>${g.totalCuts}</td>
      <td><select class="cell-input group-len-select" data-key="${key}">${options}</select></td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll(".group-len-select").forEach(sel => {
    sel.addEventListener("change", (e) => {
      groupSettings[e.target.dataset.key] = Number(e.target.value);
    });
  });
}

function getStockLengthOptions() {
  const raw = document.getElementById("stockLengthsInput").value;
  return raw.split(",").map(s => Number(s.trim())).filter(n => n > 0).sort((a, b) => a - b);
}
document.getElementById("stockLengthsInput").addEventListener("change", renderGroupSettings);

// ---------- NESTING ENGINE ----------
// Best-fit-decreasing bin packing, offcuts prioritized over new stock.
function runNesting() {
  const kerf = Number(document.getElementById("kerfInput").value) || 0;
  const minOffcut = Number(document.getElementById("minOffcutInput").value) || 0;

  const groups = {};
  bomRows.forEach(r => {
    if (!r.profile || !r.length || !r.qty) return;
    const key = groupKey(r.profile, r.grade);
    if (!groups[key]) groups[key] = { profile: r.profile, grade: r.grade, cuts: [] };
    for (let i = 0; i < r.qty; i++) {
      groups[key].cuts.push({ length: r.length, description: r.description });
    }
  });

  const result = { groupsOutput: [], errors: [], totalScrap: 0, totalNewBars: 0, totalOffcutsUsed: 0, newOffcutsCreated: [], consumedOffcutIds: [], updatedOffcuts: [] };

  Object.entries(groups).forEach(([key, g]) => {
    const stockLen = groupSettings[key] || getStockLengthOptions()[0];
    const cuts = [...g.cuts].sort((a, b) => b.length - a.length);
    const maxCutLen = Math.max(...cuts.map(c => c.length));
    if (maxCutLen + kerf > stockLen && !offcutInventory.some(o => groupKey(o.profile, o.grade) === key && o.length_mm >= maxCutLen + kerf)) {
      result.errors.push(`${g.profile} / ${g.grade}: a required cut (${fmt(maxCutLen)}mm) is longer than the chosen stock length (${fmt(stockLen)}mm) and no offcut covers it.`);
    }

    // Bins: offcut bins first (existing physical pieces), then new-stock bins opened on demand.
    const offcutBins = offcutInventory
      .filter(o => groupKey(o.profile, o.grade) === key)
      .map(o => ({ id: o.id, isOffcut: true, capacity: o.length_mm, remaining: o.length_mm, cuts: [] }));
    const newBins = [];

    cuts.forEach(cut => {
      const need = cut.length + kerf;
      // 1) best-fit among offcut bins with room
      let candidates = offcutBins.filter(b => b.remaining >= need);
      let bin;
      if (candidates.length > 0) {
        bin = candidates.reduce((best, b) => (b.remaining < best.remaining ? b : best));
      } else {
        // 2) best-fit among already-opened new bins
        candidates = newBins.filter(b => b.remaining >= need);
        if (candidates.length > 0) {
          bin = candidates.reduce((best, b) => (b.remaining < best.remaining ? b : best));
        } else {
          // 3) open a new bin
          bin = { isOffcut: false, capacity: stockLen, remaining: stockLen, cuts: [] };
          newBins.push(bin);
        }
      }
      if (bin.remaining >= need) {
        bin.cuts.push(cut);
        bin.remaining -= need;
      } else {
        result.errors.push(`${g.profile}: cut of ${fmt(cut.length)}mm could not be placed on any available length.`);
      }
    });

    // Resolve offcut bins: used (partially or fully) vs untouched
    const usedOffcutBins = offcutBins.filter(b => b.cuts.length > 0);
    usedOffcutBins.forEach(b => {
      result.consumedOffcutIds.push(b.id);
      if (b.remaining >= minOffcut) {
        result.updatedOffcuts.push({ oldId: b.id, profile: g.profile, grade: g.grade, length_mm: Math.round(b.remaining) });
      } else {
        result.totalScrap += b.remaining;
      }
    });
    result.totalOffcutsUsed += usedOffcutBins.length;

    newBins.forEach(b => {
      result.totalNewBars += 1;
      if (b.remaining >= minOffcut) {
        result.newOffcutsCreated.push({ profile: g.profile, grade: g.grade, length_mm: Math.round(b.remaining) });
      } else {
        result.totalScrap += b.remaining;
      }
    });

    result.groupsOutput.push({
      profile: g.profile, grade: g.grade, stockLen,
      bars: [...usedOffcutBins.map(b => ({ ...b, source: "offcut" })), ...newBins.map(b => ({ ...b, source: "new" }))]
    });
  });

  return result;
}

document.getElementById("runNestBtn").addEventListener("click", () => {
  nestResult = runNesting();
  setMsg("nestMsg", nestResult.errors.length
    ? nestResult.errors.map(e => "⚠ " + e).join("<br>")
    : "Nesting complete — see Cut List and Procurement tabs.",
    nestResult.errors.length ? "error" : "ok");
  renderCutList();
  renderProcurement();
});

// ---------- CUT LIST RENDER ----------
function renderCutList() {
  const content = document.getElementById("cutListContent");
  const statsEl = document.getElementById("cutListStats");
  if (!nestResult) { content.innerHTML = `<div class="empty-state">Run nesting first.</div>`; statsEl.innerHTML = ""; return; }

  const kerf = Number(document.getElementById("kerfInput").value) || 0;
  statsEl.innerHTML = `
    <div class="stat-box good"><div class="label">Offcuts reused</div><div class="value">${nestResult.totalOffcutsUsed}</div></div>
    <div class="stat-box"><div class="label">New bars to cut</div><div class="value">${nestResult.totalNewBars}</div></div>
    <div class="stat-box warn"><div class="label">Scrap generated</div><div class="value">${fmt(nestResult.totalScrap)} mm</div></div>
    <div class="stat-box"><div class="label">Blade kerf used</div><div class="value">${kerf} mm</div></div>
  `;

  let html = "";
  nestResult.groupsOutput.forEach(g => {
    html += `<div class="card"><h2>${escapeHtml(g.profile)} — ${escapeHtml(g.grade)}</h2>`;
    if (g.bars.length === 0) html += `<div class="empty-state">No cuts.</div>`;
    g.bars.forEach((bar, idx) => {
      const label = bar.source === "offcut" ? `Offcut #${String(bar.id).slice(0, 8)}` : `New bar ${idx + 1} (${fmt(bar.capacity)}mm)`;
      const pillClass = bar.source === "offcut" ? "offcut" : "new";
      let segsHtml = "";
      bar.cuts.forEach(c => {
        const pct = (c.length / bar.capacity) * 100;
        segsHtml += `<div class="bar-seg" style="width:${pct}%;" title="${fmt(c.length)}mm ${escapeHtml(c.description || "")}">${fmt(c.length)}</div>`;
      });
      const wastePct = (bar.remaining / bar.capacity) * 100;
      if (wastePct > 0.3) segsHtml += `<div class="bar-seg waste" style="width:${wastePct}%;">${fmt(Math.round(bar.remaining))}</div>`;
      html += `
        <div class="bar-block">
          <div class="bar-head">
