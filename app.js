// app.js (FULL)
// =====================================================
// FFTH Rebuild Helper (Client-side)
// Features:
// - Upload CSV (ABD Existing) + KMZ (ABD KMZ)
// - Auto-match based on normalized house number: ST_NUM <-> Placemark.name
// - Tabs Missing / Matched / All
// - Search + filter by ST_NAME
// - Bulk mode by ST_NAME (queue) for fast clicking points on map (LOCAL markers)
// - EXPORT: keeps original ABD KMZ structure unchanged; ONLY adds folder:
//   REBUILD_HELPER/MATCHED containing "tag" markers for matched rows.
// =====================================================

const $ = (id) => document.getElementById(id);

const csvFile = $("csvFile");
const kmzFile = $("kmzFile");
const btnProcess = $("btnProcess");
const btnExport = $("btnExport");
const statusEl = $("status");
const tableWrap = $("tableWrap");
const modePill = $("modePill");

const searchBox = $("searchBox");
const streetSelect = $("streetSelect");
const viewSelect = $("viewSelect");

const btnBulkStart = $("btnBulkStart");
const btnBulkSkip = $("btnBulkSkip");
const btnUndo = $("btnUndo");
const btnBulkStop = $("btnBulkStop");

const selInfo = $("selInfo");
const bulkInfo = $("bulkInfo");

let map, layerKMZ, layerAdded;
let csvRows = [];
let kmzPoints = [];
let kmlDoc = null;
let kmzKmlName = "doc.kml"; // will be set from KMZ
let matches = [];
let addedPoints = []; // local-only markers: { rowIndex, lat, lng, marker }

let currentTab = "missing";
let selectedRowKey = null;

let bulkMode = {
  active: false,
  queue: [],
  pointer: 0,
  street: ""
};

// ---------------------- Map init ----------------------
function initMap() {
  map = L.map("map").setView([-7.2575, 112.7521], 12);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 20,
    attribution: "© OpenStreetMap"
  }).addTo(map);

  layerKMZ = L.layerGroup().addTo(map);
  layerAdded = L.layerGroup().addTo(map);

  map.on("click", (e) => {
    // LOCAL marker add only (does NOT export into KMZ structure)
    const targetIdx = getTargetAddRowIndex();
    if (targetIdx === null) return;

    const m = matches[targetIdx];
    if (!m || (m.status !== "MISSING" && m.status !== "REVIEW_ADD")) return;

    const { lat, lng } = e.latlng;

    const existingAdded = addedPoints.find(x => x.rowIndex === targetIdx);
    if (existingAdded) {
      existingAdded.lat = lat;
      existingAdded.lng = lng;
      if (existingAdded.marker) existingAdded.marker.setLatLng([lat, lng]);
    } else {
      const mk = L.circleMarker([lat, lng], { radius: 7 }).addTo(layerAdded);
      mk.bindPopup(renderPopupForRow(m.row, "LOCAL CLICK"));
      addedPoints.push({ rowIndex: targetIdx, lat, lng, marker: mk });
    }

    // keep status as MISSING / REVIEW_ADD (since we are not exporting ADDED)
    // but you can still use this to guide your manual editing.
    // If you want later: export ADDED folder too, say the word.
    if (bulkMode.active) {
      bulkMode.pointer += 1;
      const nextIdx = getBulkCurrentRowIndex();
      if (nextIdx !== null) {
        selectedRowKey = nextIdx;
        zoomToRow(nextIdx);
      } else {
        stopBulk("Bulk selesai ✅");
      }
    }

    renderTable();
    refreshHUD();
    setStatus(`Local mark placed for row index=${targetIdx}. (Tidak diexport ke KMZ)`);
  });
}

function getTargetAddRowIndex() {
  if (bulkMode.active) return getBulkCurrentRowIndex();
  if (selectedRowKey === null) return null;
  const m = matches[selectedRowKey];
  if (!m) return null;
  if (m.status !== "MISSING" && m.status !== "REVIEW_ADD") return null;
  return selectedRowKey;
}

function getBulkCurrentRowIndex() {
  if (!bulkMode.active) return null;
  if (bulkMode.pointer >= bulkMode.queue.length) return null;
  return bulkMode.queue[bulkMode.pointer];
}

// ---------------------- Helpers ----------------------
function setStatus(msg) {
  statusEl.textContent = `Status:\n${msg}`;
}

function normalizeHouse(s) {
  if (s === null || s === undefined) return "";
  let t = String(s).trim().toUpperCase();
  t = t.replace(/\s+/g, "");
  t = t.replace(/[._]/g, "");
  t = t.replace(/–/g, "-").replace(/—/g, "-");
  t = t.replace(/-/g, "");
  const m = t.match(/^(\d+)(.*)$/);
  if (m) {
    const num = String(parseInt(m[1], 10));
    const suf = m[2] || "";
    if (!isNaN(Number(num))) t = num + suf;
  }
  return t;
}

function safeGet(row, key) {
  return row && Object.prototype.hasOwnProperty.call(row, key) ? row[key] : "";
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderPopupForRow(row, title) {
  const stNum = safeGet(row, "ST_NUM");
  const stName = safeGet(row, "ST_NAME");
  const rt = safeGet(row, "RT");
  const rw = safeGet(row, "RW");
  const block = safeGet(row, "BLOCK");
  return `
    <div style="font-family:Arial;font-size:13px;">
      <b>${escapeHtml(title)}</b><br/>
      <b>No:</b> ${escapeHtml(stNum || "-")}<br/>
      <b>Jalan:</b> ${escapeHtml(stName || "-")}<br/>
      <b>RT/RW:</b> ${escapeHtml(rt || "-")} / ${escapeHtml(rw || "-")}<br/>
      <b>Block:</b> ${escapeHtml(block || "-")}<br/>
    </div>
  `;
}

function zoomToRow(idx) {
  const m = matches[idx];
  if (!m) return;

  if (m.matchPoint) {
    map.setView([m.matchPoint.lat, m.matchPoint.lng], 18);
    return;
  }

  const ap = addedPoints.find(x => x.rowIndex === idx);
  if (ap) map.setView([ap.lat, ap.lng], 18);
}

// ---------------------- CSV ----------------------
async function parseCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data),
      error: (err) => reject(err)
    });
  });
}

function populateStreetSelect(rows) {
  const set = new Set();
  rows.forEach(r => {
    const st = String(safeGet(r, "ST_NAME") || "").trim();
    if (st) set.add(st);
  });
  const streets = Array.from(set).sort((a, b) => a.localeCompare(b, "id"));
  streetSelect.innerHTML =
    `<option value="">(All streets)</option>` +
    streets.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");
}

// ---------------------- KMZ / KML ----------------------
async function readKMZasKMLText(file) {
  const zip = await JSZip.loadAsync(file);
  const kmlName =
    Object.keys(zip.files).find((n) => n.toLowerCase().endsWith(".kml")) || null;
  if (!kmlName) throw new Error("Tidak menemukan file .kml di dalam KMZ.");
  const kmlText = await zip.files[kmlName].async("text");
  return { kmlText, kmlName };
}

function parseXML(text) {
  const p = new DOMParser();
  const doc = p.parseFromString(text, "application/xml");
  const parseError = doc.getElementsByTagName("parsererror");
  if (parseError && parseError.length) throw new Error("Gagal parse KML (parsererror).");
  return doc;
}

function findFolderByPath(doc, pathNames) {
  const folders = Array.from(doc.getElementsByTagName("Folder"));
  function getName(node) {
    const n = node.getElementsByTagName("name")[0];
    return n ? n.textContent.trim() : "";
  }
  for (const folder of folders) {
    const name = getName(folder).toUpperCase();
    if (name !== pathNames[0].toUpperCase()) continue;

    let cur = folder;
    let ok = true;
    for (let i = 1; i < pathNames.length; i++) {
      const target = pathNames[i].toUpperCase();
      const childFolders = Array.from(cur.children).filter((c) => c.tagName === "Folder");
      const next = childFolders.find((cf) => getName(cf).toUpperCase() === target);
      if (!next) { ok = false; break; }
      cur = next;
    }
    if (ok) return cur;
  }
  return null;
}

function extractPlacemarkPoints(folderOrDoc) {
  const placemarks = Array.from(folderOrDoc.getElementsByTagName("Placemark"));
  const out = [];
  for (const pm of placemarks) {
    const nameNode = pm.getElementsByTagName("name")[0];
    const pmName = nameNode ? nameNode.textContent.trim() : "";

    const pointNode = pm.getElementsByTagName("Point")[0];
    if (!pointNode) continue;
    const coordNode = pointNode.getElementsByTagName("coordinates")[0];
    if (!coordNode) continue;

    const coordText = coordNode.textContent.trim();
    const parts = coordText.split(",").map((x) => x.trim());
    if (parts.length < 2) continue;

    const lng = Number(parts[0]);
    const lat = Number(parts[1]);
    if (!isFinite(lat) || !isFinite(lng)) continue;

    out.push({
      name: pmName,
      norm: normalizeHouse(pmName),
      lat,
      lng,
      placemarkNode: pm
    });
  }
  return out;
}

function drawKMZPoints(points) {
  layerKMZ.clearLayers();
  const latlngs = [];
  points.forEach((p) => {
    const mk = L.circleMarker([p.lat, p.lng], { radius: 6 }).addTo(layerKMZ);
    mk.bindPopup(`<b>KMZ HP</b><br/>${escapeHtml(p.name || "-")}`);
    latlngs.push([p.lat, p.lng]);
  });
  if (latlngs.length) map.fitBounds(L.latLngBounds(latlngs).pad(0.15));
}

// ---------------------- Auto-match ----------------------
function autoMatch(rows, points) {
  const kmzIndex = new Map();
  for (const p of points) {
    if (!kmzIndex.has(p.norm)) kmzIndex.set(p.norm, []);
    kmzIndex.get(p.norm).push(p);
  }

  return rows.map((row, idx) => {
    const stNum = safeGet(row, "ST_NUM");
    const key = normalizeHouse(stNum);
    row.__idx = idx;
    row.__norm = key;

    const candidates = kmzIndex.get(key) || [];
    if (candidates.length === 1) {
      return { row, matchPoint: candidates[0], status: "MATCHED", reason: "EXACT" };
    }
    if (candidates.length > 1) {
      return { row, matchPoint: candidates[0], status: "REVIEW", reason: "DUPLICATE_KMZ" };
    }

    // numeric-only fallback
    const m = key.match(/^(\d+)/);
    if (m) {
      const numOnly = m[1];
      const c2 = kmzIndex.get(numOnly) || [];
      if (c2.length === 1) {
        return { row, matchPoint: c2[0], status: "REVIEW_ADD", reason: "NUMERIC_ONLY" };
      }
    }

    return { row, matchPoint: null, status: "MISSING", reason: "NOT_FOUND" };
  });
}

// ---------------------- Filters + Table ----------------------
function setActiveTab(tab) {
  currentTab = tab;
  document.querySelectorAll(".tab").forEach((t) => {
    t.classList.toggle("active", t.dataset.tab === tab);
  });
  renderTable();
}

function getFilteredMatches() {
  if (!matches.length) return [];

  const q = (searchBox.value || "").trim().toUpperCase();
  const street = (streetSelect.value || "").trim();
  const view = viewSelect.value;

  let arr = matches.map((m, idx) => ({...m, __k: idx}));

  if (view === "tab") {
    if (currentTab === "missing") arr = arr.filter(x => x.status === "MISSING" || x.status === "REVIEW_ADD");
    if (currentTab === "matched") arr = arr.filter(x => x.status === "MATCHED" || x.status === "REVIEW");
  } else if (view === "missing") {
    arr = arr.filter(x => x.status === "MISSING" || x.status === "REVIEW_ADD");
  } else if (view === "street_missing") {
    arr = arr.filter(x => (x.status === "MISSING" || x.status === "REVIEW_ADD") &&
      String(safeGet(x.row, "ST_NAME")||"").trim() === street
    );
  }

  if (street) {
    arr = arr.filter(x => String(safeGet(x.row, "ST_NAME")||"").trim() === street);
  }

  if (q) {
    arr = arr.filter(x => {
      const stNum = String(safeGet(x.row, "ST_NUM") || "").toUpperCase();
      const stName = String(safeGet(x.row, "ST_NAME") || "").toUpperCase();
      return stNum.includes(q) || stName.includes(q);
    });
  }

  return arr;
}

function renderTable() {
  const filtered = getFilteredMatches();

  const rowsHtml = filtered.map((m) => {
    const idx = m.__k;
    const stNum = safeGet(m.row, "ST_NUM");
    const stName = safeGet(m.row, "ST_NAME");
    const status = m.status;
    const reason = m.reason || "";
    const kmzName = m.matchPoint ? m.matchPoint.name : "-";
    const sel = (selectedRowKey === idx) ? "selected" : "";
    return `
      <tr class="${sel}" data-k="${idx}">
        <td>${escapeHtml(stNum || "")}</td>
        <td>${escapeHtml(stName || "")}</td>
        <td><b>${escapeHtml(status)}</b><br/><span style="color:#666;">${escapeHtml(reason)}</span></td>
        <td>${escapeHtml(kmzName || "-")}</td>
      </tr>
    `;
  }).join("");

  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>No Rumah (ST_NUM)</th>
          <th>Nama Jalan (ST_NAME)</th>
          <th>Status</th>
          <th>KMZ HP</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;

  tableWrap.querySelectorAll("tr[data-k]").forEach((tr) => {
    tr.addEventListener("click", () => {
      const k = Number(tr.dataset.k);
      selectedRowKey = k;
      zoomToRow(k);
      renderTable();
      refreshHUD();
    });
  });

  refreshHUD();
}

// ---------------------- Bulk controls ----------------------
function startBulk() {
  const street = (streetSelect.value || "").trim();
  if (!street) {
    setStatus("Pilih ST_NAME dulu untuk Bulk.");
    return;
  }

  const q = [];
  matches.forEach((m, idx) => {
    const st = String(safeGet(m.row, "ST_NAME") || "").trim();
    if (st !== street) return;
    if (m.status === "MISSING" || m.status === "REVIEW_ADD") q.push(idx);
  });

  if (!q.length) {
    setStatus(`Tidak ada Missing pada street: ${street}`);
    return;
  }

  bulkMode.active = true;
  bulkMode.queue = q;
  bulkMode.pointer = 0;
  bulkMode.street = street;

  selectedRowKey = q[0];
  zoomToRow(selectedRowKey);

  modePill.textContent = "BULK: Click map (LOCAL) → auto next";
  btnBulkSkip.disabled = false;
  btnUndo.disabled = false;
  btnBulkStop.disabled = false;

  setStatus(`Bulk started: ${street}\nQueue: ${q.length} items.\nKlik map untuk memberi tanda lokasi (LOCAL).`);
  renderTable();
  refreshHUD();
}

function stopBulk(msg = "Bulk stopped.") {
  bulkMode.active = false;
  bulkMode.queue = [];
  bulkMode.pointer = 0;
  bulkMode.street = "";

  modePill.textContent = "Select row → Click map to add (local)";
  btnBulkSkip.disabled = true;
  btnBulkStop.disabled = true;

  setStatus(msg);
  renderTable();
  refreshHUD();
}

function skipCurrent() {
  if (!bulkMode.active) return;
  bulkMode.pointer += 1;

  const nextIdx = getBulkCurrentRowIndex();
  if (nextIdx !== null) {
    selectedRowKey = nextIdx;
    zoomToRow(nextIdx);
    setStatus("Skipped. Lanjut item berikutnya.");
    renderTable();
    refreshHUD();
  } else {
    stopBulk("Bulk selesai ✅ (setelah skip)");
  }
}

function undoLastAdd() {
  if (!addedPoints.length) {
    setStatus("Tidak ada local mark yang bisa di-undo.");
    return;
  }
  const last = addedPoints.pop();
  if (last.marker) layerAdded.removeLayer(last.marker);

  if (bulkMode.active) {
    bulkMode.pointer = Math.max(0, bulkMode.pointer - 1);
    const cur = getBulkCurrentRowIndex();
    if (cur !== null) {
      selectedRowKey = cur;
      zoomToRow(cur);
    }
  }

  setStatus("Undo last local mark ✅");
  renderTable();
  refreshHUD();
}

// ---------------------- HUD ----------------------
function refreshHUD() {
  if (selectedRowKey === null || !matches[selectedRowKey]) {
    selInfo.textContent = "-";
  } else {
    const r = matches[selectedRowKey].row;
    selInfo.textContent = `${safeGet(r, "ST_NUM") || "-"} • ${safeGet(r, "ST_NAME") || "-"}`;
  }

  if (!bulkMode.active) {
    bulkInfo.textContent = "off";
  } else {
    const total = bulkMode.queue.length;
    const curPos = Math.min(bulkMode.pointer + 1, total);
    bulkInfo.textContent = `${bulkMode.street} (${curPos}/${total})`;
  }

  btnBulkStart.disabled = !(matches.length && streetSelect.value);
  btnUndo.disabled = !addedPoints.length;

  if (!bulkMode.active) {
    const m = (selectedRowKey !== null) ? matches[selectedRowKey] : null;
    if (m && (m.status === "MISSING" || m.status === "REVIEW_ADD")) {
      modePill.textContent = "Click map to add (local) for selected row";
    } else {
      modePill.textContent = "Select row → Click map to add (local)";
    }
  }
}

// ---------------------- EXPORT: add MATCHED folder only ----------------------
function getDocumentNode() {
  const docNode = kmlDoc.getElementsByTagName("Document")[0];
  if (!docNode) throw new Error("KML tidak punya <Document>.");
  return docNode;
}

function ensureStyle(id, kmlColorAABBGGRR, scale=1.1) {
  const docNode = getDocumentNode();
  const styles = Array.from(docNode.getElementsByTagName("Style"));
  const exists = styles.find(s => (s.getAttribute("id") || "") === id);
  if (exists) return;

  const style = kmlDoc.createElement("Style");
  style.setAttribute("id", id);

  const iconStyle = kmlDoc.createElement("IconStyle");
  const color = kmlDoc.createElement("color");
  color.textContent = kmlColorAABBGGRR; // AABBGGRR
  const sc = kmlDoc.createElement("scale");
  sc.textContent = String(scale);

  const icon = kmlDoc.createElement("Icon");
  const href = kmlDoc.createElement("href");
  href.textContent = "http://maps.google.com/mapfiles/kml/paddle/grn-blank.png";
  icon.appendChild(href);

  iconStyle.appendChild(color);
  iconStyle.appendChild(sc);
  iconStyle.appendChild(icon);
  style.appendChild(iconStyle);

  docNode.appendChild(style);
}

function ensureMatchFolderOnly() {
  const docNode = getDocumentNode();

  // REBUILD_HELPER (top)
  let helper = Array.from(docNode.getElementsByTagName("Folder")).find(f => {
    const n = f.getElementsByTagName("name")[0];
    return n && n.textContent.trim() === "REBUILD_HELPER";
  });

  if (!helper) {
    helper = kmlDoc.createElement("Folder");
    const n = kmlDoc.createElement("name");
    n.textContent = "REBUILD_HELPER";
    helper.appendChild(n);
    docNode.appendChild(helper);
  }

  // MATCHED subfolder
  let matched = Array.from(helper.children).filter(x => x.tagName === "Folder").find(f => {
    const n = f.getElementsByTagName("name")[0];
    return n && n.textContent.trim() === "MATCHED";
  });

  if (!matched) {
    matched = kmlDoc.createElement("Folder");
    const n = kmlDoc.createElement("name");
    n.textContent = "MATCHED";
    matched.appendChild(n);
    helper.appendChild(matched);
  }

  ensureStyle("ffthMatched", "ff00ff00", 1.1);
  return { matched };
}

function makePointPlacemark(nameText, lat, lng, styleUrl, extendedPairs = {}) {
  const pm = kmlDoc.createElement("Placemark");

  const name = kmlDoc.createElement("name");
  name.textContent = nameText;
  pm.appendChild(name);

  const su = kmlDoc.createElement("styleUrl");
  su.textContent = styleUrl;
  pm.appendChild(su);

  const ext = kmlDoc.createElement("ExtendedData");
  pm.appendChild(ext);
  for (const [k, v] of Object.entries(extendedPairs)) {
    if (v === undefined || v === null || String(v).trim() === "") continue;
    const d = kmlDoc.createElement("Data");
    d.setAttribute("name", k);
    const val = kmlDoc.createElement("value");
    val.textContent = String(v);
    d.appendChild(val);
    ext.appendChild(d);
  }

  const point = kmlDoc.createElement("Point");
  const coords = kmlDoc.createElement("coordinates");
  coords.textContent = `${lng},${lat},0`;
  point.appendChild(coords);
  pm.appendChild(point);

  return pm;
}

async function exportKMZ_AddMatchedFolderOnly() {
  // ✅ No edits to original HP placemarks/folders.
  // ✅ Only adds REBUILD_HELPER/MATCHED "tag" markers.

  const { matched } = ensureMatchFolderOnly();

  // clear old MATCHED placemarks (so export ulang tidak numpuk)
  Array.from(matched.children)
    .filter(n => n.tagName === "Placemark")
    .forEach(n => matched.removeChild(n));

  for (const m of matches) {
    if (m.status !== "MATCHED") continue;
    if (!m.matchPoint) continue;

    const stNum = safeGet(m.row, "ST_NUM");
    const stName = safeGet(m.row, "ST_NAME");

    const pm = makePointPlacemark(
      `MATCHED - ${stNum || ""}`,
      m.matchPoint.lat,
      m.matchPoint.lng,
      "#ffthMatched",
      { ST_NUM: stNum, ST_NAME: stName, STATUS: "MATCHED", REASON: m.reason || "EXACT" }
    );

    matched.appendChild(pm);
  }

  const xml = new XMLSerializer().serializeToString(kmlDoc);
  const zip = new JSZip();
  zip.file(kmzKmlName || "doc.kml", xml);

  const blob = await zip.generateAsync({ type: "blob" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "ABD_KMZ_UPDATED.kmz";
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---------------------- Wire UI ----------------------
function wireControls() {
  document.querySelectorAll(".tab").forEach((t) => {
    t.onclick = () => setActiveTab(t.dataset.tab);
  });

  searchBox.addEventListener("input", () => renderTable());
  streetSelect.addEventListener("change", () => { renderTable(); refreshHUD(); });
  viewSelect.addEventListener("change", () => renderTable());

  btnBulkStart.addEventListener("click", startBulk);
  btnBulkStop.addEventListener("click", () => stopBulk("Bulk stopped."));
  btnBulkSkip.addEventListener("click", skipCurrent);
  btnUndo.addEventListener("click", undoLastAdd);
}

// ---------------------- Main flow ----------------------
btnProcess.addEventListener("click", async () => {
  try {
    btnExport.disabled = true;
    selectedRowKey = null;

    addedPoints = [];
    layerAdded.clearLayers();
    layerKMZ.clearLayers();
    stopBulk("Bulk reset.");

    if (!csvFile.files[0]) throw new Error("Upload dulu ABD Existing (CSV).");
    if (!kmzFile.files[0]) throw new Error("Upload dulu ABD KMZ (KMZ).");

    setStatus("Parsing CSV...");
    csvRows = await parseCSV(csvFile.files[0]);

    if (!csvRows.length) throw new Error("CSV kosong / tidak terbaca.");
    const cols = Object.keys(csvRows[0] || {});
    if (!cols.includes("ST_NUM")) throw new Error("CSV wajib punya kolom ST_NUM (nomor rumah).");

    populateStreetSelect(csvRows);

    setStatus("Reading KMZ & parsing KML...");
    const { kmlText, kmlName } = await readKMZasKMLText(kmzFile.files[0]);
    kmzKmlName = kmlName || "doc.kml";
    kmlDoc = parseXML(kmlText);

    const homeFolder = findFolderByPath(kmlDoc, ["HP","HOME"]);
    if (homeFolder) kmzPoints = extractPlacemarkPoints(homeFolder);
    else kmzPoints = extractPlacemarkPoints(kmlDoc);

    drawKMZPoints(kmzPoints);

    setStatus(`KMZ points loaded: ${kmzPoints.length}\nAuto-matching...`);
    matches = autoMatch(csvRows, kmzPoints);

    const total = matches.length;
    const nMatched = matches.filter(x => x.status === "MATCHED").length;
    const nReview  = matches.filter(x => x.status === "REVIEW").length;
    const nReviewAdd = matches.filter(x => x.status === "REVIEW_ADD").length;
    const nMissing = matches.filter(x => x.status === "MISSING").length;

    btnExport.disabled = false;

    setStatus(
      `Done.\nCSV rows: ${total}\nMATCHED: ${nMatched}\nREVIEW: ${nReview}\nREVIEW_ADD: ${nReviewAdd}\nMISSING: ${nMissing}\n\nExport akan menambah folder REBUILD_HELPER/MATCHED saja.`
    );

    const firstMissingIdx = matches.findIndex(m => m.status === "MISSING" || m.status === "REVIEW_ADD");
    if (firstMissingIdx >= 0) {
      selectedRowKey = firstMissingIdx;
      zoomToRow(firstMissingIdx);
    }

    renderTable();
    refreshHUD();

  } catch (err) {
    setStatus(`ERROR: ${err.message || err}`);
  }
});

btnExport.addEventListener("click", async () => {
  try {
    if (!kmlDoc) throw new Error("Belum ada data untuk diexport. Klik Process dulu.");
    setStatus("Exporting KMZ (add MATCHED folder only)...");
    await exportKMZ_AddMatchedFolderOnly();
    setStatus("Export done: ABD_KMZ_UPDATED.kmz\nFolder baru: REBUILD_HELPER/MATCHED");
  } catch (err) {
    setStatus(`ERROR: ${err.message || err}`);
  }
});

// Init
initMap();
wireControls();
setActiveTab("missing");
refreshHUD();
