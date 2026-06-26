import { useState, useEffect, useCallback, useMemo } from "react";
import {
  collection, addDoc, updateDoc, doc, onSnapshot,
  query, orderBy, setDoc, getDocs, deleteDoc
} from "firebase/firestore";
import { db } from "./firebase.js";

// ─── Helpers ────────────────────────────────────────────────
function formatTime(ts) {
  if (!ts) return "—";
  return new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}
function formatDate(ts) {
  return new Date(ts).toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
}
function formatDuration(ms) {
  if (!ms || ms < 0) return "0m";
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  return h > 0 ? `${h}h ${m}m` : `${m}m`;
}
function mapsUrl(address) {
  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`;
}
// Link a Google Maps a partir de coordenadas exactas.
function geoMapsUrl(lat, lng) {
  return `https://www.google.com/maps/search/?api=1&query=${lat},${lng}`;
}

// ─── Geolocalización ────────────────────────────────────────
// Pide la ubicación del dispositivo. Devuelve { lat, lng, accuracy }.
// Requiere contexto seguro (https o localhost) y permiso del usuario.
function getCurrentPosition() {
  return new Promise((resolve, reject) => {
    if (!("geolocation" in navigator)) {
      reject(new Error("Geolocalización no soportada"));
      return;
    }
    navigator.geolocation.getCurrentPosition(
      pos => resolve({
        lat: +pos.coords.latitude.toFixed(6),
        lng: +pos.coords.longitude.toFixed(6),
        accuracy: pos.coords.accuracy != null ? Math.round(pos.coords.accuracy) : null,
      }),
      err => reject(err),
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  });
}
// Reverse geocoding gratuito (OpenStreetMap Nominatim), best-effort.
// Devuelve la dirección como texto, o "" si falla / agota tiempo.
async function reverseGeocode(lat, lng) {
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 6000);
    const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${lat}&lon=${lng}&zoom=18&accept-language=es`;
    const res = await fetch(url, { headers: { Accept: "application/json" }, signal: ctrl.signal });
    clearTimeout(t);
    if (!res.ok) return "";
    const data = await res.json();
    return data.display_name || "";
  } catch {
    return "";
  }
}

// ─── Quick Entry Types ───────────────────────────────────────
const QUICK_TYPES = [
  { id: "work",   label: "Work",   icon: "🔧", color: "#22d3a0", border: "rgba(34,211,160,0.4)",  bg: "rgba(34,211,160,0.10)" },
  { id: "travel", label: "Travel", icon: "🚗", color: "#38bdf8", border: "rgba(56,189,248,0.4)",  bg: "rgba(56,189,248,0.10)" },
  { id: "lunch",  label: "Lunch",  icon: "🍽",  color: "#fb923c", border: "rgba(251,146,60,0.4)",  bg: "rgba(251,146,60,0.10)"  },
  { id: "break",  label: "Break",  icon: "☕",  color: "#a78bfa", border: "rgba(167,139,250,0.4)", bg: "rgba(167,139,250,0.10)" },
];
const S = {
  input: {
    width: "100%", background: "#0a0f1c", border: "1px solid #1e3a5f",
    borderRadius: 10, color: "#f1f5f9", padding: "13px 16px",
    fontSize: 15, fontFamily: "inherit", outline: "none", boxSizing: "border-box",
  },
  label: { fontSize: 10, letterSpacing: "0.12em", color: "#64748b", display: "block", marginBottom: 8 },
  card: {
    background: "#0a0f1c", border: "1px solid #1e3a5f",
    borderRadius: 12, padding: "16px 20px",
  },
  btnGreen: {
    padding: "15px", borderRadius: 12, border: "1px solid rgba(34,211,160,0.4)",
    background: "rgba(34,211,160,0.12)", color: "#22d3a0", fontSize: 14,
    fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.08em",
    cursor: "pointer", width: "100%",
  },
  btnRed: {
    padding: "15px", borderRadius: 12, border: "1px solid rgba(239,68,68,0.4)",
    background: "rgba(239,68,68,0.12)", color: "#f87171", fontSize: 14,
    fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.08em",
    cursor: "pointer", width: "100%",
  },
};

// ════════════════════════════════════════════════════════════
//  LÓGICA DE REPORTES (Entradas / Salidas / Completo)
//
//  Cada documento de `timeRecords` es un TURNO con entrada (clockIn*)
//  y, si existe, salida (clockOut*). Aquí se "explota" cada turno en
//  eventos individuales para poder reportarlos por separado.
//  La ubicación se captura por GPS al fichar (clockInLat/Lng,
//  clockOutLat/Lng) y se reverse-geocodifica a dirección de texto
//  (clockInAddress/clockOutAddress). Registros antiguos sin GPS solo
//  tienen la dirección escrita a mano.
// ════════════════════════════════════════════════════════════
const REPORT_TYPES = [
  { id: "completo", label: "Completo (Entrada + Salida)" },
  { id: "entradas", label: "Entradas" },
  { id: "salidas",  label: "Salidas" },
];

// Columnas del reporte (orden fijo, reutilizado por XLSX / CSV / vista previa)
const COLUMNS = [
  "Empleado", "Fecha", "Hora", "Evento", "Actividad",
  "Cliente", "Ticket", "Ubicación / Dirección", "Coordenadas", "Mapa (GPS)", "Nota",
];
const MAP_COL = COLUMNS.indexOf("Mapa (GPS)");        // se omite en PDF
const ADDRESS_COL = COLUMNS.indexOf("Ubicación / Dirección");

function activityLabel(entryType) {
  const qt = QUICK_TYPES.find(q => q.id === (entryType || "work"));
  return qt ? qt.label : (entryType || "—");
}

// Las marcas de tiempo se guardan como string (resultado de `Date()`);
// también soportamos epoch en ms por robustez.
function parseTs(ts) {
  if (ts === null || ts === undefined || ts === "") return null;
  const d = new Date(ts);
  return Number.isNaN(d.getTime()) ? null : d;
}
function fmtDate(d) {
  return d.toLocaleDateString("es-MX", { year: "numeric", month: "2-digit", day: "2-digit" });
}
function fmtTime(d) {
  return d.toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}

// Cada turno produce 1 evento de Entrada y (si tiene salida) 1 de Salida.
function buildEvents(records) {
  const events = [];
  for (const r of records) {
    const base = {
      employeeId: r.employeeId,
      employeeName: r.employeeName || "—",
      activity: activityLabel(r.entryType),
      customer: r.customer || "",
      ticket: r.ticket || "",
      note: r.note || "",
      recordId: r.id,
    };
    const inDate = parseTs(r.clockIn);
    if (inDate) {
      events.push({
        ...base, kind: "in", eventType: "Entrada", date: inDate,
        address: r.clockInAddress || "",
        lat: r.clockInLat ?? null, lng: r.clockInLng ?? null,
      });
    }
    const outDate = parseTs(r.clockOut);
    if (outDate) {
      events.push({
        ...base, kind: "out", eventType: "Salida", date: outDate,
        address: r.clockOutAddress || r.clockInAddress || "",
        // Si no hubo GPS de salida, usa el de entrada como referencia.
        lat: r.clockOutLat ?? r.clockInLat ?? null,
        lng: r.clockOutLng ?? r.clockInLng ?? null,
      });
    }
  }
  return events;
}

// Validación del rango: ambas requeridas + inicio <= fin.
function validateRange(startStr, endStr) {
  if (!startStr || !endStr) {
    return { ok: false, error: "Selecciona la fecha inicial y la fecha final." };
  }
  const start = new Date(`${startStr}T00:00:00`);
  const end = new Date(`${endStr}T23:59:59.999`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return { ok: false, error: "Una de las fechas no es válida." };
  }
  if (start > end) {
    return { ok: false, error: "La fecha inicial no puede ser posterior a la fecha final." };
  }
  return { ok: true, start, end };
}

function filterEvents(events, { reportType, start, end, employeeId }) {
  return events
    .filter(e => {
      if (reportType === "entradas" && e.kind !== "in") return false;
      if (reportType === "salidas" && e.kind !== "out") return false;
      if (start && e.date < start) return false;
      if (end && e.date > end) return false;
      if (employeeId && employeeId !== "all" && e.employeeId !== employeeId) return false;
      return true;
    })
    .sort((a, b) => a.date - b.date);
}

function eventsToRows(events) {
  return events.map(e => {
    const hasCoords = e.lat != null && e.lng != null;
    const coords = hasCoords ? `${e.lat}, ${e.lng}` : "";
    const map = hasCoords ? geoMapsUrl(e.lat, e.lng) : (e.address ? mapsUrl(e.address) : "");
    return [
      e.employeeName,
      fmtDate(e.date),
      fmtTime(e.date),
      e.eventType,
      e.activity,
      e.customer,
      e.ticket,
      e.address,
      coords,
      map,
      e.note,
    ];
  });
}

function timestampSlug() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}
function buildFilename(reportType, ext) {
  return `reporte-asistencia-${reportType}-${timestampSlug()}.${ext}`;
}
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── Exportadores ────────────────────────────────────────────
// xlsx y jspdf son pesados: se cargan con import() dinámico (solo al
// exportar, y por formato) para no inflar el bundle inicial.

async function exportXLSX(events, reportType, meta) {
  const XLSX = await import("xlsx");
  const aoa = [];
  if (meta) {
    aoa.push([meta.title]);
    aoa.push([meta.subtitle]);
    aoa.push([]);
  }
  aoa.push(COLUMNS);
  for (const row of eventsToRows(events)) aoa.push(row);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [
    { wch: 22 }, { wch: 12 }, { wch: 11 }, { wch: 9 }, { wch: 10 },
    { wch: 18 }, { wch: 12 }, { wch: 34 }, { wch: 20 }, { wch: 42 }, { wch: 30 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Asistencia");
  XLSX.writeFile(wb, buildFilename(reportType, "xlsx"));
}

async function exportPDF(events, reportType, meta) {
  const [{ jsPDF }, { default: autoTable }] = await Promise.all([
    import("jspdf"),
    import("jspdf-autotable"),
  ]);
  const docPdf = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageWidth = docPdf.internal.pageSize.getWidth();

  docPdf.setFont("helvetica", "bold");
  docPdf.setFontSize(13);
  docPdf.setTextColor(20, 30, 55);
  docPdf.text(meta?.title || "Reporte de Asistencia", 40, 40);

  let startY = 56;
  if (meta?.subtitle) {
    docPdf.setFont("helvetica", "normal");
    docPdf.setFontSize(8);
    docPdf.setTextColor(90, 100, 120);
    const sub = docPdf.splitTextToSize(meta.subtitle, pageWidth - 80);
    docPdf.text(sub, 40, startY);
    startY += sub.length * 11 + 8;
  }

  // En PDF omitimos la columna del enlace (URL larga) y hacemos
  // clicable la celda de dirección.
  const pdfColumns = COLUMNS.filter((_, i) => i !== MAP_COL);
  const pdfBody = eventsToRows(events).map(r => r.filter((_, i) => i !== MAP_COL));
  const pdfAddressCol = ADDRESS_COL < MAP_COL ? ADDRESS_COL : ADDRESS_COL - 1;

  autoTable(docPdf, {
    startY,
    head: [pdfColumns],
    body: pdfBody,
    styles: { fontSize: 7, cellPadding: 3, overflow: "linebreak", valign: "top" },
    headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontStyle: "bold" },
    alternateRowStyles: { fillColor: [245, 248, 252] },
    columnStyles: { [pdfAddressCol]: { textColor: [14, 116, 180] } },
    didDrawCell: (data) => {
      if (data.section === "body" && data.column.index === pdfAddressCol) {
        const ev = events[data.row.index];
        const hasCoords = ev && ev.lat != null && ev.lng != null;
        if (ev && (ev.address || hasCoords)) {
          const url = hasCoords ? geoMapsUrl(ev.lat, ev.lng) : mapsUrl(ev.address);
          docPdf.link(data.cell.x, data.cell.y, data.cell.width, data.cell.height, { url });
        }
      }
    },
    didDrawPage: () => {
      const page = docPdf.internal.getNumberOfPages();
      docPdf.setFontSize(7);
      docPdf.setTextColor(150, 160, 175);
      docPdf.text(`Página ${page}`, pageWidth - 60, docPdf.internal.pageSize.getHeight() - 16);
    },
  });

  docPdf.save(buildFilename(reportType, "pdf"));
}

function exportCSV(events, reportType) {
  const escape = (v) => {
    const s = String(v ?? "");
    return /[",\n\r;]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [COLUMNS, ...eventsToRows(events)]
    .map(r => r.map(escape).join(","))
    .join("\r\n");
  // BOM para que Excel respete acentos UTF-8.
  const blob = new Blob(["﻿" + lines], { type: "text/csv;charset=utf-8;" });
  triggerDownload(blob, buildFilename(reportType, "csv"));
}

// ════════════════════════════════════════════════════════════
//  PANEL DE REPORTES (UI) — visible solo para admin
// ════════════════════════════════════════════════════════════
const PREVIEW_LIMIT = 50;

// Date -> "YYYY-MM-DD" en hora local (para <input type="date">).
function toInputDate(d) {
  const local = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

const selectStyle = { ...S.input, cursor: "pointer", WebkitAppearance: "none", appearance: "none" };
const btnBlue = {
  padding: "13px", borderRadius: 12, border: "1px solid rgba(14,165,233,0.4)",
  background: "rgba(14,165,233,0.12)", color: "#38bdf8", fontSize: 13,
  fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.06em", cursor: "pointer", width: "100%",
};
const btnSlate = {
  padding: "13px", borderRadius: 12, border: "1px solid #1e3a5f",
  background: "transparent", color: "#94a3b8", fontSize: 13,
  fontFamily: "inherit", fontWeight: 700, letterSpacing: "0.06em", cursor: "pointer", width: "100%",
};

// Nombre seguro para archivos a partir del nombre del empleado (sin acentos).
function slugify(s) {
  return String(s).normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "empleado";
}
// Botón pequeño (XLSX/PDF/CSV) para la lista individual por empleado.
function miniBtn(color, disabled) {
  return {
    padding: "6px 12px", borderRadius: 8, border: `1px solid ${color}55`,
    background: `${color}1f`, color, fontSize: 11, fontFamily: "inherit",
    fontWeight: 700, letterSpacing: "0.05em",
    cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.4 : 1,
  };
}

function ReportsPanel({ records, employees }) {
  const now = new Date();
  const monthAgo = new Date(now.getTime() - 29 * 86400000);

  const [reportType, setReportType] = useState("completo");
  const [startStr, setStartStr] = useState(toInputDate(monthAgo));
  const [endStr, setEndStr] = useState(toInputDate(now));
  const [employeeId, setEmployeeId] = useState("all");
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);
  const [busyEmp, setBusyEmp] = useState(null);

  const events = useMemo(() => buildEvents(records), [records]);
  const validation = useMemo(() => validateRange(startStr, endStr), [startStr, endStr]);
  const filtered = useMemo(() => {
    if (!validation.ok) return [];
    return filterEvents(events, { reportType, start: validation.start, end: validation.end, employeeId });
  }, [events, validation, reportType, employeeId]);
  const previewRows = useMemo(() => eventsToRows(filtered.slice(0, PREVIEW_LIMIT)), [filtered]);

  // Para "reporte individual por empleado": eventos del rango/tipo de TODOS
  // los empleados, y conteo agrupado por persona (una sola pasada).
  const rangeAllFiltered = useMemo(() => {
    if (!validation.ok) return [];
    return filterEvents(events, { reportType, start: validation.start, end: validation.end, employeeId: "all" });
  }, [events, validation, reportType]);
  const countByEmp = useMemo(() => {
    const m = {};
    for (const e of rangeAllFiltered) m[e.employeeId] = (m[e.employeeId] || 0) + 1;
    return m;
  }, [rangeAllFiltered]);

  const employeeName = employeeId === "all"
    ? "Todos los empleados"
    : (employees.find(e => e.id === employeeId)?.name || "—");

  function buildMeta() {
    const label = REPORT_TYPES.find(t => t.id === reportType)?.label || reportType;
    return {
      title: `Reporte de Asistencia — ${label}`,
      subtitle:
        `Rango: ${startStr} a ${endStr}   |   Empleado: ${employeeName}   |   ` +
        `Registros: ${filtered.length}   |   Generado: ${new Date().toLocaleString("es-MX")}`,
    };
  }

  async function run(exporter, formatName) {
    setMsg(null);
    if (!validation.ok) { setMsg({ type: "error", text: validation.error }); return; }
    if (filtered.length === 0) { setMsg({ type: "error", text: "No hay registros para los filtros seleccionados." }); return; }
    setBusy(true);
    try {
      await exporter(filtered, reportType, buildMeta());
      setMsg({ type: "success", text: `✓ ${formatName} generado (${filtered.length} registros).` });
    } catch (err) {
      setMsg({ type: "error", text: `Error al generar ${formatName}: ${err?.message || err}` });
    } finally {
      setBusy(false);
    }
  }

  // Exporta SOLO a un empleado, respetando el rango y tipo seleccionados.
  async function runForEmployee(emp, exporter, formatName) {
    setMsg(null);
    if (!validation.ok) { setMsg({ type: "error", text: validation.error }); return; }
    const empEvents = rangeAllFiltered.filter(e => e.employeeId === emp.id);
    if (empEvents.length === 0) {
      setMsg({ type: "error", text: `No hay registros de ${emp.name} en el rango seleccionado.` });
      return;
    }
    setBusy(true); setBusyEmp(emp.id);
    try {
      const label = REPORT_TYPES.find(t => t.id === reportType)?.label || reportType;
      const meta = {
        title: `Reporte de Asistencia — ${label} — ${emp.name}`,
        subtitle:
          `Empleado: ${emp.name}   |   Rango: ${startStr} a ${endStr}   |   ` +
          `Registros: ${empEvents.length}   |   Generado: ${new Date().toLocaleString("es-MX")}`,
      };
      await exporter(empEvents, `${reportType}-${slugify(emp.name)}`, meta);
      setMsg({ type: "success", text: `✓ ${formatName} de ${emp.name} generado (${empEvents.length} registros).` });
    } catch (err) {
      setMsg({ type: "error", text: `Error al generar ${formatName} de ${emp.name}: ${err?.message || err}` });
    } finally {
      setBusy(false); setBusyEmp(null);
    }
  }

  const canExport = validation.ok && filtered.length > 0 && !busy;

  return (
    <div style={{ maxWidth: 980, margin: "0 auto", display: "flex", flexDirection: "column", gap: 20 }}>

      {/* ── Filtros ─────────────────────────────────────── */}
      <div style={S.card}>
        <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 16 }}>FILTROS DEL REPORTE</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
          <div>
            <label style={S.label}>FECHA INICIAL <span style={{ color: "#f87171" }}>*</span></label>
            <input type="date" value={startStr} max={endStr || undefined}
              onChange={e => setStartStr(e.target.value)} style={S.input} />
          </div>
          <div>
            <label style={S.label}>FECHA FINAL <span style={{ color: "#f87171" }}>*</span></label>
            <input type="date" value={endStr} min={startStr || undefined}
              onChange={e => setEndStr(e.target.value)} style={S.input} />
          </div>
          <div>
            <label style={S.label}>EMPLEADO</label>
            <select value={employeeId} onChange={e => setEmployeeId(e.target.value)} style={selectStyle}>
              <option value="all">Todos los empleados</option>
              {employees.map(emp => (
                <option key={emp.id} value={emp.id}>{emp.name}</option>
              ))}
            </select>
          </div>
          <div>
            <label style={S.label}>TIPO DE REPORTE</label>
            <select value={reportType} onChange={e => setReportType(e.target.value)} style={selectStyle}>
              {REPORT_TYPES.map(t => (
                <option key={t.id} value={t.id}>{t.label}</option>
              ))}
            </select>
          </div>
        </div>
        {!validation.ok ? (
          <div style={{ marginTop: 14, fontSize: 12, color: "#f87171" }}>⚠ {validation.error}</div>
        ) : (
          <div style={{ marginTop: 14, fontSize: 12, color: "#64748b" }}>
            {filtered.length} registro(s) coinciden con los filtros · {employeeName}
          </div>
        )}
      </div>

      {/* ── Exportación ─────────────────────────────────── */}
      <div style={S.card}>
        <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 12 }}>DESCARGAR</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12 }}>
          <button onClick={() => run(exportXLSX, "Excel")} disabled={!canExport}
            style={{ ...S.btnGreen, opacity: canExport ? 1 : 0.4, cursor: canExport ? "pointer" : "not-allowed" }}>
            {busy ? "⏳ Generando…" : "⬇ Excel (.xlsx)"}
          </button>
          <button onClick={() => run(exportPDF, "PDF")} disabled={!canExport}
            style={{ ...btnBlue, opacity: canExport ? 1 : 0.4, cursor: canExport ? "pointer" : "not-allowed" }}>
            {busy ? "⏳ Generando…" : "⬇ PDF"}
          </button>
          <button onClick={() => run(exportCSV, "CSV")} disabled={!canExport}
            style={{ ...btnSlate, opacity: canExport ? 1 : 0.4, cursor: canExport ? "pointer" : "not-allowed" }}>
            {busy ? "⏳ Generando…" : "⬇ CSV"}
          </button>
        </div>
        {msg && (
          <div style={{ marginTop: 12, fontSize: 13, color: msg.type === "error" ? "#f87171" : "#22d3a0" }}>{msg.text}</div>
        )}
      </div>

      {/* ── Reporte individual por empleado ─────────────── */}
      {employees.length > 0 && (
        <div style={S.card}>
          <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 6 }}>REPORTE INDIVIDUAL POR EMPLEADO</div>
          <div style={{ fontSize: 11, color: "#475569", marginBottom: 12 }}>
            Usa el rango de fechas y el tipo de reporte de arriba. Cada botón descarga solo a esa persona.
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {employees.map(emp => {
              const n = countByEmp[emp.id] || 0;
              const disabled = !validation.ok || n === 0 || busy;
              const generating = busyEmp === emp.id;
              return (
                <div key={emp.id} style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  gap: 10, padding: "8px 0", borderBottom: "1px solid #111c30", flexWrap: "wrap",
                }}>
                  <div style={{ minWidth: 150 }}>
                    <span style={{ color: "#f1f5f9", fontSize: 13 }}>{emp.name}</span>
                    <span style={{ color: "#475569", fontSize: 11, marginLeft: 8 }}>{n} registro(s)</span>
                  </div>
                  <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    {generating ? (
                      <span style={{ color: "#64748b", fontSize: 11 }}>⏳ Generando…</span>
                    ) : (
                      <>
                        <button onClick={() => runForEmployee(emp, exportXLSX, "Excel")} disabled={disabled} style={miniBtn("#22d3a0", disabled)}>XLSX</button>
                        <button onClick={() => runForEmployee(emp, exportPDF, "PDF")} disabled={disabled} style={miniBtn("#38bdf8", disabled)}>PDF</button>
                        <button onClick={() => runForEmployee(emp, exportCSV, "CSV")} disabled={disabled} style={miniBtn("#94a3b8", disabled)}>CSV</button>
                      </>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── Vista previa ────────────────────────────────── */}
      <div style={S.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b" }}>VISTA PREVIA</div>
          <div style={{ fontSize: 11, color: "#475569" }}>
            {filtered.length > PREVIEW_LIMIT ? `Mostrando ${PREVIEW_LIMIT} de ${filtered.length}` : `${filtered.length} fila(s)`}
          </div>
        </div>
        {filtered.length === 0 ? (
          <div style={{ color: "#475569", textAlign: "center", padding: 28, fontSize: 13 }}>
            No hay registros para mostrar con los filtros actuales.
          </div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr>
                  {COLUMNS.map(c => (
                    <th key={c} style={{
                      textAlign: "left", padding: "8px 10px", whiteSpace: "nowrap",
                      color: "#64748b", borderBottom: "1px solid #1e3a5f", letterSpacing: "0.04em", fontWeight: 600,
                    }}>{c}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {previewRows.map((row, ri) => (
                  <tr key={ri}>
                    {row.map((cell, ci) => {
                      const isMap = COLUMNS[ci] === "Mapa (GPS)";
                      const isEvent = COLUMNS[ci] === "Evento";
                      return (
                        <td key={ci} style={{
                          padding: "7px 10px", borderBottom: "1px solid #111c30",
                          color: isEvent ? (cell === "Entrada" ? "#22d3a0" : "#f87171") : "#cbd5e1",
                          whiteSpace: isMap ? "nowrap" : "normal",
                          maxWidth: 260, overflow: "hidden", textOverflow: "ellipsis",
                        }}>
                          {isMap && cell ? (
                            <a href={cell} target="_blank" rel="noopener noreferrer"
                              style={{ color: "#0ea5e9", textDecoration: "none" }}>Ver mapa →</a>
                          ) : (cell || <span style={{ color: "#334155" }}>—</span>)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
//  LOGIN SCREEN
// ════════════════════════════════════════════════════════════
function LoginScreen({ employees, onLogin }) {
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  function handleDigit(d) { if (pin.length < 4) setPin(p => p + d); }
  function handleDel() { setPin(p => p.slice(0, -1)); setError(""); }

  useEffect(() => {
    if (pin.length === 4) {
      setLoading(true);
      setTimeout(() => {
        const match = employees.find(e => e.pin === pin);
        if (match) { onLogin(match); }
        else { setError("Incorrect PIN"); setPin(""); }
        setLoading(false);
      }, 300);
    }
  }, [pin, employees, onLogin]);

  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ width: 36, height: 36, background: "linear-gradient(135deg,#22d3a0,#0ea5e9)", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, marginBottom: 16 }}>⏱</div>
      <div style={{ fontWeight: 700, fontSize: 18, letterSpacing: "0.1em", color: "#f1f5f9", marginBottom: 4 }}>APS FIELDCLOCK</div>
      <div style={{ fontSize: 11, color: "#475569", letterSpacing: "0.14em", marginBottom: 40 }}>ENTER YOUR PIN</div>

      {/* PIN dots */}
      <div style={{ display: "flex", gap: 14, marginBottom: 32 }}>
        {[0, 1, 2, 3].map(i => (
          <div key={i} style={{
            width: 18, height: 18, borderRadius: "50%",
            border: "2px solid",
            borderColor: pin.length > i ? "#22d3a0" : "#1e3a5f",
            background: pin.length > i ? "#22d3a0" : "transparent",
            transition: "all 0.15s",
            boxShadow: pin.length > i ? "0 0 8px #22d3a0" : "none",
          }} />
        ))}
      </div>

      {error && <div style={{ color: "#f87171", fontSize: 13, marginBottom: 20 }}>{error}</div>}

      {/* Numpad */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 72px)", gap: 12 }}>
        {[1,2,3,4,5,6,7,8,9,"",0,"⌫"].map((d, i) => (
          <button key={i} onClick={() => d === "⌫" ? handleDel() : d !== "" && handleDigit(String(d))}
            disabled={loading || d === ""}
            style={{
              height: 72, borderRadius: 14,
              border: "1px solid",
              borderColor: d === "" ? "transparent" : "#1e3a5f",
              background: d === "" ? "transparent" : "#0a0f1c",
              color: d === "⌫" ? "#64748b" : "#f1f5f9",
              fontSize: d === "⌫" ? 20 : 22,
              fontFamily: "inherit",
              fontWeight: 500,
              cursor: d === "" ? "default" : "pointer",
              transition: "all 0.1s",
            }}>{d}</button>
        ))}
      </div>

      <div style={{ marginTop: 40, fontSize: 11, color: "#334155", textAlign: "center" }}>
        Contact your admin if you forgot your PIN
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════
//  MAIN APP
// ════════════════════════════════════════════════════════════
export default function App() {
  const [employees, setEmployees] = useState([]);
  const [records, setRecords] = useState([]);
  const [user, setUser] = useState(null); // logged-in employee obj
  const [view, setView] = useState("clock");
  const [ticket, setTicket] = useState("");
  const [customer, setCustomer] = useState("");
  const [note, setNote] = useState("");
  const [entryType, setEntryType] = useState("work"); // work | lunch | break | travel
  const [statusMsg, setStatusMsg] = useState(null);
  const [liveTime, setLiveTime] = useState(new Date());
  const [loadingDb, setLoadingDb] = useState(true);
  // Admin
  const [adminPinInput, setAdminPinInput] = useState("");
  const [adminAuthed, setAdminAuthed] = useState(false);
  const ADMIN_PIN = "9201"; // Change this in code before deploying!
  const [newEmpName, setNewEmpName] = useState("");
  const [newEmpPin, setNewEmpPin] = useState("");
  const [filterEmp, setFilterEmp] = useState("all");

  // Live clock
  useEffect(() => {
    const t = setInterval(() => setLiveTime(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  // Subscribe to Firestore employees
  useEffect(() => {
    const unsub = onSnapshot(collection(db, "employees"), snap => {
      setEmployees(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    });
    return unsub;
  }, []);

  // Subscribe to Firestore records
  useEffect(() => {
    const q = query(collection(db, "timeRecords"), orderBy("clockIn", "desc"));
    const unsub = onSnapshot(q, snap => {
      setRecords(snap.docs.map(d => ({ id: d.id, ...d.data() })));
      setLoadingDb(false);
    });
    return unsub;
  }, []);

  function flash(type, msg) {
    setStatusMsg({ type, msg });
    setTimeout(() => setStatusMsg(null), 4000);
  }

  const currentRecord = user ? records.find(r => r.employeeId === user.id && !r.clockOut) : null;

  async function handleClockIn() {
    if (!user) return;
    if (currentRecord) return flash("error", "You are already clocked in.");
    // Ubicación 100% automática por GPS (ya no se escribe a mano).
    flash("loading", "📍 Obteniendo ubicación…");
    let pos = null, locNote = "";
    try {
      pos = await getCurrentPosition();
    } catch (e) {
      locNote = e?.code === 1 ? " — sin ubicación (permiso denegado)" : " — sin ubicación";
    }
    flash("loading", "Guardando…");
    const geoAddress = pos ? await reverseGeocode(pos.lat, pos.lng) : "";
    try {
      await addDoc(collection(db, "timeRecords"), {
        employeeId: user.id,
        employeeName: user.name,
        clockIn: Date(),
        entryType,
        clockInAddress: geoAddress || null,
        clockInLat: pos ? pos.lat : null,
        clockInLng: pos ? pos.lng : null,
        clockInAccuracy: pos ? pos.accuracy : null,
        customer: entryType === "work" ? customer.trim() : "",
        ticket: entryType === "work" ? ticket.trim() : "",
        note: note.trim(),
        clockOut: null,
        clockOutAddress: null,
        clockOutLat: null,
        clockOutLng: null,
        clockOutAccuracy: null,
      });
      setTicket(""); setCustomer(""); setNote("");
      flash("success", `✓ Clocked IN — ${entryType.toUpperCase()} at ${formatTime(Date.now())}${pos ? " 📍" : locNote}`);
    } catch (e) {
      flash("error", "Save failed: " + (e.message || "Check your connection."));
    }
  }

  async function handleClockOut() {
    if (!user || !currentRecord) return flash("error", "You are not clocked in.");
    // También captura GPS al salir (antes reusaba la dirección de entrada).
    flash("loading", "📍 Obteniendo ubicación…");
    let pos = null, locNote = "";
    try {
      pos = await getCurrentPosition();
    } catch (e) {
      locNote = e?.code === 1 ? " — sin ubicación (permiso denegado)" : " — sin ubicación";
    }
    flash("loading", "Guardando…");
    const geoAddress = pos ? await reverseGeocode(pos.lat, pos.lng) : "";
    try {
      await updateDoc(doc(db, "timeRecords", currentRecord.id), {
        clockOut: Date(),
        clockOutAddress: geoAddress || currentRecord.clockInAddress || null,
        clockOutLat: pos ? pos.lat : null,
        clockOutLng: pos ? pos.lng : null,
        clockOutAccuracy: pos ? pos.accuracy : null,
      });
      flash("success", `✓ Clocked OUT — ${formatDuration(Date.now() - currentRecord.clockIn)} shift${pos ? " 📍" : locNote}`);
    } catch (e) {
      flash("error", "Save failed: " + (e.message || "Check your connection."));
    }
  }

  async function addEmployee() {
    const name = newEmpName.trim();
    const pin = newEmpPin.trim();
    if (!name || pin.length !== 4 || !/^\d+$/.test(pin)) return flash("error", "Name required + 4-digit PIN.");
    if (employees.some(e => e.pin === pin)) return flash("error", "PIN already in use.");
    await addDoc(collection(db, "employees"), { name, pin });
    setNewEmpName(""); setNewEmpPin("");
  }

  async function removeEmployee(id) {
    if (!confirm("Remove this employee?")) return;
    await deleteDoc(doc(db, "employees", id));
  }

  const activeEmployees = employees.filter(e => records.some(r => r.employeeId === e.id && !r.clockOut));
  const filteredRecords = filterEmp === "all" ? records : records.filter(r => r.employeeId === filterEmp);

  // ── Not logged in ──────────────────────────────────────────
  if (!user) {
    return (
      <div style={{ background: "#0a0f1c", minHeight: "100vh" }}>
        <LoginScreen employees={employees} onLogin={setUser} />
      </div>
    );
  }

  // ── Logged in ──────────────────────────────────────────────
  const isAdmin = user.name === "Admin" || adminAuthed;

  return (
    <div style={{ minHeight: "100vh", background: "#0a0f1c", fontFamily: "'DM Mono', monospace", color: "#e2e8f0", display: "flex", flexDirection: "column" }}>

      {/* Header */}
      <header style={{
        background: "linear-gradient(135deg,#0f172a,#1e293b)",
        borderBottom: "1px solid #1e3a5f",
        padding: "0 20px",
        display: "flex", alignItems: "center", justifyContent: "space-between",
        height: 60, position: "sticky", top: 0, zIndex: 100,
        boxShadow: "0 4px 24px rgba(0,0,0,0.4)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 30, height: 30, background: "linear-gradient(135deg,#22d3a0,#0ea5e9)", borderRadius: 7, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⏱</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 13, letterSpacing: "0.08em", color: "#f1f5f9" }}>APS FIELDCLOCK</div>
            <div style={{ fontSize: 10, color: "#64748b", letterSpacing: "0.1em" }}>{user.name.toUpperCase()}</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ fontSize: 18, fontWeight: 700, color: "#22d3a0", fontVariantNumeric: "tabular-nums" }}>
            {liveTime.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
          </div>
          <button onClick={() => { setUser(null); setAdminAuthed(false); setView("clock"); }} style={{
            background: "transparent", border: "1px solid #1e3a5f", borderRadius: 7,
            color: "#64748b", fontSize: 11, fontFamily: "inherit", padding: "5px 10px", cursor: "pointer", letterSpacing: "0.08em",
          }}>LOG OUT</button>
        </div>
      </header>

      {/* Nav */}
      <nav style={{ display: "flex", gap: 4, padding: "10px 20px 0", background: "#0a0f1c" }}>
        {[["clock", "⏱ Clock"], ["log", "📋 Log"], ...(adminAuthed ? [["reports", "📊 Reportes"]] : []), ["admin", "⚙ Admin"]].map(([id, label]) => (
          <button key={id} onClick={() => setView(id)} style={{
            padding: "7px 16px", borderRadius: "8px 8px 0 0",
            border: "1px solid", borderBottom: "none",
            borderColor: view === id ? "#1e3a5f" : "transparent",
            background: view === id ? "#0f172a" : "transparent",
            color: view === id ? "#22d3a0" : "#64748b",
            cursor: "pointer", fontSize: 12, fontFamily: "inherit",
            letterSpacing: "0.06em", fontWeight: view === id ? 600 : 400,
          }}>{label}</button>
        ))}
      </nav>

      <main style={{ flex: 1, padding: 20, background: "#0f172a", borderTop: "1px solid #1e3a5f" }}>

        {/* ── CLOCK ─────────────────────────────────────── */}
        {view === "clock" && (
          <div style={{ maxWidth: 480, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>

            {/* Active badge */}
            {currentRecord ? (
              (() => {
                const qt = QUICK_TYPES.find(q => q.id === (currentRecord.entryType || "work")) || QUICK_TYPES[0];
                return (
                  <div style={{ ...S.card, border: `1px solid ${qt.border}`, background: qt.bg.replace("0.10","0.05") }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                      <span style={{ width: 8, height: 8, borderRadius: "50%", background: qt.color, display: "inline-block", boxShadow: `0 0 8px ${qt.color}`, animation: "pulse 2s infinite" }} />
                      <span style={{ color: qt.color, fontWeight: 700, fontSize: 13, letterSpacing: "0.08em" }}>{qt.icon} {qt.label.toUpperCase()} — CLOCKED IN</span>
                    </div>
                    <div style={{ fontSize: 12, color: "#94a3b8", display: "flex", flexDirection: "column", gap: 4 }}>
                      <span>Since: <b style={{ color: "#e2e8f0" }}>{formatDate(currentRecord.clockIn)} {formatTime(currentRecord.clockIn)}</b></span>
                      {currentRecord.customer && <span>Customer: <b style={{ color: "#e2e8f0" }}>{currentRecord.customer}</b></span>}
                      {currentRecord.ticket && <span>Ticket: <b style={{ color: "#e2e8f0" }}>#{currentRecord.ticket}</b></span>}
                      {currentRecord.clockInAddress && (
                        <span>Address: <a href={mapsUrl(currentRecord.clockInAddress)} target="_blank" rel="noopener noreferrer"
                          style={{ color: "#0ea5e9" }}>{currentRecord.clockInAddress} →</a></span>
                      )}
                      {currentRecord.note && <span>Note: <b style={{ color: "#e2e8f0" }}>{currentRecord.note}</b></span>}
                      <span>Duration: <b style={{ color: qt.color }}>{formatDuration(liveTime - currentRecord.clockIn)}</b></span>
                    </div>
                  </div>
                );
              })()
            ) : (
              <div style={{ ...S.card, textAlign: "center" }}>
                <div style={{ fontSize: 12, color: "#475569" }}>You are currently <b style={{ color: "#f87171" }}>clocked out</b></div>
              </div>
            )}

            {/* Quick type selector */}
            {!currentRecord && (
              <>
                <div>
                  <label style={S.label}>CLOCK IN TYPE</label>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 8 }}>
                    {QUICK_TYPES.map(qt => (
                      <button key={qt.id} onClick={() => setEntryType(qt.id)} style={{
                        padding: "12px 6px",
                        borderRadius: 10,
                        border: `1px solid ${entryType === qt.id ? qt.border : "#1e3a5f"}`,
                        background: entryType === qt.id ? qt.bg : "transparent",
                        color: entryType === qt.id ? qt.color : "#475569",
                        fontFamily: "inherit",
                        fontSize: 11,
                        fontWeight: entryType === qt.id ? 700 : 400,
                        letterSpacing: "0.06em",
                        cursor: "pointer",
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "center",
                        gap: 5,
                        transition: "all 0.15s",
                      }}>
                        <span style={{ fontSize: 20 }}>{qt.icon}</span>
                        {qt.label.toUpperCase()}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Work-only fields */}
                {entryType === "work" && (
                  <>
                    <div>
                      <label style={S.label}>CUSTOMER NAME <span style={{ color: "#334155" }}>(optional)</span></label>
                      <input value={customer} onChange={e => setCustomer(e.target.value)} placeholder="e.g. Smith Residence" style={S.input} />
                    </div>
                    <div>
                      <label style={S.label}>SERVICE TICKET # <span style={{ color: "#334155" }}>(optional)</span></label>
                      <input value={ticket} onChange={e => setTicket(e.target.value)} placeholder="e.g. TK-1042" style={S.input} />
                    </div>
                  </>
                )}

                {/* Note — always visible */}
                <div>
                  <label style={S.label}>NOTE <span style={{ color: "#334155" }}>(optional)</span></label>
                  <input value={note} onChange={e => setNote(e.target.value)} placeholder={
                    entryType === "lunch" ? "Where are you eating?" :
                    entryType === "break" ? "Any details..." :
                    entryType === "travel" ? "Destino, vehículo, km inicial, etc." :
                    "Additional details…"
                  } style={S.input} />
                </div>

                {/* La ubicación se captura automáticamente por GPS al fichar */}
                <div style={{ fontSize: 11, color: "#475569", display: "flex", alignItems: "center", gap: 6 }}>
                  📍 Tu ubicación se guarda automáticamente al dar Clock In / Out.
                </div>
              </>
            )}

            {/* Status */}
            {statusMsg && (
              <div style={{
                padding: "12px 16px", borderRadius: 10, fontSize: 13, border: "1px solid",
                borderColor: statusMsg.type === "success" ? "rgba(34,211,160,0.4)" : statusMsg.type === "error" ? "rgba(239,68,68,0.4)" : "rgba(14,165,233,0.4)",
                background: statusMsg.type === "success" ? "rgba(34,211,160,0.08)" : statusMsg.type === "error" ? "rgba(239,68,68,0.08)" : "rgba(14,165,233,0.08)",
                color: statusMsg.type === "success" ? "#22d3a0" : statusMsg.type === "error" ? "#f87171" : "#38bdf8",
              }}>{statusMsg.type === "loading" ? "⏳ " : ""}{statusMsg.msg}</div>
            )}

            {/* Buttons */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <button onClick={handleClockIn} disabled={!!currentRecord || statusMsg?.type === "loading"} style={{ ...S.btnGreen, opacity: currentRecord ? 0.4 : 1 }}>▶ CLOCK IN</button>
              <button onClick={handleClockOut} disabled={!currentRecord || statusMsg?.type === "loading"} style={{ ...S.btnRed, opacity: !currentRecord ? 0.4 : 1 }}>■ CLOCK OUT</button>
            </div>

            {/* Other active employees */}
            {activeEmployees.filter(e => e.id !== user.id).length > 0 && (
              <div style={S.card}>
                <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 10 }}>OTHER ACTIVE EMPLOYEES</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {activeEmployees.filter(e => e.id !== user.id).map(e => (
                    <div key={e.id} style={{
                      display: "flex", alignItems: "center", gap: 6,
                      background: "rgba(34,211,160,0.08)", border: "1px solid rgba(34,211,160,0.25)",
                      borderRadius: 6, padding: "4px 10px", fontSize: 12,
                    }}>
                      <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#22d3a0", display: "inline-block" }} />
                      <span style={{ color: "#22d3a0" }}>{e.name}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── LOG ───────────────────────────────────────── */}
        {view === "log" && (
          <div style={{ maxWidth: 860, margin: "0 auto" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
              <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b" }}>TIME RECORDS — {filteredRecords.length} ENTRIES</div>
              {adminAuthed && (
                <select value={filterEmp} onChange={e => setFilterEmp(e.target.value)} style={{
                  background: "#0a0f1c", border: "1px solid #1e3a5f", borderRadius: 8,
                  color: "#e2e8f0", padding: "7px 12px", fontSize: 12, fontFamily: "inherit", outline: "none", cursor: "pointer",
                }}>
                  <option value="all">All Employees</option>
                  {employees.map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
                </select>
              )}
            </div>

            {loadingDb ? (
              <div style={{ color: "#64748b", textAlign: "center", padding: 40 }}>Loading…</div>
            ) : filteredRecords.length === 0 ? (
              <div style={{ color: "#64748b", textAlign: "center", padding: 40 }}>No records yet.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {(adminAuthed ? filteredRecords : filteredRecords.filter(r => r.employeeId === user.id)).map(r => {
                  const qt = QUICK_TYPES.find(q => q.id === (r.entryType || "work")) || QUICK_TYPES[0];
                  return (
                  <div key={r.id} style={{ ...S.card, display: "grid", gridTemplateColumns: "8px 1fr auto", gap: "0 16px", alignItems: "start" }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", marginTop: 5, background: r.clockOut ? "#334155" : qt.color, boxShadow: r.clockOut ? "none" : `0 0 6px ${qt.color}` }} />
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 6 }}>
                        {adminAuthed && <span style={{ fontWeight: 700, color: "#f1f5f9", fontSize: 14 }}>{r.employeeName}</span>}
                        <span style={{ background: qt.bg, border: `1px solid ${qt.border}`, color: qt.color, borderRadius: 5, padding: "2px 8px", fontSize: 11 }}>{qt.icon} {qt.label}</span>
                        {r.ticket && (
                          <span style={{ background: "rgba(14,165,233,0.15)", border: "1px solid rgba(14,165,233,0.3)", color: "#38bdf8", borderRadius: 5, padding: "2px 8px", fontSize: 11 }}>#{r.ticket}</span>
                        )}
                        {!r.clockOut && <span style={{ background: "rgba(34,211,160,0.15)", border: "1px solid rgba(34,211,160,0.3)", color: "#22d3a0", borderRadius: 5, padding: "2px 7px", fontSize: 10 }}>● ACTIVE</span>}
                      </div>
                      <div style={{ fontSize: 12, color: "#64748b", display: "flex", gap: 14, flexWrap: "wrap" }}>
                        <span>📅 {formatDate(r.clockIn)}</span>
                        <span>▶ {formatTime(r.clockIn)}</span>
                        {r.clockOut && <span>■ {formatTime(r.clockOut)}</span>}
                        {r.clockOut
                          ? <span style={{ color: "#94a3b8" }}>⏱ {formatDuration(r.clockOut - r.clockIn)}</span>
                          : <span style={{ color: qt.color }}>⏱ {formatDuration(liveTime - r.clockIn)} live</span>}
                      </div>
                      {r.customer && <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 4 }}>👤 {r.customer}</div>}
                      {r.clockInAddress && (
                        <div style={{ fontSize: 12, marginTop: 4 }}>
                          <a href={mapsUrl(r.clockInAddress)} target="_blank" rel="noopener noreferrer"
                            style={{ color: "#0ea5e9", textDecoration: "none" }}>📍 {r.clockInAddress} →</a>
                        </div>
                      )}
                      {r.note && <div style={{ fontSize: 12, color: "#475569", marginTop: 4 }}>📝 {r.note}</div>}
                    </div>
                    {r.clockOut && (
                      <div style={{ background: "#0f172a", border: "1px solid #1e3a5f", borderRadius: 8, padding: "6px 12px", textAlign: "center", minWidth: 60 }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "#e2e8f0" }}>{formatDuration(r.clockOut - r.clockIn)}</div>
                        <div style={{ fontSize: 9, color: "#475569", letterSpacing: "0.1em" }}>TOTAL</div>
                      </div>
                    )}
                  </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* ── REPORTES ──────────────────────────────────── */}
        {view === "reports" && adminAuthed && (
          <ReportsPanel records={records} employees={employees} />
        )}

        {/* ── ADMIN ─────────────────────────────────────── */}
        {view === "admin" && (
          <div style={{ maxWidth: 480, margin: "0 auto", display: "flex", flexDirection: "column", gap: 20 }}>
            {!adminAuthed ? (
              <div style={S.card}>
                <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 14 }}>ADMIN ACCESS</div>
                <input
                  type="password"
                  maxLength={4}
                  value={adminPinInput}
                  onChange={e => setAdminPinInput(e.target.value)}
                  placeholder="Enter admin PIN"
                  style={S.input}
                />
                <button onClick={() => {
                  if (adminPinInput === ADMIN_PIN) setAdminAuthed(true);
                  else { flash("error", "Wrong admin PIN."); setAdminPinInput(""); }
                }} style={{ ...S.btnGreen, marginTop: 12 }}>UNLOCK</button>
                {statusMsg && <div style={{ marginTop: 10, fontSize: 13, color: "#f87171" }}>{statusMsg.msg}</div>}
              </div>
            ) : (
              <>
                {/* Add employee */}
                <div style={S.card}>
                  <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 14 }}>ADD EMPLOYEE</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    <input value={newEmpName} onChange={e => setNewEmpName(e.target.value)} placeholder="Full name" style={S.input} />
                    <input value={newEmpPin} onChange={e => setNewEmpPin(e.target.value.replace(/\D/g, "").slice(0, 4))} placeholder="4-digit PIN" style={S.input} maxLength={4} />
                    <button onClick={addEmployee} style={S.btnGreen}>+ ADD EMPLOYEE</button>
                  </div>
                  {statusMsg && <div style={{ marginTop: 10, fontSize: 13, color: statusMsg.type === "error" ? "#f87171" : "#22d3a0" }}>{statusMsg.msg}</div>}
                </div>

                {/* Employee list */}
                <div style={S.card}>
                  <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 14 }}>EMPLOYEES ({employees.length})</div>
                  {employees.length === 0 ? (
                    <div style={{ color: "#334155", fontSize: 13, textAlign: "center", padding: "10px 0" }}>No employees yet.</div>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      {employees.map(e => {
                        const isActive = records.some(r => r.employeeId === e.id && !r.clockOut);
                        const empRecs = records.filter(r => r.employeeId === e.id && r.clockOut);
                        const totalMs = empRecs.reduce((s, r) => s + (r.clockOut - r.clockIn), 0);
                        return (
                          <div key={e.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderBottom: "1px solid #1e293b" }}>
                            <div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                <span style={{ width: 7, height: 7, borderRadius: "50%", background: isActive ? "#22d3a0" : "#334155", display: "inline-block", boxShadow: isActive ? "0 0 6px #22d3a0" : "none" }} />
                                <span style={{ color: "#f1f5f9", fontSize: 14 }}>{e.name}</span>
                                <span style={{ color: "#334155", fontSize: 11 }}>PIN: {e.pin}</span>
                              </div>
                              <div style={{ fontSize: 11, color: "#475569", marginTop: 3, paddingLeft: 15 }}>
                                {empRecs.length} shifts · {totalMs > 0 ? formatDuration(totalMs) : "0m"} logged
                              </div>
                            </div>
                            <button onClick={() => removeEmployee(e.id)} style={{
                              background: "transparent", border: "1px solid #2d3748", borderRadius: 6,
                              color: "#64748b", padding: "4px 10px", fontSize: 11, fontFamily: "inherit", cursor: "pointer",
                            }}>Remove</button>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* Summary */}
                <div style={S.card}>
                  <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 14 }}>TODAY'S SUMMARY</div>
                  {(() => {
                    const today = new Date().toDateString();
                    const todayRecs = records.filter(r => new Date(r.clockIn).toDateString() === today);
                    const done = todayRecs.filter(r => r.clockOut);
                    const totalMs = done.reduce((s, r) => s + (r.clockOut - r.clockIn), 0);
                    return (
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, textAlign: "center" }}>
                        {[["Active", activeEmployees.length, "#22d3a0"], ["Shifts", todayRecs.length, "#38bdf8"], ["Hours", done.length ? formatDuration(totalMs) : "0m", "#a78bfa"]].map(([l, v, c]) => (
                          <div key={l}><div style={{ fontSize: 22, fontWeight: 700, color: c }}>{v}</div><div style={{ fontSize: 9, color: "#475569", letterSpacing: "0.1em" }}>{l}</div></div>
                        ))}
                      </div>
                    );
                  })()}
                </div>

                {/* Export */}
                <div style={S.card}>
                  <div style={{ fontSize: 10, letterSpacing: "0.12em", color: "#64748b", marginBottom: 12 }}>EXPORT</div>
                  <button onClick={() => {
                    const blob = new Blob([JSON.stringify(records, null, 2)], { type: "application/json" });
                    const a = document.createElement("a");
                    a.href = URL.createObjectURL(blob); a.download = `fieldclock-${Date.now()}.json`; a.click();
                  }} style={{ background: "transparent", border: "1px solid #1e3a5f", borderRadius: 8, color: "#94a3b8", padding: "8px 16px", fontSize: 12, fontFamily: "inherit", cursor: "pointer" }}>
                    ⬇ Export All Records (JSON)
                  </button>
                  <div style={{ fontSize: 11, color: "#475569", marginTop: 10 }}>
                    Para reportes filtrados (Excel / PDF / CSV) usa la pestaña 📊 Reportes.
                  </div>
                </div>
              </>
            )}
          </div>
        )}
      </main>

      <style>{`
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
        button:active { opacity: 0.7 !important; }
        select option { background: #0a0f1c; }
        input::placeholder { color: #334155; }
        a:hover { text-decoration: underline; }
      `}</style>
    </div>
  );
}
