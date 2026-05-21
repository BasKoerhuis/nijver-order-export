// ============================================================================
// Nijver Order Export Converter
// ============================================================================
// Tech:
//   - React 18+ functionele componenten
//   - ExcelJS voor Excel-generatie met styling én image-embedding
//   - papaparse voor CSV-parsing
//   - lucide-react voor iconen
//   - Inline styles (geen CSS framework)
// ============================================================================

import React, { useState, useMemo } from 'react';
import Papa from 'papaparse';
import ExcelJS from 'exceljs';
import {
  Upload, FileCheck2, Download, AlertCircle, X,
  Building2, Users, Euro, RefreshCw, ArrowRight,
} from 'lucide-react';

// ============================================================================
// NIJVER HUISSTIJL
// Kleuren gesampled uit het Nijver-logo
// ============================================================================
const NIJVER = {
  turquoise: '#5DD3C5',
  turquoiseDark: '#3FBFAF',
  turquoiseDeep: '#18A69A',
  turquoiseLight: '#E8FAF7',
  turquoiseBorder: '#A7E8DF',
  navy: '#10142A',
};

const UI = {
  bg: '#F8FAFC',
  card: '#FFFFFF',
  text: NIJVER.navy,
  textBody: '#334155',
  textMuted: '#64748B',
  border: '#E2E8F0',
  borderMuted: '#F1F5F9',
  primary: NIJVER.turquoise,
  primaryDark: NIJVER.turquoiseDark,
  primaryDeep: NIJVER.turquoiseDeep,
  primaryLight: NIJVER.turquoiseLight,
  primaryBorder: NIJVER.turquoiseBorder,
  success: '#10B981',
  successLight: '#ECFDF5',
  successBorder: '#A7F3D0',
  danger: '#DC2626',
  dangerLight: '#FEF2F2',
};

// Google Fonts (Manrope voor UI, JetBrains Mono voor code/getallen)
const FONT_LINK_ID = 'nijver-fonts-link';
if (typeof document !== 'undefined' && !document.getElementById(FONT_LINK_ID)) {
  const link = document.createElement('link');
  link.id = FONT_LINK_ID;
  link.rel = 'stylesheet';
  link.href = 'https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap';
  document.head.appendChild(link);
}
const FONT_UI = "'Manrope', -apple-system, sans-serif";
const FONT_MONO = "'JetBrains Mono', 'SF Mono', monospace";

// ============================================================================
// EXCEL CELSTIJLEN (ExcelJS — ARGB met FF-prefix)
// ============================================================================
const XC = {
  turquoise: 'FF5DD3C5',
  turquoiseDeep: 'FF18A69A',
  turquoiseLight: 'FFE8FAF7',
  navy: 'FF10142A',
  white: 'FFFFFFFF',
  lightGray: 'FFF8FAFC',
  altRow: 'FFF1F5F9',
  borderGray: 'FFE2E8F0',
};
const XFONT = 'Arial';
const NUMFMT_EURO = '€ #,##0.00;[Red]-€ #,##0.00;-';

const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const border = () => {
  const side = { style: 'thin', color: { argb: XC.borderGray } };
  return { top: side, bottom: side, left: side, right: side };
};

const S = {
  banner: {
    fill: fill(XC.white),
    font: { bold: true, size: 20, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 2 },
  },
  bannerRight: {
    fill: fill(XC.white),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 2 },
  },
  subtitle: {
    fill: fill(XC.navy),
    font: { size: 11, color: { argb: XC.white }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 2 },
  },
  infoLabel: {
    fill: fill(XC.altRow),
    font: { bold: true, size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    border: border(),
  },
  infoValue: {
    fill: fill(XC.white),
    font: { size: 11, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    border: border(),
  },
  sectionTitle: {
    fill: fill(XC.navy),
    font: { bold: true, size: 12, color: { argb: XC.turquoise }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 2 },
  },
  tableHeader: {
    fill: fill(XC.navy),
    font: { bold: true, size: 10, color: { argb: XC.white }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 },
    border: border(),
  },
  cellText: {
    fill: fill(XC.white),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    border: border(),
  },
  cellTextAlt: {
    fill: fill(XC.lightGray),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    border: border(),
  },
  cellNum: {
    fill: fill(XC.white),
    font: { bold: true, size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: NUMFMT_EURO,
  },
  cellNumAlt: {
    fill: fill(XC.lightGray),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: NUMFMT_EURO,
  },
  cellNumGray: {
    fill: fill(XC.altRow),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: NUMFMT_EURO,
  },
  cellInt: {
    fill: fill(XC.white),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: '0',
  },
  cellIntAlt: {
    fill: fill(XC.lightGray),
    font: { size: 10, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: '0',
  },
  complexHeader: {
    fill: fill(XC.turquoiseLight),
    font: { bold: true, size: 11, color: { argb: XC.turquoiseDeep }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    border: border(),
  },
  totalText: {
    fill: fill(XC.turquoise),
    font: { bold: true, size: 12, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
    border: border(),
  },
  totalNum: {
    fill: fill(XC.turquoise),
    font: { bold: true, size: 12, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: NUMFMT_EURO,
  },
  totalInt: {
    fill: fill(XC.turquoise),
    font: { bold: true, size: 12, color: { argb: XC.navy }, name: XFONT },
    alignment: { horizontal: 'right', vertical: 'middle', indent: 1 },
    border: border(),
    numFmt: '0',
  },
};

function applyStyle(cell, style) {
  if (style.fill) cell.fill = style.fill;
  if (style.font) cell.font = style.font;
  if (style.alignment) cell.alignment = style.alignment;
  if (style.border) cell.border = style.border;
  if (style.numFmt) cell.numFmt = style.numFmt;
}
function styleRange(ws, r, c0, c1, style) {
  for (let c = c0; c <= c1; c++) applyStyle(ws.getRow(r).getCell(c), style);
}

// ============================================================================
// CSV → CATEGORIE LOGICA
// NL-SfB hoofdstukcodes mapping naar Nijver-categorieën
// ============================================================================
function chapterToCategory(code) {
  if (!code || code === 'null') return 'Overig';
  const c = parseInt(code, 10);
  if (isNaN(c)) return 'Overig';
  if ([27, 37, 46, 47].includes(c)) return 'Dak';
  if ([23, 33].includes(c)) return 'Vloer';
  if ([21, 28, 31, 41].includes(c)) return 'Gevel';
  if (c >= 50 && c <= 69) return 'Installatie';
  if ([22, 24, 32, 34, 42, 43, 44, 45].includes(c) || (c >= 70 && c <= 79)) return 'Interieur';
  return 'Overig';
}
const CATEGORIES = ['Dak', 'Gevel', 'Installatie', 'Interieur', 'Overig', 'Vloer'];

function extractCategories(complexRow) {
  const measures = {};
  for (const [key, value] of Object.entries(complexRow)) {
    if (!key.startsWith('calculation.offer.measures.')) continue;
    const parts = key.split('.');
    const idx = parts[3];
    const field = parts.slice(4).join('.');
    if (!measures[idx]) measures[idx] = {};
    measures[idx][field] = value;
  }
  const cats = {};
  for (const m of Object.values(measures)) {
    const name = m.measureName;
    if (!name || name === '' || name === 'null') continue;
    const cat = chapterToCategory(m.measureNlsfbChapterCode);
    if (!cats[cat]) cats[cat] = { total: 0, inv: 0, maint: 0 };
    cats[cat].total += parseFloat(m.measureSalesPrice) || 0;
    cats[cat].inv += parseFloat(m.measureSalesPriceInvestment) || 0;
    cats[cat].maint += parseFloat(m.measureSalesPriceMaintenance) || 0;
  }
  return cats;
}

const stripCorpSuffix = (s) => s ? String(s).split('|')[0] : '';
const formatDateDutch = (d = new Date()) => {
  const m = ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'];
  return `${d.getDate()} ${m[d.getMonth()]} ${d.getFullYear()}`;
};
const underscoreToSpace = (v) => !v ? '' : String(v).replace(/_/g, ' ');
// Fallback voor ontbrekende JSON-info: lege/null-waarden worden 'Nee'.
const orNee = (v) => {
  if (v == null) return 'Nee';
  const s = String(v).trim();
  if (!s || s.toLowerCase() === 'null') return 'Nee';
  return s.replace(/_/g, ' ');
};
// Boolean-veld ('true'/'false'/null) → 'Ja'/'Nee'
const boolToJaNee = (v) => {
  if (v == null) return 'Nee';
  const s = String(v).trim().toLowerCase();
  return (s === 'true' || s === 'ja' || s === 'yes' || s === '1') ? 'Ja' : 'Nee';
};
// Lees veld uit nieuwe pad (specifications.X) of val terug op oude pad (X)
const itemField = (c, name) =>
  c[`calculation.item.specifications.${name}`] ?? c[`calculation.item.${name}`];

// ============================================================================
// buildWorkbook — genereert beide bladen (Uitgangspunten + Prijs)
// ============================================================================
async function loadLogoBuffer() {
  const res = await fetch('/nijver-logo.jpg');
  if (!res.ok) throw new Error(`Logo laden mislukt (${res.status})`);
  return await res.arrayBuffer();
}

// Logo is 105×39 px; row 1 is 40pt (~53px) and column A is 22 chars (~155px).
// col=0.28 puts logo ~20px further right than the previous 0.15 anchor.
const LOGO_EXT = { width: 105, height: 39 };
const LOGO_TL = { col: 0.28, row: 0.18 };

function setRowValues(ws, rowIdx, values) {
  const row = ws.getRow(rowIdx);
  values.forEach((v, i) => { row.getCell(i + 1).value = v; });
}

async function buildWorkbook({ project, complexes, corporationName }) {
  const orderCode = project.project_id || '';
  const authorName = [project['author.firstName'], project['author.lastName']]
    .filter(Boolean).join(' ').trim();
  const authorEmail = project['author.email'] || '';
  const today = formatDateDutch();
  const corpName = corporationName || project.externalCorporationId || '';

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Nijver Order Export';
  wb.created = new Date();

  const logoBuffer = await loadLogoBuffer();
  const logoId = wb.addImage({ buffer: logoBuffer, extension: 'jpeg' });

  // ============================================
  // BLAD 1: Uitgangspunten
  // ============================================
  const NCOLS_U = 9;
  const wsUit = wb.addWorksheet('Uitgangspunten');
  wsUit.columns = [
    { width: 22 }, { width: 18 }, { width: 22 }, { width: 20 },
    { width: 18 }, { width: 22 }, { width: 22 }, { width: 22 }, { width: 14 },
  ];

  // Row 1: banner (logo embedded, date right-aligned)
  setRowValues(wsUit, 1, ['', '', '', '', '', '', '', '', today]);
  wsUit.getRow(1).height = 40;
  wsUit.mergeCells(1, 1, 1, NCOLS_U - 1);
  styleRange(wsUit, 1, 1, NCOLS_U - 1, S.banner);
  applyStyle(wsUit.getCell(1, NCOLS_U), S.bannerRight);
  wsUit.addImage(logoId, { tl: LOGO_TL, ext: LOGO_EXT });

  // Row 2: subtitle
  setRowValues(wsUit, 2, ['Order Export — Uitgangspunten']);
  wsUit.getRow(2).height = 24;
  wsUit.mergeCells(2, 1, 2, NCOLS_U);
  styleRange(wsUit, 2, 1, NCOLS_U, S.subtitle);

  // Rows 4-7: info block
  const infoRows = [
    ['Corporatie', corpName],
    ['Order code', orderCode],
    ['Order auteur', authorName],
    ['E-mail', authorEmail],
  ];
  infoRows.forEach((vals, i) => {
    const r = 4 + i;
    setRowValues(wsUit, r, vals);
    wsUit.mergeCells(r, 2, r, NCOLS_U);
    applyStyle(wsUit.getCell(r, 1), S.infoLabel);
    styleRange(wsUit, r, 2, NCOLS_U, S.infoValue);
  });

  // Row 9: section title
  setRowValues(wsUit, 9, ['Complexen in deze order']);
  wsUit.mergeCells(9, 1, 9, NCOLS_U);
  styleRange(wsUit, 9, 1, NCOLS_U, S.sectionTitle);

  // Row 10: table header
  setRowValues(wsUit, 10, [
    'Complex Code', 'Systematiek', 'Onderdeel van selectie',
    'Ambitie voor complex', 'Asbest verwacht?', 'Keuzes badkamer',
    'Keuzes keuken', 'Keuzes toilet', 'Aantal VHE',
  ]);
  wsUit.getRow(10).height = 32;
  styleRange(wsUit, 10, 1, NCOLS_U, S.tableHeader);

  // Data rows
  let totalVHE = 0;
  complexes.forEach((c, i) => {
    const r = 11 + i;
    const complexCode = stripCorpSuffix(c.externalComplexId || c['calculation.item.complex']);
    const totalPrice = parseFloat(c['calculation.offer.totalPrice']) || 0;
    const pricePerAsset = parseFloat(c['calculation.offer.totalPricePerAsset']) || 0;
    const vhe = pricePerAsset > 0 ? Math.round(totalPrice / pricePerAsset) : 0;
    totalVHE += vhe;
    const strategyName = c['calculation.targetStrategyName'] || '';
    const systematiek = strategyName.replace(/^Breng\s+naar\s+(de\s+)?/i, '').trim() || strategyName;
    setRowValues(wsUit, r, [
      complexCode, systematiek, 'Ja',
      c['calculation.targetEnergyLabel'] || 'Nee',
      boolToJaNee(c['calculation.item.specifications.asbestos']),
      orNee(itemField(c, 'bathroom')),
      orNee(itemField(c, 'kitchen')),
      orNee(itemField(c, 'toilet')),
      vhe,
    ]);
    const alt = i % 2 === 1;
    styleRange(wsUit, r, 1, NCOLS_U - 1, alt ? S.cellTextAlt : S.cellText);
    applyStyle(wsUit.getCell(r, NCOLS_U), alt ? S.cellIntAlt : S.cellInt);
  });

  // Total row
  const totRowU = 11 + complexes.length;
  setRowValues(wsUit, totRowU, ['Totaal', '', '', '', '', '', '', '', totalVHE]);
  styleRange(wsUit, totRowU, 1, NCOLS_U - 1, S.totalText);
  applyStyle(wsUit.getCell(totRowU, NCOLS_U), S.totalInt);

  wsUit.views = [{ state: 'frozen', xSplit: 0, ySplit: 10 }];

  // ============================================
  // BLAD 2: Prijs
  // ============================================
  const NCOLS_P = 7;
  const wsPrijs = wb.addWorksheet('Prijs');
  wsPrijs.columns = [
    { width: 28 }, { width: 18 }, { width: 24 }, { width: 24 },
    { width: 18 }, { width: 28 }, { width: 28 },
  ];

  // Row 1: banner
  setRowValues(wsPrijs, 1, ['', '', '', '', '', '', today]);
  wsPrijs.getRow(1).height = 40;
  wsPrijs.mergeCells(1, 1, 1, NCOLS_P - 1);
  styleRange(wsPrijs, 1, 1, NCOLS_P - 1, S.banner);
  applyStyle(wsPrijs.getCell(1, NCOLS_P), S.bannerRight);
  wsPrijs.addImage(logoId, { tl: LOGO_TL, ext: LOGO_EXT });

  // Row 2: subtitle
  setRowValues(wsPrijs, 2, ['Order Export — Budget per complex']);
  wsPrijs.getRow(2).height = 24;
  wsPrijs.mergeCells(2, 1, 2, NCOLS_P);
  styleRange(wsPrijs, 2, 1, NCOLS_P, S.subtitle);

  // Rows 4-7: info block
  infoRows.forEach((vals, i) => {
    const r = 4 + i;
    setRowValues(wsPrijs, r, vals);
    wsPrijs.mergeCells(r, 2, r, NCOLS_P);
    applyStyle(wsPrijs.getCell(r, 1), S.infoLabel);
    styleRange(wsPrijs, r, 2, NCOLS_P, S.infoValue);
  });

  // Row 9: section title
  setRowValues(wsPrijs, 9, ['Budget per categorie (NL-SfB)']);
  wsPrijs.mergeCells(9, 1, 9, NCOLS_P);
  styleRange(wsPrijs, 9, 1, NCOLS_P, S.sectionTitle);

  // Row 10: table header
  setRowValues(wsPrijs, 10, [
    'Categorie', 'Totaal budget', 'Investering', 'Onderhoud',
    'Budget / VHE', 'Investering / VHE', 'Onderhoud / VHE',
  ]);
  wsPrijs.getRow(10).height = 36;
  styleRange(wsPrijs, 10, 1, NCOLS_P, S.tableHeader);

  // Per-complex sections
  let cursor = 11;
  for (const c of complexes) {
    const complexCode = stripCorpSuffix(c.externalComplexId || c['calculation.item.complex']);
    const totalPrice = parseFloat(c['calculation.offer.totalPrice']) || 0;
    const pricePerAsset = parseFloat(c['calculation.offer.totalPricePerAsset']) || 0;
    const vhe = pricePerAsset > 0 ? Math.round(totalPrice / pricePerAsset) : 0;

    // Sub-header row spanning all columns
    setRowValues(wsPrijs, cursor, [`Complex ${complexCode}  ·  ${vhe} VHE`]);
    wsPrijs.mergeCells(cursor, 1, cursor, NCOLS_P);
    styleRange(wsPrijs, cursor, 1, NCOLS_P, S.complexHeader);
    cursor++;

    const cats = extractCategories(c);
    for (const cat of CATEGORIES) {
      if (!cats[cat]) continue;
      const { total, inv, maint } = cats[cat];
      setRowValues(wsPrijs, cursor, [
        cat, total, inv, maint,
        vhe > 0 ? total / vhe : 0,
        vhe > 0 ? inv / vhe : 0,
        vhe > 0 ? maint / vhe : 0,
      ]);
      // Per-kolom shading: B (Totaal budget) + E (Budget/VHE) zijn de
      // hoofdsommen → wit. C, D, F, G zijn de breakdown → lichtgrijs.
      applyStyle(wsPrijs.getCell(cursor, 1), S.cellText);     // A Categorie
      applyStyle(wsPrijs.getCell(cursor, 2), S.cellNum);      // B Totaal
      applyStyle(wsPrijs.getCell(cursor, 3), S.cellNumGray);  // C Investering
      applyStyle(wsPrijs.getCell(cursor, 4), S.cellNumGray);  // D Onderhoud
      applyStyle(wsPrijs.getCell(cursor, 5), S.cellNum);      // E Budget/VHE
      applyStyle(wsPrijs.getCell(cursor, 6), S.cellNumGray);  // F Inv/VHE
      applyStyle(wsPrijs.getCell(cursor, 7), S.cellNumGray);  // G Ond/VHE
      cursor++;
    }
  }

  // Grand totals
  let grandTotal = 0, grandInv = 0, grandMaint = 0, grandVHE = 0;
  for (const c of complexes) {
    grandTotal += parseFloat(c['calculation.offer.totalPrice']) || 0;
    grandInv += parseFloat(c['calculation.offer.investmentPrice']) || 0;
    grandMaint += parseFloat(c['calculation.offer.maintenancePrice']) || 0;
    const tp = parseFloat(c['calculation.offer.totalPrice']) || 0;
    const tpa = parseFloat(c['calculation.offer.totalPricePerAsset']) || 0;
    grandVHE += tpa > 0 ? Math.round(tp / tpa) : 0;
  }
  setRowValues(wsPrijs, cursor, [
    'TOTAAL', grandTotal, grandInv, grandMaint,
    grandVHE > 0 ? grandTotal / grandVHE : 0,
    grandVHE > 0 ? grandInv / grandVHE : 0,
    grandVHE > 0 ? grandMaint / grandVHE : 0,
  ]);
  applyStyle(wsPrijs.getCell(cursor, 1), S.totalText);
  styleRange(wsPrijs, cursor, 2, NCOLS_P, S.totalNum);

  wsPrijs.views = [{ state: 'frozen', xSplit: 0, ySplit: 10 }];

  return wb;
}

// ============================================================================
// REACT UI COMPONENTEN
// ============================================================================

function DropZone({ step, label, description, file, onFile, onClear }) {
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = React.useRef(null);
  const hasFile = !!file;
  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(e) => {
        e.preventDefault(); setIsDragging(false);
        const f = e.dataTransfer.files?.[0]; if (f) onFile(f);
      }}
      onClick={() => !hasFile && inputRef.current?.click()}
      style={{
        border: `1.5px ${hasFile ? 'solid' : 'dashed'} ${
          hasFile ? UI.successBorder : isDragging ? UI.primary : UI.border
        }`,
        borderRadius: 12, padding: '18px 20px',
        background: hasFile ? UI.successLight : isDragging ? UI.primaryLight : UI.card,
        cursor: hasFile ? 'default' : 'pointer',
        transition: 'all 0.15s ease',
        display: 'flex', alignItems: 'center', gap: 14,
      }}>
      <div style={{
        width: 40, height: 40, borderRadius: 8,
        background: hasFile ? UI.success : isDragging ? UI.primary : UI.borderMuted,
        color: hasFile || isDragging ? '#fff' : UI.textMuted,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        flexShrink: 0, fontFamily: FONT_MONO, fontWeight: 600, fontSize: 14,
      }}>{hasFile ? <FileCheck2 size={18} /> : step}</div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{
          fontSize: 11, color: hasFile ? '#047857' : UI.textMuted,
          fontWeight: 600, marginBottom: 2,
        }}>{label}</div>
        <div style={{
          fontSize: 14, fontWeight: 600, color: UI.text,
          overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
        }}>{hasFile ? file.name : description}</div>
        {hasFile && (
          <div style={{
            fontSize: 11, color: UI.textMuted, marginTop: 2, fontFamily: FONT_MONO,
          }}>{(file.size / 1024).toFixed(1)} KB · ingelezen</div>
        )}
      </div>
      {hasFile ? (
        <button onClick={(e) => { e.stopPropagation(); onClear(); }}
          style={{
            background: 'transparent', border: 'none', cursor: 'pointer',
            padding: 6, borderRadius: 6, color: UI.textMuted,
            display: 'flex', alignItems: 'center',
          }} title="Verwijderen">
          <X size={16} />
        </button>
      ) : (
        <Upload size={18} color={UI.textMuted} />
      )}
      <input ref={inputRef} type="file" accept=".csv" style={{ display: 'none' }}
        onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])} />
    </div>
  );
}

function StatCard({ icon, label, value }) {
  return (
    <div style={{
      background: UI.card, border: `1px solid ${UI.border}`,
      borderRadius: 10, padding: '14px 16px',
    }}>
      <div style={{
        display: 'flex', alignItems: 'center', gap: 7, marginBottom: 8,
        color: UI.textMuted,
      }}>
        {icon}
        <span style={{ fontSize: 11, fontWeight: 600 }}>{label}</span>
      </div>
      <div style={{
        fontSize: 22, fontWeight: 700, color: UI.text,
        fontFamily: FONT_MONO, lineHeight: 1.1, letterSpacing: '-0.01em',
      }}>{value}</div>
    </div>
  );
}

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true, delimiter: ';', skipEmptyLines: true,
      complete: (res) => resolve(res.data),
      error: reject,
    });
  });
}

// ============================================================================
// HOOFDCOMPONENT
// ============================================================================
export default function App() {
  const [projectFile, setProjectFile] = useState(null);
  const [complexesFile, setComplexesFile] = useState(null);
  const [project, setProject] = useState(null);
  const [complexes, setComplexes] = useState(null);
  const [error, setError] = useState(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [lastGenerated, setLastGenerated] = useState(null);
  const [corpName, setCorpName] = useState('');

  const handleProjectFile = async (file) => {
    setError(null); setLastGenerated(null);
    try {
      const rows = await parseCsvFile(file);
      if (!rows.length) throw new Error('Geen rijen in project.csv');
      setProjectFile(file); setProject(rows[0]);
    } catch (e) { setError(`Fout bij inlezen project.csv: ${e.message}`); }
  };

  const handleComplexesFile = async (file) => {
    setError(null); setLastGenerated(null);
    try {
      const rows = await parseCsvFile(file);
      if (!rows.length) throw new Error('Geen rijen in complexes.csv');
      setComplexesFile(file); setComplexes(rows);
    } catch (e) { setError(`Fout bij inlezen complexes.csv: ${e.message}`); }
  };

  const clearProject = () => { setProjectFile(null); setProject(null); setLastGenerated(null); };
  const clearComplexes = () => { setComplexesFile(null); setComplexes(null); setLastGenerated(null); };
  const resetAll = () => {
    setProjectFile(null); setComplexesFile(null);
    setProject(null); setComplexes(null);
    setCorpName(''); setError(null); setLastGenerated(null);
  };

  const summary = useMemo(() => {
    if (!project || !complexes) return null;
    const totalVHE = complexes.reduce((acc, c) => {
      const tp = parseFloat(c['calculation.offer.totalPrice']) || 0;
      const tpa = parseFloat(c['calculation.offer.totalPricePerAsset']) || 0;
      return acc + (tpa > 0 ? Math.round(tp / tpa) : 0);
    }, 0);
    const totalBudget = complexes.reduce((acc, c) =>
      acc + (parseFloat(c['calculation.offer.totalPrice']) || 0), 0);
    return {
      orderCode: project.project_id,
      authorName: [project['author.firstName'], project['author.lastName']].filter(Boolean).join(' '),
      authorEmail: project['author.email'],
      complexCount: complexes.length,
      totalVHE, totalBudget,
      strategy: complexes[0]?.['calculation.targetStrategyName'] || '—',
      targetLabel: complexes[0]?.['calculation.targetEnergyLabel'] || '—',
      complexes: complexes.map(c => ({
        code: stripCorpSuffix(c.externalComplexId || c['calculation.item.complex']),
        totalPrice: parseFloat(c['calculation.offer.totalPrice']) || 0,
      })),
    };
  }, [project, complexes]);

  const canGenerate = project && complexes;

  const generate = async () => {
    if (!canGenerate) return;
    setIsGenerating(true); setError(null);
    try {
      const wb = await buildWorkbook({ project, complexes, corporationName: corpName });
      const today = new Date();
      const dd = String(today.getDate()).padStart(2, '0');
      const mm = String(today.getMonth() + 1).padStart(2, '0');
      const yy = String(today.getFullYear()).slice(2);
      const author = [project['author.firstName'], project['author.lastName']]
        .filter(Boolean).join('_');
      const filename = `${dd}-${mm}-${yy}_${author || 'Export'}_Nijver_Order_Export.xlsx`;
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setLastGenerated({ filename, when: new Date() });
    } catch (e) {
      console.error(e);
      setError(`Fout bij genereren: ${e.message}`);
    } finally {
      setIsGenerating(false);
    }
  };

  const fmtEuro = (n) => new Intl.NumberFormat('nl-NL',
    { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(n);

  return (
    <div style={{
      minHeight: '100vh', background: UI.bg,
      fontFamily: FONT_UI, color: UI.text,
      padding: '28px 24px 60px',
      backgroundImage: `radial-gradient(ellipse 800px 400px at 50% -100px, ${UI.primaryLight} 0%, transparent 70%)`,
    }}>
      <div style={{ maxWidth: 880, margin: '0 auto' }}>

        {/* Top bar met logo en reset-knop */}
        <div style={{
          display: 'flex', alignItems: 'center',
          justifyContent: 'space-between', marginBottom: 40,
        }}>
          <img src="/nijver-logo.jpg" alt="Nijver"
            style={{ height: 36, display: 'block' }} />
          {(projectFile || complexesFile) && (
            <button onClick={resetAll}
              style={{
                background: 'transparent', border: `1px solid ${UI.border}`,
                color: UI.textMuted, padding: '7px 13px', borderRadius: 8,
                cursor: 'pointer', fontSize: 12, fontFamily: FONT_UI,
                display: 'flex', alignItems: 'center', gap: 6, fontWeight: 500,
              }}>
              <RefreshCw size={13} /> Opnieuw beginnen
            </button>
          )}
        </div>

        {/* Hero */}
        <div style={{ marginBottom: 32 }}>
          <div style={{
            display: 'inline-flex', alignItems: 'center', gap: 6,
            background: UI.primaryLight, color: UI.primaryDeep,
            fontSize: 11, fontWeight: 700, textTransform: 'uppercase',
            letterSpacing: '0.06em', padding: '4px 10px', borderRadius: 999,
            border: `1px solid ${UI.primaryBorder}`, marginBottom: 14,
          }}>Order Export</div>
          <h1 style={{
            fontSize: 34, fontWeight: 800, color: UI.text,
            margin: 0, lineHeight: 1.1, letterSpacing: '-0.025em',
          }}>Van CSV naar opgemaakt Excel-bestand</h1>
          <p style={{
            fontSize: 15, color: UI.textBody, margin: '10px 0 0',
            maxWidth: 620, lineHeight: 1.55,
          }}>
            Upload <span style={{
              fontFamily: FONT_MONO, fontSize: 13,
              background: UI.borderMuted, padding: '1px 6px', borderRadius: 4,
            }}>project.csv</span> en <span style={{
              fontFamily: FONT_MONO, fontSize: 13,
              background: UI.borderMuted, padding: '1px 6px', borderRadius: 4,
            }}>complexes.csv</span>. De tool aggregeert de measures per complex
            op basis van NL-SfB-codes en levert een Order Export met volledige
            Nijver-huisstijl.
          </p>
        </div>

        {/* Upload zones */}
        <div style={{ display: 'grid', gap: 10, marginBottom: 20 }}>
          <DropZone step="1" label="Stap 1 — Projectbestand"
            description="Sleep project.csv hierheen of klik om te kiezen"
            file={projectFile} onFile={handleProjectFile} onClear={clearProject} />
          <DropZone step="2" label="Stap 2 — Complexenbestand"
            description="Sleep complexes.csv hierheen of klik om te kiezen"
            file={complexesFile} onFile={handleComplexesFile} onClear={clearComplexes} />
        </div>

        {/* Error */}
        {error && (
          <div style={{
            background: UI.dangerLight, border: `1px solid ${UI.danger}`,
            borderRadius: 10, padding: '12px 16px', marginBottom: 20,
            display: 'flex', gap: 10, alignItems: 'flex-start',
            fontSize: 13, color: UI.danger,
          }}>
            <AlertCircle size={18} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>{error}</span>
          </div>
        )}

        {/* Preview card */}
        {summary && (
          <>
            <div style={{
              background: UI.card, border: `1px solid ${UI.border}`,
              borderRadius: 14, padding: '20px 22px', marginBottom: 14,
            }}>
              <div style={{
                display: 'flex', alignItems: 'center', gap: 8,
                marginBottom: 14, paddingBottom: 14,
                borderBottom: `1px solid ${UI.borderMuted}`,
              }}>
                <div style={{
                  width: 6, height: 6, borderRadius: '50%',
                  background: UI.primary, flexShrink: 0,
                }} />
                <div style={{ flex: 1 }}>
                  <div style={{
                    fontSize: 11, color: UI.textMuted,
                    fontWeight: 600, marginBottom: 3,
                  }}>Preview van de export</div>
                  <div style={{
                    fontSize: 17, fontWeight: 700,
                    fontFamily: FONT_MONO, letterSpacing: '-0.01em',
                  }}>{summary.orderCode}</div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <div style={{ fontSize: 11, color: UI.textMuted, marginBottom: 2 }}>
                    {summary.authorName}
                  </div>
                  <div style={{ fontSize: 11, color: UI.textMuted, fontFamily: FONT_MONO }}>
                    {summary.authorEmail}
                  </div>
                </div>
              </div>

              <div style={{
                display: 'flex', alignItems: 'center', gap: 10,
                background: UI.primaryLight, borderRadius: 8,
                padding: '10px 14px', marginBottom: 16,
                border: `1px solid ${UI.primaryBorder}`,
              }}>
                <ArrowRight size={14} color={UI.primaryDeep} />
                <div style={{
                  flex: 1, fontSize: 13,
                  color: UI.primaryDeep, fontWeight: 500,
                }}>{summary.strategy}</div>
                <div style={{
                  fontSize: 11, fontFamily: FONT_MONO,
                  fontWeight: 600, color: UI.primaryDeep, background: '#fff',
                  padding: '2px 8px', borderRadius: 4,
                  border: `1px solid ${UI.primaryBorder}`,
                }}>{summary.targetLabel}</div>
              </div>

              <div style={{
                display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)',
                gap: 10, marginBottom: 16,
              }}>
                <StatCard icon={<Building2 size={13} />}
                  label="Complexen" value={summary.complexCount} />
                <StatCard icon={<Users size={13} />}
                  label="Totaal VHE" value={summary.totalVHE} />
                <StatCard icon={<Euro size={13} />}
                  label="Totaal budget" value={fmtEuro(summary.totalBudget)} />
              </div>

              <div style={{
                background: UI.bg, borderRadius: 10, padding: '12px 14px',
                fontSize: 12, border: `1px solid ${UI.borderMuted}`,
              }}>
                <div style={{
                  fontSize: 11, color: UI.textMuted,
                  fontWeight: 600, marginBottom: 8,
                }}>Complexen in export</div>
                {summary.complexes.map((c, i) => (
                  <div key={i} style={{
                    display: 'flex', justifyContent: 'space-between', padding: '5px 0',
                    borderBottom: i < summary.complexes.length - 1
                      ? `1px solid ${UI.borderMuted}` : 'none',
                  }}>
                    <span style={{
                      fontFamily: FONT_MONO, color: UI.text, fontWeight: 500,
                    }}>{c.code}</span>
                    <span style={{ fontFamily: FONT_MONO, color: UI.textBody }}>
                      {fmtEuro(c.totalPrice)}
                    </span>
                  </div>
                ))}
              </div>
            </div>

            <div style={{
              background: UI.card, border: `1px solid ${UI.border}`,
              borderRadius: 14, padding: '16px 22px', marginBottom: 20,
            }}>
              <label style={{
                fontSize: 11, color: UI.text, fontWeight: 600,
                display: 'block', marginBottom: 8,
              }}>
                Corporatienaam
                <span style={{
                  color: UI.textMuted, marginLeft: 6, fontWeight: 400,
                }}>(niet aanwezig in CSV — vul handmatig in)</span>
              </label>
              <input type="text" value={corpName}
                onChange={(e) => setCorpName(e.target.value)}
                placeholder={`Bijv. Elkien Holding BV  ·  fallback: ${project.externalCorporationId}`}
                style={{
                  width: '100%', padding: '10px 14px', fontSize: 14,
                  border: `1px solid ${UI.border}`, borderRadius: 8,
                  fontFamily: FONT_UI, background: UI.bg,
                  color: UI.text, outline: 'none',
                  boxSizing: 'border-box', transition: 'border-color 0.15s',
                }}
                onFocus={(e) => e.target.style.borderColor = UI.primary}
                onBlur={(e) => e.target.style.borderColor = UI.border} />
            </div>
          </>
        )}

        {/* Generate-button / success-state */}
        {lastGenerated ? (
          <div style={{
            background: UI.text, borderRadius: 14, padding: '18px 22px',
            color: '#fff', display: 'flex', alignItems: 'center',
            justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
          }}>
            <div style={{ minWidth: 0, flex: 1 }}>
              <div style={{
                fontSize: 11, color: UI.primary, fontWeight: 700,
                textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 3,
              }}>Klaar</div>
              <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>
                Excel gegenereerd en gedownload
              </div>
              <div style={{
                fontSize: 12, color: '#94A3B8',
                fontFamily: FONT_MONO, wordBreak: 'break-all',
              }}>{lastGenerated.filename}</div>
            </div>
            <button onClick={generate}
              style={{
                background: UI.primary, border: 'none', color: UI.text,
                padding: '10px 18px', borderRadius: 8, cursor: 'pointer',
                fontSize: 13, fontWeight: 700, fontFamily: FONT_UI,
                display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0,
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = UI.primaryDark}
              onMouseLeave={(e) => e.currentTarget.style.background = UI.primary}>
              <Download size={15} /> Opnieuw downloaden
            </button>
          </div>
        ) : (
          <button disabled={!canGenerate || isGenerating} onClick={generate}
            style={{
              width: '100%', padding: '16px', fontSize: 15, fontWeight: 700,
              background: canGenerate ? UI.primary : UI.border,
              color: canGenerate ? UI.text : UI.textMuted,
              border: 'none', borderRadius: 12,
              cursor: canGenerate ? 'pointer' : 'not-allowed',
              fontFamily: FONT_UI, letterSpacing: '-0.005em',
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
              transition: 'background 0.15s ease',
              boxShadow: canGenerate ? `0 1px 2px rgba(93, 211, 197, 0.3)` : 'none',
            }}
            onMouseEnter={(e) => {
              if (canGenerate) e.currentTarget.style.background = UI.primaryDark;
            }}
            onMouseLeave={(e) => {
              if (canGenerate) e.currentTarget.style.background = UI.primary;
            }}>
            <Download size={17} />
            {isGenerating ? 'Bezig met genereren…' :
             canGenerate ? 'Genereer & download Excel' :
             'Upload beide bestanden om verder te gaan'}
          </button>
        )}
      </div>
    </div>
  );
}
