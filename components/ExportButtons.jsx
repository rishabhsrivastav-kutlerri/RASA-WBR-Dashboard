'use client';

// "Export" buttons dropped into a table-card/chart-card's header row,
// visible to every user. Deliberately data-agnostic — instead of threading
// each table's headers/rows or each chart's dataset through new props,
// these find the nearest actual <table>/<canvas> DOM node via the click
// event and read it directly. That keeps every call site to a single
// one-line addition with zero changes to how the table/chart itself is built.

const btnStyle = {
  background: '#f3f4f6',
  border: '1px solid #e5e7eb',
  color: '#6b7280',
  padding: '3px 8px',
  borderRadius: 6,
  fontSize: 13,
  lineHeight: 1,
  cursor: 'pointer',
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
};

// Standard "download" glyph (arrow into a tray) — no icon library is
// installed in this project, so it's inlined as SVG rather than adding one.
function DownloadIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </svg>
  );
}

function triggerDownload(href, filename) {
  const a = document.createElement('a');
  a.href = href;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function csvCell(s) {
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

// Expands colspan/rowspan into a full grid (repeating a merged header cell's
// text into every column/row it visually covers) before turning it into CSV
// — a plain per-row textContent scrape misaligns as soon as any header uses
// merged cells (e.g. a grouped-column header spanning several sub-columns),
// since the header row then has fewer cells than the data rows below it.
function tableToGrid(tableEl) {
  const rows = Array.from(tableEl.querySelectorAll('tr'));
  const grid = [];
  const rowSpans = {}; // col -> { text, remaining }
  rows.forEach((tr, r) => {
    grid[r] = [];
    let col = 0;
    const fillCovered = () => {
      while (rowSpans[col] && rowSpans[col].remaining > 0) {
        grid[r][col] = rowSpans[col].text;
        rowSpans[col].remaining--;
        col++;
      }
    };
    for (const cell of tr.querySelectorAll('th,td')) {
      fillCovered();
      const text = (cell.textContent || '').trim();
      const colSpan = parseInt(cell.getAttribute('colspan') || '1', 10) || 1;
      const rowSpan = parseInt(cell.getAttribute('rowspan') || '1', 10) || 1;
      for (let k = 0; k < colSpan; k++) {
        grid[r][col] = text;
        if (rowSpan > 1) rowSpans[col] = { text, remaining: rowSpan - 1 };
        col++;
      }
    }
    fillCovered();
  });
  return grid;
}

function tableToCsv(tableEl) {
  return tableToGrid(tableEl)
    .map(row => row.map(cell => csvCell(cell || '')).join(','))
    .join('\r\n');
}

// Place inside a .table-card (or pass containerClass for a different
// wrapper) — exports the first <table> found in that container.
export function ExportCsvButton({ filename, containerClass = 'table-card' }) {
  const handleClick = e => {
    const container = e.currentTarget.closest('.' + containerClass);
    const table = container?.querySelector('table');
    if (!table) return;
    const blob = new Blob([tableToCsv(table)], { type: 'text/csv;charset=utf-8;' });
    triggerDownload(URL.createObjectURL(blob), filename.endsWith('.csv') ? filename : filename + '.csv');
  };
  return <button type="button" onClick={handleClick} style={btnStyle} title="Export as CSV"><DownloadIcon /></button>;
}

// Place inside a .chart-card (or pass containerClass) — exports the first
// <canvas> found in that container as a PNG (Chart.js renders to canvas, so
// canvas.toDataURL works regardless of chart type, no chart-instance ref
// needed).
export function ExportImageButton({ filename, containerClass = 'chart-card' }) {
  const handleClick = e => {
    const container = e.currentTarget.closest('.' + containerClass);
    const canvas = container?.querySelector('canvas');
    if (!canvas) return;
    triggerDownload(canvas.toDataURL('image/png', 1), filename.endsWith('.png') ? filename : filename + '.png');
  };
  return <button type="button" onClick={handleClick} style={btnStyle} title="Export as image"><DownloadIcon /></button>;
}
