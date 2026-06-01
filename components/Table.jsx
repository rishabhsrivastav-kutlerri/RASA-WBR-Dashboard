'use client';

// Reusable table that mirrors the buildTable helper from the original HTML.
// Props:
//   headers: array of { label, cls? }   (cls === 'right' aligns right)
//   rows:    array of { _cls?, cells }  (_cls === 'total-row' highlights)
export default function Table({ headers, rows }) {
  return (
    <table>
      <thead>
        <tr>
          {headers.map((h, i) => <th key={i} className={h.cls === 'right' ? 'right' : ''}>{h.label}</th>)}
        </tr>
      </thead>
      <tbody>
        {rows.map((r, i) => (
          <tr key={i} className={r._cls || ''}>
            {r.cells.map((c, j) => {
              const cls = headers[j]?.cls === 'right' ? 'right' : '';
              if (typeof c === 'string' && c.includes('<')) {
                return <td key={j} className={cls} dangerouslySetInnerHTML={{ __html: c }} />;
              }
              return <td key={j} className={cls}>{c}</td>;
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
