'use client';

import { useEffect, useState, useRef } from 'react';
import {
  fetchAdminStatus,
  adminUploadData,
  adminUploadScorecard,
  adminDeleteData,
  adminDeleteScorecard,
  adminDownloadData,
  adminDownloadScorecard,
  fetchAdminUsers,
  adminSaveUser,
  adminDeleteUser,
  fetchAdminLocations,
  adminSaveLocation,
  adminDeleteLocation,
} from '@/lib/api';

const FILE_TYPES = ['wbr', 'loyalty', 'catering'];
const GRANS      = ['weekly', 'period', 'quarter'];

// ── small UI atoms ─────────────────────────────────────────────────────────────

function Badge({ color, children }) {
  const bg = { green: ['#dcfce7', '#15803d'], gray: ['#f3f4f6', '#6b7280'] }[color] || ['#f3f4f6', '#6b7280'];
  return (
    <span style={{ background: bg[0], color: bg[1], fontSize: 11, fontWeight: 700, borderRadius: 4, padding: '2px 7px' }}>
      {children}
    </span>
  );
}

// Cell showing one file type's state for a week row: View / Replace / Delete, or Upload.
function FileCell({ present, onView, onDelete, onUpload }) {
  const btn = (c) => ({
    background: c, color: '#fff', border: 'none', borderRadius: 5,
    padding: '2px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap',
  });
  const upBtn = {
    background: 'none', border: '1px solid #7c3aed', color: '#7c3aed',
    borderRadius: 5, padding: '2px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer',
  };

  if (!present) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, justifyContent: 'space-between', minWidth: 160 }}>
        <Badge color="gray">—</Badge>
        <button style={upBtn} onClick={onUpload}>Upload</button>
      </div>
    );
  }
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6, justifyContent: 'space-between', minWidth: 160 }}>
      <Badge color="green">✓</Badge>
      <div style={{ display: 'flex', gap: 5 }}>
        <button style={btn('#2563eb')} onClick={onView}>View</button>
        <button style={upBtn} onClick={onUpload}>Replace</button>
        <button style={btn('#dc2626')} onClick={onDelete}>Delete</button>
      </div>
    </div>
  );
}

// ── Upload modal (inline drawer) ───────────────────────────────────────────────

function UploadModal({ prefillWeek, prefillType, onClose, onDone }) {
  const [weekName, setWeekName] = useState(prefillWeek || '');
  const [fileType, setFileType] = useState(prefillType || 'wbr');
  const [file,     setFile]     = useState(null);
  const [busy,     setBusy]     = useState(false);
  const [msg,      setMsg]      = useState('');
  const fileRef = useRef();

  async function submit(e) {
    e.preventDefault();
    if (!weekName.trim() || !file) return;
    setBusy(true); setMsg('');
    try {
      await adminUploadData(weekName.trim(), fileType, file);
      setMsg('Uploaded! Pushing to GitHub — it will appear on the live site in a few minutes.');
      setTimeout(() => { onDone(); onClose(); }, 1200);
    } catch (err) {
      setMsg('Error: ' + err.message);
      setBusy(false);
    }
  }

  const inp = { background: '#f3f4f6', border: '1px solid #e5e7eb', borderRadius: 8, padding: '8px 12px', fontSize: 13, width: '100%', boxSizing: 'border-box', fontFamily: 'inherit' };
  const lbl = { fontSize: 11, fontWeight: 700, color: '#6b7280', textTransform: 'uppercase', display: 'block', marginBottom: 4 };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: 'white', borderRadius: 16, padding: 28, width: 420, boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Montserrat',sans-serif" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Upload WBR Data File</div>
        <form onSubmit={submit} style={{ display: 'grid', gap: 14 }}>
          <div>
            <label style={lbl}>Week Name</label>
            <input style={inp} value={weekName} onChange={e => setWeekName(e.target.value)} placeholder="e.g. Week of June 15" />
          </div>
          <div>
            <label style={lbl}>File Type</label>
            <select style={inp} value={fileType} onChange={e => setFileType(e.target.value)}>
              {FILE_TYPES.map(t => <option key={t} value={t}>{t.charAt(0).toUpperCase() + t.slice(1)}</option>)}
            </select>
          </div>
          <div>
            <label style={lbl}>Excel File (.xlsx)</label>
            <input ref={fileRef} type="file" accept=".xlsx" style={{ ...inp, padding: '6px 12px' }}
              onChange={e => setFile(e.target.files[0] || null)} />
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button type="submit" disabled={busy || !weekName.trim() || !file}
              style={{ background: '#7c3aed', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              {busy ? 'Uploading…' : 'Upload'}
            </button>
            <button type="button" onClick={onClose}
              style={{ background: '#e5e7eb', color: '#374151', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              Cancel
            </button>
            {msg && <span style={{ fontSize: 13, color: msg.startsWith('Error') ? '#dc2626' : '#15803d' }}>{msg}</span>}
          </div>
        </form>
      </div>
    </div>
  );
}

// ── User modal (add / edit) ───────────────────────────────────────────────────

function UserModal({ prefillEmail, prefillRole, isEdit, onClose, onDone }) {
  const [email,    setEmail]    = useState(prefillEmail || '');
  const [password, setPassword] = useState('');
  const [role,     setRole]     = useState(prefillRole  || 'user');
  const [busy,     setBusy]     = useState(false);
  const [msg,      setMsg]      = useState('');

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setMsg('');
    try {
      await adminSaveUser(email.trim(), password || undefined, role);
      setMsg(isEdit ? 'Updated! Changes live after Vercel redeploys (~40s).' : 'User added! They can log in after Vercel redeploys (~40s).');
      setTimeout(() => { onDone(); onClose(); }, 1400);
    } catch (err) {
      setMsg('Error: ' + err.message);
      setBusy(false);
    }
  }

  const inp = { background: '#f3f4f6', border: '1px solid #e5e7eb', borderRadius: 8, padding: '8px 12px', fontSize: 13, width: '100%', boxSizing: 'border-box', fontFamily: 'inherit' };
  const lbl = { fontSize: 11, fontWeight: 700, color: '#6b7280', textTransform: 'uppercase', display: 'block', marginBottom: 4 };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: 'white', borderRadius: 16, padding: 28, width: 400, boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Montserrat',sans-serif" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 18 }}>{isEdit ? 'Edit User' : 'Add New User'}</div>
        <form onSubmit={submit} style={{ display: 'grid', gap: 14 }}>
          <div>
            <label style={lbl}>Email</label>
            <input style={{ ...inp, background: isEdit ? '#f9fafb' : '#f3f4f6' }}
              value={email} onChange={e => setEmail(e.target.value)}
              placeholder="user@example.com" readOnly={isEdit} />
          </div>
          <div>
            <label style={lbl}>Password {isEdit && <span style={{ fontWeight: 400, textTransform: 'none' }}>(leave blank to keep current)</span>}</label>
            <input style={inp} type="password" value={password}
              onChange={e => setPassword(e.target.value)}
              placeholder={isEdit ? 'Leave blank to keep unchanged' : 'Min 6 characters'} />
          </div>
          <div>
            <label style={lbl}>Role</label>
            <select style={inp} value={role} onChange={e => setRole(e.target.value)}>
              <option value="user">User — can view dashboard</option>
              <option value="admin">Admin — full access + admin panel</option>
            </select>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button type="submit" disabled={busy || (!isEdit && !email.trim()) || (!isEdit && !password)}
              style={{ background: '#7c3aed', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              {busy ? 'Saving…' : isEdit ? 'Update' : 'Add User'}
            </button>
            <button type="button" onClick={onClose}
              style={{ background: '#e5e7eb', color: '#374151', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              Cancel
            </button>
            {msg && <span style={{ fontSize: 12, color: msg.startsWith('Error') ? '#dc2626' : '#15803d' }}>{msg}</span>}
          </div>
        </form>
      </div>
    </div>
  );
}

// ── Add Location modal ────────────────────────────────────────────────────────

function AddLocationModal({ onClose, onDone }) {
  const [name, setName] = useState('');
  const [busy, setBusy] = useState(false);
  const [msg, setMsg]   = useState('');

  async function submit(e) {
    e.preventDefault();
    if (!name.trim()) return;
    setBusy(true); setMsg('');
    try {
      await adminSaveLocation(name.trim(), true);
      setMsg('Location added!');
      setTimeout(() => { onDone(); onClose(); }, 800);
    } catch (err) {
      setMsg('Error: ' + err.message);
      setBusy(false);
    }
  }

  const inp = { background: '#f3f4f6', border: '1px solid #e5e7eb', borderRadius: 8, padding: '8px 12px', fontSize: 13, width: '100%', boxSizing: 'border-box', fontFamily: 'inherit' };
  const lbl = { fontSize: 11, fontWeight: 700, color: '#6b7280', textTransform: 'uppercase', display: 'block', marginBottom: 4 };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: 'white', borderRadius: 16, padding: 28, width: 380, boxShadow: '0 8px 40px rgba(0,0,0,0.18)', fontFamily: "'Montserrat',sans-serif" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Add New Location</div>
        <form onSubmit={submit} style={{ display: 'grid', gap: 14 }}>
          <div>
            <label style={lbl}>Location Name</label>
            <input style={inp} value={name} onChange={e => setName(e.target.value)} placeholder="e.g. Georgetown" autoFocus />
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button type="submit" disabled={busy || !name.trim()}
              style={{ background: '#7c3aed', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              {busy ? 'Adding…' : 'Add Location'}
            </button>
            <button type="button" onClick={onClose}
              style={{ background: '#e5e7eb', color: '#374151', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              Cancel
            </button>
            {msg && <span style={{ fontSize: 13, color: msg.startsWith('Error') ? '#dc2626' : '#15803d' }}>{msg}</span>}
          </div>
        </form>
      </div>
    </div>
  );
}

// ── Main component ─────────────────────────────────────────────────────────────

export default function AdminPanel() {
  const [status,     setStatus]     = useState(null);
  const [loading,    setLoading]    = useState(true);
  const [error,      setError]      = useState('');
  const [section,    setSection]    = useState('data');

  // Scorecard upload
  const [gran,       setGran]       = useState('weekly');
  const [scFile,     setScFile]     = useState(null);
  const [scUploading,setScUploading]= useState(false);
  const [scMsg,      setScMsg]      = useState('');
  const scFileRef = useRef();

  // Upload modal
  const [modal, setModal] = useState(null); // { week, type }

  // Users section
  const [users,      setUsers]      = useState(null);
  const [usersErr,   setUsersErr]   = useState('');
  const [userModal,  setUserModal]  = useState(null); // null | { email?, role?, isEdit }

  // Locations section
  const [locations,    setLocations]    = useState(null);
  const [locsErr,      setLocsErr]      = useState('');
  const [locModal,     setLocModal]     = useState(false);
  const [locBusy,      setLocBusy]      = useState({});

  const load = () => {
    setLoading(true);
    fetchAdminStatus()
      .then(d => { setStatus(d); setLoading(false); })
      .catch(e => { setError(e.message); setLoading(false); });
  };
  useEffect(load, []);

  const loadUsers = () => {
    fetchAdminUsers()
      .then(d => setUsers(d.users))
      .catch(e => setUsersErr(e.message));
  };
  useEffect(() => { if (section === 'users') loadUsers(); }, [section]);

  const loadLocations = () => {
    setLocsErr('');
    fetchAdminLocations()
      .then(d => setLocations(d.locations))
      .catch(e => setLocsErr(e.message));
  };
  useEffect(() => { if (section === 'locations') loadLocations(); }, [section]);

  async function delData(weekName, ft) {
    if (!confirm(`Delete ${ft} file for "${weekName}"? This commits the deletion to GitHub.`)) return;
    try { await adminDeleteData(weekName, ft); load(); }
    catch (err) { alert('Delete failed: ' + err.message); }
  }
  async function viewData(weekName, ft) {
    try { await adminDownloadData(weekName, ft); }
    catch (err) { alert('Download failed: ' + err.message); }
  }
  async function handleScorecardUpload(e) {
    e.preventDefault();
    if (!scFile) return;
    setScUploading(true); setScMsg('');
    try {
      await adminUploadScorecard(gran, scFile);
      setScMsg('Uploaded! Pushing to GitHub — it will appear on the live site in a few minutes.');
      setScFile(null);
      if (scFileRef.current) scFileRef.current.value = '';
      load();
    } catch (err) {
      setScMsg('Error: ' + err.message);
    } finally {
      setScUploading(false);
    }
  }
  async function delScorecard(g, filename) {
    if (!confirm(`Delete "${filename}"? This commits the deletion to GitHub.`)) return;
    try { await adminDeleteScorecard(g, filename); load(); }
    catch (err) { alert('Delete failed: ' + err.message); }
  }
  async function viewScorecard(g, filename) {
    try { await adminDownloadScorecard(g, filename); }
    catch (err) { alert('Download failed: ' + err.message); }
  }

  const card = { background: 'var(--card)', borderRadius: 12, padding: '20px 24px', border: '1px solid var(--border)', marginBottom: 20 };
  const btn  = (c = '#7c3aed') => ({ background: c, color: c === '#e5e7eb' ? '#374151' : '#fff', border: 'none', borderRadius: 8, padding: '8px 18px', fontSize: 13, fontWeight: 700, cursor: 'pointer', fontFamily: "'Montserrat',sans-serif" });

  return (
    <div>
      {modal && (
        <UploadModal
          prefillWeek={modal.week}
          prefillType={modal.type}
          onClose={() => setModal(null)}
          onDone={load}
        />
      )}
      {userModal && (
        <UserModal
          prefillEmail={userModal.email}
          prefillRole={userModal.role}
          isEdit={userModal.isEdit}
          onClose={() => setUserModal(null)}
          onDone={loadUsers}
        />
      )}
      {locModal && (
        <AddLocationModal
          onClose={() => setLocModal(false)}
          onDone={loadLocations}
        />
      )}

      {/* Section tabs */}
      <div style={{ display: 'flex', gap: 10, marginBottom: 20, flexWrap: 'wrap' }}>
        {[['data', 'WBR Data'], ['scorecard', 'Scorecard'], ['users', 'Users'], ['locations', 'Locations']].map(([s, label]) => (
          <button key={s} onClick={() => setSection(s)} style={btn(section === s ? '#7c3aed' : '#e5e7eb')}>
            {label}
          </button>
        ))}
        {section === 'data' && (
          <button onClick={() => setModal({ week: '', type: 'wbr' })} style={{ ...btn('#059669'), marginLeft: 'auto' }}>
            + Upload New Week
          </button>
        )}
        {section === 'users' && (
          <button onClick={() => setUserModal({ isEdit: false })} style={{ ...btn('#059669'), marginLeft: 'auto' }}>
            + Add User
          </button>
        )}
        {section === 'locations' && (
          <button onClick={() => setLocModal(true)} style={{ ...btn('#059669'), marginLeft: 'auto' }}>
            + Add Location
          </button>
        )}
        <button onClick={section === 'users' ? loadUsers : section === 'locations' ? loadLocations : load} style={btn('#6b7280')}>Refresh</button>
      </div>

      {loading && <div style={{ color: 'var(--muted)', padding: 20 }}>Loading…</div>}
      {error   && <div style={{ color: '#dc2626', padding: 10 }}>{error}</div>}

      {/* ── DATA SECTION ──────────────────────────────────────────────────── */}
      {!loading && section === 'data' && (
        <>
          <div style={{ display: 'flex', gap: 16, marginBottom: 14, alignItems: 'center', fontSize: 12, color: 'var(--muted)' }}>
            Every upload / delete here commits directly to the GitHub repo, which redeploys automatically.
          </div>

          {!status?.weeks?.length ? (
            <div style={{ ...card, color: 'var(--muted)' }}>No weeks found.</div>
          ) : (
            <div style={card}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}>
                All Weeks ({status.weeks.length})
              </div>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <thead>
                  <tr style={{ borderBottom: '2px solid var(--border)' }}>
                    {['Week', 'WBR', 'Loyalty', 'Catering'].map(h => (
                      <th key={h} style={{ textAlign: h === 'Week' ? 'left' : 'center', padding: '6px 10px', fontSize: 11, fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {status.weeks.map(w => (
                    <tr key={w.weekName} style={{ borderBottom: '1px solid var(--border)' }}>
                      <td style={{ padding: '10px 10px', fontWeight: 600, verticalAlign: 'middle' }}>{w.weekName}</td>
                      {FILE_TYPES.map(ft => (
                        <td key={ft} style={{ padding: '10px 10px', verticalAlign: 'middle' }}>
                          <FileCell
                            present={w[ft]}
                            onView={() => viewData(w.weekName, ft)}
                            onDelete={() => delData(w.weekName, ft)}
                            onUpload={() => setModal({ week: w.weekName, type: ft })}
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}

      {/* ── SCORECARD SECTION ─────────────────────────────────────────────── */}
      {!loading && section === 'scorecard' && (
        <>
          {/* Upload form */}
          <div style={card}>
            <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 16 }}>Upload Scorecard File</div>
            <form onSubmit={handleScorecardUpload} style={{ display: 'grid', gap: 12 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '160px 1fr', gap: 12 }}>
                <div>
                  <label style={{ fontSize: 11, fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', display: 'block', marginBottom: 4 }}>Granularity</label>
                  <select style={{ background: '#f3f4f6', border: '1px solid var(--border)', borderRadius: 8, padding: '8px 12px', fontSize: 13, width: '100%', fontFamily: 'inherit' }}
                    value={gran} onChange={e => setGran(e.target.value)}>
                    {GRANS.map(g => <option key={g} value={g}>{g.charAt(0).toUpperCase() + g.slice(1)}</option>)}
                  </select>
                </div>
                <div>
                  <label style={{ fontSize: 11, fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', display: 'block', marginBottom: 4 }}>Excel File (.xlsx)</label>
                  <input ref={scFileRef} type="file" accept=".xlsx"
                    style={{ background: '#f3f4f6', border: '1px solid var(--border)', borderRadius: 8, padding: '6px 12px', fontSize: 13, width: '100%', boxSizing: 'border-box', fontFamily: 'inherit' }}
                    onChange={e => setScFile(e.target.files[0] || null)} />
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <button type="submit" disabled={scUploading || !scFile} style={btn()}>
                  {scUploading ? 'Uploading…' : 'Upload'}
                </button>
                {scMsg && <span style={{ fontSize: 13, color: scMsg.startsWith('Error') ? '#dc2626' : '#15803d' }}>{scMsg}</span>}
              </div>
            </form>
          </div>

          {GRANS.map(g => {
            const files = status?.scorecards?.[g] || [];
            return (
              <div key={g} style={card}>
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 12, textTransform: 'capitalize' }}>
                  {g} Scorecards ({files.length})
                </div>
                {!files.length ? (
                  <div style={{ color: 'var(--muted)', fontSize: 13 }}>None found.</div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {files.map(f => (
                      <div key={f.filename} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f9fafb', borderRadius: 8, padding: '8px 14px' }}>
                        <span style={{ fontSize: 13 }}>{f.filename}</span>
                        <div style={{ display: 'flex', gap: 6 }}>
                          <button onClick={() => viewScorecard(g, f.filename)}
                            style={{ background: '#2563eb', color: '#fff', border: 'none', borderRadius: 6, padding: '4px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                            View
                          </button>
                          <button onClick={() => delScorecard(g, f.filename)}
                            style={{ background: '#dc2626', color: '#fff', border: 'none', borderRadius: 6, padding: '4px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                            Delete
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </>
      )}

      {/* ── LOCATIONS SECTION ─────────────────────────────────────────────── */}
      {section === 'locations' && (
        <>
          <div style={{ display: 'flex', gap: 16, marginBottom: 14, alignItems: 'center', fontSize: 12, color: 'var(--muted)' }}>
            Toggle locations open/closed. Closed locations are excluded when "Open Only" is selected on the dashboard.
          </div>

          {locsErr && <div style={{ color: '#dc2626', padding: 10, marginBottom: 10 }}>{locsErr}</div>}

          {!locations ? (
            <div style={{ color: 'var(--muted)', padding: 20 }}>Loading locations…</div>
          ) : (
            <div style={card}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}>
                All Locations ({Object.keys(locations).length})
              </div>
              {!Object.keys(locations).length ? (
                <div style={{ color: 'var(--muted)', fontSize: 13 }}>No locations found. Add one above.</div>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--border)' }}>
                      {['Location', 'Status', 'Actions'].map(h => (
                        <th key={h} style={{ textAlign: 'left', padding: '6px 10px', fontSize: 11, fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {Object.entries(locations).map(([name, cfg]) => (
                      <tr key={name} style={{ borderBottom: '1px solid var(--border)' }}>
                        <td style={{ padding: '10px 10px', fontWeight: 600 }}>{name}</td>
                        <td style={{ padding: '10px 10px' }}>
                          <span style={{
                            background: cfg.open ? '#dcfce7' : '#fee2e2',
                            color: cfg.open ? '#15803d' : '#dc2626',
                            fontSize: 11, fontWeight: 700, borderRadius: 4, padding: '2px 8px',
                          }}>
                            {cfg.open ? 'Open' : 'Closed'}
                          </span>
                        </td>
                        <td style={{ padding: '10px 10px' }}>
                          <div style={{ display: 'flex', gap: 6 }}>
                            <button
                              disabled={!!locBusy[name]}
                              onClick={async () => {
                                setLocBusy(b => ({ ...b, [name]: true }));
                                try {
                                  await adminSaveLocation(name, !cfg.open);
                                  await loadLocations();
                                } catch (err) { alert('Error: ' + err.message); }
                                setLocBusy(b => ({ ...b, [name]: false }));
                              }}
                              style={{ background: cfg.open ? '#dc2626' : '#15803d', color: '#fff', border: 'none', borderRadius: 6, padding: '4px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', opacity: locBusy[name] ? 0.6 : 1 }}>
                              {locBusy[name] ? '…' : cfg.open ? 'Mark Closed' : 'Mark Open'}
                            </button>
                            <button
                              disabled={!!locBusy[name]}
                              onClick={async () => {
                                if (!confirm(`Delete "${name}"? This cannot be undone.`)) return;
                                setLocBusy(b => ({ ...b, [name]: true }));
                                try {
                                  await adminDeleteLocation(name);
                                  await loadLocations();
                                } catch (err) { alert('Error: ' + err.message); }
                                setLocBusy(b => ({ ...b, [name]: false }));
                              }}
                              style={{ background: '#6b7280', color: '#fff', border: 'none', borderRadius: 6, padding: '4px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', opacity: locBusy[name] ? 0.6 : 1 }}>
                              Delete
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </>
      )}

      {/* ── USERS SECTION ─────────────────────────────────────────────────── */}
      {section === 'users' && (
        <>
          <div style={{ display: 'flex', gap: 16, marginBottom: 14, alignItems: 'center', fontSize: 12, color: 'var(--muted)' }}>
            Changes commit to GitHub and take effect after Vercel redeploys (~40s).
          </div>

          {usersErr && <div style={{ color: '#dc2626', padding: 10, marginBottom: 10 }}>{usersErr}</div>}

          {!users ? (
            <div style={{ color: 'var(--muted)', padding: 20 }}>Loading users…</div>
          ) : (
            <div style={card}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}>
                All Users ({users.length})
              </div>
              {!users.length ? (
                <div style={{ color: 'var(--muted)', fontSize: 13 }}>No users found.</div>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--border)' }}>
                      {['Email', 'Role', 'Actions'].map(h => (
                        <th key={h} style={{ textAlign: 'left', padding: '6px 10px', fontSize: 11, fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {users.map(u => (
                      <tr key={u.email} style={{ borderBottom: '1px solid var(--border)' }}>
                        <td style={{ padding: '10px 10px', fontWeight: 500 }}>{u.email}</td>
                        <td style={{ padding: '10px 10px' }}>
                          <span style={{
                            background: u.role === 'admin' ? '#ede9fe' : '#f3f4f6',
                            color: u.role === 'admin' ? '#7c3aed' : '#6b7280',
                            fontSize: 11, fontWeight: 700, borderRadius: 4, padding: '2px 8px',
                          }}>
                            {u.role}
                          </span>
                        </td>
                        <td style={{ padding: '10px 10px' }}>
                          <div style={{ display: 'flex', gap: 6 }}>
                            <button
                              onClick={() => setUserModal({ email: u.email, role: u.role, isEdit: true })}
                              style={{ background: '#2563eb', color: '#fff', border: 'none', borderRadius: 6, padding: '4px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                              Edit
                            </button>
                            <button
                              onClick={async () => {
                                if (!confirm(`Delete ${u.email}? This cannot be undone.`)) return;
                                try { await adminDeleteUser(u.email); loadUsers(); }
                                catch (err) { alert('Error: ' + err.message); }
                              }}
                              style={{ background: '#dc2626', color: '#fff', border: 'none', borderRadius: 6, padding: '4px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                              Delete
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}
