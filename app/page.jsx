'use client';

import { useEffect, useState } from 'react';
import { useRouter } from 'next/navigation';
import { fetchSheets, fetchWeekData } from '@/lib/api';
import Snapshot from '@/components/Snapshot';
import Sales from '@/components/Sales';
import Costs from '@/components/Costs';
import Reviews from '@/components/Reviews';
import ThirdParty from '@/components/ThirdParty';
import Bikky from '@/components/Bikky';
import Loyalty from '@/components/Loyalty';
import Marketing from '@/components/Marketing';
import CateringSales from '@/components/CateringSales';
import Scorecard from '@/components/Scorecard';

const TABS = [
  { id: 'snapshot',     label: 'Overview' },
  { id: 'sales',        label: 'Revenue & Channels' },
  { id: 'costs',        label: 'Costs' },
  { id: 'reviews',      label: 'Reviews & Ratings' },
  { id: 'thirdparty',   label: '3rd Party Delivery' },
  { id: 'bikky',        label: 'Traffic Insights' },
  { id: 'loyalty',      label: 'Loyalty' },
  { id: 'marketing',    label: 'Marketing' },
  { id: 'cateringsales',label: 'Catering' },
  { id: 'scorecard',    label: 'Leadership Scorecard' },
];

export default function DashboardPage() {
  const router = useRouter();
  const [authChecked, setAuthChecked] = useState(false);
  const [sheets, setSheets] = useState([]);
  const [week, setWeek] = useState('');
  const [data, setData] = useState(null);
  const [prevData, setPrevData] = useState(null);
  const [tab, setTab] = useState('snapshot');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Auth gate: redirect to /login if no token in localStorage.
  useEffect(() => {
    if (typeof window === 'undefined') return;
    if (!localStorage.getItem('wbr_token')) {
      router.replace('/login');
    } else {
      setAuthChecked(true);
    }
  }, [router]);

  useEffect(() => {
    if (!authChecked) return;
    fetchSheets()
      .then(list => {
        setSheets(list);
        if (list.length > 0) {
          // Default to the latest available fiscal week (highest period, then week).
          const rank = s => (s.period == null ? -1 : s.period * 100 + (s.weekInPeriod || 0));
          const latest = list.reduce((a, b) => (rank(b) >= rank(a) ? b : a));
          setWeek(latest.week);
        } else setLoading(false);
      })
      .catch(err => { setError(err.message); setLoading(false); });
  }, [authChecked]);

  useEffect(() => {
    if (!week) return;
    setLoading(true);
    // Find the immediately-previous week (chronologically) to power the
    // "Var to Last Week" comparison; fetch it alongside the current week.
    const rank = s => (s.period == null ? -1 : s.period * 100 + (s.weekInPeriod || 0));
    const cur = sheets.find(s => s.week === week);
    const curRank = cur ? rank(cur) : -1;
    const prev = sheets
      .filter(s => rank(s) < curRank)
      .sort((a, b) => rank(b) - rank(a))[0];
    Promise.all([
      fetchWeekData(week),
      prev ? fetchWeekData(prev.week).catch(() => null) : Promise.resolve(null),
    ])
      .then(([d, pd]) => { setData(d); setPrevData(pd); setLoading(false); setError(''); })
      .catch(err => { setError(err.message); setLoading(false); });
  }, [week, sheets]);

  function handleLogout() {
    localStorage.removeItem('wbr_token');
    router.replace('/login');
  }

  if (!authChecked) {
    return (
      <div className="loading-screen">
        <div className="spinner" />
      </div>
    );
  }

  if (loading && !data) {
    return (
      <div className="loading-screen">
        <div className="spinner" />
        <span>Loading dashboard…</span>
      </div>
    );
  }
  if (error && !data) {
    return (
      <div className="loading-screen">
        <span style={{ color: '#dc2626' }}>{error}</span>
      </div>
    );
  }
  if (!data) {
    return (
      <div className="loading-screen">
        <span>No weeks found in data/ folder.</span>
      </div>
    );
  }

  // Period / Week filters derived from the enriched sheet list. `week` (the
  // selected folder) stays the single source of truth that drives data loading.
  const current = sheets.find(s => s.week === week);
  const currentPeriod = current ? current.period : null;
  const periods = [...new Set(sheets.map(s => s.period))]
    .sort((a, b) => (a == null ? Infinity : a) - (b == null ? Infinity : b));
  const weeksInPeriod = sheets
    .filter(s => s.period === currentPeriod)
    .sort((a, b) => (a.weekInPeriod || 0) - (b.weekInPeriod || 0));
  const periodText = p => (p == null ? 'Other' : `Period ${p}`);
  const weekText = s => (s.weekInPeriod != null ? `Week ${s.weekInPeriod}` : s.label);

  function handlePeriodChange(val) {
    const p = val === 'null' ? null : Number(val);
    const first = sheets
      .filter(s => s.period === p)
      .sort((a, b) => (a.weekInPeriod || 0) - (b.weekInPeriod || 0))[0];
    if (first) setWeek(first.week);
  }

  return (
    <>
      <header className="header">
        <div className="brand">
          <img src="/rasa-logo.png" alt="RASA" className="brand-logo brand-logo-rasa" />
          <div className="brand-title">Weekly Business Review</div>
          {tab !== 'scorecard' && (
            <>
              <select
                className="week-selector"
                value={currentPeriod == null ? 'null' : String(currentPeriod)}
                onChange={e => handlePeriodChange(e.target.value)}
              >
                {periods.map(p => (
                  <option key={String(p)} value={p == null ? 'null' : String(p)}>{periodText(p)}</option>
                ))}
              </select>
              <select className="week-selector" value={week} onChange={e => setWeek(e.target.value)}>
                {weeksInPeriod.map(s => (
                  <option key={s.week} value={s.week}>{weekText(s)}</option>
                ))}
              </select>
            </>
          )}
        </div>
        <div className="brand-right">
          <img src="/kutlerri-logo.png" alt="Kutlerri" className="brand-logo brand-logo-kutlerri" />
          <button className="logout-btn" onClick={handleLogout}>Sign out</button>
        </div>
      </header>

      <nav className="tabs-bar">
        {TABS.map(t => (
          <button
            key={t.id}
            className={`tab-btn${tab === t.id ? ' active' : ''}`}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </nav>

      <main className="main">
        {tab === 'scorecard' ? (
          <Scorecard />
        ) : (
          <>
            {loading && <div style={{ textAlign: 'center', padding: '40px 0', color: 'var(--muted)' }}><div className="spinner" style={{ margin: '0 auto 10px' }} />Loading week…</div>}
            {!loading && tab === 'snapshot'      && <Snapshot data={data} prevData={prevData} />}
            {!loading && tab === 'sales'         && <Sales data={data} prevData={prevData} />}
            {!loading && tab === 'costs'         && <Costs data={data} />}
            {!loading && tab === 'reviews'       && <Reviews data={data} />}
            {!loading && tab === 'thirdparty'    && <ThirdParty data={data} />}
            {!loading && tab === 'bikky'         && <Bikky data={data} />}
            {!loading && tab === 'loyalty'       && <Loyalty data={data} />}
            {!loading && tab === 'marketing'     && <Marketing data={data} />}
            {!loading && tab === 'cateringsales' && <CateringSales data={data} />}
          </>
        )}
      </main>
    </>
  );
}
