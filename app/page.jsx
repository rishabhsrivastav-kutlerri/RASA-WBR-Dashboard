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

const TABS = [
  { id: 'snapshot',     label: 'Overview' },
  { id: 'sales',        label: 'Sales & Revenue' },
  { id: 'costs',        label: 'Costs' },
  { id: 'reviews',      label: 'Reviews & Ratings' },
  { id: 'thirdparty',   label: '3rd Party Delivery' },
  { id: 'bikky',        label: 'Customer Insights' },
  { id: 'loyalty',      label: 'Loyalty' },
  { id: 'marketing',    label: 'Marketing' },
  { id: 'cateringsales',label: 'Catering' },
];

export default function DashboardPage() {
  const router = useRouter();
  const [authChecked, setAuthChecked] = useState(false);
  const [sheets, setSheets] = useState([]);
  const [week, setWeek] = useState('');
  const [data, setData] = useState(null);
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
        if (list.length > 0) setWeek(list[list.length - 1].week);
        else setLoading(false);
      })
      .catch(err => { setError(err.message); setLoading(false); });
  }, [authChecked]);

  useEffect(() => {
    if (!week) return;
    setLoading(true);
    fetchWeekData(week)
      .then(d => { setData(d); setLoading(false); setError(''); })
      .catch(err => { setError(err.message); setLoading(false); });
  }, [week]);

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

  return (
    <>
      <header className="header">
        <div className="brand">
          <img src="/rasa-logo.png" alt="RASA" className="brand-logo brand-logo-rasa" />
          <div className="brand-title">Weekly Business Review</div>
          <select className="week-selector" value={week} onChange={e => setWeek(e.target.value)}>
            {sheets.map(s => (
              <option key={s.week} value={s.week}>{s.label}</option>
            ))}
          </select>
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
        {loading && <div style={{ textAlign: 'center', padding: '40px 0', color: 'var(--muted)' }}><div className="spinner" style={{ margin: '0 auto 10px' }} />Loading week…</div>}
        {!loading && tab === 'snapshot'      && <Snapshot data={data} />}
        {!loading && tab === 'sales'         && <Sales data={data} />}
        {!loading && tab === 'costs'         && <Costs data={data} />}
        {!loading && tab === 'reviews'       && <Reviews data={data} />}
        {!loading && tab === 'thirdparty'    && <ThirdParty data={data} />}
        {!loading && tab === 'bikky'         && <Bikky data={data} />}
        {!loading && tab === 'loyalty'       && <Loyalty data={data} />}
        {!loading && tab === 'marketing'     && <Marketing data={data} />}
        {!loading && tab === 'cateringsales' && <CateringSales data={data} />}
      </main>
    </>
  );
}
