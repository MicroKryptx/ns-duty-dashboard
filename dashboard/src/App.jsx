import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  Activity,
  AlertCircle,
  ArrowLeft,
  ArrowUpDown,
  Calendar,
  CalendarRange,
  CheckCircle,
  Cloud,
  CloudOff,
  Database,
  Download,
  FileSpreadsheet,
  FileText,
  Moon,
  RefreshCw,
  Search,
  Shield,
  SlidersHorizontal,
  Sun,
} from 'lucide-react';

function getStoredTheme() {
  try {
    return localStorage.getItem('duty-dashboard-theme') || 'dark';
  } catch {
    return 'dark';
  }
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  try {
    localStorage.setItem('duty-dashboard-theme', theme);
  } catch {
    // Ignore private-mode storage failures.
  }
}

function getStoredRecentNames() {
  try {
    return JSON.parse(localStorage.getItem('duty-dashboard-recent-names') || '[]');
  } catch {
    return [];
  }
}

function saveRecentName(name) {
  const normalized = name.trim().toUpperCase();
  const next = [normalized, ...getStoredRecentNames().filter((item) => item !== normalized)].slice(0, 6);
  try {
    localStorage.setItem('duty-dashboard-recent-names', JSON.stringify(next));
  } catch {
    // Ignore storage failures.
  }
  return next;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function formatRefreshTime(isoString) {
  if (!isoString) return 'Never';
  const d = new Date(isoString);
  return `${d.toLocaleDateString()} at ${d.toLocaleTimeString()}`;
}

function getSourceLabel(source) {
  if (source === 'google_sheets') return 'Google Sheets';
  if (source === 'cached') return 'Cached workbook';
  if (source === 'local_fallback') return 'Local fallback';
  return 'Preparing data';
}

function getSourceIcon(source) {
  if (source === 'google_sheets') return <Cloud size={12} style={{ color: 'var(--success)' }} />;
  if (source === 'cached') return <Database size={12} style={{ color: 'var(--warning)' }} />;
  if (source === 'local_fallback') return <CloudOff size={12} style={{ color: 'var(--danger)' }} />;
  return <Database size={12} style={{ color: 'var(--text-muted)' }} />;
}

function csvValue(value) {
  const text = String(value ?? '');
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function downloadDutiesCsv(rows, name) {
  const headers = ['#', 'Date', 'Day', 'Day Category', 'Duty Type', 'Entry Details', 'Sheet', 'Row'];
  const body = rows.map((row, index) => [
    index + 1,
    row.date_label,
    row.day,
    row.day_type,
    row.duty_type,
    row.duty_display,
    row.sheet,
    row.row,
  ]);
  const csv = [headers, ...body].map((line) => line.map(csvValue).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${name.replace(/\s+/g, '_').toLowerCase()}_duties.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function DataFreshnessBadge({ dataStatus }) {
  if (!dataStatus) return null;
  const isOffline = dataStatus.source && dataStatus.source !== 'google_sheets';

  return (
    <div className={`freshness-badge ${isOffline ? 'offline' : ''}`}>
      {getSourceIcon(dataStatus.source)}
      <span>
        {getSourceLabel(dataStatus.source)}
        {dataStatus.last_refreshed && <> - {formatRefreshTime(dataStatus.last_refreshed)}</>}
      </span>
    </div>
  );
}

function DataStatusNotice({ dataStatus }) {
  if (!dataStatus) return null;

  const messages = [];
  if (dataStatus.warming) messages.push('Preparing the duty index. This usually happens once after a cold start.');
  if (dataStatus.refreshing && dataStatus.ready) messages.push('Refreshing Google Sheets in the background. Current data remains usable.');
  if (dataStatus.is_stale && dataStatus.ready) messages.push('Using cached data while the service refreshes.');
  if (dataStatus.error) messages.push(`Last refresh issue: ${dataStatus.error}`);
  if (messages.length === 0) return null;

  return (
    <div className={`status-notice ${dataStatus.error ? 'error' : ''}`}>
      <AlertCircle size={15} />
      <span>{messages[0]}</span>
    </div>
  );
}

function Toast({ toast, onDismiss }) {
  useEffect(() => {
    if (!toast) return undefined;
    const timer = setTimeout(onDismiss, 5000);
    return () => clearTimeout(timer);
  }, [toast, onDismiss]);

  if (!toast) return null;

  return (
    <div className={`toast ${toast.type}`}>
      {toast.type === 'success' ? (
        <CheckCircle size={18} style={{ color: 'var(--success)', flexShrink: 0 }} />
      ) : (
        <AlertCircle size={18} style={{ color: 'var(--danger)', flexShrink: 0 }} />
      )}
      <span>{toast.message}</span>
    </div>
  );
}

function ThemeToggle({ theme, onToggle }) {
  return (
    <button className="theme-toggle" onClick={onToggle} title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
      {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
      <span>{theme === 'dark' ? 'Light' : 'Dark'}</span>
    </button>
  );
}

function NameSelection({
  names,
  months,
  loading,
  error,
  onRetry,
  onSelect,
  theme,
  onToggleTheme,
  dataStatus,
  recentNames,
}) {
  const [selectedName, setSelectedName] = useState('');
  const [startMonth, setStartMonth] = useState('');
  const [endMonth, setEndMonth] = useState('');

  useEffect(() => {
    if (months.length > 0 && !startMonth) {
      setStartMonth(months[0].key);
      setEndMonth(months[months.length - 1].key);
    }
  }, [months, startMonth]);

  const startIdx = months.findIndex((m) => m.key === startMonth);
  const endIdx = months.findIndex((m) => m.key === endMonth);
  const isRangeInvalid = startIdx >= 0 && endIdx >= 0 && startIdx > endIdx;

  const readyToProceed = selectedName && startMonth && endMonth && !isRangeInvalid;

  return (
    <div className="container">
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '20px' }}>
        <ThemeToggle theme={theme} onToggle={onToggleTheme} />
      </div>

      <div className="selection-screen">
        <div className="selection-card">
          <Shield size={48} style={{ color: 'var(--primary)', marginBottom: '20px' }} />
          <h2>FOE Duty Dashboard</h2>
          <DataFreshnessBadge dataStatus={dataStatus} />
          <DataStatusNotice dataStatus={dataStatus} />
          <p>Select a personnel and date range to load their consolidated duty report</p>

          {loading ? (
            <div style={{ padding: '40px 0' }}>
              <RefreshCw size={36} style={{ color: 'var(--primary)', animation: 'spin 1s linear infinite' }} />
              <p style={{ marginTop: '16px', color: 'var(--text-muted)', fontSize: '0.9rem' }}>
                Preparing cached duty data for fast lookups...
              </p>
            </div>
          ) : error ? (
            <div style={{ padding: '20px 0' }}>
              <AlertCircle size={36} style={{ color: 'var(--danger)', marginBottom: '12px' }} />
              <p style={{ color: 'var(--danger)', marginBottom: '16px', fontSize: '0.9rem' }}>{error}</p>
              <button className="selection-go-btn" onClick={onRetry}>
                Retry
              </button>
            </div>
          ) : (
            <>
              {recentNames.length > 0 && (
                <div className="recent-row">
                  {recentNames.map((name) => (
                    <button key={name} type="button" className="recent-pill" onClick={() => setSelectedName(name)}>
                      {name}
                    </button>
                  ))}
                </div>
              )}

              <div className="name-select-wrapper">
                <label className="selection-label">
                  <FileText size={14} style={{ color: 'var(--primary)' }} />
                  Personnel Name
                </label>
                <select
                  className="filter-select"
                  value={selectedName}
                  onChange={(e) => setSelectedName(e.target.value)}
                  style={{ width: '100%', padding: '14px 18px', fontSize: '1.05rem' }}
                >
                  <option value="">Choose a name ({names.length} found)</option>
                  {names.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>

              <div className="date-range-row">
                <div className="date-range-field">
                  <label className="selection-label">
                    <CalendarRange size={14} style={{ color: 'var(--success)' }} />
                    Start Month
                  </label>
                  <select
                    className="filter-select"
                    value={startMonth}
                    onChange={(e) => setStartMonth(e.target.value)}
                    style={{ width: '100%', padding: '12px 14px' }}
                  >
                    {months.map((month) => (
                      <option key={month.key} value={month.key}>{month.label}</option>
                    ))}
                  </select>
                </div>

                <div className="date-range-separator">to</div>

                <div className="date-range-field">
                  <label className="selection-label">
                    <CalendarRange size={14} style={{ color: 'var(--danger)' }} />
                    End Month
                  </label>
                  <select
                    className="filter-select"
                    value={endMonth}
                    onChange={(e) => setEndMonth(e.target.value)}
                    style={{ width: '100%', padding: '12px 14px' }}
                  >
                    {months.map((month) => (
                      <option key={month.key} value={month.key}>{month.label}</option>
                    ))}
                  </select>
                </div>
              </div>

              {isRangeInvalid && (
                <div className="validation-error">
                  <AlertCircle size={16} />
                  <span>Start month cannot be after end month.</span>
                </div>
              )}

              <button
                className="selection-go-btn"
                disabled={!readyToProceed}
                onClick={() => onSelect(selectedName, startMonth, endMonth)}
              >
                {selectedName ? `Load Dashboard for ${selectedName}` : 'Select a person to continue'}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function App() {
  const [theme, setTheme] = useState(getStoredTheme);
  const [screen, setScreen] = useState('select');
  const [names, setNames] = useState([]);
  const [months, setMonths] = useState([]);
  const [dataStatus, setDataStatus] = useState(null);
  const [initLoading, setInitLoading] = useState(true);
  const [initError, setInitError] = useState(null);
  const [activeName, setActiveName] = useState('');
  const [activeStart, setActiveStart] = useState('');
  const [activeEnd, setActiveEnd] = useState('');
  const [urlSelectionApplied, setUrlSelectionApplied] = useState(false);
  const [recentNames, setRecentNames] = useState(getStoredRecentNames);

  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [toast, setToast] = useState(null);

  const [searchTerm, setSearchTerm] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [dayTypeFilter, setDayTypeFilter] = useState('ALL');
  const [dutyTypeFilter, setDutyTypeFilter] = useState('ALL');
  const [sortOrder, setSortOrder] = useState('asc');

  useEffect(() => { applyTheme(theme); }, [theme]);
  const toggleTheme = () => setTheme((value) => (value === 'dark' ? 'light' : 'dark'));

  useEffect(() => {
    let cancelled = false;
    let retryTimer = null;

    const loadBootstrap = async () => {
      try {
        const res = await fetch('/api/bootstrap');
        const json = await res.json();
        if (cancelled) return;

        setDataStatus(json.data_status || null);
        setNames(json.names || []);
        setMonths(json.months || []);
        setInitError(null);
        setInitLoading(!json.ready);

        if (!json.ready) {
          retryTimer = setTimeout(loadBootstrap, 1500);
        }
      } catch (err) {
        if (!cancelled) {
          setInitError(err.message);
          setInitLoading(false);
        }
      }
    };

    loadBootstrap();
    return () => {
      cancelled = true;
      if (retryTimer) clearTimeout(retryTimer);
    };
  }, []);

  const fetchDuties = useCallback(async (name, start, end, options = {}) => {
    if (!options.silent) setLoading(true);
    try {
      const res = await fetch('/api/duties', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, start, end }),
      });

      const json = await res.json();
      if (json.data_status) setDataStatus(json.data_status);

      if (!res.ok || json.error || json.valid === false) {
        setToast({ type: 'error', message: json.error || 'Could not load duties.' });
        return false;
      }

      setData(json);
      setToast({ type: 'success', message: `Loaded ${json.duties?.length || 0} duties for ${name}` });
      return true;
    } catch (err) {
      setToast({ type: 'error', message: err.message });
      return false;
    } finally {
      if (!options.silent) setLoading(false);
    }
  }, []);

  const handleSelectName = useCallback((name, start, end) => {
    const normalizedName = name.trim().toUpperCase();
    setActiveName(normalizedName);
    setActiveStart(start);
    setActiveEnd(end);
    setScreen('dashboard');
    setSearchTerm('');
    setMonthFilter('ALL');
    setDayTypeFilter('ALL');
    setDutyTypeFilter('ALL');
    setSortOrder('asc');
    setRecentNames(saveRecentName(normalizedName));

    const params = new URLSearchParams({ name: normalizedName, start, end });
    window.history.replaceState(null, '', `${window.location.pathname}?${params.toString()}`);
    fetchDuties(normalizedName, start, end);
  }, [fetchDuties]);

  useEffect(() => {
    if (urlSelectionApplied || initLoading || names.length === 0 || months.length === 0) return;
    setUrlSelectionApplied(true);

    const params = new URLSearchParams(window.location.search);
    const name = (params.get('name') || '').trim().toUpperCase();
    const start = params.get('start') || '';
    const end = params.get('end') || '';
    if (!name || !start || !end) return;
    if (!names.includes(name)) return;
    if (!months.some((month) => month.key === start) || !months.some((month) => month.key === end)) return;

    handleSelectName(name, start, end);
  }, [handleSelectName, initLoading, months, names, urlSelectionApplied]);

  const handleRefresh = async () => {
    if (!activeName || refreshing) return;
    setRefreshing(true);
    try {
      const refreshRes = await fetch('/api/refresh', { method: 'POST' });
      const refreshJson = await refreshRes.json();
      if (refreshJson.data_status) setDataStatus(refreshJson.data_status);
      setToast({ type: 'success', message: refreshJson.message || 'Refresh started in the background.' });

      for (let attempt = 0; attempt < 40; attempt += 1) {
        await delay(1500);
        const statusRes = await fetch('/api/status');
        const statusJson = await statusRes.json();
        setDataStatus(statusJson);
        if (!statusJson.refreshing) break;
      }

      await fetchDuties(activeName, activeStart, activeEnd, { silent: true });
    } catch (err) {
      setToast({ type: 'error', message: `Refresh failed: ${err.message}` });
    } finally {
      setRefreshing(false);
    }
  };

  const handleBack = () => {
    setScreen('select');
    setData(null);
    setActiveName('');
    window.history.replaceState(null, '', window.location.pathname);
  };

  const duties = data?.duties || [];

  const uniqueMonths = useMemo(() => {
    const parseKey = (key) => {
      const [monthName, yearText] = key.split(' ');
      const monthIndex = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'].indexOf(monthName);
      return Number(yearText) * 12 + monthIndex;
    };
    return Array.from(new Set(duties.map((duty) => duty.month_key))).sort((a, b) => parseKey(a) - parseKey(b));
  }, [duties]);

  const monthLabelsMap = useMemo(() => (
    duties.reduce((acc, duty) => {
      acc[duty.month_key] = duty.month_label;
      return acc;
    }, {})
  ), [duties]);

  const filteredDuties = useMemo(() => {
    const search = searchTerm.toLowerCase();
    return duties
      .filter((duty) => {
        const matchesSearch =
          duty.duty_display.toLowerCase().includes(search) ||
          duty.sheet.toLowerCase().includes(search) ||
          duty.day.toLowerCase().includes(search);
        const matchesMonth = monthFilter === 'ALL' || duty.month_key === monthFilter;
        const matchesDayType = dayTypeFilter === 'ALL' || duty.day_type === dayTypeFilter;
        const matchesDutyType = dutyTypeFilter === 'ALL' || duty.duty_type === dutyTypeFilter;
        return matchesSearch && matchesMonth && matchesDayType && matchesDutyType;
      })
      .sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
      });
  }, [dayTypeFilter, duties, dutyTypeFilter, monthFilter, searchTerm, sortOrder]);

  const liveStats = useMemo(() => {
    const total = filteredDuties.length;
    const guard = filteredDuties.filter((duty) => duty.duty_type === 'Guard').length;
    const bds = filteredDuties.filter((duty) => duty.duty_type === 'BDS').length;
    const weekdays = filteredDuties.filter((duty) => duty.day_type === 'Weekday').length;
    const fridays = filteredDuties.filter((duty) => duty.day_type === 'Friday').length;
    const weekends = filteredDuties.filter((duty) => duty.day_type === 'Weekend').length;
    return { total, guard, bds, weekdays, fridays, weekends };
  }, [filteredDuties]);

  if (screen === 'select') {
    return (
      <>
        <NameSelection
          names={names}
          months={months}
          loading={initLoading}
          error={initError}
          onRetry={() => window.location.reload()}
          onSelect={handleSelectName}
          theme={theme}
          onToggleTheme={toggleTheme}
          dataStatus={dataStatus}
          recentNames={recentNames.filter((name) => names.includes(name))}
        />
        <Toast toast={toast} onDismiss={() => setToast(null)} />
      </>
    );
  }

  if (loading && !data) {
    return (
      <div className="container">
        <div className="empty-state" style={{ minHeight: '60vh', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', gap: '20px' }}>
          <RefreshCw size={48} style={{ color: 'var(--primary)', animation: 'spin 1s linear infinite' }} />
          <p>Generating duty report for <strong>{activeName}</strong>...</p>
          <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Reading from the cached duty index</p>
        </div>
      </div>
    );
  }

  if (!data) {
    return (
      <div className="container" style={{ padding: '60px 20px', maxWidth: '680px' }}>
        <div className="stat-card" style={{ padding: '40px', textAlign: 'center', border: '1px solid rgba(239, 68, 68, 0.2)' }}>
          <FileSpreadsheet size={64} style={{ color: 'var(--danger)', margin: '0 auto 20px' }} />
          <h2 style={{ fontSize: '1.8rem', fontWeight: 800, marginBottom: '16px' }}>No Duty Data</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '24px', fontSize: '0.95rem' }}>
            Could not load duties for {activeName}. Please try another person or date range.
          </p>
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
            <button className="back-btn" onClick={handleBack}>
              <ArrowLeft size={14} /> Back
            </button>
            <button className="refresh-btn" onClick={() => fetchDuties(activeName, activeStart, activeEnd)}>
              <RefreshCw size={14} /> Retry
            </button>
          </div>
        </div>
        <Toast toast={toast} onDismiss={() => setToast(null)} />
      </div>
    );
  }

  const { metadata } = data;

  return (
    <div className="container">
      <header>
        <div className="header-title-row">
          <div>
            <h1>Consolidated Duty Dashboard</h1>
            <p>Interactive tracking system for military duties</p>

            <div className="meta-badges">
              <div className="meta-badge">
                <FileText size={14} style={{ color: 'var(--primary)' }} />
                <span>Personnel: <strong>{metadata.name}</strong></span>
              </div>
              {metadata.gen && (
                <div className="meta-badge">
                  <Shield size={14} style={{ color: 'var(--warning)' }} />
                  <span>Gen: <strong>{metadata.gen}</strong></span>
                </div>
              )}
              <div className="meta-badge">
                <Calendar size={14} style={{ color: 'var(--success)' }} />
                <span>Range: <strong>{metadata.range}</strong></span>
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '10px' }}>
            <div className="header-actions">
              <button className="back-btn" onClick={handleBack}>
                <ArrowLeft size={14} /> Change Person
              </button>
              <button className="back-btn" onClick={() => downloadDutiesCsv(filteredDuties, metadata.name)} disabled={filteredDuties.length === 0}>
                <Download size={14} /> Export CSV
              </button>
              <button className="refresh-btn" onClick={handleRefresh} disabled={refreshing}>
                <RefreshCw size={14} className={refreshing ? 'spin' : ''} />
                {refreshing ? 'Refreshing...' : 'Refresh Data'}
              </button>
              <ThemeToggle theme={theme} onToggle={toggleTheme} />
            </div>
            <div style={{ textAlign: 'right' }}>
              <DataFreshnessBadge dataStatus={dataStatus} />
              <DataStatusNotice dataStatus={dataStatus} />
              <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', marginTop: '4px' }}>
                Report generated: {new Date(metadata.generated_at).toLocaleTimeString()}
              </div>
            </div>
          </div>
        </div>
      </header>

      <section className="stats-grid">
        <div className="stat-card primary-card">
          <span className="stat-label">Total Duties</span>
          <div className="stat-value">{liveStats.total}</div>
          <span className="stat-desc">{liveStats.guard} Guard - {liveStats.bds} BDS</span>
        </div>

        <div className="stat-card bds-card">
          <span className="stat-label">BDS Duties</span>
          <div className="stat-value" style={{ color: 'var(--stat-bds-color)' }}>{liveStats.bds}</div>
          <span className="stat-desc">
            {liveStats.total > 0 ? ((liveStats.bds / liveStats.total) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>

        <div className="stat-card weekday-card">
          <span className="stat-label">Weekdays (Mon-Thu)</span>
          <div className="stat-value" style={{ color: 'var(--stat-weekday-color)' }}>{liveStats.weekdays}</div>
          <span className="stat-desc">
            {liveStats.total > 0 ? ((liveStats.weekdays / liveStats.total) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>

        <div className="stat-card friday-card">
          <span className="stat-label">Fridays</span>
          <div className="stat-value" style={{ color: 'var(--stat-friday-color)' }}>{liveStats.fridays}</div>
          <span className="stat-desc">
            {liveStats.total > 0 ? ((liveStats.fridays / liveStats.total) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>

        <div className="stat-card weekend-card">
          <span className="stat-label">Weekends (Sat-Sun)</span>
          <div className="stat-value" style={{ color: 'var(--stat-weekend-color)' }}>{liveStats.weekends}</div>
          <span className="stat-desc">
            {liveStats.total > 0 ? ((liveStats.weekends / liveStats.total) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>
      </section>

      <section className="filter-panel">
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '18px', borderBottom: '1px solid var(--filter-divider)', paddingBottom: '10px' }}>
          <SlidersHorizontal size={16} style={{ color: 'var(--primary)' }} />
          <h3 style={{ fontSize: '0.95rem', fontWeight: 700, letterSpacing: '0.02em', textTransform: 'uppercase' }}>Interactive Filter Panel</h3>
        </div>

        <div className="filter-grid">
          <div className="filter-group">
            <label htmlFor="search">Search Keywords</label>
            <div style={{ position: 'relative' }}>
              <input
                id="search"
                type="text"
                placeholder="Search entry, day..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="search-input"
                style={{ width: '100%', paddingLeft: '36px' }}
              />
              <Search size={16} style={{ position: 'absolute', left: '12px', top: '13px', color: 'var(--text-muted)' }} />
            </div>
          </div>

          <div className="filter-group">
            <label htmlFor="month">Filter by Month</label>
            <select id="month" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} className="filter-select">
              <option value="ALL">All Months ({uniqueMonths.length})</option>
              {uniqueMonths.map((monthKey) => (
                <option key={monthKey} value={monthKey}>{monthLabelsMap[monthKey]}</option>
              ))}
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="dayType">Type of Day</label>
            <select id="dayType" value={dayTypeFilter} onChange={(e) => setDayTypeFilter(e.target.value)} className="filter-select">
              <option value="ALL">All Day Types</option>
              <option value="Weekday">Weekday (Mon-Thu)</option>
              <option value="Friday">Friday</option>
              <option value="Weekend">Weekend (Sat-Sun)</option>
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="dutyType">Duty Type</label>
            <select id="dutyType" value={dutyTypeFilter} onChange={(e) => setDutyTypeFilter(e.target.value)} className="filter-select">
              <option value="ALL">All Duties</option>
              <option value="Guard">Guard Duty (G)</option>
              <option value="BDS">BDS Duty</option>
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="sort">Sort Date</label>
            <button
              id="sort"
              onClick={() => setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')}
              className="filter-select"
              style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', textAlign: 'left' }}
            >
              <span>{sortOrder === 'asc' ? 'Oldest First' : 'Newest First'}</span>
              <ArrowUpDown size={14} style={{ color: 'var(--text-muted)' }} />
            </button>
          </div>
        </div>
      </section>

      <section className="data-container">
        {filteredDuties.length === 0 ? (
          <div className="empty-state">
            <Activity size={48} style={{ color: 'var(--text-muted)', marginBottom: '16px', opacity: 0.5 }} />
            <p style={{ fontWeight: 600, color: 'var(--text)' }}>No Matching Records Found</p>
            <p style={{ fontSize: '0.85rem' }}>Try clearing or relaxing your active search and filters</p>
          </div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Date</th>
                  <th>Day</th>
                  <th>Day Category</th>
                  <th>Duty Type</th>
                  <th>Entry Details</th>
                  <th>Source Reference</th>
                </tr>
              </thead>
              <tbody>
                {filteredDuties.map((duty, index) => (
                  <tr key={`${duty.date}-${duty.duty_type}-${duty.number}`}>
                    <td style={{ fontWeight: 600, color: 'var(--text-muted)' }}>{index + 1}</td>
                    <td style={{ fontWeight: 500 }}>{duty.date_label}</td>
                    <td>{duty.day}</td>
                    <td>
                      <span className={`badge badge-${duty.day_type.toLowerCase()}`}>
                        {duty.day_type === 'Weekday' ? 'Weekday (Mon-Thu)' : duty.day_type === 'Friday' ? 'Friday' : 'Weekend (Sat-Sun)'}
                      </span>
                    </td>
                    <td>
                      <span className={`badge badge-${duty.duty_type.toLowerCase()}`}>
                        {duty.duty_type}
                      </span>
                    </td>
                    <td style={{ fontWeight: 600 }}>{duty.duty_display}</td>
                    <td style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                      Sheet: {duty.sheet} (Row: {duty.row})
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      <footer>
        <p>2026 FOE Duty Checker dashboard - React + Vite + Flask</p>
      </footer>

      <Toast toast={toast} onDismiss={() => setToast(null)} />
    </div>
  );
}

export default App;
