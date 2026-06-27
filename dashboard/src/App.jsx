import React, { useState, useEffect, useCallback } from 'react';
import {
  Calendar,
  Shield,
  Activity,
  Search,
  RefreshCw,
  FileText,
  SlidersHorizontal,
  FileSpreadsheet,
  ArrowUpDown,
  Sun,
  Moon,
  ArrowLeft,
  CheckCircle,
  AlertCircle,
  CalendarRange,
  Cloud,
  CloudOff,
  Database,
} from 'lucide-react';

// ── Helpers ──────────────────────────────────────────────────────────────────

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
  } catch { /* ignore */ }
}

function formatRefreshTime(isoString) {
  if (!isoString) return 'Never';
  const d = new Date(isoString);
  return d.toLocaleDateString() + ' at ' + d.toLocaleTimeString();
}

function getSourceLabel(source) {
  if (source === 'google_sheets') return 'Google Sheets (Live)';
  if (source === 'cached') return 'Cached (Offline)';
  if (source === 'local_fallback') return 'Local File (Offline)';
  return 'Unknown';
}

function getSourceIcon(source) {
  if (source === 'google_sheets') return <Cloud size={12} style={{ color: 'var(--success)' }} />;
  if (source === 'cached') return <Database size={12} style={{ color: 'var(--warning)' }} />;
  if (source === 'local_fallback') return <CloudOff size={12} style={{ color: 'var(--danger)' }} />;
  return <Database size={12} style={{ color: 'var(--text-muted)' }} />;
}

// ── Data Freshness Badge ─────────────────────────────────────────────────────

function DataFreshnessBadge({ dataStatus }) {
  if (!dataStatus) return null;

  const isOffline = dataStatus.source && dataStatus.source !== 'google_sheets';

  return (
    <div className={`freshness-badge ${isOffline ? 'offline' : ''}`}>
      {getSourceIcon(dataStatus.source)}
      <span>
        {getSourceLabel(dataStatus.source)}
        {dataStatus.last_refreshed && (
          <> · {formatRefreshTime(dataStatus.last_refreshed)}</>
        )}
      </span>
    </div>
  );
}

// ── Toast Component ──────────────────────────────────────────────────────────

function Toast({ toast, onDismiss }) {
  useEffect(() => {
    if (!toast) return;
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

// ── Theme Toggle Button ──────────────────────────────────────────────────────

function ThemeToggle({ theme, onToggle }) {
  return (
    <button className="theme-toggle" onClick={onToggle} title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
      {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
      <span>{theme === 'dark' ? 'Light' : 'Dark'}</span>
    </button>
  );
}

// ── Name Selection Screen ────────────────────────────────────────────────────

function NameSelection({ names, months, loading, error, onSelect, theme, onToggleTheme, dataStatus }) {
  const [selectedName, setSelectedName] = useState('');
  const [startMonth, setStartMonth] = useState('');
  const [endMonth, setEndMonth] = useState('');
  const [validating, setValidating] = useState(false);
  const [validationError, setValidationError] = useState('');

  // Set default start/end once months load
  useEffect(() => {
    if (months.length > 0 && !startMonth) {
      setStartMonth(months[0].key);
      setEndMonth(months[months.length - 1].key);
    }
  }, [months, startMonth]);

  // Clear validation error when inputs change
  useEffect(() => {
    setValidationError('');
  }, [selectedName, startMonth, endMonth]);

  // Validate date range ordering
  const startIdx = months.findIndex((m) => m.key === startMonth);
  const endIdx = months.findIndex((m) => m.key === endMonth);
  const isRangeInvalid = startIdx >= 0 && endIdx >= 0 && startIdx > endIdx;

  const handleProceed = async () => {
    if (!selectedName || !startMonth || !endMonth) return;
    if (isRangeInvalid) {
      setValidationError('Start month cannot be after end month.');
      return;
    }

    setValidating(true);
    setValidationError('');

    try {
      const res = await fetch('/api/validate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: selectedName,
          start: startMonth,
          end: endMonth,
        }),
      });

      const json = await res.json();

      if (json.error) {
        setValidationError(json.error);
        setValidating(false);
        return;
      }

      if (!json.valid) {
        setValidationError(json.error || `${selectedName} was not found in the selected date range.`);
        setValidating(false);
        return;
      }

      // Validation passed — proceed
      onSelect(selectedName, startMonth, endMonth, json.details);
    } catch (err) {
      setValidationError(`Connection error: ${err.message}`);
    } finally {
      setValidating(false);
    }
  };

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
          <p>Select a personnel and date range to load their consolidated duty report</p>

          {loading ? (
            <div style={{ padding: '40px 0' }}>
              <RefreshCw size={36} style={{ color: 'var(--primary)', animation: 'spin 1s linear infinite' }} />
              <p style={{ marginTop: '16px', color: 'var(--text-muted)', fontSize: '0.9rem' }}>
                Loading personnel and months from Excel workbook...
              </p>
            </div>
          ) : error ? (
            <div style={{ padding: '20px 0' }}>
              <AlertCircle size={36} style={{ color: 'var(--danger)', marginBottom: '12px' }} />
              <p style={{ color: 'var(--danger)', marginBottom: '16px', fontSize: '0.9rem' }}>{error}</p>
              <button className="selection-go-btn" onClick={() => window.location.reload()}>
                Retry
              </button>
            </div>
          ) : (
            <>
              {/* Name Selector */}
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
                  <option value="">— Choose a name ({names.length} found) —</option>
                  {names.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>

              {/* Date Range Selectors */}
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
                    {months.map((m) => (
                      <option key={m.key} value={m.key}>{m.label}</option>
                    ))}
                  </select>
                </div>

                <div className="date-range-separator">→</div>

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
                    {months.map((m) => (
                      <option key={m.key} value={m.key}>{m.label}</option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Validation Error */}
              {(validationError || isRangeInvalid) && (
                <div className="validation-error">
                  <AlertCircle size={16} />
                  <span>{validationError || 'Start month cannot be after end month.'}</span>
                </div>
              )}

              {/* Submit Button */}
              <button
                className="selection-go-btn"
                disabled={!selectedName || !startMonth || !endMonth || isRangeInvalid || validating}
                onClick={handleProceed}
              >
                {validating ? (
                  <>
                    <RefreshCw size={16} style={{ animation: 'spin 1s linear infinite', marginRight: '8px', display: 'inline' }} />
                    Validating...
                  </>
                ) : selectedName ? (
                  `Load Dashboard for ${selectedName}`
                ) : (
                  'Select a person to continue'
                )}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Main App ─────────────────────────────────────────────────────────────────

function App() {
  // Theme state
  const [theme, setTheme] = useState(getStoredTheme);

  // App-level states
  const [screen, setScreen] = useState('select');
  const [names, setNames] = useState([]);
  const [months, setMonths] = useState([]);
  const [dataStatus, setDataStatus] = useState(null);
  const [initLoading, setInitLoading] = useState(true);
  const [initError, setInitError] = useState(null);
  const [activeName, setActiveName] = useState('');
  const [activeStart, setActiveStart] = useState('');
  const [activeEnd, setActiveEnd] = useState('');

  // Dashboard data
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState(null);

  // Filters
  const [searchTerm, setSearchTerm] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [dayTypeFilter, setDayTypeFilter] = useState('ALL');
  const [dutyTypeFilter, setDutyTypeFilter] = useState('ALL');
  const [sortOrder, setSortOrder] = useState('asc');

  // Apply theme on mount and changes
  useEffect(() => { applyTheme(theme); }, [theme]);
  const toggleTheme = () => setTheme((t) => (t === 'dark' ? 'light' : 'dark'));

  // ── Fetch names + months on mount ──
  useEffect(() => {
    Promise.all([
      fetch('/api/names').then((r) => {
        if (!r.ok) throw new Error('Failed to load names');
        return r.json();
      }),
      fetch('/api/months').then((r) => {
        if (!r.ok) throw new Error('Failed to load months');
        return r.json();
      }),
    ])
      .then(([namesData, monthsData]) => {
        setNames(namesData.names || []);
        setMonths(monthsData.months || []);
        if (namesData.data_status) setDataStatus(namesData.data_status);
        setInitLoading(false);
      })
      .catch((err) => {
        setInitError(err.message);
        setInitLoading(false);
      });
  }, []);

  // ── Fetch duties for a person ──
  const fetchDuties = useCallback(async (name, start, end) => {
    setLoading(true);
    try {
      const res = await fetch('/api/duties', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, start, end }),
      });

      const json = await res.json();

      if (json.error) {
        setToast({ type: 'error', message: json.error });
        setLoading(false);
        return;
      }

      setData(json);
      if (json.data_status) setDataStatus(json.data_status);
      setToast({ type: 'success', message: `Loaded ${json.duties?.length || 0} duties for ${name}` });
    } catch (err) {
      setToast({ type: 'error', message: err.message });
    } finally {
      setLoading(false);
    }
  }, []);

  // ── Handle name selection (after validation passed) ──
  const handleSelectName = (name, start, end, _validationDetails) => {
    setActiveName(name);
    setActiveStart(start);
    setActiveEnd(end);
    setScreen('dashboard');
    // Reset filters
    setSearchTerm('');
    setMonthFilter('ALL');
    setDayTypeFilter('ALL');
    setDutyTypeFilter('ALL');
    setSortOrder('asc');
    fetchDuties(name, start, end);
  };

  // ── Handle refresh (re-download from Google Sheets, then re-fetch duties) ──
  const handleRefresh = async () => {
    if (!activeName) return;
    setLoading(true);
    try {
      // Step 1: Force re-download from Google Sheets
      const refreshRes = await fetch('/api/refresh', { method: 'POST' });
      const refreshJson = await refreshRes.json();

      if (refreshJson.data_status) setDataStatus(refreshJson.data_status);

      if (refreshJson.status === 'error') {
        setToast({ type: 'error', message: refreshJson.message });
        setLoading(false);
        return;
      }

      const isOffline = refreshJson.status?.startsWith('fallback');
      if (isOffline) {
        setToast({ type: 'error', message: refreshJson.message });
      } else {
        setToast({ type: 'success', message: 'Refreshed data from Google Sheets' });
      }

      // Step 2: Re-fetch duties with fresh data
      await fetchDuties(activeName, activeStart, activeEnd);
    } catch (err) {
      setToast({ type: 'error', message: `Refresh failed: ${err.message}` });
      setLoading(false);
    }
  };

  // ── Handle back ──
  const handleBack = () => {
    setScreen('select');
    setData(null);
    setActiveName('');
  };

  // ── Render selection screen ──
  if (screen === 'select') {
    return (
      <>
        <NameSelection
          names={names}
          months={months}
          loading={initLoading}
          error={initError}
          onSelect={handleSelectName}
          theme={theme}
          onToggleTheme={toggleTheme}
          dataStatus={dataStatus}
        />
        <Toast toast={toast} onDismiss={() => setToast(null)} />
      </>
    );
  }

  // ── Dashboard loading ──
  if (loading && !data) {
    return (
      <div className="container">
        <div className="empty-state" style={{ minHeight: '60vh', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', gap: '20px' }}>
          <RefreshCw size={48} style={{ color: 'var(--primary)', animation: 'spin 1s linear infinite' }} />
          <p>Generating duty report for <strong>{activeName}</strong>...</p>
          <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Scanning Excel workbook across all monthly sheets</p>
        </div>
      </div>
    );
  }

  // ── No data fallback ──
  if (!data) {
    return (
      <div className="container" style={{ padding: '60px 20px', maxWidth: '680px' }}>
        <div className="stat-card" style={{ padding: '40px', textAlign: 'center', border: '1px solid rgba(239, 68, 68, 0.2)' }}>
          <FileSpreadsheet size={64} style={{ color: 'var(--danger)', margin: '0 auto 20px' }} />
          <h2 style={{ fontSize: '1.8rem', fontWeight: 800, marginBottom: '16px' }}>No Duty Data</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '24px', fontSize: '0.95rem' }}>
            Could not load duties for {activeName}. Please try again.
          </p>
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
            <button className="back-btn" onClick={handleBack}>
              <ArrowLeft size={14} /> Back
            </button>
            <button className="refresh-btn" onClick={handleRefresh}>
              <RefreshCw size={14} /> Retry
            </button>
          </div>
        </div>
        <Toast toast={toast} onDismiss={() => setToast(null)} />
      </div>
    );
  }

  // ── Dashboard with data ──
  const { metadata, duties } = data;

  const uniqueMonths = Array.from(new Set(duties.map((d) => d.month_key))).sort((a, b) => {
    const parseKey = (k) => {
      const parts = k.split(' ');
      const mIdx = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'].indexOf(parts[0]);
      const yr = parseInt(parts[1]);
      return yr * 12 + mIdx;
    };
    return parseKey(a) - parseKey(b);
  });

  const monthLabelsMap = duties.reduce((acc, d) => {
    acc[d.month_key] = d.month_label;
    return acc;
  }, {});

  let filteredDuties = duties.filter((d) => {
    const matchesSearch =
      d.duty_display.toLowerCase().includes(searchTerm.toLowerCase()) ||
      d.sheet.toLowerCase().includes(searchTerm.toLowerCase()) ||
      d.day.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesMonth = monthFilter === 'ALL' || d.month_key === monthFilter;
    const matchesDayType = dayTypeFilter === 'ALL' || d.day_type === dayTypeFilter;
    const matchesDutyType = dutyTypeFilter === 'ALL' || d.duty_type === dutyTypeFilter;
    return matchesSearch && matchesMonth && matchesDayType && matchesDutyType;
  });

  filteredDuties = [...filteredDuties].sort((a, b) => {
    const dateA = new Date(a.date);
    const dateB = new Date(b.date);
    return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
  });

  const liveTotal = filteredDuties.length;
  const liveGuard = filteredDuties.filter((d) => d.duty_type === 'Guard').length;
  const liveBds = filteredDuties.filter((d) => d.duty_type === 'BDS').length;
  const liveWeekdays = filteredDuties.filter((d) => d.day_type === 'Weekday').length;
  const liveFridays = filteredDuties.filter((d) => d.day_type === 'Friday').length;
  const liveWeekends = filteredDuties.filter((d) => d.day_type === 'Weekend').length;

  return (
    <div className="container">
      {/* Header */}
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
              <button className="refresh-btn" onClick={handleRefresh} disabled={loading}>
                <RefreshCw size={14} className={loading ? 'spin' : ''} />
                {loading ? 'Refreshing...' : 'Refresh Data'}
              </button>
              <ThemeToggle theme={theme} onToggle={toggleTheme} />
            </div>
            <div style={{ textAlign: 'right' }}>
              <DataFreshnessBadge dataStatus={dataStatus} />
              <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', marginTop: '4px' }}>
                Report generated: {new Date(metadata.generated_at).toLocaleTimeString()}
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Stats Grid */}
      <section className="stats-grid">
        <div className="stat-card primary-card">
          <span className="stat-label">Total Duties</span>
          <div className="stat-value">{liveTotal}</div>
          <span className="stat-desc">{liveGuard} Guard · {liveBds} BDS</span>
        </div>

        <div className="stat-card bds-card">
          <span className="stat-label">BDS Duties</span>
          <div className="stat-value" style={{ color: 'var(--stat-bds-color)' }}>{liveBds}</div>
          <span className="stat-desc">
            {liveTotal > 0 ? ((liveBds / liveTotal) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>

        <div className="stat-card weekday-card">
          <span className="stat-label">Weekdays (Mon-Thu)</span>
          <div className="stat-value" style={{ color: 'var(--stat-weekday-color)' }}>{liveWeekdays}</div>
          <span className="stat-desc">
            {liveTotal > 0 ? ((liveWeekdays / liveTotal) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>

        <div className="stat-card friday-card">
          <span className="stat-label">Fridays</span>
          <div className="stat-value" style={{ color: 'var(--stat-friday-color)' }}>{liveFridays}</div>
          <span className="stat-desc">
            {liveTotal > 0 ? ((liveFridays / liveTotal) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>

        <div className="stat-card weekend-card">
          <span className="stat-label">Weekends (Sat-Sun)</span>
          <div className="stat-value" style={{ color: 'var(--stat-weekend-color)' }}>{liveWeekends}</div>
          <span className="stat-desc">
            {liveTotal > 0 ? ((liveWeekends / liveTotal) * 100).toFixed(1) : 0}% of filtered total
          </span>
        </div>
      </section>

      {/* Filters */}
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
              {uniqueMonths.map((mKey) => (
                <option key={mKey} value={mKey}>{monthLabelsMap[mKey]}</option>
              ))}
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="dayType">Type of Day</label>
            <select id="dayType" value={dayTypeFilter} onChange={(e) => setDayTypeFilter(e.target.value)} className="filter-select">
              <option value="ALL">All Day Types</option>
              <option value="Weekday">Weekday (Mon–Thu)</option>
              <option value="Friday">Friday</option>
              <option value="Weekend">Weekend (Sat–Sun)</option>
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

      {/* Data Table */}
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
                {filteredDuties.map((d, index) => (
                  <tr key={`${d.date}-${d.duty_type}-${d.number}`}>
                    <td style={{ fontWeight: 600, color: 'var(--text-muted)' }}>{index + 1}</td>
                    <td style={{ fontWeight: 500 }}>{d.date_label}</td>
                    <td>{d.day}</td>
                    <td>
                      <span className={`badge badge-${d.day_type.toLowerCase()}`}>
                        {d.day_type === 'Weekday' ? 'Weekday (Mon-Thu)' : d.day_type === 'Friday' ? 'Friday' : 'Weekend (Sat-Sun)'}
                      </span>
                    </td>
                    <td>
                      <span className={`badge badge-${d.duty_type.toLowerCase()}`}>
                        {d.duty_type}
                      </span>
                    </td>
                    <td style={{ fontWeight: 600 }}>{d.duty_display}</td>
                    <td style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                      Sheet: {d.sheet} (Row: {d.row})
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      <footer>
        <p>© 2026 FOE Duty Checker dashboard · Built dynamically with React + Vite + Flask</p>
      </footer>

      <Toast toast={toast} onDismiss={() => setToast(null)} />
    </div>
  );
}

export default App;
