import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  Activity,
  AlertCircle,
  ArrowLeft,
  ArrowUpDown,
  Award,
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
  Users,
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
  const headers = ['#', 'Date', 'Day', 'Day Category', 'Duty Type', 'Entry Details', 'Points', 'Sheet', 'Row'];
  const body = rows.map((row, index) => [
    index + 1,
    row.date_label,
    row.day,
    row.day_type,
    row.duty_type,
    row.duty_display,
    (row.points ?? 0).toFixed(2),
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

function DutyPointsSection({ points }) {
  const [view, setView] = useState('simplified');

  if (!points) return null;

  const progressPercent = Math.min(
    ((points.remaining_points / points.points_per_off) * 100),
    100
  );

  return (
    <section className="points-section">
      <div className="points-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <Award size={20} style={{ color: 'var(--warning)' }} />
          <h3>Duty Points & Compensation</h3>
        </div>
        <div className="points-view-toggle">
          <button
            className={view === 'simplified' ? 'active' : ''}
            onClick={() => setView('simplified')}
          >
            Simplified
          </button>
          <button
            className={view === 'detailed' ? 'active' : ''}
            onClick={() => setView('detailed')}
          >
            Detailed
          </button>
        </div>
      </div>

      {view === 'simplified' && (
        <div className="points-simplified">
          <div className="points-summary-cards">
            <div className="points-big-card">
              <span className="points-big-value">
                {points.grand_total_points}
              </span>
              <span className="points-big-label">Total Points</span>
            </div>
            <div className="points-big-card offs">
              <span className="points-big-value">
                {points.full_offs}
              </span>
              <span className="points-big-label">Off Days Earned</span>
            </div>
          </div>

          <div className="points-progress-section">
            <div className="points-progress-label">
              <span>Progress to next off</span>
              <span>{points.remaining_points} / {points.points_per_off} pts</span>
            </div>
            <div className="points-progress-track">
              <div
                className="points-progress-fill"
                style={{ width: `${progressPercent}%` }}
              />
            </div>
            <span className="points-progress-hint">
              {points.points_to_next_off > 0
                ? `${points.points_to_next_off} more points needed for next off`
                : 'Off day threshold reached!'}
            </span>
          </div>

          <div className="points-quick-row">
            <div className="points-quick-item">
              <Shield size={14} style={{ color: 'var(--info)' }} />
              <span>Guard: <strong>{points.guard_total_points} pts</strong> ({points.guard_total_duties} duties)</span>
            </div>
            <div className="points-quick-item">
              <Activity size={14} style={{ color: 'var(--success)' }} />
              <span>BDS: <strong>{points.bds_total_points} pts</strong> ({points.bds_total_duties} duties)</span>
            </div>
          </div>
        </div>
      )}

      {view === 'detailed' && (
        <div className="points-detailed">
          <table className="points-table">
            <thead>
              <tr>
                <th>Duty Type</th>
                <th>Day Category</th>
                <th>Duties</th>
                <th>Pts/Duty</th>
                <th>Subtotal</th>
              </tr>
            </thead>
            <tbody>
              <tr className="points-row guard">
                <td rowSpan={3}>
                  <span className="badge badge-guard">Guard</span>
                </td>
                <td>Weekday (Mon-Fri)</td>
                <td>{points.breakdown.Guard.weekday.count}</td>
                <td>0.20</td>
                <td>{points.breakdown.Guard.weekday.points.toFixed(2)}</td>
              </tr>
              <tr className="points-row guard">
                <td>Saturday</td>
                <td>{points.breakdown.Guard.saturday.count}</td>
                <td>0.40</td>
                <td>{points.breakdown.Guard.saturday.points.toFixed(2)}</td>
              </tr>
              <tr className="points-row guard">
                <td>Sunday</td>
                <td>{points.breakdown.Guard.sunday.count}</td>
                <td>0.30</td>
                <td>{points.breakdown.Guard.sunday.points.toFixed(2)}</td>
              </tr>
              <tr className="points-subtotal">
                <td colSpan={4}>Guard Subtotal</td>
                <td><strong>{points.guard_total_points.toFixed(2)}</strong></td>
              </tr>

              <tr className="points-row bds">
                <td rowSpan={3}>
                  <span className="badge badge-bds">BDS</span>
                </td>
                <td>Weekday (Mon-Fri)</td>
                <td>{points.breakdown.BDS.weekday.count}</td>
                <td>0.00</td>
                <td>{points.breakdown.BDS.weekday.points.toFixed(2)}</td>
              </tr>
              <tr className="points-row bds">
                <td>Saturday</td>
                <td>{points.breakdown.BDS.saturday.count}</td>
                <td>0.25</td>
                <td>{points.breakdown.BDS.saturday.points.toFixed(2)}</td>
              </tr>
              <tr className="points-row bds">
                <td>Sunday</td>
                <td>{points.breakdown.BDS.sunday.count}</td>
                <td>0.25</td>
                <td>{points.breakdown.BDS.sunday.points.toFixed(2)}</td>
              </tr>
              <tr className="points-subtotal">
                <td colSpan={4}>BDS Subtotal</td>
                <td><strong>{points.bds_total_points.toFixed(2)}</strong></td>
              </tr>

              <tr className="points-grand-total">
                <td colSpan={4}>Grand Total</td>
                <td><strong>{points.grand_total_points.toFixed(2)}</strong></td>
              </tr>
              <tr className="points-grand-total">
                <td colSpan={4}>Estimated Off Days (&divide; {points.points_per_off})</td>
                <td><strong>{points.full_offs} offs</strong> + {points.remaining_points.toFixed(2)} pts remaining</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
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
  if (dataStatus.warming) messages.push(dataStatus.progress?.message || 'Preparing the duty index. This usually happens once after a cold start.');
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

function IndexProgress({ dataStatus, compact = false }) {
  const progress = dataStatus?.progress;
  if (!progress) return null;

  const isActive = dataStatus?.warming || dataStatus?.refreshing || ['loading_workbook', 'downloading', 'scanning_sheets', 'writing_index'].includes(progress.phase);
  if (!isActive && progress.phase !== 'error' && progress.phase !== 'download_failed') return null;

  const percent = typeof progress.percent === 'number'
    ? Math.max(0, Math.min(100, progress.percent))
    : null;
  const sheetText = progress.total_sheets > 0
    ? `${progress.completed_sheets} / ${progress.total_sheets} sheets`
    : '';

  return (
    <div className={`index-progress ${compact ? 'compact' : ''}`}>
      <div className="index-progress-meta">
        <span>{progress.message || 'Preparing duty index...'}</span>
        <strong>{percent === null ? 'Working...' : `${percent}%`}</strong>
      </div>
      <div className={`index-progress-track ${percent === null ? 'indeterminate' : ''}`}>
        <div
          className="index-progress-fill"
          style={percent === null ? undefined : { width: `${percent}%` }}
        />
      </div>
      {(sheetText || progress.current_sheet) && (
        <div className="index-progress-detail">
          {sheetText}
          {sheetText && progress.current_sheet ? ' - ' : ''}
          {progress.current_sheet || ''}
        </div>
      )}
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
  onCompare,
  theme,
  onToggleTheme,
  dataStatus,
  recentNames,
}) {
  const [mode, setMode] = useState('single');
  const [selectedName, setSelectedName] = useState('');
  const [selectedNameB, setSelectedNameB] = useState('');
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

  const readyToProceed = mode === 'single'
    ? (selectedName && startMonth && endMonth && !isRangeInvalid)
    : (selectedName && selectedNameB && selectedName !== selectedNameB && startMonth && endMonth && !isRangeInvalid);

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
              <IndexProgress dataStatus={dataStatus} />
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
              <div className="compare-mode-toggle">
                <button className={mode === 'single' ? 'active' : ''} onClick={() => setMode('single')}>
                  <FileText size={14} style={{ marginRight: '6px', verticalAlign: '-2px' }} />Single View
                </button>
                <button className={mode === 'compare' ? 'active' : ''} onClick={() => setMode('compare')}>
                  <Users size={14} style={{ marginRight: '6px', verticalAlign: '-2px' }} />Compare Mode
                </button>
              </div>

              {recentNames.length > 0 && mode === 'single' && (
                <div className="recent-row">
                  {recentNames.map((name) => (
                    <button key={name} type="button" className="recent-pill" onClick={() => setSelectedName(name)}>
                      {name}
                    </button>
                  ))}
                </div>
              )}

              {mode === 'single' ? (
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
              ) : (
                <>
                  <div className="compare-name-grid">
                    <div className="name-select-wrapper">
                      <label className="selection-label">
                        <span style={{ color: 'var(--primary)', fontWeight: 700 }}>A</span>
                        &nbsp;Person A
                      </label>
                      <select
                        className="filter-select"
                        value={selectedName}
                        onChange={(e) => setSelectedName(e.target.value)}
                        style={{ width: '100%', padding: '12px 14px' }}
                      >
                        <option value="">Select Person A</option>
                        {names.map((name) => (
                          <option key={name} value={name}>{name}</option>
                        ))}
                      </select>
                    </div>
                    <div className="name-select-wrapper">
                      <label className="selection-label">
                        <span style={{ color: 'var(--info)', fontWeight: 700 }}>B</span>
                        &nbsp;Person B
                      </label>
                      <select
                        className="filter-select"
                        value={selectedNameB}
                        onChange={(e) => setSelectedNameB(e.target.value)}
                        style={{ width: '100%', padding: '12px 14px' }}
                      >
                        <option value="">Select Person B</option>
                        {names.filter((n) => n !== selectedName).map((name) => (
                          <option key={name} value={name}>{name}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  {selectedName && selectedNameB && selectedName === selectedNameB && (
                    <div className="validation-error">
                      <AlertCircle size={16} />
                      <span>Please select two different people to compare.</span>
                    </div>
                  )}
                </>
              )}

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
                onClick={() => {
                  if (mode === 'compare') {
                    onCompare(selectedName, selectedNameB, startMonth, endMonth);
                  } else {
                    onSelect(selectedName, startMonth, endMonth);
                  }
                }}
              >
                {mode === 'compare'
                  ? (readyToProceed ? `Compare ${selectedName} vs ${selectedNameB}` : 'Select two people to compare')
                  : (selectedName ? `Load Dashboard for ${selectedName}` : 'Select a person to continue')
                }
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function ComparisonDiffBadge({ a, b, label }) {
  const diff = a - b;
  if (diff === 0) return <span className="compare-diff equal">=</span>;
  return (
    <span className={`compare-diff ${diff > 0 ? 'more' : 'less'}`}>
      {diff > 0 ? `+${diff}` : diff} {label || ''}
    </span>
  );
}

function ComparisonView({ data, onBack, theme, onToggleTheme }) {
  if (!data || data.length !== 2) return null;

  const [a, b] = data;
  const nameA = a.metadata?.name || 'Person A';
  const nameB = b.metadata?.name || 'Person B';

  const statRows = [
    { label: 'Total Duties', keyA: a.summary.total_duties, keyB: b.summary.total_duties },
    { label: 'Guard', keyA: a.summary.total_guard, keyB: b.summary.total_guard },
    { label: 'BDS', keyA: a.summary.total_bds, keyB: b.summary.total_bds },
    { label: 'Weekdays', keyA: a.summary.weekdays, keyB: b.summary.weekdays },
    { label: 'Fridays', keyA: a.summary.fridays, keyB: b.summary.fridays },
    { label: 'Weekends', keyA: a.summary.weekends, keyB: b.summary.weekends },
  ];

  const mergedDuties = [
    ...(a.duties || []).map((d) => ({ ...d, _person: 'a', _name: nameA })),
    ...(b.duties || []).map((d) => ({ ...d, _person: 'b', _name: nameB })),
  ].sort((x, y) => new Date(x.date) - new Date(y.date));

  return (
    <div className="compare-container container">
      <div className="compare-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <button className="back-btn" onClick={onBack}>
            <ArrowLeft size={16} />
            <span>Back</span>
          </button>
          <h1>
            <Users size={22} style={{ verticalAlign: '-4px', marginRight: '8px' }} />
            {nameA} vs {nameB}
          </h1>
        </div>
        <ThemeToggle theme={theme} onToggle={onToggleTheme} />
      </div>

      <div className="compare-stat-grid">
        <div className="compare-person-card">
          <h3><span className="person-label-a">{nameA}</span></h3>
          {statRows.map((row) => (
            <div className="compare-stat-row" key={row.label}>
              <span className="stat-name">{row.label}</span>
              <span className="stat-val">
                {row.keyA}
                <ComparisonDiffBadge a={row.keyA} b={row.keyB} />
              </span>
            </div>
          ))}
        </div>

        <div className="compare-person-card">
          <h3><span className="person-label-b">{nameB}</span></h3>
          {statRows.map((row) => (
            <div className="compare-stat-row" key={row.label}>
              <span className="stat-name">{row.label}</span>
              <span className="stat-val">
                {row.keyB}
                <ComparisonDiffBadge a={row.keyB} b={row.keyA} />
              </span>
            </div>
          ))}
        </div>
      </div>

      {a.points && b.points && (
        <section className="compare-points-section">
          <h3>
            <Award size={16} style={{ color: 'var(--warning)', verticalAlign: '-2px', marginRight: '8px' }} />
            Points Comparison
          </h3>
          <div className="compare-points-grid">
            <div className="compare-points-card person-a">
              <span className="compare-points-value">{a.points.grand_total_points}</span>
              <span className="compare-points-sub">{nameA} · {a.points.full_offs} off(s) earned</span>
            </div>
            <div className="compare-points-card person-b">
              <span className="compare-points-value">{b.points.grand_total_points}</span>
              <span className="compare-points-sub">{nameB} · {b.points.full_offs} off(s) earned</span>
            </div>
          </div>
        </section>
      )}

      <section className="compare-timeline-section">
        <h3>
          <Calendar size={16} style={{ color: 'var(--primary)', verticalAlign: '-2px', marginRight: '8px' }} />
          Merged Timeline ({mergedDuties.length} duties)
        </h3>
        <div style={{ overflowX: 'auto' }}>
          <table>
            <thead>
              <tr>
                <th>Person</th>
                <th>Date</th>
                <th>Day</th>
                <th>Duty Type</th>
                <th>Entry Details</th>
                <th>Points</th>
              </tr>
            </thead>
            <tbody>
              {mergedDuties.map((duty, i) => (
                <tr key={`${duty._person}-${duty.date}-${duty.duty_type}-${i}`} className={`${duty._person === 'a' ? 'person-a-row' : 'person-b-row'}`}>
                  <td>
                    <span className={`person-tag ${duty._person}`}>{duty._name}</span>
                  </td>
                  <td style={{ fontWeight: 500 }}>{duty.date_label}</td>
                  <td>{duty.day}</td>
                  <td>
                    <span className={`badge badge-${duty.duty_type.toLowerCase()}`}>{duty.duty_type}</span>
                  </td>
                  <td style={{ fontWeight: 600 }}>{duty.duty_display}</td>
                  <td style={{ fontWeight: 600, color: duty.points > 0 ? 'var(--warning)' : 'var(--text-muted)' }}>
                    {duty.points > 0 ? `+${duty.points.toFixed(2)}` : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <footer>
        <p>2026 FOE Duty Checker dashboard - React + Vite + Flask</p>
      </footer>
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
  const [compareData, setCompareData] = useState(null);
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

  const handleRefresh = () => {
    if (!activeName || refreshing) return;
    setRefreshing(true);

    const es = new EventSource('/api/refresh/stream');

    es.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        
        if (data.phase === 'rate_limited') {
          es.close();
          setRefreshing(false);
          setToast({ type: 'error', message: data.message || 'Refresh is on cooldown.' });
          return;
        }

        setDataStatus((prev) => ({
          ...prev,
          refreshing: data.refreshing,
          ready: data.ready,
          progress: {
            phase: data.phase,
            percent: data.percent,
            message: data.message,
            completed_sheets: data.completed_sheets,
            total_sheets: data.total_sheets
          }
        }));

        if (data.phase === 'done' || !data.refreshing) {
          es.close();
          fetchDuties(activeName, activeStart, activeEnd, { silent: true });
          setRefreshing(false);
          setToast({ type: 'success', message: 'Data refresh complete.' });
        }
      } catch (err) {
        console.error('SSE parse error:', err);
      }
    };

    es.onerror = () => {
      es.close();
      setRefreshing(false);
      setToast({ type: 'error', message: 'Refresh stream connection lost.' });
      fetchDuties(activeName, activeStart, activeEnd, { silent: true });
    };
  };

  const handleBack = () => {
    setScreen('select');
    setData(null);
    setCompareData(null);
    setActiveName('');
    window.history.replaceState(null, '', window.location.pathname);
  };

  const handleCompare = useCallback(async (nameA, nameB, start, end) => {
    setLoading(true);
    setScreen('compare');
    try {
      const res = await fetch('/api/compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ names: [nameA, nameB], start, end }),
      });
      const json = await res.json();
      if (json.data_status) setDataStatus(json.data_status);
      if (json.error) {
        setToast({ type: 'error', message: json.error });
        setScreen('select');
      } else {
        setCompareData(json.comparison);
      }
    } catch (err) {
      setToast({ type: 'error', message: `Compare failed: ${err.message}` });
      setScreen('select');
    } finally {
      setLoading(false);
    }
  }, []);

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
          onCompare={handleCompare}
          theme={theme}
          onToggleTheme={toggleTheme}
          dataStatus={dataStatus}
          recentNames={recentNames.filter((name) => names.includes(name))}
        />
        <Toast toast={toast} onDismiss={() => setToast(null)} />
      </>
    );
  }

  if (screen === 'compare') {
    if (loading) {
      return (
        <div className="container" style={{ textAlign: 'center', padding: '80px 0' }}>
          <RefreshCw size={36} style={{ color: 'var(--primary)', animation: 'spin 1s linear infinite' }} />
          <p style={{ marginTop: '16px', color: 'var(--text-muted)' }}>Loading comparison data...</p>
        </div>
      );
    }
    return (
      <>
        <ComparisonView
          data={compareData}
          onBack={handleBack}
          theme={theme}
          onToggleTheme={toggleTheme}
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
              <IndexProgress dataStatus={dataStatus} compact />
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

      <DutyPointsSection points={data?.points} />

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
                  <th>Points</th>
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
                    <td style={{ fontWeight: 600, color: duty.points > 0 ? 'var(--warning)' : 'var(--text-muted)' }}>
                      {duty.points > 0 ? `+${duty.points.toFixed(2)}` : '—'}
                    </td>
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
