import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const demoUser = {
  username: 'admin',
  password: '1234',
  displayName: 'AP User',
}

const reportMeta = {
  ap: {
    title: 'Accounts Payable',
    subtitle: 'Purchase invoice and credit memo review from SAP Business One.',
    source: 'PU = OPCH, PC = ORPC',
    endpoint: '/api/ap',
    filters: ['ALL', 'PU', 'PC'],
  },
  cash: {
    title: 'Incoming Cash',
    subtitle: 'Incoming payment and outgoing payment cash movement from SAP Business One.',
    source: 'RC = ORCT/RCT3, PS = OVPM/VPM3',
    endpoint: '/api/incoming-cash',
    filters: ['ALL', 'RC', 'PS'],
  },
}

function money(value) {
  return new Intl.NumberFormat('th-TH', {
    style: 'currency',
    currency: 'THB',
    maximumFractionDigits: 2,
  }).format(value)
}

function isoToDisplayDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value)
  if (!match) return ''

  const [, year, month, day] = match
  return `${day}/${month}/${year}`
}

function buildExcelRows(report, rows) {
  if (report === 'ap') {
    return rows.map((row) => ({
      Type: row.type,
      'Vendor Code': row.vendorCode,
      'Vendor Name': row.vendorName,
      'Doc Number': row.docNumber,
      'Posting Date': row.postingDate,
      'Due Date': row.dueDate,
      'BP Ref Number': row.bpRef,
      Remarks: row.remarks,
      Amount: Number(row.amount || 0),
      Balance: Number(row.balance || 0),
      Tax: Number(row.tax || 0),
    }))
  }

  return rows.map((row) => ({
    Brancod: row.channel,
    'Doc Date': row.docDate,
    'Doc No': row.docNumber,
    'Card Code': row.cardCode,
    'Branch Name': row.cardName,
    Brand: row.brand,
    'Branch code': row.whsCode,
    'Journal Memo': row.journalMemo,
    'Account Code': row.acctCode,
    'Account Name': row.acctName,
    Amount: Number(row.amount || 0),
  }))
}

function App() {
  const [user, setUser] = useState(() => {
    const savedUser = localStorage.getItem('appsapb1-user')
    return savedUser ? JSON.parse(savedUser) : null
  })
  const [form, setForm] = useState({ username: '', password: '' })
  const [error, setError] = useState('')
  const [activeView, setActiveView] = useState('ap')
  const [filter, setFilter] = useState('ALL')
  const [dateRange, setDateRange] = useState({ st: '2023-11-01', ed: '2023-12-31' })
  const [reportRows, setReportRows] = useState([])
  const [loading, setLoading] = useState(false)
  const [loadError, setLoadError] = useState('')
  const [hasLoadedReport, setHasLoadedReport] = useState(false)
  const [dbConfig, setDbConfig] = useState({
    server: '',
    port: '1433',
    database: '',
    user: '',
    password: '',
    encrypt: false,
    trustServerCertificate: true,
    hasPassword: false,
  })
  const [dbStatus, setDbStatus] = useState('')

  useEffect(() => {
    if (!user) return

    async function loadDbConfig() {
      try {
        const response = await fetch('/api/db-config')
        const data = await response.json()
        setDbConfig((current) => ({ ...current, ...data, password: '' }))
      } catch (configError) {
        setDbStatus(`Cannot load database config: ${configError.message}`)
      }
    }

    loadDbConfig()
  }, [user])

  const filteredRows = useMemo(() => {
    return reportRows.filter((row) => {
      const matchesType = filter === 'ALL' || row.type === filter
      return matchesType
    })
  }, [filter, reportRows])

  const totals = useMemo(
    () =>
      filteredRows.reduce(
        (summary, row) => ({
          amount: summary.amount + Number(row.amount || 0),
          balance: summary.balance + Number(row.balance || 0),
          tax: summary.tax + Number(row.tax || 0),
        }),
        { amount: 0, balance: 0, tax: 0 },
      ),
    [filteredRows],
  )

  function selectReport(nextReport) {
    setActiveView(nextReport)
    setFilter('ALL')
    setReportRows([])
    setLoadError('')
    setHasLoadedReport(false)
  }

  async function loadReport() {
    setLoading(true)
    setLoadError('')
    setHasLoadedReport(true)

    try {
      if (!dateRange.st || !dateRange.ed) {
        throw new Error('กรุณาเลือกวันที่เริ่มต้นและวันที่สิ้นสุด')
      }

      const params = new URLSearchParams(dateRange)
      const response = await fetch(`${reportMeta[activeView].endpoint}?${params.toString()}`)

      if (!response.ok) {
        const body = await response.json().catch(() => ({}))
        throw new Error(body.detail || body.message || 'Load report data failed')
      }

      const data = await response.json()
      setReportRows(data.rows || [])
    } catch (fetchError) {
      setLoadError(fetchError.message)
      setReportRows([])
    } finally {
      setLoading(false)
    }
  }

  function exportToExcel() {
    if (!filteredRows.length) return

    const excelRows = buildExcelRows(activeView, filteredRows)
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['APPSAPB1'],
      [reportMeta[activeView].title],
      [reportMeta[activeView].subtitle],
      [`Date: ${isoToDisplayDate(dateRange.st)} - ${isoToDisplayDate(dateRange.ed)}`],
      [],
    ])
    XLSX.utils.sheet_add_json(worksheet, excelRows, { origin: 'A6' })
    const workbook = XLSX.utils.book_new()
    const sheetName = activeView === 'ap' ? 'Accounts Payable' : 'Incoming Cash'
    const fileName = `${sheetName.replaceAll(' ', '_')}_${dateRange.st}_${dateRange.ed}.xlsx`

    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
    XLSX.writeFile(workbook, fileName)
  }

  function handleLogin(event) {
    event.preventDefault()

    if (form.username === demoUser.username && form.password === demoUser.password) {
      const nextUser = { displayName: demoUser.displayName, username: demoUser.username }
      localStorage.setItem('appsapb1-user', JSON.stringify(nextUser))
      setUser(nextUser)
      setError('')
      return
    }

    setError('Username หรือ Password ไม่ถูกต้อง')
  }

  function handleLogout() {
    localStorage.removeItem('appsapb1-user')
    setUser(null)
    setForm({ username: '', password: '' })
  }

  async function handleSaveDbConfig(event) {
    event.preventDefault()
    setDbStatus('Saving database server...')

    try {
      const response = await fetch('/api/db-config', {
        body: JSON.stringify(dbConfig),
        headers: { 'Content-Type': 'application/json' },
        method: 'POST',
      })
      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.message || 'Save database config failed')
      }

      setDbConfig((current) => ({ ...current, ...data.config, password: '' }))
      setDbStatus('Saved. Testing connection...')
      await handleTestConnection()
    } catch (saveError) {
      setDbStatus(`Save failed: ${saveError.message}`)
    }
  }

  function updateDateRange(key, value) {
    setDateRange((current) => ({ ...current, [key]: value }))
    setReportRows([])
    setLoadError('')
    setHasLoadedReport(false)
  }

  async function handleTestConnection() {
    const response = await fetch('/api/health')
    const data = await response.json().catch(() => ({}))

    if (!response.ok) {
      const message = data.message || 'Connection failed'
      setDbStatus(`Connection failed: ${message}`)
      throw new Error(message)
    }

    setDbStatus('Connection OK')
  }

  if (!user) {
    return (
      <main className="login-page">
        <section className="login-brand" aria-label="Application summary">
          <div className="brand-mark">AP</div>
          <p className="eyebrow">SAP Business One</p>
          <h1>APPSAPB1</h1>
          <p>
            Enterprise Accounts Payable workspace for reviewing purchase invoices,
            credit memos, balances, and tax totals.
          </p>
        </section>
        <section className="login-panel" aria-labelledby="login-title">
          <div>
            <p className="eyebrow">APPSAPB1</p>
            <h1 id="login-title">Login</h1>
            <p className="muted">เข้าสู่ระบบเพื่อดูรายงาน AP จาก SAP Business One</p>
          </div>

          <form className="login-form" onSubmit={handleLogin}>
            <label>
              Username
              <input
                autoComplete="username"
                onChange={(event) => setForm({ ...form, username: event.target.value })}
                placeholder="admin"
                type="text"
                value={form.username}
              />
            </label>

            <label>
              Password
              <input
                autoComplete="current-password"
                onChange={(event) => setForm({ ...form, password: event.target.value })}
                placeholder="1234"
                type="password"
                value={form.password}
              />
            </label>

            {error && <p className="error">{error}</p>}

            <button type="submit">Sign in</button>
          </form>

          <p className="hint">Demo user: admin / 1234</p>
        </section>
      </main>
    )
  }

  return (
    <main className="enterprise-shell">
      <aside className="sidebar">
        <div className="sidebar-brand">
          <div className="brand-mark">AP</div>
          <div>
            <strong>APPSAPB1</strong>
            <span>SAP B1 Reporting</span>
          </div>
        </div>
        <nav className="side-nav" aria-label="Main navigation">
          <button
            className={activeView === 'ap' ? 'active' : ''}
            onClick={() => selectReport('ap')}
            type="button"
          >
            Accounts Payable
          </button>
          <button
            className={activeView === 'cash' ? 'active' : ''}
            onClick={() => selectReport('cash')}
            type="button"
          >
            Incoming Cash
          </button>
          <button
            className={activeView === 'database' ? 'active' : ''}
            type="button"
            onClick={() => setActiveView('database')}
          >
            Database Server
          </button>
        </nav>
        <div className="sidebar-footer">
          <span>Signed in as</span>
          <strong>{user.displayName}</strong>
        </div>
      </aside>

      <section className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">APPSAPB1</p>
          <h1>{activeView === 'database' ? 'Database Server' : reportMeta[activeView].title}</h1>
          <p className="muted">
            {activeView === 'database'
              ? 'Manage SAP SQL Server connection settings.'
              : reportMeta[activeView].subtitle}
          </p>
        </div>
        <div className="user-box">
          <button
            className="secondary"
            onClick={() => setActiveView('database')}
            type="button"
          >
            Database Server
          </button>
          <button className="secondary" onClick={handleLogout} type="button">
            Logout
          </button>
        </div>
      </header>

      {activeView === 'database' ? (
        <section className="db-panel" aria-labelledby="db-title">
          <div>
            <p className="eyebrow">Connection</p>
            <h2 id="db-title">Edit Database Server</h2>
            <p className="muted">แก้ค่าการเชื่อมต่อ SAP SQL Server แล้วกด Save & Test</p>
          </div>

          <form className="db-form" onSubmit={handleSaveDbConfig}>
            <label>
              Server
              <input
                onChange={(event) => setDbConfig({ ...dbConfig, server: event.target.value })}
                placeholder="SERVER\\INSTANCE หรือ IP"
                value={dbConfig.server}
              />
            </label>
            <label>
              Port
              <input
                onChange={(event) => setDbConfig({ ...dbConfig, port: event.target.value })}
                placeholder="1433"
                value={dbConfig.port}
              />
            </label>
            <label>
              Database
              <input
                onChange={(event) => setDbConfig({ ...dbConfig, database: event.target.value })}
                placeholder="SBO_PFC_PE"
                value={dbConfig.database}
              />
            </label>
            <label>
              SQL User
              <input
                autoComplete="username"
                onChange={(event) => setDbConfig({ ...dbConfig, user: event.target.value })}
                value={dbConfig.user}
              />
            </label>
            <label>
              SQL Password
              <input
                autoComplete="new-password"
                onChange={(event) => setDbConfig({ ...dbConfig, password: event.target.value })}
                placeholder={dbConfig.hasPassword ? 'Leave blank to keep current password' : ''}
                type="password"
                value={dbConfig.password}
              />
            </label>
            <label className="check-row">
              <input
                checked={dbConfig.encrypt}
                onChange={(event) => setDbConfig({ ...dbConfig, encrypt: event.target.checked })}
                type="checkbox"
              />
              Encrypt
            </label>
            <label className="check-row">
              <input
                checked={dbConfig.trustServerCertificate}
                onChange={(event) =>
                  setDbConfig({ ...dbConfig, trustServerCertificate: event.target.checked })
                }
                type="checkbox"
              />
              Trust Server Certificate
            </label>
            <div className="db-actions">
              <button type="submit">Save & Test</button>
              <button
                className="secondary"
                onClick={() => {
                  setDbStatus('Testing connection...')
                  handleTestConnection().catch(() => {})
                }}
                type="button"
              >
                Test Current
              </button>
            </div>
          </form>

          {dbStatus && <p className="table-state">{dbStatus}</p>}
        </section>
      ) : (
        <>

      <section className="toolbar" aria-label="AP filters">
        <div>
          <strong>Source</strong>
          <span>{reportMeta[activeView].source}</span>
        </div>
        <div className="date-filters">
          <label>
            From
            <span className="date-combo">
              <input readOnly type="text" value={isoToDisplayDate(dateRange.st)} />
              <input
                aria-label="Select from date"
                className="date-native"
                onChange={(event) => updateDateRange('st', event.target.value)}
                type="date"
                value={dateRange.st}
              />
            </span>
          </label>
          <label>
            To
            <span className="date-combo">
              <input readOnly type="text" value={isoToDisplayDate(dateRange.ed)} />
              <input
                aria-label="Select to date"
                className="date-native"
                onChange={(event) => updateDateRange('ed', event.target.value)}
                type="date"
                value={dateRange.ed}
              />
            </span>
          </label>
          <button onClick={loadReport} type="button">
            ค้นหา
          </button>
          <button
            className="secondary"
            disabled={!filteredRows.length}
            onClick={exportToExcel}
            type="button"
          >
            Export to Excel
          </button>
        </div>
        <div className="segmented" role="group" aria-label="Document type">
          {reportMeta[activeView].filters.map((item) => (
            <button
              className={filter === item ? 'active' : ''}
              key={item}
              onClick={() => setFilter(item)}
              type="button"
            >
              {item}
            </button>
          ))}
        </div>
      </section>

      <section className="summary-grid" aria-label="AP summary">
        <article>
          <span>{activeView === 'cash' ? 'Cash Amount' : 'Total Amount'}</span>
          <strong>{money(totals.amount)}</strong>
        </article>
        <article>
          <span>{activeView === 'cash' ? 'Receipts / RC' : 'Balance Amount'}</span>
          <strong>
            {activeView === 'cash'
              ? money(filteredRows.filter((row) => row.type === 'RC').reduce((sum, row) => sum + Number(row.amount || 0), 0))
              : money(totals.balance)}
          </strong>
        </article>
        <article>
          <span>{activeView === 'cash' ? 'Payments / PS' : 'Tax'}</span>
          <strong>
            {activeView === 'cash'
              ? money(filteredRows.filter((row) => row.type === 'PS').reduce((sum, row) => sum + Number(row.amount || 0), 0))
              : money(totals.tax)}
          </strong>
        </article>
      </section>

      <section className="table-wrap" aria-label={`${reportMeta[activeView].title} report table`}>
        {!hasLoadedReport && !loading && (
          <p className="table-state">เลือกช่วงวันที่ แล้วกด “ค้นหา” เพื่อโหลดรายงาน</p>
        )}
        {loading && <p className="table-state">Loading {reportMeta[activeView].title} data...</p>}
        {loadError && <p className="table-state error">Database error: {loadError}</p>}
        {hasLoadedReport && !loading && !loadError && filteredRows.length === 0 && (
          <p className="table-state">ไม่พบข้อมูลตามช่วงวันที่ที่เลือก</p>
        )}
        {activeView === 'ap' ? (
        <table>
          <thead>
            <tr>
              <th>Type</th>
              <th>Vendor Code</th>
              <th>Vendor Name</th>
              <th>Doc Number</th>
              <th>Posting Date</th>
              <th>Due Date</th>
              <th>BP Ref Number</th>
              <th>Remarks</th>
              <th className="number">Amount</th>
              <th className="number">Balance</th>
              <th className="number">Tax</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row) => (
              <tr key={row.id}>
                <td>
                  <span className={`badge ${row.type.toLowerCase()}`}>{row.type}</span>
                </td>
                <td>{row.vendorCode}</td>
                <td>{row.vendorName}</td>
                <td>{row.docNumber}</td>
                <td>{row.postingDate}</td>
                <td>{row.dueDate}</td>
                <td>{row.bpRef}</td>
                <td>{row.remarks}</td>
                <td className="number">{money(row.amount)}</td>
                <td className="number">{money(row.balance)}</td>
                <td className="number">{money(row.tax)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        ) : (
          <table>
            <thead>
              <tr>
                <th>Brancod</th>
                <th>Doc Date</th>
                <th>Doc No</th>
                <th>Card Code</th>
                <th>Branch Name</th>
                <th>Brand</th>
                <th>Branch code</th>
                <th>Journal Memo</th>
                <th>Account Code</th>
                <th>Account Name</th>
                <th className="number">Amount</th>
              </tr>
            </thead>
            <tbody>
              {filteredRows.map((row) => (
                <tr key={row.id}>
                  <td>{row.channel}</td>
                  <td>{row.docDate}</td>
                  <td>
                    <span className={`badge ${row.type.toLowerCase()}`}>{row.docNumber}</span>
                  </td>
                  <td>{row.cardCode}</td>
                  <td>{row.cardName}</td>
                  <td>{row.brand}</td>
                  <td>{row.whsCode}</td>
                  <td>{row.journalMemo}</td>
                  <td>{row.acctCode}</td>
                  <td>{row.acctName}</td>
                  <td className="number">{money(row.amount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>
        </>
      )}
      </section>
    </main>
  )
}

export default App
