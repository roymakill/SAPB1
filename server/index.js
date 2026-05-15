import dotenv from 'dotenv'
import express from 'express'
import fs from 'node:fs/promises'
import path from 'node:path'
import sql from 'mssql'

dotenv.config()

const app = express()
const port = Number(process.env.API_PORT || 3001)
const envPath = path.resolve(process.cwd(), '.env')

app.use(express.json())

function buildDbConfig() {
  return {
    server: process.env.DB_SERVER,
    port: Number(process.env.DB_PORT || 1433),
    database: process.env.DB_DATABASE,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    options: {
      encrypt: process.env.DB_ENCRYPT === 'true',
      trustServerCertificate: process.env.DB_TRUST_CERT !== 'false',
    },
    pool: {
      max: 10,
      min: 0,
      idleTimeoutMillis: 30000,
    },
  }
}

let poolPromise
let activePool

async function getPool() {
  if (!poolPromise) {
    poolPromise = new sql.ConnectionPool(buildDbConfig()).connect()
    activePool = await poolPromise
  }

  return poolPromise
}

async function resetPool() {
  const pool = activePool
  poolPromise = undefined
  activePool = undefined

  if (pool) {
    await pool.close()
  }
}

function toSapDate(value) {
  if (!value) return null
  const normalized = String(value).replaceAll('-', '')
  return /^\d{8}$/.test(normalized) ? normalized : null
}

function publicDbConfig() {
  return {
    server: process.env.DB_SERVER || '',
    port: process.env.DB_PORT || '1433',
    database: process.env.DB_DATABASE || '',
    user: process.env.DB_USER || '',
    encrypt: process.env.DB_ENCRYPT === 'true',
    trustServerCertificate: process.env.DB_TRUST_CERT !== 'false',
    hasPassword: Boolean(process.env.DB_PASSWORD),
  }
}

function envValue(value) {
  return JSON.stringify(String(value ?? ''))
}

async function saveDbConfig(nextConfig) {
  const currentPassword = process.env.DB_PASSWORD || ''
  const nextPassword = nextConfig.password ? String(nextConfig.password) : currentPassword
  const env = {
    API_PORT: String(port),
    DB_SERVER: nextConfig.server || '',
    DB_PORT: nextConfig.port || '1433',
    DB_DATABASE: nextConfig.database || '',
    DB_USER: nextConfig.user || '',
    DB_PASSWORD: nextPassword,
    DB_ENCRYPT: nextConfig.encrypt ? 'true' : 'false',
    DB_TRUST_CERT: nextConfig.trustServerCertificate === false ? 'false' : 'true',
  }

  const content = `${Object.entries(env)
    .map(([key, value]) => `${key}=${envValue(value)}`)
    .join('\n')}\n`

  await fs.writeFile(envPath, content, 'utf8')
  Object.assign(process.env, env)
}

app.get('/api/db-config', (_request, response) => {
  response.json(publicDbConfig())
})

app.post('/api/db-config', async (request, response) => {
  try {
    await saveDbConfig(request.body || {})
    await resetPool()
    response.json({ ok: true, config: publicDbConfig() })
  } catch (error) {
    response.status(500).json({ ok: false, message: error.message })
  }
})

app.get('/api/health', async (_request, response) => {
  try {
    const pool = await getPool()
    await pool.request().query('SELECT 1 AS ok')
    response.json({ ok: true })
  } catch (error) {
    response.status(500).json({ ok: false, message: error.message })
  }
})

app.get('/api/ap', async (request, response) => {
  const startDate = toSapDate(request.query.st) || '19000101'
  const endDate = toSapDate(request.query.ed) || '20991231'

  try {
    const pool = await getPool()
    const result = await pool
      .request()
      .input('st', sql.NVarChar(8), startDate)
      .input('ed', sql.NVarChar(8), endDate)
      .query(`
        SELECT
          'PU' AS [type],
          [CardCode] AS [vendorCode],
          [CardName] AS [vendorName],
          'PU' + CONVERT(char(9), [DocNum]) AS [docNumber],
          CONVERT(NVARCHAR, [DocDate], 103) AS [postingDate],
          CONVERT(NVARCHAR, [DocDueDate], 103) AS [dueDate],
          [Comments] AS [remarks],
          [NumAtCard] AS [bpRef],
          [DocTotal] AS [amount],
          [DocTotal] - [PaidToDate] AS [balance],
          [VatSum] AS [tax],
          [DocDate] AS [sortDate],
          [DocNum] AS [sortNumber]
        FROM [OPCH]
        WHERE CONVERT(NVARCHAR, [DocDate], 112) >= @st
          AND CONVERT(NVARCHAR, [DocDate], 112) <= @ed

        UNION ALL

        SELECT
          'PC' AS [type],
          [CardCode] AS [vendorCode],
          [CardName] AS [vendorName],
          'PC' + CONVERT(char(9), [DocNum]) AS [docNumber],
          CONVERT(NVARCHAR, [DocDate], 103) AS [postingDate],
          CONVERT(NVARCHAR, [DocDueDate], 103) AS [dueDate],
          [Comments] AS [remarks],
          [NumAtCard] AS [bpRef],
          [DocTotal] * -1 AS [amount],
          ([DocTotal] - [PaidToDate]) * -1 AS [balance],
          [VatSum] * -1 AS [tax],
          [DocDate] AS [sortDate],
          [DocNum] AS [sortNumber]
        FROM [ORPC]
        WHERE CONVERT(NVARCHAR, [DocDate], 112) >= @st
          AND CONVERT(NVARCHAR, [DocDate], 112) <= @ed

        ORDER BY [sortDate], [sortNumber]
      `)

    response.json({
      rows: result.recordset.map(({ sortDate, sortNumber, ...row }, index) => ({
        id: `${row.type}-${sortNumber}-${index}`,
        ...row,
      })),
    })
  } catch (error) {
    response.status(500).json({
      message: 'Cannot load AP data from SAP database',
      detail: error.message,
    })
  }
})

app.get('/api/incoming-cash', async (request, response) => {
  const startDate = request.query.st || '1900-01-01'
  const endDate = request.query.ed || '2099-12-31'

  try {
    const pool = await getPool()
    const result = await pool
      .request()
      .input('st', sql.Date, startDate)
      .input('ed', sql.Date, endDate)
      .query(`
        SELECT
          'RC' AS [type],
          'SHOP' AS [channel],
          CONVERT(NVARCHAR, rc.DocDate, 103) AS [docDate],
          '100' + SUBSTRING(rc.CardCode, 4, 3) AS [brand],
          RIGHT(rc.CardCode, 8) AS [whsCode],
          rc.CardCode AS [cardCode],
          rc.CardName AS [cardName],
          'RC ' + CONVERT(nvarchar, rc.DocNum, 9) AS [docNumber],
          rc.JrnlMemo AS [journalMemo],
          ac.Segment_0 AS [acctCode],
          ac.AcctName AS [acctName],
          ct.CreditSum AS [amount],
          rc.DocDate AS [sortDate],
          rc.DocNum AS [sortNumber]
        FROM ORCT AS rc
        INNER JOIN RCT3 AS ct ON rc.DocEntry = ct.DocNum
        INNER JOIN OACT AS ac ON ct.CreditAcct = ac.AcctCode
        WHERE rc.DocDate >= @st
          AND rc.DocDate <= @ed

        UNION ALL

        SELECT
          'PS' AS [type],
          'SHOP' AS [channel],
          CONVERT(NVARCHAR, rc.DocDate, 103) AS [docDate],
          '100' + SUBSTRING(rc.CardCode, 4, 3) AS [brand],
          RIGHT(rc.CardCode, 8) AS [whsCode],
          rc.CardCode AS [cardCode],
          rc.CardName AS [cardName],
          'PS ' + CONVERT(nvarchar, rc.DocNum, 9) AS [docNumber],
          rc.JrnlMemo AS [journalMemo],
          ac.Segment_0 AS [acctCode],
          ac.AcctName AS [acctName],
          ct.CreditSum * -1 AS [amount],
          rc.DocDate AS [sortDate],
          rc.DocNum AS [sortNumber]
        FROM OVPM AS rc
        INNER JOIN VPM3 AS ct ON rc.DocEntry = ct.DocNum
        INNER JOIN OACT AS ac ON ct.CreditAcct = ac.AcctCode
        WHERE rc.DocDate >= @st
          AND rc.DocDate <= @ed

        ORDER BY [sortDate], [sortNumber]
      `)

    response.json({
      rows: result.recordset.map(({ sortDate, sortNumber, ...row }, index) => ({
        id: `${row.type}-${sortNumber}-${index}`,
        ...row,
      })),
    })
  } catch (error) {
    response.status(500).json({
      message: 'Cannot load Incoming Cash data from SAP database',
      detail: error.message,
    })
  }
})

app.listen(port, () => {
  console.log(`APPSAPB1 API running at http://127.0.0.1:${port}`)
})
