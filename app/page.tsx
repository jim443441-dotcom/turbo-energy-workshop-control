'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { isSupabaseConfigured, supabase } from '@/lib/supabase'
import { addOfflineAction, getOfflineQueue, readSnapshot, saveSnapshot, setOfflineQueue } from '@/lib/offline'

type Role = 'admin' | 'user'
type Session = { username: string; role: Role }
type Row = Record<string, any>
type TableName =
  | 'fleet_machines'
  | 'breakdowns'
  | 'repairs'
  | 'services'
  | 'spares'
  | 'spares_orders'
  | 'personnel'
  | 'work_logs'
  | 'tyres'
  | 'batteries'
  | 'photo_logs'

type AppData = Record<TableName, Row[]>

const TABLES: TableName[] = [
  'fleet_machines',
  'breakdowns',
  'repairs',
  'services',
  'spares',
  'spares_orders',
  'personnel',
  'work_logs',
  'tyres',
  'batteries',
  'photo_logs'
]

const emptyData: AppData = {
  fleet_machines: [],
  breakdowns: [],
  repairs: [],
  services: [],
  spares: [],
  spares_orders: [],
  personnel: [],
  work_logs: [],
  tyres: [],
  batteries: [],
  photo_logs: []
}

const loginUsers = [
  { username: 'admin', password: 'admin123', role: 'admin' as Role },
  { username: 'admin workshop', password: 'workshop123', role: 'admin' as Role },
  { username: 'foreman', password: 'foreman123', role: 'user' as Role },
  { username: 'user', password: 'user123', role: 'user' as Role }
]

const nav = [
  ['dashboard', 'Dashboard'],
  ['fleet', 'Fleet Register'],
  ['breakdowns', 'Breakdowns'],
  ['repairs', 'Repairs'],
  ['services', 'Services'],
  ['spares', 'Spares & Orders'],
  ['personnel', 'Personnel'],
  ['tyres', 'Tyres'],
  ['batteries', 'Batteries'],
  ['photos', 'Photos'],
  ['reports', 'Reports'],
  ['admin', 'Admin']
]

const nowIso = () => new Date().toISOString()
const today = () => new Date().toISOString().slice(0, 10)
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
const n = (value: any) => Number(value || 0)
const one = (value: any) => Math.round(n(value) * 10) / 10

function lowerKey(value: string) {
  return String(value || '')
    .toLowerCase()
    .replace(/[\s_\-\.\/]+/g, '')
    .trim()
}

function getCell(row: Row, names: string[]) {
  const map: Row = {}
  Object.keys(row || {}).forEach((key) => {
    map[lowerKey(key)] = row[key]
  })
  for (const name of names) {
    const value = map[lowerKey(name)]
    if (value !== undefined && value !== null && String(value).trim() !== '') return value
  }
  return ''
}

function cleanRow(row: Row) {
  const output: Row = {}
  Object.entries(row).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') output[key] = value
  })
  return output
}

function statusClass(status: string) {
  const s = String(status || '').toLowerCase()
  if (s.includes('available') || s.includes('online') || s.includes('complete') || s.includes('in stock')) return 'good'
  if (s.includes('due') || s.includes('order') || s.includes('pending') || s.includes('progress')) return 'warn'
  if (s.includes('down') || s.includes('offline') || s.includes('critical') || s.includes('repair')) return 'bad'
  return 'neutral'
}

function readExcelRows(file: File): Promise<Row[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        resolve(XLSX.utils.sheet_to_json(sheet, { defval: '' }))
      } catch (error) {
        reject(error)
      }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

function mapFleet(row: Row): Row {
  const hours = getCell(row, ['hours', 'current hours', 'hour meter', 'hmr', 'machine hours'])
  const lastService = getCell(row, ['last service hours', 'last service', 'last service hmr'])
  const interval = getCell(row, ['service interval', 'interval', 'service interval hours']) || 250
  return cleanRow({
    id: uid(),
    fleet_no: String(getCell(row, ['fleet', 'fleet no', 'fleet number', 'unit', 'unit no', 'machine', 'machine no', 'equipment']) || '').trim(),
    machine_type: String(getCell(row, ['type', 'machine type', 'equipment type', 'category']) || '').trim(),
    make_model: String(getCell(row, ['make', 'model', 'make model', 'description']) || '').trim(),
    reg_no: String(getCell(row, ['reg', 'registration', 'registration no', 'plate']) || '').trim(),
    department: String(getCell(row, ['department', 'section', 'allocation']) || '').trim(),
    location: String(getCell(row, ['location', 'site', 'area']) || '').trim(),
    hours: one(hours),
    mileage: one(getCell(row, ['mileage', 'km', 'kilometres'])),
    status: String(getCell(row, ['status', 'availability', 'online status']) || 'Available').trim(),
    service_interval_hours: one(interval),
    last_service_hours: one(lastService),
    next_service_hours: one(lastService) + one(interval),
    notes: String(getCell(row, ['notes', 'remarks', 'comment']) || '').trim(),
    updated_at: nowIso()
  })
}

function mapService(row: Row): Row {
  return cleanRow({
    id: uid(),
    fleet_no: String(getCell(row, ['fleet', 'fleet no', 'machine', 'machine no', 'unit']) || '').trim(),
    service_type: String(getCell(row, ['service type', 'type', 'service']) || 'Scheduled service').trim(),
    scheduled_hours: one(getCell(row, ['scheduled hours', 'due hours', 'next service hours', 'hmr'])),
    due_date: String(getCell(row, ['due date', 'date', 'service date']) || '').trim(),
    completed_date: '',
    completed_hours: 0,
    status: String(getCell(row, ['status']) || 'Due').trim(),
    technician: String(getCell(row, ['technician', 'fitter', 'mechanic']) || '').trim(),
    notes: String(getCell(row, ['notes', 'remarks']) || '').trim(),
    created_at: nowIso()
  })
}

function SectionTitle({ title, subtitle }: { title: string; subtitle?: string }) {
  return (
    <div className="section-title">
      <div>
        <h2>{title}</h2>
        {subtitle && <p>{subtitle}</p>}
      </div>
    </div>
  )
}

function StatCard({ label, value, note }: { label: string; value: any; note?: string }) {
  return (
    <div className="stat-card">
      <span>{label}</span>
      <strong>{value}</strong>
      {note && <small>{note}</small>}
    </div>
  )
}

function Badge({ value }: { value: string }) {
  return <span className={`badge ${statusClass(value)}`}>{value || 'Open'}</span>
}

function MiniTable({ rows, columns, empty }: { rows: Row[]; columns: { key: string; label: string }[]; empty?: string }) {
  if (!rows.length) return <div className="empty">{empty || 'No records yet.'}</div>
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>{columns.map((col) => <th key={col.key}>{col.label}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={row.id || i}>
              {columns.map((col) => (
                <td key={col.key}>{col.key === 'status' ? <Badge value={row[col.key]} /> : String(row[col.key] ?? '')}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export default function Home() {
  const [session, setSession] = useState<Session | null>(null)
  const [loginName, setLoginName] = useState('admin')
  const [loginPass, setLoginPass] = useState('admin123')
  const [tab, setTab] = useState('dashboard')
  const [data, setData] = useState<AppData>(emptyData)
  const [online, setOnline] = useState(true)
  const [queueCount, setQueueCount] = useState(0)
  const [message, setMessage] = useState('')
  const [search, setSearch] = useState('')
  const [reportFleet, setReportFleet] = useState('')

  const [fleetForm, setFleetForm] = useState<Row>({ fleet_no: '', machine_type: '', make_model: '', department: '', location: '', hours: 0, status: 'Available', service_interval_hours: 250, last_service_hours: 0 })
  const [breakdownForm, setBreakdownForm] = useState<Row>({ fleet_no: '', category: '', fault: '', reported_by: '', assigned_to: '', start_time: '', status: 'Open', cause: '', action_taken: '', delay_reason: '' })
  const [repairForm, setRepairForm] = useState<Row>({ fleet_no: '', job_card_no: '', fault: '', assigned_to: '', start_time: '', end_time: '', status: 'In progress', parts_used: '', notes: '' })
  const [serviceForm, setServiceForm] = useState<Row>({ fleet_no: '', service_type: '250hr service', scheduled_hours: 0, due_date: '', status: 'Due', technician: '', notes: '' })
  const [spareForm, setSpareForm] = useState<Row>({ part_no: '', description: '', section: '', machine_type: '', stock_qty: 0, min_qty: 1, supplier: '', lead_time_days: 0, order_status: 'In stock', notes: '' })
  const [orderForm, setOrderForm] = useState<Row>({ part_no: '', description: '', qty: 1, requested_by: '', priority: 'Normal', status: 'Requested', machine_fleet_no: '', notes: '' })
  const [personForm, setPersonForm] = useState<Row>({ name: '', section: '', role: '', phone: '', employment_status: 'Active', leave_start: '', leave_end: '', leave_reason: '' })
  const [workForm, setWorkForm] = useState<Row>({ person_name: '', fleet_no: '', section: '', work_date: today(), hours_worked: 0, work_type: 'Repair', notes: '' })
  const [tyreForm, setTyreForm] = useState<Row>({ serial_no: '', fleet_no: '', position: '', brand: '', size: '', fitment_date: today(), fitment_hours: 0, expected_life_hours: 5000, status: 'Fitted', notes: '' })
  const [batteryForm, setBatteryForm] = useState<Row>({ serial_no: '', fleet_no: '', volts: '12V', cca: '', fitment_date: today(), warranty_end: '', expected_life_months: 18, status: 'Fitted', notes: '' })
  const [photoForm, setPhotoForm] = useState<Row>({ fleet_no: '', linked_type: 'Breakdown', caption: '', image_data: '' })

  const isAdmin = session?.role === 'admin'

  useEffect(() => {
    const savedSession = localStorage.getItem('turbo-workshop-session')
    if (savedSession) setSession(JSON.parse(savedSession))
    setOnline(typeof navigator === 'undefined' ? true : navigator.onLine)
    setQueueCount(getOfflineQueue().length)
    const snapshot = readSnapshot<AppData>()
    if (snapshot) setData(snapshot)
    loadAll()

    const up = () => {
      setOnline(true)
      syncOfflineQueue()
    }
    const down = () => setOnline(false)
    window.addEventListener('online', up)
    window.addEventListener('offline', down)
    return () => {
      window.removeEventListener('online', up)
      window.removeEventListener('offline', down)
    }
  }, [])

  useEffect(() => {
    saveSnapshot(data)
  }, [data])

  async function loadAll() {
    if (!supabase || !isSupabaseConfigured || (typeof navigator !== 'undefined' && !navigator.onLine)) return
    const next: AppData = { ...emptyData }
    for (const table of TABLES) {
      const orderCol = table === 'fleet_machines' ? 'fleet_no' : 'created_at'
      const { data: rows, error } = await supabase.from(table).select('*').order(orderCol, { ascending: false })
      if (!error && rows) next[table] = rows as Row[]
    }
    setData(next)
    saveSnapshot(next)
  }

  async function syncOfflineQueue() {
    if (!supabase || !isSupabaseConfigured || !navigator.onLine) return
    const queue = getOfflineQueue()
    if (!queue.length) return
    const remaining = []
    for (const action of queue) {
      try {
        if (action.type === 'delete') {
          const { error } = await supabase.from(action.table).delete().eq('id', action.payload.id)
          if (error) throw error
        } else if (action.type === 'upsert') {
          const { error } = await supabase.from(action.table).upsert(action.payload, { onConflict: action.table === 'fleet_machines' ? 'fleet_no' : 'id' })
          if (error) throw error
        } else {
          const { error } = await supabase.from(action.table).insert(action.payload)
          if (error) throw error
        }
      } catch {
        remaining.push(action)
      }
    }
    setOfflineQueue(remaining)
    setQueueCount(remaining.length)
    await loadAll()
    setMessage(remaining.length ? `${remaining.length} offline records still waiting to sync.` : 'Offline records synced successfully.')
  }

  async function saveRow(table: TableName, row: Row, mode: 'insert' | 'upsert' = 'insert') {
    const payload = cleanRow({ id: row.id || uid(), ...row, created_at: row.created_at || nowIso(), updated_at: nowIso() })
    setData((prev) => {
      const existing = prev[table] || []
      const key = table === 'fleet_machines' ? 'fleet_no' : 'id'
      const found = existing.some((item) => item[key] && item[key] === payload[key])
      const rows = found ? existing.map((item) => (item[key] === payload[key] ? { ...item, ...payload } : item)) : [payload, ...existing]
      return { ...prev, [table]: rows }
    })

    if (!supabase || !isSupabaseConfigured || !navigator.onLine) {
      addOfflineAction({ table, type: mode, payload })
      setQueueCount(getOfflineQueue().length)
      setMessage('Saved offline. It will sync when internet is back.')
      return
    }

    if (mode === 'upsert') {
      const { error } = await supabase.from(table).upsert(payload, { onConflict: table === 'fleet_machines' ? 'fleet_no' : 'id' })
      if (error) {
        addOfflineAction({ table, type: mode, payload })
        setMessage(`Saved locally because Supabase rejected it: ${error.message}`)
      } else setMessage('Saved online successfully.')
    } else {
      const { error } = await supabase.from(table).insert(payload)
      if (error) {
        addOfflineAction({ table, type: mode, payload })
        setMessage(`Saved locally because Supabase rejected it: ${error.message}`)
      } else setMessage('Saved online successfully.')
    }
    setQueueCount(getOfflineQueue().length)
  }

  async function deleteRow(table: TableName, id: string) {
    if (!isAdmin) return setMessage('Only admin can delete records.')
    setData((prev) => ({ ...prev, [table]: prev[table].filter((row) => row.id !== id) }))
    if (!supabase || !isSupabaseConfigured || !navigator.onLine) {
      addOfflineAction({ table, type: 'delete', payload: { id } })
      setQueueCount(getOfflineQueue().length)
      return
    }
    await supabase.from(table).delete().eq('id', id)
  }

  function handleLogin() {
    const user = loginUsers.find((u) => u.username.toLowerCase() === loginName.toLowerCase().trim() && u.password === loginPass)
    if (!user) return setMessage('Wrong username or password.')
    const next = { username: user.username, role: user.role }
    setSession(next)
    localStorage.setItem('turbo-workshop-session', JSON.stringify(next))
    setMessage('Login successful.')
  }

  function logout() {
    localStorage.removeItem('turbo-workshop-session')
    setSession(null)
  }

  async function uploadFleet(file?: File) {
    if (!file) return
    const rows = await readExcelRows(file)
    const mapped = rows.map(mapFleet).filter((row) => row.fleet_no)
    for (const row of mapped) await saveRow('fleet_machines', row, 'upsert')
    setMessage(`${mapped.length} fleet machines uploaded.`)
  }

  async function uploadService(file?: File) {
    if (!file) return
    const rows = await readExcelRows(file)
    const mapped = rows.map(mapService).filter((row) => row.fleet_no)
    for (const row of mapped) await saveRow('services', row, 'insert')
    setMessage(`${mapped.length} service schedule records uploaded.`)
  }

  function attachPhoto(file?: File) {
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => setPhotoForm((prev) => ({ ...prev, image_data: String(reader.result || ''), file_name: file.name }))
    reader.readAsDataURL(file)
  }

  const filteredFleet = useMemo(() => {
    const q = search.toLowerCase()
    return data.fleet_machines.filter((m) => [m.fleet_no, m.machine_type, m.department, m.location, m.status].join(' ').toLowerCase().includes(q))
  }, [data.fleet_machines, search])

  const openBreakdowns = data.breakdowns.filter((b) => !String(b.status || '').toLowerCase().includes('complete') && !String(b.status || '').toLowerCase().includes('closed'))
  const repairsOpen = data.repairs.filter((r) => !String(r.status || '').toLowerCase().includes('complete') && !String(r.status || '').toLowerCase().includes('closed'))
  const servicesDue = data.services.filter((s) => String(s.status || '').toLowerCase().includes('due') || n(s.scheduled_hours) <= n(data.fleet_machines.find((m) => m.fleet_no === s.fleet_no)?.hours))
  const sparesLow = data.spares.filter((s) => n(s.stock_qty) <= n(s.min_qty))
  const tyresDue = data.tyres.filter((t) => {
    const machine = data.fleet_machines.find((m) => m.fleet_no === t.fleet_no)
    return machine && n(machine.hours) >= n(t.fitment_hours) + n(t.expected_life_hours)
  })
  const batteriesDue = data.batteries.filter((b) => {
    if (!b.fitment_date) return false
    const fitted = new Date(b.fitment_date)
    const due = new Date(fitted)
    due.setMonth(due.getMonth() + n(b.expected_life_months || 18))
    return new Date() >= due
  })

  const reportRows = useMemo(() => {
    const fleet = reportFleet.trim().toLowerCase()
    const match = (row: Row) => !fleet || String(row.fleet_no || '').toLowerCase().includes(fleet)
    return {
      breakdowns: data.breakdowns.filter(match),
      repairs: data.repairs.filter(match),
      services: data.services.filter(match),
      tyres: data.tyres.filter(match),
      batteries: data.batteries.filter(match),
      work: data.work_logs.filter(match)
    }
  }, [data, reportFleet])

  if (!session) {
    return (
      <main className="login-shell">
        <div className="login-card">
          <div className="brand-mark">TE</div>
          <h1>Turbo Energy Workshop Control</h1>
          <p>Fleet register, breakdowns, repairs, services, spares, tyres, batteries, personnel and records.</p>
          <label>Username</label>
          <input value={loginName} onChange={(e) => setLoginName(e.target.value)} />
          <label>Password</label>
          <input type="password" value={loginPass} onChange={(e) => setLoginPass(e.target.value)} />
          <button className="primary" onClick={handleLogin}>Sign in</button>
          <div className="login-help">
            <b>Starter logins:</b><br />
            admin / admin123<br />
            admin workshop / workshop123<br />
            foreman / foreman123<br />
            user / user123
          </div>
          {message && <div className="message">{message}</div>}
        </div>
      </main>
    )
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark small">TE</div>
          <div>
            <h1>Turbo Energy Workshop Control</h1>
            <p>Live and offline workshop tracking system</p>
          </div>
        </div>
        <div className="status-strip">
          <span className={online ? 'dot online' : 'dot offline'}></span>
          {online ? 'Online' : 'Offline'}
          <span>Queue: {queueCount}</span>
          <span>{session.username} ({session.role})</span>
          <button className="ghost" onClick={logout}>Logout</button>
        </div>
      </header>

      <nav className="tabs">
        {nav.map(([key, label]) => (
          <button key={key} className={tab === key ? 'active' : ''} onClick={() => setTab(key)}>{label}</button>
        ))}
      </nav>

      {message && <div className="message">{message}</div>}

      {tab === 'dashboard' && (
        <section className="panel">
          <SectionTitle title="Workshop Dashboard" subtitle="Main live view of fleet and workshop pressure points." />
          <div className="stats-grid">
            <StatCard label="Total Fleet" value={data.fleet_machines.length} note="Machines in register" />
            <StatCard label="Available / Online" value={data.fleet_machines.filter((m) => ['available', 'online'].some((x) => String(m.status || '').toLowerCase().includes(x))).length} note="Current active units" />
            <StatCard label="Open Breakdowns" value={openBreakdowns.length} note="Not completed" />
            <StatCard label="Open Repairs" value={repairsOpen.length} note="Workshop jobs" />
            <StatCard label="Services Due" value={servicesDue.length} note="By uploaded schedule" />
            <StatCard label="Spares Low" value={sparesLow.length} note="Below minimum stock" />
            <StatCard label="Tyres Due" value={tyresDue.length} note="Based on expected life" />
            <StatCard label="Batteries Due" value={batteriesDue.length} note="Based on fitment date" />
          </div>
          <div className="two-col">
            <div className="card">
              <h3>Machines currently down / under repair</h3>
              <MiniTable rows={[...openBreakdowns, ...repairsOpen].slice(0, 10)} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'fault', label: 'Fault' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'status', label: 'Status' }]} />
            </div>
            <div className="card">
              <h3>Stock, tyre and battery reminders</h3>
              <div className="reminder-list">
                {sparesLow.slice(0, 6).map((s) => <p key={s.id}><b>{s.part_no}</b> {s.description} stock {s.stock_qty} / min {s.min_qty}</p>)}
                {tyresDue.slice(0, 6).map((t) => <p key={t.id}><b>Tyre {t.serial_no}</b> due on {t.fleet_no}</p>)}
                {batteriesDue.slice(0, 6).map((b) => <p key={b.id}><b>Battery {b.serial_no}</b> due on {b.fleet_no}</p>)}
                {!sparesLow.length && !tyresDue.length && !batteriesDue.length && <p>No stock, tyre or battery reminders.</p>}
              </div>
            </div>
          </div>
        </section>
      )}

      {tab === 'fleet' && (
        <section className="panel">
          <SectionTitle title="Fleet Register" subtitle="Upload Excel or add machines manually. Search by fleet number, department, status or type." />
          <div className="toolbar">
            <input className="search" placeholder="Search fleet..." value={search} onChange={(e) => setSearch(e.target.value)} />
            {isAdmin && <label className="upload">Upload Fleet Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadFleet(e.target.files?.[0])} /></label>}
          </div>
          {isAdmin && <div className="form-grid">
            <input placeholder="Fleet No" value={fleetForm.fleet_no} onChange={(e) => setFleetForm({ ...fleetForm, fleet_no: e.target.value })} />
            <input placeholder="Machine Type" value={fleetForm.machine_type} onChange={(e) => setFleetForm({ ...fleetForm, machine_type: e.target.value })} />
            <input placeholder="Make / Model" value={fleetForm.make_model} onChange={(e) => setFleetForm({ ...fleetForm, make_model: e.target.value })} />
            <input placeholder="Department" value={fleetForm.department} onChange={(e) => setFleetForm({ ...fleetForm, department: e.target.value })} />
            <input placeholder="Location" value={fleetForm.location} onChange={(e) => setFleetForm({ ...fleetForm, location: e.target.value })} />
            <input type="number" placeholder="Hours" value={fleetForm.hours} onChange={(e) => setFleetForm({ ...fleetForm, hours: e.target.value })} />
            <select value={fleetForm.status} onChange={(e) => setFleetForm({ ...fleetForm, status: e.target.value })}><option>Available</option><option>Online</option><option>Offline</option><option>Down</option><option>Under Repair</option></select>
            <input type="number" placeholder="Service interval" value={fleetForm.service_interval_hours} onChange={(e) => setFleetForm({ ...fleetForm, service_interval_hours: e.target.value })} />
            <input type="number" placeholder="Last service hours" value={fleetForm.last_service_hours} onChange={(e) => setFleetForm({ ...fleetForm, last_service_hours: e.target.value, next_service_hours: n(e.target.value) + n(fleetForm.service_interval_hours) })} />
            <button className="primary" onClick={() => saveRow('fleet_machines', { ...fleetForm, next_service_hours: n(fleetForm.last_service_hours) + n(fleetForm.service_interval_hours) }, 'upsert')}>Save Machine</button>
          </div>}
          <MiniTable rows={filteredFleet} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'machine_type', label: 'Type' }, { key: 'make_model', label: 'Make / Model' }, { key: 'department', label: 'Department' }, { key: 'location', label: 'Location' }, { key: 'hours', label: 'Hours' }, { key: 'next_service_hours', label: 'Next Service' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'breakdowns' && (
        <section className="panel">
          <SectionTitle title="Breakdowns" subtitle="Record faults, downtime, cause, delay and who is working on the machine." />
          <div className="form-grid">
            <input placeholder="Fleet No" value={breakdownForm.fleet_no} onChange={(e) => setBreakdownForm({ ...breakdownForm, fleet_no: e.target.value })} />
            <input placeholder="Category" value={breakdownForm.category} onChange={(e) => setBreakdownForm({ ...breakdownForm, category: e.target.value })} />
            <input placeholder="Fault reported" value={breakdownForm.fault} onChange={(e) => setBreakdownForm({ ...breakdownForm, fault: e.target.value })} />
            <input placeholder="Reported by" value={breakdownForm.reported_by} onChange={(e) => setBreakdownForm({ ...breakdownForm, reported_by: e.target.value })} />
            <input placeholder="Assigned to" value={breakdownForm.assigned_to} onChange={(e) => setBreakdownForm({ ...breakdownForm, assigned_to: e.target.value })} />
            <input type="datetime-local" value={breakdownForm.start_time} onChange={(e) => setBreakdownForm({ ...breakdownForm, start_time: e.target.value })} />
            <select value={breakdownForm.status} onChange={(e) => setBreakdownForm({ ...breakdownForm, status: e.target.value })}><option>Open</option><option>In progress</option><option>Awaiting spares</option><option>Complete</option></select>
            <input placeholder="Cause" value={breakdownForm.cause} onChange={(e) => setBreakdownForm({ ...breakdownForm, cause: e.target.value })} />
            <input placeholder="Action taken" value={breakdownForm.action_taken} onChange={(e) => setBreakdownForm({ ...breakdownForm, action_taken: e.target.value })} />
            <input placeholder="Delay reason" value={breakdownForm.delay_reason} onChange={(e) => setBreakdownForm({ ...breakdownForm, delay_reason: e.target.value })} />
            <button className="primary" onClick={() => saveRow('breakdowns', { ...breakdownForm, downtime_hours: breakdownForm.start_time ? one((Date.now() - new Date(breakdownForm.start_time).getTime()) / 36e5) : 0 })}>Save Breakdown</button>
          </div>
          <MiniTable rows={data.breakdowns} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'category', label: 'Category' }, { key: 'fault', label: 'Fault' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'downtime_hours', label: 'Hours Down' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'repairs' && (
        <section className="panel">
          <SectionTitle title="Repairs and Job Cards" subtitle="Record workshop repair jobs, job card number, parts used and finish details." />
          <div className="form-grid">
            <input placeholder="Fleet No" value={repairForm.fleet_no} onChange={(e) => setRepairForm({ ...repairForm, fleet_no: e.target.value })} />
            <input placeholder="Job Card No" value={repairForm.job_card_no} onChange={(e) => setRepairForm({ ...repairForm, job_card_no: e.target.value })} />
            <input placeholder="Fault / repair" value={repairForm.fault} onChange={(e) => setRepairForm({ ...repairForm, fault: e.target.value })} />
            <input placeholder="Assigned to" value={repairForm.assigned_to} onChange={(e) => setRepairForm({ ...repairForm, assigned_to: e.target.value })} />
            <input type="datetime-local" value={repairForm.start_time} onChange={(e) => setRepairForm({ ...repairForm, start_time: e.target.value })} />
            <input type="datetime-local" value={repairForm.end_time} onChange={(e) => setRepairForm({ ...repairForm, end_time: e.target.value })} />
            <select value={repairForm.status} onChange={(e) => setRepairForm({ ...repairForm, status: e.target.value })}><option>In progress</option><option>Awaiting spares</option><option>Complete</option></select>
            <input placeholder="Parts used" value={repairForm.parts_used} onChange={(e) => setRepairForm({ ...repairForm, parts_used: e.target.value })} />
            <input placeholder="Notes" value={repairForm.notes} onChange={(e) => setRepairForm({ ...repairForm, notes: e.target.value })} />
            <button className="primary" onClick={() => saveRow('repairs', { ...repairForm, hours_worked: repairForm.start_time && repairForm.end_time ? one((new Date(repairForm.end_time).getTime() - new Date(repairForm.start_time).getTime()) / 36e5) : 0 })}>Save Repair</button>
          </div>
          <MiniTable rows={data.repairs} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'fault', label: 'Repair' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'hours_worked', label: 'Hours' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'services' && (
        <section className="panel">
          <SectionTitle title="Service Schedule" subtitle="Upload an Excel service plan and the app will show due reminders." />
          <div className="toolbar">
            {isAdmin && <label className="upload">Upload Service Schedule<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadService(e.target.files?.[0])} /></label>}
          </div>
          <div className="form-grid">
            <input placeholder="Fleet No" value={serviceForm.fleet_no} onChange={(e) => setServiceForm({ ...serviceForm, fleet_no: e.target.value })} />
            <input placeholder="Service type" value={serviceForm.service_type} onChange={(e) => setServiceForm({ ...serviceForm, service_type: e.target.value })} />
            <input type="number" placeholder="Due hours" value={serviceForm.scheduled_hours} onChange={(e) => setServiceForm({ ...serviceForm, scheduled_hours: e.target.value })} />
            <input type="date" value={serviceForm.due_date} onChange={(e) => setServiceForm({ ...serviceForm, due_date: e.target.value })} />
            <select value={serviceForm.status} onChange={(e) => setServiceForm({ ...serviceForm, status: e.target.value })}><option>Due</option><option>Planned</option><option>Completed</option><option>Overdue</option></select>
            <input placeholder="Technician" value={serviceForm.technician} onChange={(e) => setServiceForm({ ...serviceForm, technician: e.target.value })} />
            <input placeholder="Notes" value={serviceForm.notes} onChange={(e) => setServiceForm({ ...serviceForm, notes: e.target.value })} />
            <button className="primary" onClick={() => saveRow('services', serviceForm)}>Save Service</button>
          </div>
          <MiniTable rows={data.services} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'scheduled_hours', label: 'Due Hours' }, { key: 'due_date', label: 'Due Date' }, { key: 'technician', label: 'Technician' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'spares' && (
        <section className="panel">
          <SectionTitle title="Spares Stock and Ordering" subtitle="See stock availability, minimum levels, order requests and lead time." />
          <div className="two-col">
            <div className="card">
              <h3>Stock item</h3>
              <div className="form-grid compact">
                <input placeholder="Part No" value={spareForm.part_no} onChange={(e) => setSpareForm({ ...spareForm, part_no: e.target.value })} />
                <input placeholder="Description" value={spareForm.description} onChange={(e) => setSpareForm({ ...spareForm, description: e.target.value })} />
                <input placeholder="Section" value={spareForm.section} onChange={(e) => setSpareForm({ ...spareForm, section: e.target.value })} />
                <input placeholder="Machine type" value={spareForm.machine_type} onChange={(e) => setSpareForm({ ...spareForm, machine_type: e.target.value })} />
                <input type="number" placeholder="Stock Qty" value={spareForm.stock_qty} onChange={(e) => setSpareForm({ ...spareForm, stock_qty: e.target.value })} />
                <input type="number" placeholder="Minimum Qty" value={spareForm.min_qty} onChange={(e) => setSpareForm({ ...spareForm, min_qty: e.target.value })} />
                <input placeholder="Supplier" value={spareForm.supplier} onChange={(e) => setSpareForm({ ...spareForm, supplier: e.target.value })} />
                <button className="primary" onClick={() => saveRow('spares', spareForm)}>Save Stock</button>
              </div>
            </div>
            <div className="card">
              <h3>Order request</h3>
              <div className="form-grid compact">
                <input placeholder="Part No" value={orderForm.part_no} onChange={(e) => setOrderForm({ ...orderForm, part_no: e.target.value })} />
                <input placeholder="Description" value={orderForm.description} onChange={(e) => setOrderForm({ ...orderForm, description: e.target.value })} />
                <input type="number" placeholder="Qty" value={orderForm.qty} onChange={(e) => setOrderForm({ ...orderForm, qty: e.target.value })} />
                <input placeholder="Machine Fleet No" value={orderForm.machine_fleet_no} onChange={(e) => setOrderForm({ ...orderForm, machine_fleet_no: e.target.value })} />
                <input placeholder="Requested by" value={orderForm.requested_by} onChange={(e) => setOrderForm({ ...orderForm, requested_by: e.target.value })} />
                <select value={orderForm.priority} onChange={(e) => setOrderForm({ ...orderForm, priority: e.target.value })}><option>Normal</option><option>Urgent</option><option>Critical</option></select>
                <button className="primary" onClick={() => saveRow('spares_orders', orderForm)}>Request Order</button>
              </div>
            </div>
          </div>
          <h3>Stock availability</h3>
          <MiniTable rows={data.spares} columns={[{ key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'section', label: 'Section' }, { key: 'stock_qty', label: 'Stock' }, { key: 'min_qty', label: 'Min' }, { key: 'order_status', label: 'Status' }]} />
          <h3>Order requests</h3>
          <MiniTable rows={data.spares_orders} columns={[{ key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'qty', label: 'Qty' }, { key: 'requested_by', label: 'Requested By' }, { key: 'priority', label: 'Priority' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'personnel' && (
        <section className="panel">
          <SectionTitle title="Personnel, Leave and Hours" subtitle="Break people into sections, record leave and summarise work hours by person and machine." />
          <div className="two-col">
            <div className="card">
              <h3>Personnel register</h3>
              <div className="form-grid compact">
                <input placeholder="Name" value={personForm.name} onChange={(e) => setPersonForm({ ...personForm, name: e.target.value })} />
                <input placeholder="Section" value={personForm.section} onChange={(e) => setPersonForm({ ...personForm, section: e.target.value })} />
                <input placeholder="Role" value={personForm.role} onChange={(e) => setPersonForm({ ...personForm, role: e.target.value })} />
                <input placeholder="Phone" value={personForm.phone} onChange={(e) => setPersonForm({ ...personForm, phone: e.target.value })} />
                <select value={personForm.employment_status} onChange={(e) => setPersonForm({ ...personForm, employment_status: e.target.value })}><option>Active</option><option>On leave</option><option>Suspended</option><option>Left company</option></select>
                <input type="date" value={personForm.leave_start} onChange={(e) => setPersonForm({ ...personForm, leave_start: e.target.value })} />
                <input type="date" value={personForm.leave_end} onChange={(e) => setPersonForm({ ...personForm, leave_end: e.target.value })} />
                <input placeholder="Leave reason" value={personForm.leave_reason} onChange={(e) => setPersonForm({ ...personForm, leave_reason: e.target.value })} />
                <button className="primary" onClick={() => saveRow('personnel', personForm)}>Save Person</button>
              </div>
            </div>
            <div className="card">
              <h3>Work hours log</h3>
              <div className="form-grid compact">
                <input placeholder="Person name" value={workForm.person_name} onChange={(e) => setWorkForm({ ...workForm, person_name: e.target.value })} />
                <input placeholder="Fleet No" value={workForm.fleet_no} onChange={(e) => setWorkForm({ ...workForm, fleet_no: e.target.value })} />
                <input placeholder="Section" value={workForm.section} onChange={(e) => setWorkForm({ ...workForm, section: e.target.value })} />
                <input type="date" value={workForm.work_date} onChange={(e) => setWorkForm({ ...workForm, work_date: e.target.value })} />
                <input type="number" placeholder="Hours worked" value={workForm.hours_worked} onChange={(e) => setWorkForm({ ...workForm, hours_worked: e.target.value })} />
                <select value={workForm.work_type} onChange={(e) => setWorkForm({ ...workForm, work_type: e.target.value })}><option>Repair</option><option>Service</option><option>Breakdown</option><option>Inspection</option><option>Tyre</option><option>Battery</option></select>
                <input placeholder="Notes" value={workForm.notes} onChange={(e) => setWorkForm({ ...workForm, notes: e.target.value })} />
                <button className="primary" onClick={() => saveRow('work_logs', { ...workForm, hours_worked: one(workForm.hours_worked) })}>Save Hours</button>
              </div>
            </div>
          </div>
          <MiniTable rows={data.personnel} columns={[{ key: 'name', label: 'Name' }, { key: 'section', label: 'Section' }, { key: 'role', label: 'Role' }, { key: 'phone', label: 'Phone' }, { key: 'leave_start', label: 'Leave Start' }, { key: 'leave_end', label: 'Leave End' }, { key: 'employment_status', label: 'Status' }]} />
          <h3>Hours worked</h3>
          <MiniTable rows={data.work_logs} columns={[{ key: 'person_name', label: 'Person' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'section', label: 'Section' }, { key: 'work_date', label: 'Date' }, { key: 'hours_worked', label: 'Hours' }, { key: 'work_type', label: 'Type' }]} />
        </section>
      )}

      {tab === 'tyres' && (
        <section className="panel">
          <SectionTitle title="Tyre Register" subtitle="Track tyre serial numbers, fitment dates, machine, position and expected replacement hours." />
          <div className="form-grid">
            <input placeholder="Tyre Serial No" value={tyreForm.serial_no} onChange={(e) => setTyreForm({ ...tyreForm, serial_no: e.target.value })} />
            <input placeholder="Fleet No" value={tyreForm.fleet_no} onChange={(e) => setTyreForm({ ...tyreForm, fleet_no: e.target.value })} />
            <input placeholder="Position" value={tyreForm.position} onChange={(e) => setTyreForm({ ...tyreForm, position: e.target.value })} />
            <input placeholder="Brand" value={tyreForm.brand} onChange={(e) => setTyreForm({ ...tyreForm, brand: e.target.value })} />
            <input placeholder="Size" value={tyreForm.size} onChange={(e) => setTyreForm({ ...tyreForm, size: e.target.value })} />
            <input type="date" value={tyreForm.fitment_date} onChange={(e) => setTyreForm({ ...tyreForm, fitment_date: e.target.value })} />
            <input type="number" placeholder="Fitment hours" value={tyreForm.fitment_hours} onChange={(e) => setTyreForm({ ...tyreForm, fitment_hours: e.target.value })} />
            <input type="number" placeholder="Expected life hours" value={tyreForm.expected_life_hours} onChange={(e) => setTyreForm({ ...tyreForm, expected_life_hours: e.target.value })} />
            <select value={tyreForm.status} onChange={(e) => setTyreForm({ ...tyreForm, status: e.target.value })}><option>Fitted</option><option>In stock</option><option>Removed</option><option>Scrapped</option><option>Needs replacement</option></select>
            <button className="primary" onClick={() => saveRow('tyres', { ...tyreForm, estimated_replacement_hours: n(tyreForm.fitment_hours) + n(tyreForm.expected_life_hours) })}>Save Tyre</button>
          </div>
          <MiniTable rows={data.tyres} columns={[{ key: 'serial_no', label: 'Serial' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'position', label: 'Position' }, { key: 'fitment_date', label: 'Fitted' }, { key: 'fitment_hours', label: 'Fit Hrs' }, { key: 'estimated_replacement_hours', label: 'Replace Hrs' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'batteries' && (
        <section className="panel">
          <SectionTitle title="Battery Register" subtitle="Track battery serial numbers, where fitted, condition, warranty and replacement reminders." />
          <div className="form-grid">
            <input placeholder="Battery Serial No" value={batteryForm.serial_no} onChange={(e) => setBatteryForm({ ...batteryForm, serial_no: e.target.value })} />
            <input placeholder="Fleet No" value={batteryForm.fleet_no} onChange={(e) => setBatteryForm({ ...batteryForm, fleet_no: e.target.value })} />
            <input placeholder="Volts" value={batteryForm.volts} onChange={(e) => setBatteryForm({ ...batteryForm, volts: e.target.value })} />
            <input placeholder="CCA" value={batteryForm.cca} onChange={(e) => setBatteryForm({ ...batteryForm, cca: e.target.value })} />
            <input type="date" value={batteryForm.fitment_date} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_date: e.target.value })} />
            <input type="date" value={batteryForm.warranty_end} onChange={(e) => setBatteryForm({ ...batteryForm, warranty_end: e.target.value })} />
            <input type="number" placeholder="Expected life months" value={batteryForm.expected_life_months} onChange={(e) => setBatteryForm({ ...batteryForm, expected_life_months: e.target.value })} />
            <select value={batteryForm.status} onChange={(e) => setBatteryForm({ ...batteryForm, status: e.target.value })}><option>Fitted</option><option>In stock</option><option>Charging</option><option>Removed</option><option>Scrapped</option><option>Needs replacement</option></select>
            <input placeholder="Notes" value={batteryForm.notes} onChange={(e) => setBatteryForm({ ...batteryForm, notes: e.target.value })} />
            <button className="primary" onClick={() => saveRow('batteries', batteryForm)}>Save Battery</button>
          </div>
          <MiniTable rows={data.batteries} columns={[{ key: 'serial_no', label: 'Serial' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'volts', label: 'Volts' }, { key: 'cca', label: 'CCA' }, { key: 'fitment_date', label: 'Fitted' }, { key: 'warranty_end', label: 'Warranty' }, { key: 'status', label: 'Status' }]} />
        </section>
      )}

      {tab === 'photos' && (
        <section className="panel">
          <SectionTitle title="Photos and Evidence" subtitle="Upload photos from phone or laptop against a machine, breakdown, tyre, battery or repair." />
          <div className="form-grid">
            <input placeholder="Fleet No" value={photoForm.fleet_no} onChange={(e) => setPhotoForm({ ...photoForm, fleet_no: e.target.value })} />
            <select value={photoForm.linked_type} onChange={(e) => setPhotoForm({ ...photoForm, linked_type: e.target.value })}><option>Breakdown</option><option>Repair</option><option>Service</option><option>Tyre</option><option>Battery</option><option>Inspection</option></select>
            <input placeholder="Caption" value={photoForm.caption} onChange={(e) => setPhotoForm({ ...photoForm, caption: e.target.value })} />
            <input type="file" accept="image/*" capture="environment" onChange={(e) => attachPhoto(e.target.files?.[0])} />
            <button className="primary" onClick={() => saveRow('photo_logs', photoForm)}>Save Photo</button>
          </div>
          <div className="photo-grid">
            {data.photo_logs.map((p) => (
              <div className="photo-card" key={p.id}>
                {p.image_data ? <img src={p.image_data} alt={p.caption || p.fleet_no} /> : <div className="empty-img">No image</div>}
                <b>{p.fleet_no}</b>
                <span>{p.linked_type}</span>
                <p>{p.caption}</p>
              </div>
            ))}
          </div>
        </section>
      )}

      {tab === 'reports' && (
        <section className="panel printable">
          <SectionTitle title="Reports" subtitle="Build a quick machine history report and print/save as PDF from the browser." />
          <div className="toolbar no-print">
            <input className="search" placeholder="Fleet number for report..." value={reportFleet} onChange={(e) => setReportFleet(e.target.value)} />
            <button className="primary" onClick={() => window.print()}>Print / Save PDF</button>
          </div>
          <div className="report-header">
            <h2>Machine History Report {reportFleet ? `- ${reportFleet}` : ''}</h2>
            <p>Generated: {new Date().toLocaleString()}</p>
          </div>
          <h3>Breakdowns</h3>
          <MiniTable rows={reportRows.breakdowns} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'fault', label: 'Fault' }, { key: 'cause', label: 'Cause' }, { key: 'downtime_hours', label: 'Down Hrs' }, { key: 'status', label: 'Status' }]} />
          <h3>Repairs</h3>
          <MiniTable rows={reportRows.repairs} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'fault', label: 'Repair' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'hours_worked', label: 'Hours' }, { key: 'status', label: 'Status' }]} />
          <h3>Services</h3>
          <MiniTable rows={reportRows.services} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'scheduled_hours', label: 'Due Hrs' }, { key: 'due_date', label: 'Due Date' }, { key: 'status', label: 'Status' }]} />
          <h3>Personnel Hours</h3>
          <MiniTable rows={reportRows.work} columns={[{ key: 'person_name', label: 'Person' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'work_date', label: 'Date' }, { key: 'hours_worked', label: 'Hours' }, { key: 'work_type', label: 'Type' }]} />
        </section>
      )}

      {tab === 'admin' && (
        <section className="panel">
          <SectionTitle title="Admin Controls" subtitle="Upload data, sync offline records, and manage admin-only actions." />
          {!isAdmin && <div className="message warn-msg">You are signed in as a user. Admin controls are locked.</div>}
          <div className="stats-grid">
            <StatCard label="Supabase" value={isSupabaseConfigured ? 'Configured' : 'Not set'} note="Uses env variables" />
            <StatCard label="Internet" value={online ? 'Online' : 'Offline'} note="Offline queue active" />
            <StatCard label="Pending Queue" value={queueCount} note="Records waiting to upload" />
            <StatCard label="Local Snapshot" value="Enabled" note="Works on phone/laptop browser" />
          </div>
          <div className="toolbar">
            <button className="primary" onClick={syncOfflineQueue}>Sync Offline Queue</button>
            <button className="ghost" onClick={loadAll}>Reload From Supabase</button>
            {isAdmin && <label className="upload">Upload Fleet Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadFleet(e.target.files?.[0])} /></label>}
            {isAdmin && <label className="upload">Upload Service Schedule<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadService(e.target.files?.[0])} /></label>}
          </div>
          <div className="card">
            <h3>Important</h3>
            <p>This starter login gate is practical for workshop testing. For strict company security, upgrade the login to Supabase Auth before giving access outside the workshop.</p>
            <p>Photos are stored as browser image data in the database for this quick version. For heavy daily photo use, move photos to Supabase Storage in the next upgrade.</p>
          </div>
        </section>
      )}
    </main>
  )
}
