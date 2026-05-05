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
  | 'parts_books'
  | 'personnel'
  | 'leave_records'
  | 'work_logs'
  | 'tyres'
  | 'tyre_measurements'
  | 'tyre_orders'
  | 'batteries'
  | 'battery_orders'
  | 'photo_logs'
  | 'operations'

type AppData = Record<TableName, Row[]>

const TABLES: TableName[] = [
  'fleet_machines',
  'breakdowns',
  'repairs',
  'services',
  'spares',
  'spares_orders',
  'parts_books',
  'personnel',
  'leave_records',
  'work_logs',
  'tyres',
  'tyre_measurements',
  'tyre_orders',
  'batteries',
  'battery_orders',
  'photo_logs',
  'operations'
]

const emptyData: AppData = {
  fleet_machines: [],
  breakdowns: [],
  repairs: [],
  services: [],
  spares: [],
  spares_orders: [],
  parts_books: [],
  personnel: [],
  leave_records: [],
  work_logs: [],
  tyres: [],
  tyre_measurements: [],
  tyre_orders: [],
  batteries: [],
  battery_orders: [],
  photo_logs: [],
  operations: []
}

const loginUsers = [
  { username: 'admin', password: 'admin123', role: 'admin' as Role },
  { username: 'admin workshop', password: 'workshop123', role: 'admin' as Role },
  { username: 'foreman', password: 'foreman123', role: 'user' as Role },
  { username: 'user', password: 'user123', role: 'user' as Role }
]

const nav = [
  ['dashboard', 'Dashboard'],
  ['fleet', 'Fleet'],
  ['breakdowns', 'Breakdowns'],
  ['repairs', 'Repairs'],
  ['services', 'Services'],
  ['spares', 'Spares & Orders'],
  ['personnel', 'Personnel'],
  ['tyres', 'Tyres'],
  ['batteries', 'Batteries'],
  ['operations', 'Operations'],
  ['photos', 'Photos'],
  ['reports', 'Reports'],
  ['admin', 'Admin']
]

const shiftNames = ['Shift 1', 'Shift 2', 'Shift 3']
const machineGroups = ['Truck', 'Yellow Machine', 'Light Vehicle', 'Plant', 'Bus', 'Trailer']
const workshopSections = ['Mechanical', 'Auto Electrical', 'Tyre Bay', 'Boiler Shop', 'Stores', 'Panel Beating', 'Wash Bay', 'Operations']
const orderStages = ['Requested', 'Awaiting Funding', 'Funded', 'Awaiting Delivery', 'Delivered', 'Cancelled']
const orderTypes = ['Not selected', 'Local Order', 'International Order', 'Quotation Required']
const priorities = ['Low', 'Normal', 'High', 'Critical']
const tyrePositions = ['FL', 'FR', 'R1L', 'R1R', 'R2L', 'R2R', 'R3L', 'R3R', 'Spare', 'Loader Front L', 'Loader Front R', 'Loader Rear L', 'Loader Rear R']
const batteryVoltages = ['12V', '24V', '36V', '48V', '60V', '72V']
const personnelRoles = ['Manager', 'Foreman', 'Chargehand', 'Diesel Plant Fitter', 'Assistant', 'Tyre Fitter', 'Boiler Maker', 'Auto Electrical', 'Panel Beater', 'Wash Bay Attendant', 'Driver', 'Diesel Mechanic']
const leaveTypes = ['Annual Leave', 'Sick Leave', 'Off Duty', 'Training', 'Unpaid Leave', 'AWOL']
const operationMachineTypes = ['Truck', 'ADT', 'FEL', 'Excavator', 'Dozer', 'Grader', 'Compactor', 'TLB', 'LDV', 'Bus']

const nowIso = () => new Date().toISOString()
const today = () => new Date().toISOString().slice(0, 10)
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
const n = (value: any) => Number(value || 0)
const one = (value: any) => Math.round(n(value) * 10) / 10

function dateOnly(value: any) {
  if (!value) return ''
  const d = new Date(value)
  if (Number.isNaN(d.getTime())) return String(value).slice(0, 10)
  return d.toISOString().slice(0, 10)
}

function daysInclusive(start: any, end: any) {
  if (!start || !end) return 0
  const s = new Date(start)
  const e = new Date(end)
  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return 0
  return Math.max(0, Math.floor((e.getTime() - s.getTime()) / 86400000) + 1)
}

function isBetweenDates(start: any, end: any, d = today()) {
  if (!start && !end) return false
  const t = new Date(d).getTime()
  const s = start ? new Date(start).getTime() : t
  const e = end ? new Date(end).getTime() : t
  return !Number.isNaN(s) && !Number.isNaN(e) && t >= s && t <= e
}

function isFutureDate(value: any) {
  if (!value) return false
  return new Date(value).getTime() > new Date(today()).getTime()
}

function isPastDate(value: any) {
  if (!value) return false
  return new Date(value).getTime() < new Date(today()).getTime()
}

function lowerKey(value: string) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '')
}

function getCell(row: Row, names: string[]) {
  const keys = Object.keys(row)
  for (const name of names) {
    const wanted = lowerKey(name)
    const found = keys.find((key) => lowerKey(key) === wanted)
    if (found !== undefined && row[found] !== undefined && row[found] !== null) return row[found]
  }
  return ''
}

function cleanRow(row: Row) {
  const next: Row = {}
  Object.entries(row).forEach(([key, value]) => {
    if (value === undefined || value === null) return
    if (typeof value === 'string' && value.trim() === '') return
    next[key] = value
  })
  return next
}

function statusClass(status: any) {
  const s = String(status || '').toLowerCase()
  if (s.includes('critical') || s.includes('open') || s.includes('break') || s.includes('awaiting') || s.includes('low') || s.includes('due')) return 'danger'
  if (s.includes('funded') || s.includes('progress') || s.includes('request') || s.includes('quote') || s.includes('leave')) return 'warn'
  if (s.includes('complete') || s.includes('delivered') || s.includes('available') || s.includes('active') || s.includes('stock') || s.includes('approved')) return 'good'
  return 'neutral'
}

function uniqueBy(rows: Row[], keyName: string) {
  const seen = new Set<string>()
  return rows.filter((row) => {
    const key = String(row[keyName] || '').trim()
    if (!key || seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function safeJsonParse(value: any, fallback: any) {
  try {
    if (!value) return fallback
    if (typeof value !== 'string') return value
    return JSON.parse(value)
  } catch {
    return fallback
  }
}

async function readExcelRows(file: File): Promise<Row[]> {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]
  return XLSX.utils.sheet_to_json(sheet, { defval: '' }) as Row[]
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result || ''))
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
}

function arrayBufferToText(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer)
  let output = ''
  const max = Math.min(bytes.length, 1800000)
  for (let i = 0; i < max; i++) {
    const code = bytes[i]
    output += code >= 32 && code <= 126 ? String.fromCharCode(code) : ' '
  }
  return output
}

function extractPartNumbers(text: string) {
  const matches = String(text || '').toUpperCase().match(/\b[A-Z0-9]{2,}[-/.]?[A-Z0-9]{2,}[-/.]?[A-Z0-9]{0,}\b/g) || []
  const banned = new Set(['HTTP', 'HTTPS', 'DATE', 'PAGE', 'PART', 'NUMBER', 'MODEL', 'SERIAL', 'PDF', 'WWW'])
  return Array.from(new Set(matches.map((x) => x.replace(/[.,;:]$/g, '')).filter((x) => x.length >= 5 && !banned.has(x)))).slice(0, 250)
}

function mapFleet(row: Row): Row {
  const fleet_no = String(getCell(row, ['fleet_no', 'fleet', 'fleet number', 'machine', 'unit']) || '').trim()
  return cleanRow({
    id: fleet_no || uid(),
    fleet_no,
    machine_type: getCell(row, ['machine_type', 'type', 'machine type']),
    make_model: getCell(row, ['make_model', 'make model', 'model']),
    reg_no: getCell(row, ['reg_no', 'registration', 'reg']),
    department: getCell(row, ['department', 'dept']),
    location: getCell(row, ['location', 'site']),
    hours: Number(getCell(row, ['hours', 'hm', 'hour meter']) || 0),
    mileage: Number(getCell(row, ['mileage', 'km', 'odometer']) || 0),
    status: getCell(row, ['status']) || 'Available',
    service_interval_hours: Number(getCell(row, ['service_interval_hours', 'service interval']) || 250),
    last_service_hours: Number(getCell(row, ['last_service_hours', 'last service']) || 0),
    next_service_hours: Number(getCell(row, ['next_service_hours', 'next service']) || 0),
    notes: getCell(row, ['notes'])
  })
}

function mapPersonnel(row: Row): Row {
  const employee_no = String(getCell(row, ['employee_no', 'employee number', 'emp no', 'clock no']) || '').trim()
  const name = String(getCell(row, ['name', 'employee', 'person']) || '').trim()
  return cleanRow({
    id: employee_no || name || uid(),
    employee_no,
    name,
    shift: getCell(row, ['shift']) || 'Shift 1',
    section: getCell(row, ['section', 'department']) || 'Workshop',
    role: getCell(row, ['role', 'trade', 'position']) || 'Diesel Plant Fitter',
    staff_group: getCell(row, ['staff_group', 'group']) || 'Workshop Staff',
    phone: getCell(row, ['phone', 'cell']),
    employment_status: getCell(row, ['employment_status', 'status']) || 'Active',
    leave_balance_days: Number(getCell(row, ['leave_balance_days', 'leave balance']) || 30),
    leave_taken_days: Number(getCell(row, ['leave_taken_days', 'leave taken']) || 0),
    leave_start: dateOnly(getCell(row, ['leave_start', 'leave start'])),
    leave_end: dateOnly(getCell(row, ['leave_end', 'leave end'])),
    leave_reason: getCell(row, ['leave_reason', 'leave reason']),
    notes: getCell(row, ['notes'])
  })
}

function mapService(row: Row): Row {
  return cleanRow({
    id: uid(),
    fleet_no: getCell(row, ['fleet_no', 'fleet', 'machine']),
    service_type: getCell(row, ['service_type', 'service type']) || '250hr service',
    due_date: dateOnly(getCell(row, ['due_date', 'due date'])),
    scheduled_hours: Number(getCell(row, ['scheduled_hours', 'due hours', 'hours']) || 0),
    completed_hours: Number(getCell(row, ['completed_hours', 'done hours']) || 0),
    job_card_no: getCell(row, ['job_card_no', 'job card']),
    supervisor: getCell(row, ['supervisor']),
    technician: getCell(row, ['technician', 'fitter']),
    status: getCell(row, ['status']) || 'Due',
    notes: getCell(row, ['notes'])
  })
}

function mapLeave(row: Row): Row {
  const start = dateOnly(getCell(row, ['start_date', 'leave_start', 'start']))
  const end = dateOnly(getCell(row, ['end_date', 'leave_end', 'end']))
  return cleanRow({
    id: uid(),
    employee_no: getCell(row, ['employee_no', 'emp no']),
    person_name: getCell(row, ['person_name', 'name', 'employee']),
    shift: getCell(row, ['shift']) || 'Shift 1',
    section: getCell(row, ['section']),
    role: getCell(row, ['role']),
    leave_type: getCell(row, ['leave_type', 'leave type']) || 'Annual Leave',
    start_date: start,
    end_date: end,
    days: Number(getCell(row, ['days']) || daysInclusive(start, end)),
    reason: getCell(row, ['reason']),
    status: getCell(row, ['status']) || 'Approved'
  })
}

function SectionTitle({ title, subtitle, right }: { title: string; subtitle?: string; right?: any }) {
  return (
    <div className="section-title">
      <div>
        <h2>{title}</h2>
        {subtitle && <p>{subtitle}</p>}
      </div>
      {right && <div className="section-actions">{right}</div>}
    </div>
  )
}

function Badge({ value }: { value: any }) {
  return <span className={`badge ${statusClass(value)}`}>{String(value || 'Not set')}</span>
}

function StatCard({ title, value, detail, tone = 'blue', onClick }: { title: string; value: any; detail?: string; tone?: string; onClick?: () => void }) {
  return (
    <button type="button" className={`stat-card ${tone}`} onClick={onClick}>
      <span>{title}</span>
      <b>{value}</b>
      {detail && <small>{detail}</small>}
    </button>
  )
}

function Field({ label, children }: { label: string; children?: any }) {
  return <label className="field"><span>{label}</span>{children}</label>
}

function MiniTable({ rows, columns, empty, onRow }: { rows: Row[]; columns: { key: string; label: string; render?: (row: Row) => any }[]; empty?: string; onRow?: (row: Row) => void }) {
  if (!rows.length) return <div className="empty">{empty || 'No records yet.'}</div>
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>{columns.map((col) => <th key={col.key}>{col.label}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id || JSON.stringify(row)} onClick={() => onRow?.(row)} className={onRow ? 'clickable-row' : ''}>
              {columns.map((col) => <td key={col.key}>{col.render ? col.render(row) : String(row[col.key] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function Modal({ title, children, onClose }: { title: string; children?: any; onClose: () => void }) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head"><h3>{title}</h3><button onClick={onClose}>Close</button></div>
        <div className="modal-body">{children}</div>
      </div>
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
  const [modal, setModal] = useState<{ title: string; body: any } | null>(null)
  const [reportFleet, setReportFleet] = useState('')

  const [fleetForm, setFleetForm] = useState<Row>({ fleet_no: '', machine_type: '', make_model: '', department: '', location: '', hours: 0, mileage: 0, status: 'Available', service_interval_hours: 250, last_service_hours: 0 })
  const [breakdownForm, setBreakdownForm] = useState<Row>({ fleet_no: '', category: 'Mechanical', fault: '', reported_by: '', assigned_to: '', start_time: '', status: 'Open', spare_eta: '', expected_available_date: '', cause: '', action_taken: '', preventative_recommendation: '' })
  const [repairForm, setRepairForm] = useState<Row>({ fleet_no: '', job_card_no: '', fault: '', assigned_to: '', start_time: '', end_time: '', status: 'In progress', parts_used: '', notes: '' })
  const [serviceForm, setServiceForm] = useState<Row>({ fleet_no: '', service_type: '250hr service', due_date: today(), scheduled_hours: 0, completed_date: '', completed_hours: 0, job_card_no: '', supervisor: '', technician: '', status: 'Due', notes: '' })
  const [spareForm, setSpareForm] = useState<Row>({ part_no: '', description: '', section: 'Stores', machine_group: 'Yellow Machine', machine_type: '', stock_qty: 0, min_qty: 1, supplier: '', shelf_location: '', lead_time_days: 0, order_status: 'In stock' })
  const [orderForm, setOrderForm] = useState<Row>({ request_date: today(), machine_fleet_no: '', machine_group: 'Yellow Machine', workshop_section: 'Mechanical', requested_by: '', part_no: '', description: '', qty: 1, spares_items: '', priority: 'Normal', workflow_stage: 'Requested', status: 'Requested', order_type: 'Not selected', funding_status: 'Not funded', eta: '', part_numbers: '', pdf_book_ref: '', notes: '' })
  const [partsBookForm, setPartsBookForm] = useState<Row>({ book_title: '', machine_group: 'Yellow Machine', machine_type: '', uploaded_by: '', file_name: '', mime_type: '', file_size: 0, file_data: '', extracted_part_numbers: '', extracted_text: '', notes: '' })
  const [sparesSearch, setSparesSearch] = useState('')
  const [sparesStage, setSparesStage] = useState('All')
  const [personForm, setPersonForm] = useState<Row>({ employee_no: '', name: '', shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', staff_group: 'Workshop Staff', phone: '', employment_status: 'Active', leave_balance_days: 30, leave_taken_days: 0 })
  const [leaveForm, setLeaveForm] = useState<Row>({ person_name: '', employee_no: '', leave_type: 'Annual Leave', start_date: today(), end_date: today(), reason: '', status: 'Approved' })
  const [bulkPersonForm, setBulkPersonForm] = useState<Row>({ shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', qty: 1 })
  const [workForm, setWorkForm] = useState<Row>({ person_name: '', fleet_no: '', section: 'Mechanical', work_date: today(), hours_worked: 0, work_type: 'Repair', notes: '' })
  const [tyreForm, setTyreForm] = useState<Row>({ make: '', serial_no: '', company_no: '', production_date: '', fitted_date: today(), fleet_no: '', fitted_by: '', position: 'FL', tyre_type: '', size: '', fitment_hours: 0, fitment_mileage: 0, current_tread_mm: 0, measurement_due_date: '', stock_status: 'Fitted', status: 'Fitted', fitment_photo: '', notes: '' })
  const [tyreMeasurementForm, setTyreMeasurementForm] = useState<Row>({ tyre_serial_no: '', fleet_no: '', position: '', measurement_date: today(), measured_by: '', tread_depth_mm: 0, pressure_psi: 0, condition: 'Good', notes: '' })
  const [tyreOrderForm, setTyreOrderForm] = useState<Row>({ request_date: today(), fleet_no: '', tyre_type: '', size: '', qty: 1, requested_by: '', reason: '', priority: 'Normal', status: 'Requested', quote_status: 'Need Quote', eta: '', notes: '' })
  const [batteryForm, setBatteryForm] = useState<Row>({ make: '', serial_no: '', production_date: '', fitment_date: today(), fleet_no: '', fitted_by: '', charging_voltage: '24V', volts: '12V', fitment_hours: 0, fitment_mileage: 0, maintenance_due_date: '', stock_status: 'Fitted', status: 'Fitted', fitment_photo: '', notes: '' })
  const [batteryOrderForm, setBatteryOrderForm] = useState<Row>({ request_date: today(), fleet_no: '', battery_type: '', voltage: '12V', qty: 1, requested_by: '', reason: '', priority: 'Normal', status: 'Requested', eta: '', notes: '' })
  const [photoForm, setPhotoForm] = useState<Row>({ fleet_no: '', linked_type: 'Breakdown', caption: '', image_data: '' })
  const [operationForm, setOperationForm] = useState<Row>({ operator_name: '', employee_no: '', machine_group: 'Truck', machine_type: 'Truck', fleet_no: '', shift: 'Shift 1', score: 100, offence: '', damage_report: '', supervisor: '', event_date: today(), photo_data: '', notes: '' })

  const isAdmin = session?.role === 'admin'

  useEffect(() => {
    const savedSession = localStorage.getItem('turbo-workshop-session')
    if (savedSession) setSession(JSON.parse(savedSession))
    setOnline(typeof navigator === 'undefined' ? true : navigator.onLine)
    setQueueCount(getOfflineQueue().length)
    const snapshot = readSnapshot<AppData>()
    if (snapshot) setData({ ...emptyData, ...snapshot })
    loadAll()

    const up = () => { setOnline(true); syncOfflineQueue() }
    const down = () => setOnline(false)
    window.addEventListener('online', up)
    window.addEventListener('offline', down)
    return () => {
      window.removeEventListener('online', up)
      window.removeEventListener('offline', down)
    }
  }, [])

  useEffect(() => { saveSnapshot(data) }, [data])

  async function loadAll() {
    if (!supabase || !isSupabaseConfigured || (typeof navigator !== 'undefined' && !navigator.onLine)) return
    const next: AppData = { ...emptyData }
    for (const table of TABLES) {
      const orderCol = table === 'fleet_machines' ? 'fleet_no' : 'created_at'
      const { data: rows } = await supabase.from(table).select('*').order(orderCol, { ascending: false })
      if (rows) next[table] = rows as Row[]
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
    setMessage(remaining.length ? `${remaining.length} records still waiting to sync.` : 'Offline records synced successfully.')
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

    const request = mode === 'upsert'
      ? supabase.from(table).upsert(payload, { onConflict: table === 'fleet_machines' ? 'fleet_no' : 'id' })
      : supabase.from(table).insert(payload)
    const { error } = await request
    if (error) {
      addOfflineAction({ table, type: mode, payload })
      setQueueCount(getOfflineQueue().length)
      setMessage(`Saved locally because Supabase rejected it: ${error.message}`)
    } else {
      setMessage('Saved online successfully.')
    }
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
    setMessage(`${mapped.length} service records uploaded.`)
  }

  async function uploadPersonnel(file?: File) {
    if (!file) return
    const rows = await readExcelRows(file)
    const people = rows.map(mapPersonnel).filter((row) => row.name)
    const leave = rows.map(mapLeave).filter((row) => row.person_name && row.start_date)
    for (const row of people) await saveRow('personnel', row, 'upsert')
    for (const row of leave) await saveRow('leave_records', row, 'insert')
    setMessage(`${people.length} personnel records and ${leave.length} leave records uploaded.`)
  }

  async function attachPhoto(file: File | undefined, setter: (value: string) => void) {
    if (!file) return
    const dataUrl = await fileToDataUrl(file)
    setter(dataUrl)
  }

  async function attachPartsBook(file?: File) {
    if (!file) return
    const buffer = await file.arrayBuffer()
    const dataUrl = await fileToDataUrl(file)
    const rawText = arrayBufferToText(buffer)
    const detected = extractPartNumbers(`${file.name} ${partsBookForm.notes || ''} ${rawText}`)
    setPartsBookForm((prev) => ({
      ...prev,
      file_name: file.name,
      mime_type: file.type || 'application/pdf',
      file_size: file.size,
      file_data: dataUrl,
      extracted_part_numbers: detected.join('\n'),
      extracted_count: detected.length,
      extracted_text: rawText.slice(0, 7000)
    }))
    setMessage(`${detected.length} possible part numbers found. You can still type or add part numbers manually.`)
  }

  function addPartToOrder(partNo: string) {
    const current = String(orderForm.part_numbers || '')
    const parts = new Set(current.split(/[\n,; ]+/).map((x) => x.trim()).filter(Boolean))
    parts.add(partNo)
    setOrderForm((prev) => ({ ...prev, part_numbers: Array.from(parts).join('\n'), part_no: prev.part_no || partNo }))
    setMessage(`${partNo} added to order sheet.`)
  }

  async function saveLeaveRecord() {
    const days = Number(leaveForm.days || daysInclusive(leaveForm.start_date, leaveForm.end_date))
    const payload = { ...leaveForm, days }
    await saveRow('leave_records', payload)
    const person = data.personnel.find((p) => String(p.name || '').toLowerCase() === String(leaveForm.person_name || '').toLowerCase() || String(p.employee_no || '') === String(leaveForm.employee_no || ''))
    if (person && !String(leaveForm.leave_type || '').toLowerCase().includes('off')) {
      await saveRow('personnel', { ...person, leave_taken_days: one(n(person.leave_taken_days) + days), leave_balance_days: one(n(person.leave_balance_days) - days), employment_status: isBetweenDates(leaveForm.start_date, leaveForm.end_date) ? 'On Leave' : person.employment_status }, 'upsert')
    }
  }

  async function addBulkPersonnel() {
    const qty = Math.max(1, Number(bulkPersonForm.qty || 1))
    for (let i = 1; i <= qty; i++) {
      await saveRow('personnel', {
        employee_no: `VAC-${bulkPersonForm.shift}-${bulkPersonForm.role}-${i}-${Date.now()}`,
        name: `Vacant ${bulkPersonForm.role} ${i}`,
        shift: bulkPersonForm.shift,
        section: bulkPersonForm.section,
        role: bulkPersonForm.role,
        staff_group: bulkPersonForm.role === 'Manager' || bulkPersonForm.role === 'Foreman' ? 'Management' : 'Workshop Staff',
        employment_status: 'Vacant',
        leave_balance_days: 0,
        leave_taken_days: 0
      })
    }
  }

  function moveOrder(table: TableName, order: Row, status: string, extra: Row = {}) {
    saveRow(table, { ...order, status, workflow_stage: status, ...extra }, 'upsert')
  }

  function showPdfBook(book: Row) {
    if (!book.file_data) {
      setMessage('This parts book was indexed before PDF viewing was added. Upload it again to open diagrams inside the app.')
      return
    }
    setModal({
      title: book.book_title || book.file_name || 'Parts book',
      body: (
        <div className="pdf-modal-grid">
          <div className="pdf-frame-wrap"><iframe src={book.file_data} title={book.book_title || 'Parts Book'} /></div>
          <div className="pdf-side">
            <h4>Extracted part numbers</h4>
            <div className="part-chip-box tall">
              {String(book.extracted_part_numbers || '').split(/[\n,; ]+/).map((p) => p.trim()).filter(Boolean).slice(0, 250).map((part) => (
                <button key={part} className="part-chip" onClick={() => addPartToOrder(part)}>{part}</button>
              ))}
            </div>
            <p className="muted">Click a part number to place it on the spares order sheet. Use the PDF diagram on the left to confirm item number and page.</p>
          </div>
        </div>
      )
    })
  }

  const openBreakdowns = data.breakdowns.filter((b) => !['closed', 'complete', 'available'].some((x) => String(b.status || '').toLowerCase().includes(x)))
  const openRepairs = data.repairs.filter((r) => !['closed', 'complete'].some((x) => String(r.status || '').toLowerCase().includes(x)))
  const servicesDue = data.services.filter((s) => !String(s.status || '').toLowerCase().includes('complete'))
  const servicesToday = servicesDue.filter((s) => dateOnly(s.due_date) === today())
  const sparesLow = data.spares.filter((s) => n(s.stock_qty) <= n(s.min_qty))
  const awaitingFunding = data.spares_orders.filter((o) => String(o.workflow_stage || o.status || '').toLowerCase().includes('awaiting funding'))
  const batteryRequests = data.battery_orders.filter((o) => !['delivered', 'closed', 'cancelled'].some((x) => String(o.status || '').toLowerCase().includes(x)))
  const tyreRequests = data.tyre_orders.filter((o) => !['delivered', 'closed', 'cancelled'].some((x) => String(o.status || '').toLowerCase().includes(x)))
  const currentLeave = data.leave_records.filter((l) => isBetweenDates(l.start_date, l.end_date))
  const upcomingLeave = data.leave_records.filter((l) => isFutureDate(l.start_date))
  const pastLeave = data.leave_records.filter((l) => isPastDate(l.end_date))
  const activePersonnel = data.personnel.filter((p) => String(p.employment_status || '').toLowerCase().includes('active'))
  const peopleOff = data.personnel.filter((p) => String(p.employment_status || '').toLowerCase().includes('leave') || String(p.employment_status || '').toLowerCase().includes('off') || currentLeave.some((l) => l.person_name === p.name))
  const foremenOnDuty = data.personnel.filter((p) => ['foreman', 'chargehand'].some((x) => String(p.role || '').toLowerCase().includes(x)) && !peopleOff.some((off) => off.id === p.id))
  const currentShift = activePersonnel.reduce((acc: Record<string, number>, p) => { acc[p.shift || 'Shift 1'] = (acc[p.shift || 'Shift 1'] || 0) + 1; return acc }, {})
  const partsInBooks = Array.from(new Set(data.parts_books.flatMap((b) => String(b.extracted_part_numbers || '').split(/[\n,; ]+/).map((x) => x.trim()).filter(Boolean)))).slice(0, 150)
  const filteredFleet = data.fleet_machines.filter((r) => !search || [r.fleet_no, r.machine_type, r.department, r.location, r.status].join(' ').toLowerCase().includes(search.toLowerCase()))
  const filteredSpareOrders = data.spares_orders.filter((o) => {
    const text = [o.machine_fleet_no, o.part_no, o.description, o.spares_items, o.requested_by, o.workflow_stage, o.status, o.order_type, o.machine_group].join(' ').toLowerCase()
    const qOk = !sparesSearch || text.includes(sparesSearch.toLowerCase())
    const sOk = sparesStage === 'All' || String(o.workflow_stage || o.status || '') === sparesStage || String(o.order_type || '') === sparesStage || String(o.machine_group || '') === sparesStage
    return qOk && sOk
  })
  const departmentReport = useMemo(() => {
    const groups: Record<string, { total: number; breakdowns: number; services: number }> = {}
    data.fleet_machines.forEach((m) => {
      const dep = m.department || 'Unallocated'
      groups[dep] ||= { total: 0, breakdowns: 0, services: 0 }
      groups[dep].total += 1
      groups[dep].breakdowns += openBreakdowns.some((b) => b.fleet_no === m.fleet_no) ? 1 : 0
      groups[dep].services += servicesDue.some((s) => s.fleet_no === m.fleet_no) ? 1 : 0
    })
    return Object.entries(groups).map(([department, vals]) => ({ id: department, department, ...vals, availability: vals.total ? one(((vals.total - vals.breakdowns) / vals.total) * 100) : 0 }))
  }, [data, openBreakdowns, servicesDue])
  const typeReport = useMemo(() => {
    const groups: Record<string, { total: number; breakdowns: number }> = {}
    data.fleet_machines.forEach((m) => {
      const type = m.machine_type || 'Unknown'
      groups[type] ||= { total: 0, breakdowns: 0 }
      groups[type].total += 1
      groups[type].breakdowns += openBreakdowns.some((b) => b.fleet_no === m.fleet_no) ? 1 : 0
    })
    return Object.entries(groups).map(([machine_type, vals]) => ({ id: machine_type, machine_type, ...vals, availability: vals.total ? one(((vals.total - vals.breakdowns) / vals.total) * 100) : 0 }))
  }, [data, openBreakdowns])
  const repeatedFaults = useMemo(() => {
    const counts: Record<string, number> = {}
    data.breakdowns.forEach((b) => {
      const key = `${b.fleet_no || 'NO FLEET'} · ${b.category || b.fault || 'Fault'}`
      counts[key] = (counts[key] || 0) + 1
    })
    return Object.entries(counts).filter(([, count]) => count > 1).map(([fault, count]) => ({ id: fault, fault, count, recommendation: 'Check root cause and add preventative maintenance task.' }))
  }, [data.breakdowns])
  const reportRows = {
    breakdowns: data.breakdowns.filter((r) => !reportFleet || String(r.fleet_no || '').toLowerCase().includes(reportFleet.toLowerCase())),
    services: data.services.filter((r) => !reportFleet || String(r.fleet_no || '').toLowerCase().includes(reportFleet.toLowerCase())),
    tyres: data.tyres.filter((r) => !reportFleet || String(r.fleet_no || '').toLowerCase().includes(reportFleet.toLowerCase())),
    batteries: data.batteries.filter((r) => !reportFleet || String(r.fleet_no || '').toLowerCase().includes(reportFleet.toLowerCase()))
  }

  function renderSparesOrderCards(rows: Row[]) {
    if (!rows.length) return <div className="empty">No orders in this section yet.</div>
    return <div className="kanban-cards">{rows.map((order) => (
      <div className="kanban-card" key={order.id}>
        <div className="card-row"><b>{order.machine_fleet_no || 'No fleet'}</b><Badge value={order.workflow_stage || order.status} /></div>
        <p>{order.part_no || 'Parts request'} · Qty {order.qty || 1}</p>
        {order.description && <small>{order.description}</small>}
        {order.spares_items && <pre>{order.spares_items}</pre>}
        <div className="meta-grid compact">
          <span>By: {order.requested_by || '-'}</span><span>Area: {order.machine_group || '-'}</span><span>ETA: {dateOnly(order.eta) || '-'}</span><span>{order.order_type || 'No order type'}</span>
        </div>
        <div className="button-row wrap">
          <button onClick={() => moveOrder('spares_orders', order, 'Awaiting Funding', { funding_status: 'Awaiting Funding' })}>Awaiting funding</button>
          <button onClick={() => moveOrder('spares_orders', order, 'Funded', { funding_status: 'Funded', funded_date: today() })}>Funded</button>
          <button onClick={() => { const eta = prompt('Delivery ETA date YYYY-MM-DD', dateOnly(order.eta) || today()); if (eta) moveOrder('spares_orders', order, 'Awaiting Delivery', { eta, workflow_stage: 'Awaiting Delivery' }) }}>Delivery ETA</button>
          <button onClick={() => moveOrder('spares_orders', order, order.workflow_stage || 'Requested', { order_type: 'Local Order' })}>Local</button>
          <button onClick={() => moveOrder('spares_orders', order, order.workflow_stage || 'Requested', { order_type: 'International Order' })}>International</button>
          <button onClick={() => moveOrder('spares_orders', order, 'Delivered', { delivered_date: today() })}>Delivered</button>
        </div>
      </div>
    ))}</div>
  }

  function renderPersonnelShift(shift: string) {
    const people = data.personnel.filter((p) => String(p.shift || 'Shift 1') === shift)
    return <details className="drop-panel" open={shift === 'Shift 1'} key={shift}>
      <summary><span>{shift}</span><Badge value={`${people.length} people / slots`} /></summary>
      <div className="role-count-grid">
        {personnelRoles.map((role) => <button key={role} type="button"><span>{role}</span><b>{people.filter((p) => String(p.role || '') === role).length}</b></button>)}
      </div>
      <MiniTable rows={people} columns={[{ key: 'employee_no', label: 'Emp No' }, { key: 'name', label: 'Name' }, { key: 'role', label: 'Role' }, { key: 'section', label: 'Section' }, { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }, { key: 'leave_balance_days', label: 'Leave Bal' }]} empty={`No people in ${shift}.`} />
    </details>
  }

  if (!session) {
    return <main className="login-shell">
      <div className="login-card glass">
        <div className="brand-mark">TE</div>
        <h1>Turbo Energy Workshop Control</h1>
        <p>Dark control system for fleet, breakdowns, spares, tyres, batteries, personnel, operations and reports.</p>
        <Field label="Username"><input value={loginName} onChange={(e) => setLoginName(e.target.value)} /></Field>
        <Field label="Password"><input type="password" value={loginPass} onChange={(e) => setLoginPass(e.target.value)} /></Field>
        <button className="primary wide" onClick={handleLogin}>Sign in</button>
        <div className="login-help">admin / admin123 · admin workshop / workshop123 · foreman / foreman123 · user / user123</div>
        {message && <div className="message">{message}</div>}
      </div>
    </main>
  }

  return <main className="app-shell">
    <aside className="sidebar">
      <div className="logo-block"><div className="brand-mark small">TE</div><div><b>Turbo Energy</b><span>Workshop Control</span></div></div>
      <nav>{nav.map(([key, label]) => <button key={key} className={tab === key ? 'active' : ''} onClick={() => setTab(key)}>{label}</button>)}</nav>
      <div className="sync-box"><Badge value={online ? 'Online' : 'Offline'} /><span>Queue: {queueCount}</span><button onClick={syncOfflineQueue}>Sync</button></div>
    </aside>

    <section className="workspace">
      <header className="topbar">
        <div><h1>{nav.find((x) => x[0] === tab)?.[1]}</h1><p>{today()} · Logged in as {session.username}</p></div>
        <div className="top-actions"><button onClick={() => loadAll()}>Refresh</button><button onClick={() => window.print()}>Print</button><button className="danger-btn" onClick={logout}>Logout</button></div>
      </header>
      {message && <div className="message floating">{message}</div>}

      {tab === 'dashboard' && <div className="page-grid">
        <SectionTitle title="Workshop control room" subtitle="Click any card to open the live list behind it." />
        <div className="stats-grid">
          <StatCard title="Total Fleet" value={data.fleet_machines.length} detail="machines registered" onClick={() => setTab('fleet')} />
          <StatCard title="Current Shift Working" value={Object.entries(currentShift).map(([s, c]) => `${s}: ${c}`).join(' · ') || 'No shift loaded'} detail={`${activePersonnel.length} active people`} tone="green" onClick={() => setTab('personnel')} />
          <StatCard title="People Off" value={peopleOff.length} detail="leave/off schedule" tone="orange" onClick={() => setTab('personnel')} />
          <StatCard title="Foremen On Duty" value={foremenOnDuty.length} detail="foremen + chargehands" tone="blue" onClick={() => setTab('personnel')} />
          <StatCard title="Service Today" value={servicesToday.length} detail="machines due today" tone="purple" onClick={() => setTab('services')} />
          <StatCard title="Breakdowns" value={uniqueBy(openBreakdowns, 'fleet_no').length} detail="machines down" tone="red" onClick={() => setTab('breakdowns')} />
          <StatCard title="Being Worked On" value={uniqueBy([...openRepairs, ...openBreakdowns], 'fleet_no').length} detail="repair/breakdown activity" tone="orange" onClick={() => setTab('repairs')} />
          <StatCard title="Admin Requests" value={awaitingFunding.length + tyreRequests.length + batteryRequests.length} detail="spares/tyres/batteries" tone="red" onClick={() => setTab('spares')} />
        </div>
        <div className="grid-2">
          <div className="card"><h3>Machines on breakdown</h3><MiniTable rows={openBreakdowns.slice(0, 8)} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'category', label: 'Fault Type' }, { key: 'assigned_to', label: 'Worked By' }, { key: 'spare_eta', label: 'Spares ETA' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} empty="No active breakdowns." /></div>
          <div className="card"><h3>Machines on service today</h3><MiniTable rows={servicesToday.slice(0, 8)} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'supervisor', label: 'Supervisor' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} empty="No services due today." /></div>
          <div className="card"><h3>Foremen / chargehands on duty</h3><MiniTable rows={foremenOnDuty.slice(0, 10)} columns={[{ key: 'name', label: 'Name' }, { key: 'shift', label: 'Shift' }, { key: 'role', label: 'Role' }, { key: 'section', label: 'Section' }]} empty="No foremen loaded yet." /></div>
          <div className="card"><h3>People off / leave</h3><MiniTable rows={currentLeave.slice(0, 10)} columns={[{ key: 'person_name', label: 'Name' }, { key: 'leave_type', label: 'Type' }, { key: 'start_date', label: 'Start' }, { key: 'end_date', label: 'End' }, { key: 'days', label: 'Days' }]} empty="No current leave records." /></div>
        </div>
      </div>}

      {tab === 'fleet' && <div className="page-grid">
        <SectionTitle title="Fleet register" subtitle="Upload fleet register or add machines manually." right={<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadFleet(e.target.files?.[0])} />} />
        <div className="card form-card"><div className="form-grid">
          <Field label="Fleet number"><input value={fleetForm.fleet_no} onChange={(e) => setFleetForm({ ...fleetForm, fleet_no: e.target.value })} /></Field>
          <Field label="Machine type"><input value={fleetForm.machine_type} onChange={(e) => setFleetForm({ ...fleetForm, machine_type: e.target.value })} /></Field>
          <Field label="Make/model"><input value={fleetForm.make_model} onChange={(e) => setFleetForm({ ...fleetForm, make_model: e.target.value })} /></Field>
          <Field label="Department"><input value={fleetForm.department} onChange={(e) => setFleetForm({ ...fleetForm, department: e.target.value })} /></Field>
          <Field label="Location"><input value={fleetForm.location} onChange={(e) => setFleetForm({ ...fleetForm, location: e.target.value })} /></Field>
          <Field label="Hours"><input type="number" value={fleetForm.hours} onChange={(e) => setFleetForm({ ...fleetForm, hours: Number(e.target.value) })} /></Field>
          <Field label="Mileage"><input type="number" value={fleetForm.mileage} onChange={(e) => setFleetForm({ ...fleetForm, mileage: Number(e.target.value) })} /></Field>
          <Field label="Status"><select value={fleetForm.status} onChange={(e) => setFleetForm({ ...fleetForm, status: e.target.value })}><option>Available</option><option>Breakdown</option><option>Service</option><option>Major Repair</option></select></Field>
        </div><button className="primary" onClick={() => saveRow('fleet_machines', fleetForm, 'upsert')}>Save machine</button></div>
        <div className="card"><div className="card-head"><h3>Fleet list</h3><input placeholder="Search fleet..." value={search} onChange={(e) => setSearch(e.target.value)} /></div><MiniTable rows={filteredFleet} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'machine_type', label: 'Type' }, { key: 'make_model', label: 'Model' }, { key: 'department', label: 'Department' }, { key: 'hours', label: 'Hours' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div>
      </div>}

      {tab === 'spares' && <div className="page-grid spares-page">
        <SectionTitle title="Spares & Orders dark control" subtitle="Request spares, move orders through funding, local/international order and delivery ETA, and open PDF parts books with diagrams." />
        <div className="stats-grid mini"><StatCard title="Requests" value={data.spares_orders.filter((o) => String(o.workflow_stage || o.status) === 'Requested').length} onClick={() => setSparesStage('Requested')} /><StatCard title="Awaiting Funding" value={awaitingFunding.length} tone="red" onClick={() => setSparesStage('Awaiting Funding')} /><StatCard title="Awaiting Delivery" value={data.spares_orders.filter((o) => String(o.workflow_stage || '').includes('Delivery')).length} tone="orange" onClick={() => setSparesStage('Awaiting Delivery')} /><StatCard title="Low Stock" value={sparesLow.length} tone="red" /></div>
        <div className="grid-2">
          <div className="card form-card accent-card"><h3>New spares request sheet</h3><div className="form-grid">
            <Field label="Fleet number"><input list="fleetList" value={orderForm.machine_fleet_no} onChange={(e) => setOrderForm({ ...orderForm, machine_fleet_no: e.target.value })} /></Field>
            <Field label="Date"><input type="date" value={dateOnly(orderForm.request_date)} onChange={(e) => setOrderForm({ ...orderForm, request_date: e.target.value })} /></Field>
            <Field label="Requested by"><input value={orderForm.requested_by} onChange={(e) => setOrderForm({ ...orderForm, requested_by: e.target.value })} /></Field>
            <Field label="Machine area"><select value={orderForm.machine_group} onChange={(e) => setOrderForm({ ...orderForm, machine_group: e.target.value })}>{machineGroups.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Workshop section"><select value={orderForm.workshop_section} onChange={(e) => setOrderForm({ ...orderForm, workshop_section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Priority"><select value={orderForm.priority} onChange={(e) => setOrderForm({ ...orderForm, priority: e.target.value })}>{priorities.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Part number"><input value={orderForm.part_no} onChange={(e) => setOrderForm({ ...orderForm, part_no: e.target.value })} /></Field>
            <Field label="Qty"><input type="number" value={orderForm.qty} onChange={(e) => setOrderForm({ ...orderForm, qty: Number(e.target.value) })} /></Field>
            <Field label="Order type"><select value={orderForm.order_type} onChange={(e) => setOrderForm({ ...orderForm, order_type: e.target.value })}>{orderTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="ETA"><input type="date" value={dateOnly(orderForm.eta)} onChange={(e) => setOrderForm({ ...orderForm, eta: e.target.value })} /></Field>
          </div>
          <Field label="Description"><input value={orderForm.description} onChange={(e) => setOrderForm({ ...orderForm, description: e.target.value })} /></Field>
          <Field label="Spares list / parts needed"><textarea rows={5} value={orderForm.spares_items} onChange={(e) => setOrderForm({ ...orderForm, spares_items: e.target.value })} /></Field>
          <Field label="Part numbers from book"><textarea rows={3} value={orderForm.part_numbers} onChange={(e) => setOrderForm({ ...orderForm, part_numbers: e.target.value })} /></Field>
          <button className="primary" onClick={() => saveRow('spares_orders', orderForm)}>Save spares request + notify admin</button>
          </div>
          <div className="card parts-book-card"><h3>PDF parts books</h3><p className="muted">Upload PDF, save it, then open it to see diagrams. Click extracted part numbers into the order sheet.</p><div className="form-grid">
            <Field label="Book title"><input value={partsBookForm.book_title} onChange={(e) => setPartsBookForm({ ...partsBookForm, book_title: e.target.value })} /></Field>
            <Field label="Area"><select value={partsBookForm.machine_group} onChange={(e) => setPartsBookForm({ ...partsBookForm, machine_group: e.target.value })}>{machineGroups.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Machine type"><input value={partsBookForm.machine_type} onChange={(e) => setPartsBookForm({ ...partsBookForm, machine_type: e.target.value })} /></Field>
            <Field label="Uploaded by"><input value={partsBookForm.uploaded_by} onChange={(e) => setPartsBookForm({ ...partsBookForm, uploaded_by: e.target.value })} /></Field>
          </div>
          <input type="file" accept="application/pdf,.pdf" onChange={(e) => attachPartsBook(e.target.files?.[0])} />
          {partsBookForm.file_name && <div className="file-note"><b>{partsBookForm.file_name}</b> · {partsBookForm.extracted_count || 0} possible part numbers</div>}
          <div className="part-chip-box">{String(partsBookForm.extracted_part_numbers || '').split(/[\n,; ]+/).filter(Boolean).slice(0, 80).map((p) => <button key={p} className="part-chip" onClick={() => addPartToOrder(p)}>{p}</button>)}</div>
          <button className="primary" onClick={() => saveRow('parts_books', partsBookForm)}>Save parts book</button>
          </div>
        </div>
        <div className="card"><div className="card-head"><h3>Order board</h3><div className="filters"><input placeholder="Search orders..." value={sparesSearch} onChange={(e) => setSparesSearch(e.target.value)} /><select value={sparesStage} onChange={(e) => setSparesStage(e.target.value)}><option>All</option>{orderStages.map((x) => <option key={x}>{x}</option>)}{orderTypes.map((x) => <option key={x}>{x}</option>)}{machineGroups.map((x) => <option key={x}>{x}</option>)}</select></div></div>
          <div className="kanban-grid">{orderStages.slice(0, 5).map((stage) => <details className="drop-panel kanban-drop" key={stage} open={['Requested','Awaiting Funding','Awaiting Delivery'].includes(stage)}><summary><span>{stage}</span><Badge value={filteredSpareOrders.filter((o) => String(o.workflow_stage || o.status) === stage).length} /></summary>{renderSparesOrderCards(filteredSpareOrders.filter((o) => String(o.workflow_stage || o.status) === stage))}</details>)}</div>
        </div>
        <div className="grid-2"><div className="card"><h3>Saved PDF parts books</h3><MiniTable rows={data.parts_books} onRow={showPdfBook} columns={[{ key: 'book_title', label: 'Book' }, { key: 'machine_group', label: 'Area' }, { key: 'machine_type', label: 'Machine' }, { key: 'file_name', label: 'File' }, { key: 'extracted_count', label: 'Parts' }, { key: 'open', label: 'Open', render: () => <button>Open PDF</button> }]} empty="No PDF parts books uploaded." /></div><div className="card"><h3>Stock availability</h3><MiniTable rows={data.spares} columns={[{ key: 'part_no', label: 'Part' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'min_qty', label: 'Min' }, { key: 'order_status', label: 'Status', render: (r) => <Badge value={n(r.stock_qty) <= n(r.min_qty) ? 'Low stock' : r.order_status} /> }]} empty="No stock loaded." /></div></div>
      </div>}

      {tab === 'batteries' && <div className="page-grid batteries-page">
        <SectionTitle title="Battery control" subtitle="Battery serial tracking, fitment pictures, maintenance reminders, stock and ordering against machines." />
        <div className="stats-grid mini"><StatCard title="Fitted Batteries" value={data.batteries.filter((b) => String(b.status || b.stock_status).toLowerCase().includes('fitted')).length} /><StatCard title="Stock Batteries" value={data.batteries.filter((b) => String(b.stock_status || '').toLowerCase().includes('stock')).length} tone="green" /><StatCard title="Maintenance Due" value={data.batteries.filter((b) => dateOnly(b.maintenance_due_date) && new Date(b.maintenance_due_date) <= new Date()).length} tone="orange" /><StatCard title="Battery Orders" value={batteryRequests.length} tone="red" /></div>
        <div className="grid-2">
          <div className="card form-card"><h3>Battery fitment / stock record</h3><div className="form-grid">
            <Field label="Battery make"><input value={batteryForm.make} onChange={(e) => setBatteryForm({ ...batteryForm, make: e.target.value })} /></Field>
            <Field label="Serial number"><input value={batteryForm.serial_no} onChange={(e) => setBatteryForm({ ...batteryForm, serial_no: e.target.value })} /></Field>
            <Field label="Production date"><input type="date" value={dateOnly(batteryForm.production_date)} onChange={(e) => setBatteryForm({ ...batteryForm, production_date: e.target.value })} /></Field>
            <Field label="Date fitted"><input type="date" value={dateOnly(batteryForm.fitment_date)} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_date: e.target.value })} /></Field>
            <Field label="Fleet number"><input list="fleetList" value={batteryForm.fleet_no} onChange={(e) => setBatteryForm({ ...batteryForm, fleet_no: e.target.value })} /></Field>
            <Field label="Fitted by"><input value={batteryForm.fitted_by} onChange={(e) => setBatteryForm({ ...batteryForm, fitted_by: e.target.value })} /></Field>
            <Field label="Battery volts"><select value={batteryForm.volts} onChange={(e) => setBatteryForm({ ...batteryForm, volts: e.target.value })}>{batteryVoltages.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Charging system voltage"><input value={batteryForm.charging_voltage} onChange={(e) => setBatteryForm({ ...batteryForm, charging_voltage: e.target.value })} /></Field>
            <Field label="Hours when fitted"><input type="number" value={batteryForm.fitment_hours} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_hours: Number(e.target.value) })} /></Field>
            <Field label="Mileage when fitted"><input type="number" value={batteryForm.fitment_mileage} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_mileage: Number(e.target.value) })} /></Field>
            <Field label="Maintenance due"><input type="date" value={dateOnly(batteryForm.maintenance_due_date)} onChange={(e) => setBatteryForm({ ...batteryForm, maintenance_due_date: e.target.value })} /></Field>
            <Field label="Status"><select value={batteryForm.stock_status} onChange={(e) => setBatteryForm({ ...batteryForm, stock_status: e.target.value, status: e.target.value })}><option>Fitted</option><option>In Stock</option><option>Removed</option><option>Scrapped</option><option>On Charge</option></select></Field>
          </div><Field label="Battery fitment picture"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setBatteryForm((p) => ({ ...p, fitment_photo: v })))} /></Field>{batteryForm.fitment_photo && <img className="thumb" src={batteryForm.fitment_photo} alt="Battery fitment" />}<button className="primary" onClick={() => saveRow('batteries', batteryForm)}>Save battery record</button></div>
          <div className="card form-card"><h3>Battery order request</h3><div className="form-grid">
            <Field label="Fleet number"><input list="fleetList" value={batteryOrderForm.fleet_no} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, fleet_no: e.target.value })} /></Field>
            <Field label="Battery type"><input value={batteryOrderForm.battery_type} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, battery_type: e.target.value })} /></Field>
            <Field label="Voltage"><select value={batteryOrderForm.voltage} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, voltage: e.target.value })}>{batteryVoltages.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Qty"><input type="number" value={batteryOrderForm.qty} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, qty: Number(e.target.value) })} /></Field>
            <Field label="Requested by"><input value={batteryOrderForm.requested_by} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, requested_by: e.target.value })} /></Field>
            <Field label="Priority"><select value={batteryOrderForm.priority} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, priority: e.target.value })}>{priorities.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="ETA"><input type="date" value={dateOnly(batteryOrderForm.eta)} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, eta: e.target.value })} /></Field>
            <Field label="Status"><select value={batteryOrderForm.status} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, status: e.target.value })}>{orderStages.map((x) => <option key={x}>{x}</option>)}</select></Field>
          </div><Field label="Reason"><textarea rows={4} value={batteryOrderForm.reason} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, reason: e.target.value })} /></Field><button className="primary" onClick={() => saveRow('battery_orders', batteryOrderForm)}>Submit battery request + notify admin</button></div>
        </div>
        <div className="grid-2"><div className="card"><h3>All batteries</h3><MiniTable rows={data.batteries} columns={[{ key: 'serial_no', label: 'Serial' }, { key: 'make', label: 'Make' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'fitment_date', label: 'Fitted' }, { key: 'fitment_hours', label: 'Hours' }, { key: 'maintenance_due_date', label: 'Maint Due' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.stock_status || r.status} /> }]} /></div><div className="card"><h3>Battery orders / admin alerts</h3><MiniTable rows={data.battery_orders} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'battery_type', label: 'Battery' }, { key: 'qty', label: 'Qty' }, { key: 'requested_by', label: 'By' }, { key: 'priority', label: 'Priority' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div></div>
      </div>}

      {tab === 'tyres' && <div className="page-grid tyres-page">
        <SectionTitle title="Tyre control" subtitle="Tyre serial/company number tracking, repairs, fitment photos, measurements, reminders, stock and order requests." />
        <div className="stats-grid mini"><StatCard title="Fitted Tyres" value={data.tyres.filter((t) => String(t.status || t.stock_status).toLowerCase().includes('fitted')).length} /><StatCard title="Stock Tyres" value={data.tyres.filter((t) => String(t.stock_status || '').toLowerCase().includes('stock')).length} tone="green" /><StatCard title="Measurements Due" value={data.tyres.filter((t) => dateOnly(t.measurement_due_date) && new Date(t.measurement_due_date) <= new Date()).length} tone="orange" /><StatCard title="Tyre Orders" value={tyreRequests.length} tone="red" /></div>
        <div className="grid-2">
          <div className="card form-card"><h3>Tyre fitment / stock record</h3><div className="form-grid">
            <Field label="Tyre make"><input value={tyreForm.make} onChange={(e) => setTyreForm({ ...tyreForm, make: e.target.value })} /></Field>
            <Field label="Serial number"><input value={tyreForm.serial_no} onChange={(e) => setTyreForm({ ...tyreForm, serial_no: e.target.value })} /></Field>
            <Field label="Company number"><input value={tyreForm.company_no} onChange={(e) => setTyreForm({ ...tyreForm, company_no: e.target.value })} /></Field>
            <Field label="Production date"><input type="date" value={dateOnly(tyreForm.production_date)} onChange={(e) => setTyreForm({ ...tyreForm, production_date: e.target.value })} /></Field>
            <Field label="Date fitted"><input type="date" value={dateOnly(tyreForm.fitted_date)} onChange={(e) => setTyreForm({ ...tyreForm, fitted_date: e.target.value })} /></Field>
            <Field label="Fleet number"><input list="fleetList" value={tyreForm.fleet_no} onChange={(e) => setTyreForm({ ...tyreForm, fleet_no: e.target.value })} /></Field>
            <Field label="Position"><select value={tyreForm.position} onChange={(e) => setTyreForm({ ...tyreForm, position: e.target.value })}>{tyrePositions.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Fitted by"><input value={tyreForm.fitted_by} onChange={(e) => setTyreForm({ ...tyreForm, fitted_by: e.target.value })} /></Field>
            <Field label="Tyre type"><input value={tyreForm.tyre_type} onChange={(e) => setTyreForm({ ...tyreForm, tyre_type: e.target.value })} /></Field>
            <Field label="Size"><input value={tyreForm.size} onChange={(e) => setTyreForm({ ...tyreForm, size: e.target.value })} /></Field>
            <Field label="Hours fitted"><input type="number" value={tyreForm.fitment_hours} onChange={(e) => setTyreForm({ ...tyreForm, fitment_hours: Number(e.target.value) })} /></Field>
            <Field label="Mileage fitted"><input type="number" value={tyreForm.fitment_mileage} onChange={(e) => setTyreForm({ ...tyreForm, fitment_mileage: Number(e.target.value) })} /></Field>
            <Field label="Current tread mm"><input type="number" value={tyreForm.current_tread_mm} onChange={(e) => setTyreForm({ ...tyreForm, current_tread_mm: Number(e.target.value) })} /></Field>
            <Field label="Measurement due"><input type="date" value={dateOnly(tyreForm.measurement_due_date)} onChange={(e) => setTyreForm({ ...tyreForm, measurement_due_date: e.target.value })} /></Field>
            <Field label="Status"><select value={tyreForm.stock_status} onChange={(e) => setTyreForm({ ...tyreForm, stock_status: e.target.value, status: e.target.value })}><option>Fitted</option><option>In Stock</option><option>Repaired</option><option>Removed</option><option>Scrapped</option></select></Field>
          </div><Field label="Fitment picture"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setTyreForm((p) => ({ ...p, fitment_photo: v })))} /></Field>{tyreForm.fitment_photo && <img className="thumb" src={tyreForm.fitment_photo} alt="Tyre fitment" />}<button className="primary" onClick={() => saveRow('tyres', tyreForm)}>Save tyre record</button></div>
          <div className="card form-card"><h3>Tyre measurement / repair</h3><div className="form-grid">
            <Field label="Tyre serial"><input value={tyreMeasurementForm.tyre_serial_no} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, tyre_serial_no: e.target.value })} /></Field>
            <Field label="Fleet"><input list="fleetList" value={tyreMeasurementForm.fleet_no} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, fleet_no: e.target.value })} /></Field>
            <Field label="Position"><select value={tyreMeasurementForm.position} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, position: e.target.value })}>{tyrePositions.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Date"><input type="date" value={dateOnly(tyreMeasurementForm.measurement_date)} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, measurement_date: e.target.value })} /></Field>
            <Field label="Measured / repaired by"><input value={tyreMeasurementForm.measured_by} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, measured_by: e.target.value })} /></Field>
            <Field label="Tread depth mm"><input type="number" value={tyreMeasurementForm.tread_depth_mm} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, tread_depth_mm: Number(e.target.value) })} /></Field>
            <Field label="Pressure PSI"><input type="number" value={tyreMeasurementForm.pressure_psi} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, pressure_psi: Number(e.target.value) })} /></Field>
            <Field label="Condition"><select value={tyreMeasurementForm.condition} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, condition: e.target.value })}><option>Good</option><option>Monitor</option><option>Repair Done</option><option>Replace Soon</option><option>Scrap</option></select></Field>
          </div><Field label="Reason / notes"><textarea rows={4} value={tyreMeasurementForm.notes} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, notes: e.target.value })} /></Field><button className="primary" onClick={() => saveRow('tyre_measurements', tyreMeasurementForm)}>Save measurement / repair</button></div>
        </div>
        <div className="card form-card"><h3>Tyre order / quote request</h3><div className="form-grid"><Field label="Fleet"><input list="fleetList" value={tyreOrderForm.fleet_no} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, fleet_no: e.target.value })} /></Field><Field label="Tyre type"><input value={tyreOrderForm.tyre_type} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, tyre_type: e.target.value })} /></Field><Field label="Size"><input value={tyreOrderForm.size} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, size: e.target.value })} /></Field><Field label="Qty"><input type="number" value={tyreOrderForm.qty} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, qty: Number(e.target.value) })} /></Field><Field label="Requested by"><input value={tyreOrderForm.requested_by} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, requested_by: e.target.value })} /></Field><Field label="Priority"><select value={tyreOrderForm.priority} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, priority: e.target.value })}>{priorities.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Status"><select value={tyreOrderForm.status} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, status: e.target.value })}>{orderStages.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="ETA"><input type="date" value={dateOnly(tyreOrderForm.eta)} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, eta: e.target.value })} /></Field></div><Field label="Reason"><textarea rows={3} value={tyreOrderForm.reason} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, reason: e.target.value })} /></Field><button className="primary" onClick={() => saveRow('tyre_orders', tyreOrderForm)}>Submit tyre order + notify admin</button></div>
        <div className="grid-2"><div className="card"><h3>Tyres by fleet</h3><MiniTable rows={data.tyres} columns={[{ key: 'serial_no', label: 'Serial' }, { key: 'company_no', label: 'Company No' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'position', label: 'Pos' }, { key: 'current_tread_mm', label: 'Tread' }, { key: 'measurement_due_date', label: 'Measure Due' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.stock_status || r.status} /> }]} /></div><div className="card"><h3>Tyre orders / quotes</h3><MiniTable rows={data.tyre_orders} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'tyre_type', label: 'Type' }, { key: 'size', label: 'Size' }, { key: 'qty', label: 'Qty' }, { key: 'requested_by', label: 'By' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div></div>
      </div>}

      {tab === 'services' && <div className="page-grid services-page">
        <SectionTitle title="Services" subtitle="See past services, job cards, supervisors, technicians and complete service history by fleet." right={<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadService(e.target.files?.[0])} />} />
        <div className="card form-card"><h3>Service job card</h3><div className="form-grid"><Field label="Fleet"><input list="fleetList" value={serviceForm.fleet_no} onChange={(e) => setServiceForm({ ...serviceForm, fleet_no: e.target.value })} /></Field><Field label="Service type"><input value={serviceForm.service_type} onChange={(e) => setServiceForm({ ...serviceForm, service_type: e.target.value })} /></Field><Field label="Due date"><input type="date" value={dateOnly(serviceForm.due_date)} onChange={(e) => setServiceForm({ ...serviceForm, due_date: e.target.value })} /></Field><Field label="Scheduled hours"><input type="number" value={serviceForm.scheduled_hours} onChange={(e) => setServiceForm({ ...serviceForm, scheduled_hours: Number(e.target.value) })} /></Field><Field label="Completed date"><input type="date" value={dateOnly(serviceForm.completed_date)} onChange={(e) => setServiceForm({ ...serviceForm, completed_date: e.target.value })} /></Field><Field label="Completed hours"><input type="number" value={serviceForm.completed_hours} onChange={(e) => setServiceForm({ ...serviceForm, completed_hours: Number(e.target.value) })} /></Field><Field label="Job card"><input value={serviceForm.job_card_no} onChange={(e) => setServiceForm({ ...serviceForm, job_card_no: e.target.value })} /></Field><Field label="Supervisor"><input value={serviceForm.supervisor} onChange={(e) => setServiceForm({ ...serviceForm, supervisor: e.target.value })} /></Field><Field label="Technician"><input value={serviceForm.technician} onChange={(e) => setServiceForm({ ...serviceForm, technician: e.target.value })} /></Field><Field label="Status"><select value={serviceForm.status} onChange={(e) => setServiceForm({ ...serviceForm, status: e.target.value })}><option>Due</option><option>In progress</option><option>Completed</option><option>Deferred</option></select></Field></div><Field label="Service notes"><textarea rows={4} value={serviceForm.notes} onChange={(e) => setServiceForm({ ...serviceForm, notes: e.target.value })} /></Field><button className="primary" onClick={() => saveRow('services', serviceForm)}>Save service</button></div>
        <div className="grid-2"><div className="card"><h3>Due / active services</h3><MiniTable rows={servicesDue} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'due_date', label: 'Due' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'supervisor', label: 'Supervisor' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div><div className="card"><h3>Past service history</h3><MiniTable rows={data.services.filter((s) => String(s.status || '').toLowerCase().includes('complete'))} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'completed_date', label: 'Completed' }, { key: 'completed_hours', label: 'Hours' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'technician', label: 'Technician' }]} empty="No completed services yet." /></div></div>
      </div>}

      {tab === 'breakdowns' && <div className="page-grid breakdown-page">
        <SectionTitle title="Breakdown control" subtitle="Track downtime, spares ETA, availability ETA, repeated faults and preventative maintenance suggestions." />
        <div className="card form-card"><h3>Breakdown event</h3><div className="form-grid"><Field label="Fleet"><input list="fleetList" value={breakdownForm.fleet_no} onChange={(e) => setBreakdownForm({ ...breakdownForm, fleet_no: e.target.value })} /></Field><Field label="Fault category"><input value={breakdownForm.category} onChange={(e) => setBreakdownForm({ ...breakdownForm, category: e.target.value })} /></Field><Field label="Reported by"><input value={breakdownForm.reported_by} onChange={(e) => setBreakdownForm({ ...breakdownForm, reported_by: e.target.value })} /></Field><Field label="Worked by"><input value={breakdownForm.assigned_to} onChange={(e) => setBreakdownForm({ ...breakdownForm, assigned_to: e.target.value })} /></Field><Field label="Start time"><input type="datetime-local" value={breakdownForm.start_time} onChange={(e) => setBreakdownForm({ ...breakdownForm, start_time: e.target.value })} /></Field><Field label="Spares ETA"><input type="date" value={dateOnly(breakdownForm.spare_eta)} onChange={(e) => setBreakdownForm({ ...breakdownForm, spare_eta: e.target.value })} /></Field><Field label="Expected available"><input type="date" value={dateOnly(breakdownForm.expected_available_date)} onChange={(e) => setBreakdownForm({ ...breakdownForm, expected_available_date: e.target.value })} /></Field><Field label="Status"><select value={breakdownForm.status} onChange={(e) => setBreakdownForm({ ...breakdownForm, status: e.target.value })}><option>Open</option><option>Awaiting Spares</option><option>In progress</option><option>Available</option><option>Closed</option></select></Field></div><Field label="Fault"><textarea rows={3} value={breakdownForm.fault} onChange={(e) => setBreakdownForm({ ...breakdownForm, fault: e.target.value })} /></Field><Field label="Action / preventative maintenance suggestion"><textarea rows={3} value={breakdownForm.preventative_recommendation} onChange={(e) => setBreakdownForm({ ...breakdownForm, preventative_recommendation: e.target.value })} /></Field><button className="primary" onClick={() => saveRow('breakdowns', breakdownForm)}>Save breakdown</button></div>
        <div className="grid-2"><div className="card"><h3>Active breakdowns</h3><MiniTable rows={openBreakdowns} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'category', label: 'Type' }, { key: 'assigned_to', label: 'Worked By' }, { key: 'spare_eta', label: 'Spares ETA' }, { key: 'expected_available_date', label: 'Available ETA' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div><div className="card"><h3>Repeated fault suggestions</h3><MiniTable rows={repeatedFaults} columns={[{ key: 'fault', label: 'Repeated fault' }, { key: 'count', label: 'Times' }, { key: 'recommendation', label: 'Action' }]} empty="No repeated faults yet." /></div></div>
      </div>}

      {tab === 'repairs' && <div className="page-grid"><SectionTitle title="Repairs and job cards" subtitle="Record who worked on what, when and what parts were used." /><div className="card form-card"><div className="form-grid"><Field label="Fleet"><input list="fleetList" value={repairForm.fleet_no} onChange={(e) => setRepairForm({ ...repairForm, fleet_no: e.target.value })} /></Field><Field label="Job card"><input value={repairForm.job_card_no} onChange={(e) => setRepairForm({ ...repairForm, job_card_no: e.target.value })} /></Field><Field label="Assigned to"><input value={repairForm.assigned_to} onChange={(e) => setRepairForm({ ...repairForm, assigned_to: e.target.value })} /></Field><Field label="Start"><input type="datetime-local" value={repairForm.start_time} onChange={(e) => setRepairForm({ ...repairForm, start_time: e.target.value })} /></Field><Field label="End"><input type="datetime-local" value={repairForm.end_time} onChange={(e) => setRepairForm({ ...repairForm, end_time: e.target.value })} /></Field><Field label="Status"><select value={repairForm.status} onChange={(e) => setRepairForm({ ...repairForm, status: e.target.value })}><option>In progress</option><option>Awaiting Spares</option><option>Completed</option><option>Closed</option></select></Field></div><Field label="Fault"><textarea rows={3} value={repairForm.fault} onChange={(e) => setRepairForm({ ...repairForm, fault: e.target.value })} /></Field><Field label="Parts used"><textarea rows={3} value={repairForm.parts_used} onChange={(e) => setRepairForm({ ...repairForm, parts_used: e.target.value })} /></Field><button className="primary" onClick={() => saveRow('repairs', repairForm)}>Save repair</button></div><div className="card"><h3>Repair list</h3><MiniTable rows={data.repairs} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'assigned_to', label: 'Worked By' }, { key: 'fault', label: 'Fault' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div></div>}

      {tab === 'personnel' && <div className="page-grid personnel-page"><SectionTitle title="Personnel, shifts and leave" subtitle="Shift 1, Shift 2, Shift 3, managers, foremen, leave balances, off schedule and history." right={<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadPersonnel(e.target.files?.[0])} />} /><div className="stats-grid mini"><StatCard title="Managers" value={data.personnel.filter((p) => String(p.role).includes('Manager')).length} /><StatCard title="Foremen" value={data.personnel.filter((p) => String(p.role).includes('Foreman')).length} /><StatCard title="Current Leave" value={currentLeave.length} tone="orange" /><StatCard title="Upcoming Leave" value={upcomingLeave.length} /></div><div className="grid-2"><div className="card form-card"><h3>Add employee / slot</h3><div className="form-grid"><Field label="Emp No"><input value={personForm.employee_no} onChange={(e) => setPersonForm({ ...personForm, employee_no: e.target.value })} /></Field><Field label="Name"><input value={personForm.name} onChange={(e) => setPersonForm({ ...personForm, name: e.target.value })} /></Field><Field label="Shift"><select value={personForm.shift} onChange={(e) => setPersonForm({ ...personForm, shift: e.target.value })}>{shiftNames.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Section"><select value={personForm.section} onChange={(e) => setPersonForm({ ...personForm, section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Role"><select value={personForm.role} onChange={(e) => setPersonForm({ ...personForm, role: e.target.value })}>{personnelRoles.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Status"><select value={personForm.employment_status} onChange={(e) => setPersonForm({ ...personForm, employment_status: e.target.value })}><option>Active</option><option>On Leave</option><option>Off Duty</option><option>Vacant</option><option>Suspended</option></select></Field><Field label="Leave balance"><input type="number" value={personForm.leave_balance_days} onChange={(e) => setPersonForm({ ...personForm, leave_balance_days: Number(e.target.value) })} /></Field><Field label="Leave taken"><input type="number" value={personForm.leave_taken_days} onChange={(e) => setPersonForm({ ...personForm, leave_taken_days: Number(e.target.value) })} /></Field></div><button className="primary" onClick={() => saveRow('personnel', personForm, 'upsert')}>Save person</button></div><div className="card form-card"><h3>Add leave / off schedule</h3><div className="form-grid"><Field label="Person"><input value={leaveForm.person_name} onChange={(e) => setLeaveForm({ ...leaveForm, person_name: e.target.value })} /></Field><Field label="Emp No"><input value={leaveForm.employee_no} onChange={(e) => setLeaveForm({ ...leaveForm, employee_no: e.target.value })} /></Field><Field label="Type"><select value={leaveForm.leave_type} onChange={(e) => setLeaveForm({ ...leaveForm, leave_type: e.target.value })}>{leaveTypes.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Start"><input type="date" value={dateOnly(leaveForm.start_date)} onChange={(e) => setLeaveForm({ ...leaveForm, start_date: e.target.value })} /></Field><Field label="End"><input type="date" value={dateOnly(leaveForm.end_date)} onChange={(e) => setLeaveForm({ ...leaveForm, end_date: e.target.value })} /></Field><Field label="Status"><select value={leaveForm.status} onChange={(e) => setLeaveForm({ ...leaveForm, status: e.target.value })}><option>Approved</option><option>Pending</option><option>Cancelled</option></select></Field></div><Field label="Reason"><textarea rows={3} value={leaveForm.reason} onChange={(e) => setLeaveForm({ ...leaveForm, reason: e.target.value })} /></Field><button className="primary" onClick={saveLeaveRecord}>Save leave and deduct days</button></div></div><div className="card form-card"><h3>Add manpower by quantity</h3><div className="form-grid"><Field label="Shift"><select value={bulkPersonForm.shift} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, shift: e.target.value })}>{shiftNames.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Role"><select value={bulkPersonForm.role} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, role: e.target.value })}>{personnelRoles.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Section"><select value={bulkPersonForm.section} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Qty"><input type="number" value={bulkPersonForm.qty} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, qty: Number(e.target.value) })} /></Field></div><button onClick={addBulkPersonnel}>Add vacant slots</button></div>{shiftNames.map(renderPersonnelShift)}<div className="grid-2"><div className="card"><h3>Current leave</h3><MiniTable rows={currentLeave} columns={[{ key: 'person_name', label: 'Name' }, { key: 'leave_type', label: 'Type' }, { key: 'start_date', label: 'Start' }, { key: 'end_date', label: 'End' }, { key: 'days', label: 'Days' }]} /></div><div className="card"><h3>Past leave history</h3><MiniTable rows={pastLeave} columns={[{ key: 'person_name', label: 'Name' }, { key: 'leave_type', label: 'Type' }, { key: 'start_date', label: 'Start' }, { key: 'end_date', label: 'End' }, { key: 'days', label: 'Days' }]} /></div></div></div>}

      {tab === 'operations' && <div className="page-grid operations-page"><SectionTitle title="Operations and operator scores" subtitle="Separate operators by truck/yellow machines, record offences, damage, abuse photos, rough conditions and score history." /><div className="card form-card"><h3>Operator event / score sheet</h3><div className="form-grid"><Field label="Operator name"><input value={operationForm.operator_name} onChange={(e) => setOperationForm({ ...operationForm, operator_name: e.target.value })} /></Field><Field label="Employee No"><input value={operationForm.employee_no} onChange={(e) => setOperationForm({ ...operationForm, employee_no: e.target.value })} /></Field><Field label="Machine group"><select value={operationForm.machine_group} onChange={(e) => setOperationForm({ ...operationForm, machine_group: e.target.value })}><option>Truck</option><option>Yellow Machine</option></select></Field><Field label="Machine type"><select value={operationForm.machine_type} onChange={(e) => setOperationForm({ ...operationForm, machine_type: e.target.value })}>{operationMachineTypes.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Fleet"><input list="fleetList" value={operationForm.fleet_no} onChange={(e) => setOperationForm({ ...operationForm, fleet_no: e.target.value })} /></Field><Field label="Shift"><select value={operationForm.shift} onChange={(e) => setOperationForm({ ...operationForm, shift: e.target.value })}>{shiftNames.map((x) => <option key={x}>{x}</option>)}</select></Field><Field label="Score"><input type="number" value={operationForm.score} onChange={(e) => setOperationForm({ ...operationForm, score: Number(e.target.value) })} /></Field><Field label="Supervisor"><input value={operationForm.supervisor} onChange={(e) => setOperationForm({ ...operationForm, supervisor: e.target.value })} /></Field><Field label="Event date"><input type="date" value={dateOnly(operationForm.event_date)} onChange={(e) => setOperationForm({ ...operationForm, event_date: e.target.value })} /></Field></div><Field label="Offence / damage / abuse / rough working conditions"><textarea rows={4} value={operationForm.offence} onChange={(e) => setOperationForm({ ...operationForm, offence: e.target.value })} /></Field><Field label="Supervisor photo upload"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setOperationForm((p) => ({ ...p, photo_data: v })))} /></Field>{operationForm.photo_data && <img className="thumb" src={operationForm.photo_data} alt="Operator event" />}<button className="primary" onClick={() => saveRow('operations', operationForm)}>Save operator event</button></div><div className="grid-2"><div className="card"><h3>Truck operators</h3><MiniTable rows={data.operations.filter((o) => String(o.machine_group).includes('Truck'))} columns={[{ key: 'operator_name', label: 'Operator' }, { key: 'machine_type', label: 'Machine Type' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'score', label: 'Score' }, { key: 'offence', label: 'Event' }, { key: 'supervisor', label: 'Supervisor' }]} /></div><div className="card"><h3>Yellow machine operators</h3><MiniTable rows={data.operations.filter((o) => String(o.machine_group).includes('Yellow'))} columns={[{ key: 'operator_name', label: 'Operator' }, { key: 'machine_type', label: 'Machine Type' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'score', label: 'Score' }, { key: 'offence', label: 'Event' }, { key: 'supervisor', label: 'Supervisor' }]} /></div></div></div>}

      {tab === 'photos' && <div className="page-grid"><SectionTitle title="Photo evidence" subtitle="Upload photos against machines, breakdowns, tyre fitment, battery fitment or operator abuse/conditions." /><div className="card form-card"><div className="form-grid"><Field label="Fleet"><input list="fleetList" value={photoForm.fleet_no} onChange={(e) => setPhotoForm({ ...photoForm, fleet_no: e.target.value })} /></Field><Field label="Linked type"><select value={photoForm.linked_type} onChange={(e) => setPhotoForm({ ...photoForm, linked_type: e.target.value })}><option>Breakdown</option><option>Repair</option><option>Tyre</option><option>Battery</option><option>Operator</option><option>Service</option></select></Field><Field label="Caption"><input value={photoForm.caption} onChange={(e) => setPhotoForm({ ...photoForm, caption: e.target.value })} /></Field><Field label="Photo"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setPhotoForm((p) => ({ ...p, image_data: v })))} /></Field></div>{photoForm.image_data && <img className="thumb" src={photoForm.image_data} alt="Preview" />}<button className="primary" onClick={() => saveRow('photo_logs', photoForm)}>Save photo</button></div><div className="photo-grid">{data.photo_logs.map((p) => <div className="photo-card" key={p.id}>{p.image_data && <img src={p.image_data} alt={p.caption || 'photo'} />}<b>{p.fleet_no}</b><span>{p.linked_type}</span><p>{p.caption}</p></div>)}</div></div>}

      {tab === 'reports' && <div className="page-grid reports-page"><SectionTitle title="Reports" subtitle="Machine availability, breakdowns per department, fleet type, service history and submitted records." right={<input placeholder="Filter fleet..." value={reportFleet} onChange={(e) => setReportFleet(e.target.value)} />} /><div className="stats-grid mini"><StatCard title="Fleet Availability" value={`${departmentReport.length ? one(((data.fleet_machines.length - uniqueBy(openBreakdowns, 'fleet_no').length) / Math.max(1, data.fleet_machines.length)) * 100) : 0}%`} /><StatCard title="Breakdowns" value={openBreakdowns.length} tone="red" /><StatCard title="Services Due" value={servicesDue.length} tone="orange" /><StatCard title="Orders / Requests" value={data.spares_orders.length + data.tyre_orders.length + data.battery_orders.length} /></div><div className="grid-2"><div className="card"><h3>Breakdown per department</h3><MiniTable rows={departmentReport} columns={[{ key: 'department', label: 'Department' }, { key: 'total', label: 'Fleet' }, { key: 'breakdowns', label: 'Down' }, { key: 'services', label: 'Services' }, { key: 'availability', label: 'Avail %' }]} /></div><div className="card"><h3>Breakdown per fleet type</h3><MiniTable rows={typeReport} columns={[{ key: 'machine_type', label: 'Fleet Type' }, { key: 'total', label: 'Fleet' }, { key: 'breakdowns', label: 'Down' }, { key: 'availability', label: 'Avail %' }]} /></div><div className="card"><h3>Machine breakdown records</h3><MiniTable rows={reportRows.breakdowns} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'category', label: 'Type' }, { key: 'fault', label: 'Fault' }, { key: 'assigned_to', label: 'Worked By' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div><div className="card"><h3>Service history report</h3><MiniTable rows={reportRows.services} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'supervisor', label: 'Supervisor' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></div></div></div>}

      {tab === 'admin' && <div className="page-grid"><SectionTitle title="Admin control" subtitle="System sync, local queue and delete actions." /><div className="card"><h3>Connection</h3><p>Supabase: {isSupabaseConfigured ? 'Configured' : 'Not configured'} · Browser: {online ? 'Online' : 'Offline'} · Queue: {queueCount}</p><button className="primary" onClick={syncOfflineQueue}>Sync offline queue</button><button onClick={() => { localStorage.removeItem('turbo-workshop-snapshot'); setMessage('Local snapshot cleared. Refresh to reload from Supabase.') }}>Clear local snapshot</button></div><div className="card"><h3>Admin alerts</h3><MiniTable rows={[...data.spares_orders, ...data.tyre_orders, ...data.battery_orders].filter((o) => String(o.status || o.workflow_stage || '').toLowerCase().includes('request') || String(o.status || o.workflow_stage || '').toLowerCase().includes('awaiting'))} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'machine_fleet_no', label: 'Machine' }, { key: 'requested_by', label: 'Requested By' }, { key: 'reason', label: 'Reason' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status || r.workflow_stage} /> }]} /></div></div>}

      <datalist id="fleetList">{data.fleet_machines.map((m) => <option key={m.fleet_no} value={m.fleet_no}>{m.machine_type}</option>)}</datalist>
      {modal && <Modal title={modal.title} onClose={() => setModal(null)}>{modal.body}</Modal>}
    </section>
  </main>
}
