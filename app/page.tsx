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
  | 'hoses'
  | 'hose_requests'
  | 'tools'
  | 'tool_orders'
  | 'employee_tools'
  | 'fabrication_stock'
  | 'fabrication_requests'
  | 'fabrication_projects'

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
  'operations',
  'hoses',
  'hose_requests',
  'tools',
  'tool_orders',
  'employee_tools',
  'fabrication_stock',
  'fabrication_requests',
  'fabrication_projects'
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
  operations: [],
  hoses: [],
  hose_requests: [],
  tools: [],
  tool_orders: [],
  employee_tools: [],
  fabrication_stock: [],
  fabrication_requests: [],
  fabrication_projects: []
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
  ['operations', 'Operations'],
  ['hoses', 'Hoses'],
  ['tools', 'Tools'],
  ['fabrication', 'Boiler/Panel'],
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
const hoseTypes = ['Hydraulic Hose', 'Fuel Hose', 'Air Hose', 'Water Hose', 'Grease Hose', 'Suction Hose', 'Return Hose', 'Brake Hose']
const fittingTypes = ['Straight', '90 Degree Elbow', '45 Degree Elbow', 'Banjo', 'Flange', 'JIC', 'BSP', 'Metric', 'NPT', 'Quick Coupler']
const hoseStages = ['Requested', 'Awaiting Admin Approval', 'Approved', 'Offsite Repair', 'In Workshop', 'Awaiting Fittings', 'Completed', 'Rejected']
const toolCategories = ['Hand Tools', 'Power Tools', 'Diagnostic Tools', 'Lifting Tools', 'Measuring Tools', 'Welding Tools', 'Special Tools', 'Consumables']
const fabricationDepartments = ['Boiler Shop', 'Panel Beaters']
const materialTypes = ['Steel Plate', 'Flat Bar', 'Angle Iron', 'Channel', 'Pipe', 'Welding Rods', 'Gas', 'Paint', 'Panel Beating Consumables', 'Bolts/Nuts', 'Consumables']

const nowIso = () => new Date().toISOString()
const today = () => new Date().toISOString().slice(0, 10)
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
const n = (value: any) => Number(value || 0)
const one = (value: any) => Math.round(n(value) * 10) / 10

function cleanRow(row: Row) {
  const next: Row = {}
  Object.entries(row).forEach(([key, value]) => {
    if (value === undefined || value === null) return
    if (typeof value === 'string' && value.trim() === '') return
    next[key] = value
  })
  return next
}

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
  const max = Math.min(bytes.length, 1500000)
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

function safeFileName(name: string) {
  return String(name || 'parts-book.pdf').replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 120)
}

function statusClass(status: any) {
  const s = String(status || '').toLowerCase()
  if (s.includes('critical') || s.includes('open') || s.includes('break') || s.includes('awaiting') || s.includes('low') || s.includes('due')) return 'danger'
  if (s.includes('funded') || s.includes('progress') || s.includes('request') || s.includes('quote') || s.includes('leave')) return 'warn'
  if (s.includes('complete') || s.includes('delivered') || s.includes('available') || s.includes('active') || s.includes('stock') || s.includes('approved')) return 'good'
  return 'neutral'
}

function stripHeavyFileData(input: AppData): AppData {
  const next = { ...input } as AppData
  next.parts_books = (next.parts_books || []).map(({ file_data, object_url, ...book }) => book)
  return next
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

function Badge({ value }: { value: any }) {
  return <span className={`badge ${statusClass(value)}`}>{String(value || 'Not set')}</span>
}

function Field({ label, children }: { label: string; children: any }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  )
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

function StatCard({ title, value, detail, tone = 'blue', onClick }: { title: string; value: any; detail?: string; tone?: string; onClick?: () => void }) {
  return (
    <button type="button" className={`stat-card ${tone}`} onClick={onClick}>
      <span>{title}</span>
      <b>{value}</b>
      {detail && <small>{detail}</small>}
    </button>
  )
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

function Modal({ title, children, onClose }: { title: string; children: any; onClose: () => void }) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>{title}</h3>
          <button onClick={onClose}>Close</button>
        </div>
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
  const [sparesSearch, setSparesSearch] = useState('')
  const [sparesStage, setSparesStage] = useState('All')

  const [fleetForm, setFleetForm] = useState<Row>({ fleet_no: '', machine_type: '', make_model: '', department: '', location: '', hours: 0, mileage: 0, status: 'Available', service_interval_hours: 250, last_service_hours: 0 })
  const [breakdownForm, setBreakdownForm] = useState<Row>({ fleet_no: '', category: 'Mechanical', fault: '', reported_by: '', assigned_to: '', start_time: '', status: 'Open', spare_eta: '', expected_available_date: '', cause: '', action_taken: '', preventative_recommendation: '' })
  const [repairForm, setRepairForm] = useState<Row>({ fleet_no: '', job_card_no: '', fault: '', assigned_to: '', start_time: '', end_time: '', status: 'In progress', parts_used: '', notes: '' })
  const [serviceForm, setServiceForm] = useState<Row>({ fleet_no: '', service_type: '250hr service', due_date: today(), scheduled_hours: 0, completed_date: '', completed_hours: 0, job_card_no: '', supervisor: '', technician: '', status: 'Due', notes: '' })
  const [spareForm, setSpareForm] = useState<Row>({ part_no: '', description: '', section: 'Stores', machine_group: 'Yellow Machine', machine_type: '', stock_qty: 0, min_qty: 1, supplier: '', shelf_location: '', lead_time_days: 0, order_status: 'In stock' })
  const [orderForm, setOrderForm] = useState<Row>({ request_date: today(), machine_fleet_no: '', machine_group: 'Yellow Machine', workshop_section: 'Mechanical', requested_by: '', part_no: '', description: '', qty: 1, spares_items: '', priority: 'Normal', workflow_stage: 'Requested', status: 'Requested', order_type: 'Not selected', funding_status: 'Not funded', eta: '', part_numbers: '', pdf_book_ref: '', notes: '' })
  const [partsBookForm, setPartsBookForm] = useState<Row>({ book_title: '', machine_group: 'Yellow Machine', machine_type: '', uploaded_by: '', file_name: '', mime_type: '', file_size: 0, file_url: '', storage_path: '', object_url: '', extracted_part_numbers: '', extracted_text: '', notes: '' })
  const [personForm, setPersonForm] = useState<Row>({ employee_no: '', name: '', shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', staff_group: 'Workshop Staff', phone: '', employment_status: 'Active', leave_balance_days: 30, leave_taken_days: 0 })
  const [leaveForm, setLeaveForm] = useState<Row>({ person_name: '', employee_no: '', leave_type: 'Annual Leave', start_date: today(), end_date: today(), reason: '', status: 'Approved' })
  const [bulkPersonForm, setBulkPersonForm] = useState<Row>({ shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', qty: 1 })
  const [tyreForm, setTyreForm] = useState<Row>({ make: '', serial_no: '', company_no: '', production_date: '', fitted_date: today(), fleet_no: '', fitted_by: '', position: 'FL', tyre_type: '', size: '', fitment_hours: 0, fitment_mileage: 0, current_tread_mm: 0, measurement_due_date: '', stock_status: 'Fitted', status: 'Fitted', fitment_photo: '', notes: '' })
  const [tyreMeasurementForm, setTyreMeasurementForm] = useState<Row>({ tyre_serial_no: '', fleet_no: '', position: '', measurement_date: today(), measured_by: '', tread_depth_mm: 0, pressure_psi: 0, condition: 'Good', notes: '' })
  const [tyreOrderForm, setTyreOrderForm] = useState<Row>({ request_date: today(), fleet_no: '', tyre_type: '', size: '', qty: 1, requested_by: '', reason: '', priority: 'Normal', status: 'Requested', quote_status: 'Need Quote', eta: '', notes: '' })
  const [batteryForm, setBatteryForm] = useState<Row>({ make: '', serial_no: '', production_date: '', fitment_date: today(), fleet_no: '', fitted_by: '', charging_voltage: '24V', volts: '12V', fitment_hours: 0, fitment_mileage: 0, maintenance_due_date: '', stock_status: 'Fitted', status: 'Fitted', fitment_photo: '', notes: '' })
  const [batteryOrderForm, setBatteryOrderForm] = useState<Row>({ request_date: today(), fleet_no: '', battery_type: '', voltage: '12V', qty: 1, requested_by: '', reason: '', priority: 'Normal', status: 'Requested', eta: '', notes: '' })
  const [photoForm, setPhotoForm] = useState<Row>({ fleet_no: '', linked_type: 'Breakdown', caption: '', image_data: '' })
  const [operationForm, setOperationForm] = useState<Row>({ operator_name: '', employee_no: '', machine_group: 'Truck', machine_type: 'Truck', fleet_no: '', shift: 'Shift 1', score: 100, offence: '', damage_report: '', supervisor: '', event_date: today(), photo_data: '', notes: '' })
  const [hoseStockForm, setHoseStockForm] = useState<Row>({ hose_no: '', hose_type: 'Hydraulic Hose', size: '', length: '', fitting_a: 'Straight', fitting_b: 'Straight', stock_qty: 0, min_qty: 1, shelf_location: '', status: 'In stock', notes: '' })
  const [hoseRequestForm, setHoseRequestForm] = useState<Row>({ request_date: today(), fleet_no: '', requested_by: '', hose_type: 'Hydraulic Hose', hose_no: '', size: '', length: '', fittings: '', qty: 1, repair_type: 'New hose', offsite_repair: 'No', admin_approval: 'Pending', status: 'Requested', eta: '', reason: '', notes: '' })
  const [toolForm, setToolForm] = useState<Row>({ tool_name: '', category: 'Hand Tools', brand: '', model: '', serial_no: '', stock_qty: 0, workshop_section: 'Mechanical', condition: 'Good', status: 'In stock', notes: '' })
  const [toolOrderForm, setToolOrderForm] = useState<Row>({ request_date: today(), tool_name: '', category: 'Hand Tools', brand: '', qty: 1, requested_by: '', reason: '', priority: 'Normal', status: 'Requested', eta: '', notes: '' })
  const [employeeToolForm, setEmployeeToolForm] = useState<Row>({ employee_name: '', employee_no: '', tool_name: '', brand: '', serial_no: '', issue_date: today(), issued_by: '', condition_issued: 'Good', return_date: '', condition_returned: '', status: 'Issued', notes: '' })
  const [fabStockForm, setFabStockForm] = useState<Row>({ department: 'Boiler Shop', material_type: 'Steel Plate', description: '', size_spec: '', stock_qty: 0, min_qty: 1, unit: 'pcs', supplier: '', location: '', status: 'In stock', notes: '' })
  const [fabRequestForm, setFabRequestForm] = useState<Row>({ request_date: today(), department: 'Boiler Shop', fleet_no: '', requested_by: '', material_type: 'Steel Plate', description: '', qty: 1, project_name: '', status: 'Requested', eta: '', reason: '', notes: '' })
  const [fabProjectForm, setFabProjectForm] = useState<Row>({ department: 'Boiler Shop', project_name: '', fleet_no: '', supervisor: '', start_date: today(), target_date: '', status: 'Ongoing', material_used: '', progress_percent: 0, notes: '' })

  const isAdmin = session?.role === 'admin'

  useEffect(() => {
    const savedSession = localStorage.getItem('turbo-workshop-session')
    if (savedSession) setSession(JSON.parse(savedSession))

    const snapshot = readSnapshot<AppData>()
    if (snapshot) setData({ ...emptyData, ...snapshot })

    setOnline(typeof navigator === 'undefined' ? true : navigator.onLine)
    setQueueCount(getOfflineQueue().length)
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

  useEffect(() => {
    saveSnapshot(stripHeavyFileData(data))
  }, [data])

  async function loadAll() {
    if (!supabase || !isSupabaseConfigured || (typeof navigator !== 'undefined' && !navigator.onLine)) return

    const next: AppData = { ...emptyData }

    for (const table of TABLES) {
      try {
        const orderCol = table === 'fleet_machines' ? 'fleet_no' : 'created_at'
        const { data: rows } = await supabase.from(table as any).select('*').order(orderCol, { ascending: false })
        if (rows) next[table] = rows as Row[]
      } catch {
        next[table] = []
      }
    }

    setData(next)
    saveSnapshot(stripHeavyFileData(next))
  }

  async function syncOfflineQueue() {
    if (!supabase || !isSupabaseConfigured || !navigator.onLine) return

    const queue = getOfflineQueue()
    if (!queue.length) return

    const remaining = []

    for (const action of queue) {
      try {
        if (action.type === 'delete') {
          const { error } = await supabase.from(action.table as any).delete().eq('id', action.payload.id)
          if (error) throw error
        } else if (action.type === 'upsert') {
          const { error } = await supabase.from(action.table as any).upsert(action.payload, { onConflict: action.table === 'fleet_machines' ? 'fleet_no' : 'id' })
          if (error) throw error
        } else {
          const { error } = await supabase.from(action.table as any).insert(action.payload)
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
    const payload = cleanRow({
      id: row.id || uid(),
      ...row,
      created_at: row.created_at || nowIso(),
      updated_at: nowIso()
    })

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
      ? supabase.from(table as any).upsert(payload, { onConflict: table === 'fleet_machines' ? 'fleet_no' : 'id' })
      : supabase.from(table as any).insert(payload)

    const { error } = await request

    if (error) {
      addOfflineAction({ table, type: mode, payload })
      setQueueCount(getOfflineQueue().length)
      setMessage(`Saved locally because Supabase rejected it: ${error.message}`)
    } else {
      setMessage('Saved online successfully.')
    }
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
    for (const row of mapped) await saveRow('services', row)
    setMessage(`${mapped.length} service records uploaded.`)
  }

  async function uploadPersonnel(file?: File) {
    if (!file) return
    const rows = await readExcelRows(file)
    const people = rows.map(mapPersonnel).filter((row) => row.name)
    const leave = rows.map(mapLeave).filter((row) => row.person_name && row.start_date)

    for (const row of people) await saveRow('personnel', row, 'upsert')
    for (const row of leave) await saveRow('leave_records', row)

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
    const rawText = arrayBufferToText(buffer)
    const detected = extractPartNumbers(`${file.name} ${partsBookForm.notes || ''} ${rawText}`)

    let fileUrl = ''
    let storagePath = ''
    let objectUrl = ''

    try {
      if (supabase && isSupabaseConfigured && navigator.onLine) {
        storagePath = `${Date.now()}-${safeFileName(file.name)}`
        const { error } = await supabase.storage.from('parts-books').upload(storagePath, file, {
          cacheControl: '3600',
          upsert: true,
          contentType: file.type || 'application/pdf'
        })

        if (error) throw error

        const { data: publicData } = supabase.storage.from('parts-books').getPublicUrl(storagePath)
        fileUrl = publicData.publicUrl
      } else {
        objectUrl = URL.createObjectURL(file)
      }
    } catch (err: any) {
      objectUrl = URL.createObjectURL(file)
      setMessage(`PDF loaded for this session only. Create a public Supabase Storage bucket called parts-books later. ${err?.message || ''}`)
    }

    setPartsBookForm((prev) => ({
      ...prev,
      file_name: file.name,
      mime_type: file.type || 'application/pdf',
      file_size: file.size,
      file_url: fileUrl,
      storage_path: storagePath,
      object_url: objectUrl,
      extracted_part_numbers: detected.join('\n'),
      extracted_count: detected.length,
      extracted_text: rawText.slice(0, 7000)
    }))

    if (fileUrl) setMessage(`${detected.length} possible part numbers found. PDF uploaded.`)
    if (!fileUrl && objectUrl) setMessage(`${detected.length} possible part numbers found. PDF can open this session.`)
  }

  function openPartsBook(book: Row) {
    const url = book.file_url || book.object_url
    const parts = String(book.extracted_part_numbers || '').split('\n').map((x) => x.trim()).filter(Boolean)

    setModal({
      title: book.book_title || book.file_name || 'Parts book',
      body: (
        <div className="pdf-modal-grid">
          <div className="pdf-frame-wrap">
            {url ? <iframe src={url} title="Parts book PDF" /> : <div className="empty">No PDF URL saved. Create Supabase Storage bucket called parts-books for saved PDF viewing.</div>}
          </div>
          <div className="pdf-side">
            <h3>Detected part numbers</h3>
            <p className="muted">Click a part number to add it to the spares request.</p>
            <div className="part-chip-box tall">
              {parts.length ? parts.map((part) => (
                <button key={part} className="part-chip" onClick={() => addPartToOrder(part)}>{part}</button>
              )) : <span className="muted">No part numbers detected.</span>}
            </div>
          </div>
        </div>
      )
    })
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

    const person = data.personnel.find((p) =>
      String(p.name || '').toLowerCase() === String(leaveForm.person_name || '').toLowerCase() ||
      String(p.employee_no || '') === String(leaveForm.employee_no || '')
    )

    if (person && !String(leaveForm.leave_type || '').toLowerCase().includes('off')) {
      await saveRow('personnel', {
        ...person,
        leave_taken_days: one(n(person.leave_taken_days) + days),
        leave_balance_days: one(n(person.leave_balance_days) - days),
        employment_status: isBetweenDates(leaveForm.start_date, leaveForm.end_date) ? 'On Leave' : person.employment_status
      }, 'upsert')
    }

    setLeaveForm({ person_name: '', employee_no: '', leave_type: 'Annual Leave', start_date: today(), end_date: today(), reason: '', status: 'Approved' })
  }

  async function addBulkPeople() {
    const qty = Math.max(1, Number(bulkPersonForm.qty || 1))
    for (let i = 1; i <= qty; i++) {
      await saveRow('personnel', {
        employee_no: '',
        name: `Vacant ${bulkPersonForm.role} ${i}`,
        shift: bulkPersonForm.shift,
        section: bulkPersonForm.section,
        role: bulkPersonForm.role,
        staff_group: 'Vacant Slot',
        employment_status: 'Vacant',
        leave_balance_days: 0,
        leave_taken_days: 0
      })
    }
    setMessage(`${qty} ${bulkPersonForm.role} slot(s) added to ${bulkPersonForm.shift}.`)
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

  const metrics = useMemo(() => {
    const openBreakdowns = data.breakdowns.filter((r) => !['Closed', 'Completed'].includes(String(r.status || '')))
    const openRepairs = data.repairs.filter((r) => !['Closed', 'Completed'].includes(String(r.status || '')))
    const servicesDue = data.services.filter((r) => !['Completed', 'Closed'].includes(String(r.status || '')))
    const lowSpares = data.spares.filter((r) => n(r.stock_qty) <= n(r.min_qty))
    const peopleOff = data.personnel.filter((p) => {
      const s = String(p.employment_status || '').toLowerCase()
      return s.includes('leave') || s.includes('off') || s.includes('sick') || s.includes('training')
    })
    const foremen = data.personnel.filter((p) => String(p.role || '').toLowerCase().includes('foreman') && !String(p.employment_status || '').toLowerCase().includes('leave'))
    const tyreDue = data.tyres.filter((t) => isPastDate(t.measurement_due_date))
    const batteryDue = data.batteries.filter((b) => isPastDate(b.maintenance_due_date))
    const hosePending = data.hose_requests.filter((h) => !['Completed', 'Rejected'].includes(String(h.status || '')))
    const toolOrders = data.tool_orders.filter((t) => !['Delivered', 'Cancelled'].includes(String(t.status || '')))
    const fabProjects = data.fabrication_projects.filter((p) => !['Completed', 'Closed'].includes(String(p.status || '')))

    return {
      totalFleet: data.fleet_machines.length,
      openBreakdowns,
      openRepairs,
      servicesDue,
      lowSpares,
      peopleOff,
      foremen,
      tyreDue,
      batteryDue,
      hosePending,
      toolOrders,
      fabProjects
    }
  }, [data])

  const filteredFleet = useMemo(() => data.fleet_machines.filter((r) => JSON.stringify(r).toLowerCase().includes(search.toLowerCase())), [data.fleet_machines, search])

  const filteredOrders = useMemo(() => {
    return data.spares_orders.filter((r) => {
      const matchesText = JSON.stringify(r).toLowerCase().includes(sparesSearch.toLowerCase())
      const matchesStage = sparesStage === 'All' || String(r.workflow_stage || r.status) === sparesStage
      return matchesText && matchesStage
    })
  }, [data.spares_orders, sparesSearch, sparesStage])

  function renderDashboard() {
    return (
      <div className="page-grid">
        <SectionTitle
          title="Workshop Dashboard"
          subtitle="Current shift, foremen, breakdowns, services, spares, hoses, tools, tyres, batteries and projects in one live control room."
          right={<div className="page-pill-row"><span className="page-pill blue">Current shift working</span><span className="page-pill orange">{today()}</span></div>}
        />

        <div className="stats-grid">
          <StatCard title="Total fleet" value={metrics.totalFleet} detail="Registered machines" tone="blue" onClick={() => setTab('fleet')} />
          <StatCard title="Machines on breakdown" value={metrics.openBreakdowns.length} detail="Open breakdown records" tone="red" onClick={() => setTab('breakdowns')} />
          <StatCard title="Machines being worked on" value={metrics.openRepairs.length} detail="Active repairs" tone="orange" onClick={() => setTab('repairs')} />
          <StatCard title="Services due" value={metrics.servicesDue.length} detail="Service today/due" tone="purple" onClick={() => setTab('services')} />
          <StatCard title="People off" value={metrics.peopleOff.length} detail="Leave/sick/off duty" tone="yellow" onClick={() => setTab('personnel')} />
          <StatCard title="Foremen on duty" value={metrics.foremen.length} detail="Loaded foremen" tone="green" onClick={() => setTab('personnel')} />
          <StatCard title="Hose requests" value={metrics.hosePending.length} detail="Pending hoses/fittings" tone="blue" onClick={() => setTab('hoses')} />
          <StatCard title="Tool orders" value={metrics.toolOrders.length} detail="Requested tools" tone="orange" onClick={() => setTab('tools')} />
        </div>

        <div className="grid-2">
          <div className="card">
            <div className="card-head">
              <h3>Machines on breakdown / under repair</h3>
              <button onClick={() => setTab('breakdowns')}>Open breakdowns</button>
            </div>
            <MiniTable
              rows={metrics.openBreakdowns.slice(0, 8)}
              empty="No open breakdowns."
              columns={[
                { key: 'fleet_no', label: 'Fleet' },
                { key: 'fault', label: 'Fault' },
                { key: 'assigned_to', label: 'Worked by' },
                { key: 'spare_eta', label: 'Spares ETA' },
                { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
              ]}
            />
          </div>

          <div className="card">
            <div className="card-head">
              <h3>Machines on service today</h3>
              <button onClick={() => setTab('services')}>Open services</button>
            </div>
            <MiniTable
              rows={metrics.servicesDue.slice(0, 8)}
              empty="No services due."
              columns={[
                { key: 'fleet_no', label: 'Fleet' },
                { key: 'service_type', label: 'Service' },
                { key: 'due_date', label: 'Due date' },
                { key: 'supervisor', label: 'Supervisor' },
                { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
              ]}
            />
          </div>

          <div className="card">
            <div className="card-head">
              <h3>Foremen on duty</h3>
              <button onClick={() => setTab('personnel')}>Personnel</button>
            </div>
            <MiniTable
              rows={metrics.foremen.slice(0, 8)}
              empty="No foremen loaded yet."
              columns={[
                { key: 'name', label: 'Name' },
                { key: 'shift', label: 'Shift' },
                { key: 'section', label: 'Section' },
                { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }
              ]}
            />
          </div>

          <div className="card">
            <div className="card-head">
              <h3>Workshop alerts</h3>
              <button onClick={() => setTab('reports')}>Reports</button>
            </div>
            <div className="alert-strip">
              <span className="page-pill red">Low spares: {metrics.lowSpares.length}</span>
              <span className="page-pill orange">Tyres due: {metrics.tyreDue.length}</span>
              <span className="page-pill blue">Batteries due: {metrics.batteryDue.length}</span>
              <span className="page-pill green">Projects: {metrics.fabProjects.length}</span>
            </div>
          </div>
        </div>
      </div>
    )
  }

  function renderFleet() {
    return (
      <div className="page-grid">
        <SectionTitle title="Fleet Register" subtitle="Upload or add machines. Fleet numbers feed breakdowns, repairs, services, tyres, batteries, spares, hoses and reports." />
        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add machine</h3>
            <div className="form-grid">
              <Field label="Fleet number"><input value={fleetForm.fleet_no} onChange={(e) => setFleetForm({ ...fleetForm, fleet_no: e.target.value })} /></Field>
              <Field label="Machine type"><input value={fleetForm.machine_type} onChange={(e) => setFleetForm({ ...fleetForm, machine_type: e.target.value })} /></Field>
              <Field label="Make/model"><input value={fleetForm.make_model} onChange={(e) => setFleetForm({ ...fleetForm, make_model: e.target.value })} /></Field>
              <Field label="Department"><input value={fleetForm.department} onChange={(e) => setFleetForm({ ...fleetForm, department: e.target.value })} /></Field>
              <Field label="Location"><input value={fleetForm.location} onChange={(e) => setFleetForm({ ...fleetForm, location: e.target.value })} /></Field>
              <Field label="Hours"><input type="number" value={fleetForm.hours} onChange={(e) => setFleetForm({ ...fleetForm, hours: e.target.value })} /></Field>
              <Field label="Mileage"><input type="number" value={fleetForm.mileage} onChange={(e) => setFleetForm({ ...fleetForm, mileage: e.target.value })} /></Field>
              <Field label="Status"><select value={fleetForm.status} onChange={(e) => setFleetForm({ ...fleetForm, status: e.target.value })}><option>Available</option><option>Breakdown</option><option>Service</option><option>Major Repair</option></select></Field>
            </div>
            <button className="primary" onClick={() => saveRow('fleet_machines', fleetForm, 'upsert')}>Save machine</button>
          </div>

          <div className="card form-card">
            <h3>Upload fleet Excel/CSV</h3>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadFleet(e.target.files?.[0])} />
            <p className="small-note">Columns: fleet_no, machine_type, make_model, department, location, hours, mileage, status.</p>
          </div>
        </div>

        <div className="card">
          <div className="card-head">
            <h3>Fleet list</h3>
            <input placeholder="Search fleet..." value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <MiniTable
            rows={filteredFleet}
            columns={[
              { key: 'fleet_no', label: 'Fleet' },
              { key: 'machine_type', label: 'Type' },
              { key: 'make_model', label: 'Model' },
              { key: 'department', label: 'Department' },
              { key: 'hours', label: 'Hours', render: (r) => one(r.hours) },
              { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
            ]}
          />
        </div>
      </div>
    )
  }

  function renderBreakdowns() {
    return (
      <div className="page-grid breakdown-page">
        <SectionTitle title="Breakdowns" subtitle="Record breakdowns, spares ETA, expected availability, repeated faults and who worked on the machine." />
        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add breakdown</h3>
            <div className="form-grid">
              <Field label="Fleet number"><input value={breakdownForm.fleet_no} onChange={(e) => setBreakdownForm({ ...breakdownForm, fleet_no: e.target.value })} /></Field>
              <Field label="Category"><select value={breakdownForm.category} onChange={(e) => setBreakdownForm({ ...breakdownForm, category: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Fault reported"><input value={breakdownForm.fault} onChange={(e) => setBreakdownForm({ ...breakdownForm, fault: e.target.value })} /></Field>
              <Field label="Reported by"><input value={breakdownForm.reported_by} onChange={(e) => setBreakdownForm({ ...breakdownForm, reported_by: e.target.value })} /></Field>
              <Field label="Worked by / assigned to"><input value={breakdownForm.assigned_to} onChange={(e) => setBreakdownForm({ ...breakdownForm, assigned_to: e.target.value })} /></Field>
              <Field label="Start time"><input type="datetime-local" value={breakdownForm.start_time} onChange={(e) => setBreakdownForm({ ...breakdownForm, start_time: e.target.value })} /></Field>
              <Field label="Spares ETA"><input type="date" value={breakdownForm.spare_eta} onChange={(e) => setBreakdownForm({ ...breakdownForm, spare_eta: e.target.value })} /></Field>
              <Field label="Expected available"><input type="date" value={breakdownForm.expected_available_date} onChange={(e) => setBreakdownForm({ ...breakdownForm, expected_available_date: e.target.value })} /></Field>
              <Field label="Status"><select value={breakdownForm.status} onChange={(e) => setBreakdownForm({ ...breakdownForm, status: e.target.value })}><option>Open</option><option>Awaiting Spares</option><option>Being Worked On</option><option>Completed</option><option>Closed</option></select></Field>
              <Field label="Cause"><input value={breakdownForm.cause} onChange={(e) => setBreakdownForm({ ...breakdownForm, cause: e.target.value })} /></Field>
            </div>
            <Field label="Action taken"><textarea rows={3} value={breakdownForm.action_taken} onChange={(e) => setBreakdownForm({ ...breakdownForm, action_taken: e.target.value })} /></Field>
            <Field label="Preventative maintenance suggestion"><textarea rows={3} value={breakdownForm.preventative_recommendation} onChange={(e) => setBreakdownForm({ ...breakdownForm, preventative_recommendation: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('breakdowns', breakdownForm)}>Save breakdown</button>
          </div>

          <div className="card">
            <h3>Breakdown pressure points</h3>
            <div className="alert-strip">
              <span className="page-pill red">Awaiting spares: {data.breakdowns.filter((x) => String(x.status || '').includes('Spares')).length}</span>
              <span className="page-pill orange">Being worked on: {data.breakdowns.filter((x) => String(x.status || '').includes('Worked')).length}</span>
              <span className="page-pill blue">Open: {data.breakdowns.filter((x) => String(x.status || '') === 'Open').length}</span>
            </div>
          </div>
        </div>

        <div className="card">
          <div className="card-head"><h3>Breakdown register</h3><input placeholder="Search breakdowns..." value={search} onChange={(e) => setSearch(e.target.value)} /></div>
          <MiniTable
            rows={data.breakdowns.filter((r) => JSON.stringify(r).toLowerCase().includes(search.toLowerCase()))}
            columns={[
              { key: 'fleet_no', label: 'Fleet' },
              { key: 'fault', label: 'Fault' },
              { key: 'category', label: 'Area' },
              { key: 'assigned_to', label: 'Worked by' },
              { key: 'spare_eta', label: 'Spares ETA' },
              { key: 'expected_available_date', label: 'Available ETA' },
              { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
            ]}
          />
        </div>
      </div>
    )
  }

  function renderRepairs() {
    return (
      <div className="page-grid">
        <SectionTitle title="Repairs" subtitle="Record repair job cards, faults, fitters, parts used and repair history." />
        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add repair job</h3>
            <div className="form-grid">
              <Field label="Fleet number"><input value={repairForm.fleet_no} onChange={(e) => setRepairForm({ ...repairForm, fleet_no: e.target.value })} /></Field>
              <Field label="Job card number"><input value={repairForm.job_card_no} onChange={(e) => setRepairForm({ ...repairForm, job_card_no: e.target.value })} /></Field>
              <Field label="Fault"><input value={repairForm.fault} onChange={(e) => setRepairForm({ ...repairForm, fault: e.target.value })} /></Field>
              <Field label="Assigned to"><input value={repairForm.assigned_to} onChange={(e) => setRepairForm({ ...repairForm, assigned_to: e.target.value })} /></Field>
              <Field label="Start"><input type="datetime-local" value={repairForm.start_time} onChange={(e) => setRepairForm({ ...repairForm, start_time: e.target.value })} /></Field>
              <Field label="End"><input type="datetime-local" value={repairForm.end_time} onChange={(e) => setRepairForm({ ...repairForm, end_time: e.target.value })} /></Field>
              <Field label="Status"><select value={repairForm.status} onChange={(e) => setRepairForm({ ...repairForm, status: e.target.value })}><option>In progress</option><option>Awaiting Spares</option><option>Completed</option><option>Closed</option></select></Field>
              <Field label="Parts used"><input value={repairForm.parts_used} onChange={(e) => setRepairForm({ ...repairForm, parts_used: e.target.value })} /></Field>
            </div>
            <Field label="Notes"><textarea rows={3} value={repairForm.notes} onChange={(e) => setRepairForm({ ...repairForm, notes: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('repairs', repairForm)}>Save repair</button>
          </div>

          <div className="card">
            <h3>Repair summary</h3>
            <div className="stats-grid mini">
              <StatCard title="Open" value={data.repairs.filter((x) => !['Completed', 'Closed'].includes(String(x.status || ''))).length} tone="orange" />
              <StatCard title="Awaiting spares" value={data.repairs.filter((x) => String(x.status || '').includes('Spares')).length} tone="red" />
              <StatCard title="Completed" value={data.repairs.filter((x) => String(x.status || '') === 'Completed').length} tone="green" />
              <StatCard title="Job cards" value={data.repairs.filter((x) => x.job_card_no).length} tone="blue" />
            </div>
          </div>
        </div>

        <div className="card">
          <h3>Repair history</h3>
          <MiniTable rows={data.repairs} columns={[
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'job_card_no', label: 'Job card' },
            { key: 'fault', label: 'Fault' },
            { key: 'assigned_to', label: 'Worked by' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} />
        </div>
      </div>
    )
  }

  function renderServices() {
    return (
      <div className="page-grid services-page">
        <SectionTitle title="Services" subtitle="Upload service schedules and record completed services, supervisors, job cards and past service history." />
        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add service record</h3>
            <div className="form-grid">
              <Field label="Fleet number"><input value={serviceForm.fleet_no} onChange={(e) => setServiceForm({ ...serviceForm, fleet_no: e.target.value })} /></Field>
              <Field label="Service type"><input value={serviceForm.service_type} onChange={(e) => setServiceForm({ ...serviceForm, service_type: e.target.value })} /></Field>
              <Field label="Due date"><input type="date" value={serviceForm.due_date} onChange={(e) => setServiceForm({ ...serviceForm, due_date: e.target.value })} /></Field>
              <Field label="Scheduled hours"><input type="number" value={serviceForm.scheduled_hours} onChange={(e) => setServiceForm({ ...serviceForm, scheduled_hours: e.target.value })} /></Field>
              <Field label="Completed date"><input type="date" value={serviceForm.completed_date} onChange={(e) => setServiceForm({ ...serviceForm, completed_date: e.target.value })} /></Field>
              <Field label="Completed hours"><input type="number" value={serviceForm.completed_hours} onChange={(e) => setServiceForm({ ...serviceForm, completed_hours: e.target.value })} /></Field>
              <Field label="Job card"><input value={serviceForm.job_card_no} onChange={(e) => setServiceForm({ ...serviceForm, job_card_no: e.target.value })} /></Field>
              <Field label="Supervisor"><input value={serviceForm.supervisor} onChange={(e) => setServiceForm({ ...serviceForm, supervisor: e.target.value })} /></Field>
              <Field label="Technician"><input value={serviceForm.technician} onChange={(e) => setServiceForm({ ...serviceForm, technician: e.target.value })} /></Field>
              <Field label="Status"><select value={serviceForm.status} onChange={(e) => setServiceForm({ ...serviceForm, status: e.target.value })}><option>Due</option><option>In Progress</option><option>Completed</option><option>Missed</option></select></Field>
            </div>
            <Field label="Notes"><textarea rows={3} value={serviceForm.notes} onChange={(e) => setServiceForm({ ...serviceForm, notes: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('services', serviceForm)}>Save service</button>
          </div>

          <div className="card form-card">
            <h3>Upload service schedule</h3>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadService(e.target.files?.[0])} />
            <p className="small-note">Columns: fleet_no, service_type, due_date, scheduled_hours, job_card_no, supervisor, technician, status.</p>
          </div>
        </div>

        <div className="card">
          <h3>Past and due service history</h3>
          <MiniTable rows={data.services} columns={[
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'service_type', label: 'Service' },
            { key: 'due_date', label: 'Due' },
            { key: 'completed_date', label: 'Completed' },
            { key: 'job_card_no', label: 'Job card' },
            { key: 'supervisor', label: 'Supervisor' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} />
        </div>
      </div>
    )
  }

  function renderSpares() {
    const partNumbers = String(partsBookForm.extracted_part_numbers || '').split('\n').map((x) => x.trim()).filter(Boolean)

    return (
      <div className="page-grid spares-page">
        <SectionTitle title="Spares & Orders" subtitle="Request sheet, funding stages, local/international orders, delivery ETA and PDF parts book viewer." />

        <div className="stats-grid">
          <StatCard title="Requests" value={data.spares_orders.filter((x) => String(x.workflow_stage || x.status) === 'Requested').length} tone="blue" />
          <StatCard title="Awaiting funding" value={data.spares_orders.filter((x) => String(x.workflow_stage || x.status) === 'Awaiting Funding').length} tone="red" />
          <StatCard title="Funded" value={data.spares_orders.filter((x) => String(x.workflow_stage || x.status) === 'Funded').length} tone="green" />
          <StatCard title="Awaiting delivery" value={data.spares_orders.filter((x) => String(x.workflow_stage || x.status) === 'Awaiting Delivery').length} tone="orange" />
        </div>

        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Spares request sheet</h3>
            <div className="form-grid">
              <Field label="Date"><input type="date" value={orderForm.request_date} onChange={(e) => setOrderForm({ ...orderForm, request_date: e.target.value })} /></Field>
              <Field label="Fleet number"><input value={orderForm.machine_fleet_no} onChange={(e) => setOrderForm({ ...orderForm, machine_fleet_no: e.target.value })} /></Field>
              <Field label="Requested by"><input value={orderForm.requested_by} onChange={(e) => setOrderForm({ ...orderForm, requested_by: e.target.value })} /></Field>
              <Field label="Machine group"><select value={orderForm.machine_group} onChange={(e) => setOrderForm({ ...orderForm, machine_group: e.target.value })}>{machineGroups.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Workshop section"><select value={orderForm.workshop_section} onChange={(e) => setOrderForm({ ...orderForm, workshop_section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Order type"><select value={orderForm.order_type} onChange={(e) => setOrderForm({ ...orderForm, order_type: e.target.value })}>{orderTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Priority"><select value={orderForm.priority} onChange={(e) => setOrderForm({ ...orderForm, priority: e.target.value })}>{priorities.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Workflow stage"><select value={orderForm.workflow_stage} onChange={(e) => setOrderForm({ ...orderForm, workflow_stage: e.target.value, status: e.target.value })}>{orderStages.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Part number"><input value={orderForm.part_no} onChange={(e) => setOrderForm({ ...orderForm, part_no: e.target.value })} /></Field>
              <Field label="Quantity"><input type="number" value={orderForm.qty} onChange={(e) => setOrderForm({ ...orderForm, qty: e.target.value })} /></Field>
              <Field label="ETA"><input type="date" value={orderForm.eta} onChange={(e) => setOrderForm({ ...orderForm, eta: e.target.value })} /></Field>
              <Field label="Description"><input value={orderForm.description} onChange={(e) => setOrderForm({ ...orderForm, description: e.target.value })} /></Field>
            </div>
            <Field label="Spares list / multiple items"><textarea rows={5} value={orderForm.spares_items} onChange={(e) => setOrderForm({ ...orderForm, spares_items: e.target.value })} /></Field>
            <Field label="Part numbers from book"><textarea rows={5} value={orderForm.part_numbers} onChange={(e) => setOrderForm({ ...orderForm, part_numbers: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('spares_orders', orderForm)}>Save spares request</button>
          </div>

          <div className="card form-card">
            <h3>PDF parts books</h3>
            <div className="form-grid">
              <Field label="Book title"><input value={partsBookForm.book_title} onChange={(e) => setPartsBookForm({ ...partsBookForm, book_title: e.target.value })} /></Field>
              <Field label="Machine group"><select value={partsBookForm.machine_group} onChange={(e) => setPartsBookForm({ ...partsBookForm, machine_group: e.target.value })}>{machineGroups.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Machine type"><input value={partsBookForm.machine_type} onChange={(e) => setPartsBookForm({ ...partsBookForm, machine_type: e.target.value })} /></Field>
              <Field label="Uploaded by"><input value={partsBookForm.uploaded_by} onChange={(e) => setPartsBookForm({ ...partsBookForm, uploaded_by: e.target.value })} /></Field>
            </div>
            <input type="file" accept="application/pdf,.pdf" onChange={(e) => attachPartsBook(e.target.files?.[0])} />
            {partsBookForm.file_name && <div className="file-note">Loaded: {partsBookForm.file_name} — {partNumbers.length} possible part numbers detected.</div>}
            <div className="part-chip-box">
              {partNumbers.slice(0, 40).map((part) => <button key={part} className="part-chip" onClick={() => addPartToOrder(part)}>{part}</button>)}
            </div>
            <button className="primary" onClick={() => {
              const { object_url, file_data, ...bookPayload } = partsBookForm
              saveRow('parts_books', bookPayload)
            }}>Save parts book</button>
          </div>
        </div>

        <details className="drop-panel" open>
          <summary>Order board by workflow stage</summary>
          <div className="filters">
            <input placeholder="Search orders..." value={sparesSearch} onChange={(e) => setSparesSearch(e.target.value)} />
            <select value={sparesStage} onChange={(e) => setSparesStage(e.target.value)}><option>All</option>{orderStages.map((x) => <option key={x}>{x}</option>)}</select>
          </div>
          <div className="kanban-grid">
            {orderStages.slice(0, 5).map((stage) => (
              <div className="card kanban-drop" key={stage}>
                <h3>{stage}</h3>
                <div className="kanban-cards">
                  {filteredOrders.filter((o) => String(o.workflow_stage || o.status) === stage).map((order) => (
                    <div className="kanban-card" key={order.id}>
                      <div className="card-row"><b>{order.machine_fleet_no || 'No fleet'}</b><Badge value={order.priority} /></div>
                      <span>{order.description || order.part_no || 'Spares request'}</span>
                      <div className="meta-grid compact">
                        <span>By: {order.requested_by || '-'}</span>
                        <span>ETA: {order.eta || '-'}</span>
                        <span>{order.machine_group || '-'}</span>
                        <span>{order.order_type || '-'}</span>
                      </div>
                      <div className="button-row wrap">
                        {orderStages.map((nextStage) => <button key={nextStage} onClick={() => saveRow('spares_orders', { ...order, workflow_stage: nextStage, status: nextStage }, 'upsert')}>{nextStage}</button>)}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </details>

        <div className="grid-2">
          <div className="card">
            <h3>Saved PDF parts books</h3>
            <MiniTable
              rows={data.parts_books}
              empty="No parts books saved yet."
              onRow={openPartsBook}
              columns={[
                { key: 'book_title', label: 'Book' },
                { key: 'machine_group', label: 'Group' },
                { key: 'machine_type', label: 'Machine' },
                { key: 'file_name', label: 'File' },
                { key: 'extracted_count', label: 'Parts found' }
              ]}
            />
          </div>

          <div className="card form-card">
            <h3>Spares stock levels</h3>
            <div className="form-grid">
              <Field label="Part number"><input value={spareForm.part_no} onChange={(e) => setSpareForm({ ...spareForm, part_no: e.target.value })} /></Field>
              <Field label="Description"><input value={spareForm.description} onChange={(e) => setSpareForm({ ...spareForm, description: e.target.value })} /></Field>
              <Field label="Stock qty"><input type="number" value={spareForm.stock_qty} onChange={(e) => setSpareForm({ ...spareForm, stock_qty: e.target.value })} /></Field>
              <Field label="Min qty"><input type="number" value={spareForm.min_qty} onChange={(e) => setSpareForm({ ...spareForm, min_qty: e.target.value })} /></Field>
              <Field label="Supplier"><input value={spareForm.supplier} onChange={(e) => setSpareForm({ ...spareForm, supplier: e.target.value })} /></Field>
              <Field label="Shelf"><input value={spareForm.shelf_location} onChange={(e) => setSpareForm({ ...spareForm, shelf_location: e.target.value })} /></Field>
            </div>
            <button className="primary" onClick={() => saveRow('spares', spareForm)}>Save stock item</button>
          </div>
        </div>
      </div>
    )
  }

  function renderPersonnel() {
    const currentLeave = data.leave_records.filter((l) => isBetweenDates(l.start_date, l.end_date))
    const upcomingLeave = data.leave_records.filter((l) => isFutureDate(l.start_date))
    const pastLeave = data.leave_records.filter((l) => isPastDate(l.end_date))

    return (
      <div className="page-grid personnel-page">
        <SectionTitle title="Personnel, Shifts and Leave" subtitle="Shift 1, Shift 2, Shift 3, managers, foremen, trades, off schedule and leave history." />

        <div className="stats-grid">
          <StatCard title="Employees" value={data.personnel.length} tone="blue" />
          <StatCard title="Current leave/off" value={currentLeave.length} tone="orange" />
          <StatCard title="Upcoming leave" value={upcomingLeave.length} tone="purple" />
          <StatCard title="Past leave records" value={pastLeave.length} tone="green" />
        </div>

        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add employee / slot</h3>
            <div className="form-grid">
              <Field label="Employee no"><input value={personForm.employee_no} onChange={(e) => setPersonForm({ ...personForm, employee_no: e.target.value })} /></Field>
              <Field label="Name"><input value={personForm.name} onChange={(e) => setPersonForm({ ...personForm, name: e.target.value })} /></Field>
              <Field label="Shift"><select value={personForm.shift} onChange={(e) => setPersonForm({ ...personForm, shift: e.target.value })}>{shiftNames.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Role"><select value={personForm.role} onChange={(e) => setPersonForm({ ...personForm, role: e.target.value })}>{personnelRoles.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Section"><select value={personForm.section} onChange={(e) => setPersonForm({ ...personForm, section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Status"><select value={personForm.employment_status} onChange={(e) => setPersonForm({ ...personForm, employment_status: e.target.value })}><option>Active</option><option>On Leave</option><option>Off Duty</option><option>Sick Leave</option><option>Training</option><option>Vacant</option></select></Field>
              <Field label="Leave balance"><input type="number" value={personForm.leave_balance_days} onChange={(e) => setPersonForm({ ...personForm, leave_balance_days: e.target.value })} /></Field>
              <Field label="Leave taken"><input type="number" value={personForm.leave_taken_days} onChange={(e) => setPersonForm({ ...personForm, leave_taken_days: e.target.value })} /></Field>
            </div>
            <button className="primary" onClick={() => saveRow('personnel', personForm, 'upsert')}>Save employee</button>
          </div>

          <div className="card form-card">
            <h3>Upload personnel / leave Excel</h3>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => uploadPersonnel(e.target.files?.[0])} />
            <h3>Add manpower by quantity</h3>
            <div className="form-grid">
              <Field label="Shift"><select value={bulkPersonForm.shift} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, shift: e.target.value })}>{shiftNames.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Role"><select value={bulkPersonForm.role} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, role: e.target.value })}>{personnelRoles.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Section"><select value={bulkPersonForm.section} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Quantity"><input type="number" value={bulkPersonForm.qty} onChange={(e) => setBulkPersonForm({ ...bulkPersonForm, qty: e.target.value })} /></Field>
            </div>
            <button onClick={addBulkPeople}>Add slots</button>
          </div>
        </div>

        <div className="card form-card">
          <h3>Leave / off schedule</h3>
          <div className="form-grid">
            <Field label="Employee"><input list="people-list" value={leaveForm.person_name} onChange={(e) => setLeaveForm({ ...leaveForm, person_name: e.target.value })} /></Field>
            <Field label="Employee no"><input value={leaveForm.employee_no} onChange={(e) => setLeaveForm({ ...leaveForm, employee_no: e.target.value })} /></Field>
            <Field label="Leave type"><select value={leaveForm.leave_type} onChange={(e) => setLeaveForm({ ...leaveForm, leave_type: e.target.value })}>{leaveTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Start date"><input type="date" value={leaveForm.start_date} onChange={(e) => setLeaveForm({ ...leaveForm, start_date: e.target.value })} /></Field>
            <Field label="End date"><input type="date" value={leaveForm.end_date} onChange={(e) => setLeaveForm({ ...leaveForm, end_date: e.target.value })} /></Field>
            <Field label="Status"><select value={leaveForm.status} onChange={(e) => setLeaveForm({ ...leaveForm, status: e.target.value })}><option>Approved</option><option>Pending</option><option>Rejected</option></select></Field>
          </div>
          <Field label="Reason"><textarea rows={2} value={leaveForm.reason} onChange={(e) => setLeaveForm({ ...leaveForm, reason: e.target.value })} /></Field>
          <button className="primary" onClick={saveLeaveRecord}>Save leave/off record and deduct leave</button>
        </div>

        <datalist id="people-list">{data.personnel.map((p) => <option key={p.id} value={p.name} />)}</datalist>

        {shiftNames.map((shift) => (
          <details className="drop-panel" open key={shift}>
            <summary>{shift} personnel</summary>
            <MiniTable rows={data.personnel.filter((p) => p.shift === shift)} columns={[
              { key: 'name', label: 'Name' },
              { key: 'role', label: 'Role' },
              { key: 'section', label: 'Section' },
              { key: 'leave_balance_days', label: 'Leave bal' },
              { key: 'leave_taken_days', label: 'Leave taken' },
              { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }
            ]} />
          </details>
        ))}

        <div className="grid-2">
          <div className="card">
            <h3>Current leave / off</h3>
            <MiniTable rows={currentLeave} columns={[
              { key: 'person_name', label: 'Name' },
              { key: 'leave_type', label: 'Type' },
              { key: 'start_date', label: 'Start' },
              { key: 'end_date', label: 'End' },
              { key: 'days', label: 'Days' },
              { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
            ]} />
          </div>
          <div className="card">
            <h3>Upcoming leave</h3>
            <MiniTable rows={upcomingLeave} columns={[
              { key: 'person_name', label: 'Name' },
              { key: 'leave_type', label: 'Type' },
              { key: 'start_date', label: 'Start' },
              { key: 'end_date', label: 'End' },
              { key: 'days', label: 'Days' }
            ]} />
          </div>
        </div>
      </div>
    )
  }

  function renderTyres() {
    return (
      <div className="page-grid tyres-page">
        <SectionTitle title="Tyres" subtitle="Track tyre make, serial, company number, fitment, fleet, hours/mileage, measurements, repairs, stock and tyre orders." />

        <div className="stats-grid">
          <StatCard title="Tyres" value={data.tyres.length} tone="blue" />
          <StatCard title="Measurements due" value={metrics.tyreDue.length} tone="red" />
          <StatCard title="Tyre orders" value={data.tyre_orders.length} tone="orange" />
          <StatCard title="Measurements" value={data.tyre_measurements.length} tone="green" />
        </div>

        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add tyre record</h3>
            <div className="form-grid">
              <Field label="Make"><input value={tyreForm.make} onChange={(e) => setTyreForm({ ...tyreForm, make: e.target.value })} /></Field>
              <Field label="Serial number"><input value={tyreForm.serial_no} onChange={(e) => setTyreForm({ ...tyreForm, serial_no: e.target.value })} /></Field>
              <Field label="Company number"><input value={tyreForm.company_no} onChange={(e) => setTyreForm({ ...tyreForm, company_no: e.target.value })} /></Field>
              <Field label="Production date"><input type="date" value={tyreForm.production_date} onChange={(e) => setTyreForm({ ...tyreForm, production_date: e.target.value })} /></Field>
              <Field label="Date fitted"><input type="date" value={tyreForm.fitted_date} onChange={(e) => setTyreForm({ ...tyreForm, fitted_date: e.target.value })} /></Field>
              <Field label="Fleet number"><input list="fleet-list" value={tyreForm.fleet_no} onChange={(e) => setTyreForm({ ...tyreForm, fleet_no: e.target.value })} /></Field>
              <Field label="Fitted by"><input value={tyreForm.fitted_by} onChange={(e) => setTyreForm({ ...tyreForm, fitted_by: e.target.value })} /></Field>
              <Field label="Position"><select value={tyreForm.position} onChange={(e) => setTyreForm({ ...tyreForm, position: e.target.value })}>{tyrePositions.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Tyre type"><input value={tyreForm.tyre_type} onChange={(e) => setTyreForm({ ...tyreForm, tyre_type: e.target.value })} /></Field>
              <Field label="Size"><input value={tyreForm.size} onChange={(e) => setTyreForm({ ...tyreForm, size: e.target.value })} /></Field>
              <Field label="Fitment hours"><input type="number" value={tyreForm.fitment_hours} onChange={(e) => setTyreForm({ ...tyreForm, fitment_hours: e.target.value })} /></Field>
              <Field label="Fitment mileage"><input type="number" value={tyreForm.fitment_mileage} onChange={(e) => setTyreForm({ ...tyreForm, fitment_mileage: e.target.value })} /></Field>
              <Field label="Current tread mm"><input type="number" value={tyreForm.current_tread_mm} onChange={(e) => setTyreForm({ ...tyreForm, current_tread_mm: e.target.value })} /></Field>
              <Field label="Measurement reminder"><input type="date" value={tyreForm.measurement_due_date} onChange={(e) => setTyreForm({ ...tyreForm, measurement_due_date: e.target.value })} /></Field>
              <Field label="Stock status"><select value={tyreForm.stock_status} onChange={(e) => setTyreForm({ ...tyreForm, stock_status: e.target.value })}><option>Fitted</option><option>In Stock</option><option>Scrapped</option><option>Repair</option></select></Field>
              <Field label="Photo"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setTyreForm({ ...tyreForm, fitment_photo: v }))} /></Field>
            </div>
            <Field label="Notes"><textarea rows={3} value={tyreForm.notes} onChange={(e) => setTyreForm({ ...tyreForm, notes: e.target.value })} /></Field>
            {tyreForm.fitment_photo && <img className="thumb" src={tyreForm.fitment_photo} alt="Tyre fitment" />}
            <button className="primary" onClick={() => saveRow('tyres', tyreForm)}>Save tyre</button>
          </div>

          <div className="card form-card">
            <h3>Tyre measurement / repair</h3>
            <div className="form-grid">
              <Field label="Tyre serial"><input value={tyreMeasurementForm.tyre_serial_no} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, tyre_serial_no: e.target.value })} /></Field>
              <Field label="Fleet number"><input value={tyreMeasurementForm.fleet_no} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, fleet_no: e.target.value })} /></Field>
              <Field label="Position"><input value={tyreMeasurementForm.position} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, position: e.target.value })} /></Field>
              <Field label="Date"><input type="date" value={tyreMeasurementForm.measurement_date} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, measurement_date: e.target.value })} /></Field>
              <Field label="Measured/repaired by"><input value={tyreMeasurementForm.measured_by} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, measured_by: e.target.value })} /></Field>
              <Field label="Tread depth mm"><input type="number" value={tyreMeasurementForm.tread_depth_mm} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, tread_depth_mm: e.target.value })} /></Field>
              <Field label="Pressure PSI"><input type="number" value={tyreMeasurementForm.pressure_psi} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, pressure_psi: e.target.value })} /></Field>
              <Field label="Condition"><select value={tyreMeasurementForm.condition} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, condition: e.target.value })}><option>Good</option><option>Worn</option><option>Cut</option><option>Puncture</option><option>Repaired</option><option>Scrap</option></select></Field>
            </div>
            <Field label="Reason / repair notes"><textarea rows={3} value={tyreMeasurementForm.notes} onChange={(e) => setTyreMeasurementForm({ ...tyreMeasurementForm, notes: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('tyre_measurements', tyreMeasurementForm)}>Save measurement/repair</button>

            <h3>Tyre order request</h3>
            <div className="form-grid">
              <Field label="Fleet"><input value={tyreOrderForm.fleet_no} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, fleet_no: e.target.value })} /></Field>
              <Field label="Tyre type"><input value={tyreOrderForm.tyre_type} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, tyre_type: e.target.value })} /></Field>
              <Field label="Size"><input value={tyreOrderForm.size} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, size: e.target.value })} /></Field>
              <Field label="Qty"><input type="number" value={tyreOrderForm.qty} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, qty: e.target.value })} /></Field>
              <Field label="Requested by"><input value={tyreOrderForm.requested_by} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, requested_by: e.target.value })} /></Field>
              <Field label="Status"><select value={tyreOrderForm.status} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, status: e.target.value })}><option>Requested</option><option>Need Quote</option><option>Approved</option><option>Ordered</option><option>Delivered</option></select></Field>
            </div>
            <Field label="Reason"><textarea rows={2} value={tyreOrderForm.reason} onChange={(e) => setTyreOrderForm({ ...tyreOrderForm, reason: e.target.value })} /></Field>
            <button onClick={() => saveRow('tyre_orders', tyreOrderForm)}>Notify admin / save tyre order</button>
          </div>
        </div>

        <div className="grid-2">
          <div className="card"><h3>Tyre register</h3><MiniTable rows={data.tyres} columns={[
            { key: 'serial_no', label: 'Serial' },
            { key: 'company_no', label: 'Company no' },
            { key: 'make', label: 'Make' },
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'position', label: 'Position' },
            { key: 'measurement_due_date', label: 'Reminder' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status || r.stock_status} /> }
          ]} /></div>
          <div className="card"><h3>Tyre orders</h3><MiniTable rows={data.tyre_orders} columns={[
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'tyre_type', label: 'Type' },
            { key: 'size', label: 'Size' },
            { key: 'qty', label: 'Qty' },
            { key: 'requested_by', label: 'By' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
        </div>
      </div>
    )
  }

  function renderBatteries() {
    return (
      <div className="page-grid batteries-page">
        <SectionTitle title="Batteries" subtitle="Track battery make, serial, production date, fitment, charging voltage, stock and battery orders." />

        <div className="stats-grid">
          <StatCard title="Batteries" value={data.batteries.length} tone="blue" />
          <StatCard title="Maintenance due" value={metrics.batteryDue.length} tone="red" />
          <StatCard title="Battery orders" value={data.battery_orders.length} tone="orange" />
          <StatCard title="In stock" value={data.batteries.filter((b) => String(b.stock_status || '').includes('Stock')).length} tone="green" />
        </div>

        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add battery</h3>
            <div className="form-grid">
              <Field label="Make"><input value={batteryForm.make} onChange={(e) => setBatteryForm({ ...batteryForm, make: e.target.value })} /></Field>
              <Field label="Serial number"><input value={batteryForm.serial_no} onChange={(e) => setBatteryForm({ ...batteryForm, serial_no: e.target.value })} /></Field>
              <Field label="Production date"><input type="date" value={batteryForm.production_date} onChange={(e) => setBatteryForm({ ...batteryForm, production_date: e.target.value })} /></Field>
              <Field label="Date fitted"><input type="date" value={batteryForm.fitment_date} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_date: e.target.value })} /></Field>
              <Field label="Fleet number"><input list="fleet-list" value={batteryForm.fleet_no} onChange={(e) => setBatteryForm({ ...batteryForm, fleet_no: e.target.value })} /></Field>
              <Field label="Fitted by"><input value={batteryForm.fitted_by} onChange={(e) => setBatteryForm({ ...batteryForm, fitted_by: e.target.value })} /></Field>
              <Field label="Battery voltage"><select value={batteryForm.volts} onChange={(e) => setBatteryForm({ ...batteryForm, volts: e.target.value })}>{batteryVoltages.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Charging system voltage"><select value={batteryForm.charging_voltage} onChange={(e) => setBatteryForm({ ...batteryForm, charging_voltage: e.target.value })}>{batteryVoltages.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Fitment hours"><input type="number" value={batteryForm.fitment_hours} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_hours: e.target.value })} /></Field>
              <Field label="Fitment mileage"><input type="number" value={batteryForm.fitment_mileage} onChange={(e) => setBatteryForm({ ...batteryForm, fitment_mileage: e.target.value })} /></Field>
              <Field label="Maintenance reminder"><input type="date" value={batteryForm.maintenance_due_date} onChange={(e) => setBatteryForm({ ...batteryForm, maintenance_due_date: e.target.value })} /></Field>
              <Field label="Stock status"><select value={batteryForm.stock_status} onChange={(e) => setBatteryForm({ ...batteryForm, stock_status: e.target.value })}><option>Fitted</option><option>In Stock</option><option>Charging</option><option>Failed</option><option>Scrapped</option></select></Field>
              <Field label="Fitment photo"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setBatteryForm({ ...batteryForm, fitment_photo: v }))} /></Field>
              <Field label="Status"><select value={batteryForm.status} onChange={(e) => setBatteryForm({ ...batteryForm, status: e.target.value })}><option>Fitted</option><option>In Stock</option><option>Maintenance Due</option><option>Failed</option><option>Replaced</option></select></Field>
            </div>
            <Field label="Notes"><textarea rows={3} value={batteryForm.notes} onChange={(e) => setBatteryForm({ ...batteryForm, notes: e.target.value })} /></Field>
            {batteryForm.fitment_photo && <img className="thumb" src={batteryForm.fitment_photo} alt="Battery fitment" />}
            <button className="primary" onClick={() => saveRow('batteries', batteryForm)}>Save battery</button>
          </div>

          <div className="card form-card">
            <h3>Battery order request</h3>
            <div className="form-grid">
              <Field label="Fleet number"><input value={batteryOrderForm.fleet_no} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, fleet_no: e.target.value })} /></Field>
              <Field label="Battery type"><input value={batteryOrderForm.battery_type} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, battery_type: e.target.value })} /></Field>
              <Field label="Voltage"><select value={batteryOrderForm.voltage} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, voltage: e.target.value })}>{batteryVoltages.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Qty"><input type="number" value={batteryOrderForm.qty} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, qty: e.target.value })} /></Field>
              <Field label="Requested by"><input value={batteryOrderForm.requested_by} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, requested_by: e.target.value })} /></Field>
              <Field label="Priority"><select value={batteryOrderForm.priority} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, priority: e.target.value })}>{priorities.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Status"><select value={batteryOrderForm.status} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, status: e.target.value })}><option>Requested</option><option>Approved</option><option>Ordered</option><option>Delivered</option><option>Cancelled</option></select></Field>
              <Field label="ETA"><input type="date" value={batteryOrderForm.eta} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, eta: e.target.value })} /></Field>
            </div>
            <Field label="Reason"><textarea rows={3} value={batteryOrderForm.reason} onChange={(e) => setBatteryOrderForm({ ...batteryOrderForm, reason: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('battery_orders', batteryOrderForm)}>Notify admin / save battery order</button>
          </div>
        </div>

        <div className="grid-2">
          <div className="card"><h3>Battery register</h3><MiniTable rows={data.batteries} columns={[
            { key: 'serial_no', label: 'Serial' },
            { key: 'make', label: 'Make' },
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'volts', label: 'Volts' },
            { key: 'charging_voltage', label: 'Charging' },
            { key: 'maintenance_due_date', label: 'Reminder' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
          <div className="card"><h3>Battery orders</h3><MiniTable rows={data.battery_orders} columns={[
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'battery_type', label: 'Type' },
            { key: 'voltage', label: 'Voltage' },
            { key: 'qty', label: 'Qty' },
            { key: 'requested_by', label: 'By' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
        </div>
      </div>
    )
  }

  function renderOperations() {
    return (
      <div className="page-grid operations-page">
        <SectionTitle title="Operations" subtitle="Separate operators by truck/yellow machine type, record history, offences, damages, abuse photos, conditions and score." />
        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Operator score sheet</h3>
            <div className="form-grid">
              <Field label="Operator name"><input value={operationForm.operator_name} onChange={(e) => setOperationForm({ ...operationForm, operator_name: e.target.value })} /></Field>
              <Field label="Employee no"><input value={operationForm.employee_no} onChange={(e) => setOperationForm({ ...operationForm, employee_no: e.target.value })} /></Field>
              <Field label="Machine group"><select value={operationForm.machine_group} onChange={(e) => setOperationForm({ ...operationForm, machine_group: e.target.value })}><option>Truck</option><option>Yellow Machine</option></select></Field>
              <Field label="Machine type"><select value={operationForm.machine_type} onChange={(e) => setOperationForm({ ...operationForm, machine_type: e.target.value })}>{operationMachineTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Fleet number"><input value={operationForm.fleet_no} onChange={(e) => setOperationForm({ ...operationForm, fleet_no: e.target.value })} /></Field>
              <Field label="Shift"><select value={operationForm.shift} onChange={(e) => setOperationForm({ ...operationForm, shift: e.target.value })}>{shiftNames.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Score"><input type="number" value={operationForm.score} onChange={(e) => setOperationForm({ ...operationForm, score: e.target.value })} /></Field>
              <Field label="Supervisor"><input value={operationForm.supervisor} onChange={(e) => setOperationForm({ ...operationForm, supervisor: e.target.value })} /></Field>
              <Field label="Event date"><input type="date" value={operationForm.event_date} onChange={(e) => setOperationForm({ ...operationForm, event_date: e.target.value })} /></Field>
              <Field label="Photo"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setOperationForm({ ...operationForm, photo_data: v }))} /></Field>
            </div>
            <Field label="Offence"><textarea rows={2} value={operationForm.offence} onChange={(e) => setOperationForm({ ...operationForm, offence: e.target.value })} /></Field>
            <Field label="Damage / abuse report / rough conditions"><textarea rows={3} value={operationForm.damage_report} onChange={(e) => setOperationForm({ ...operationForm, damage_report: e.target.value })} /></Field>
            {operationForm.photo_data && <img className="thumb" src={operationForm.photo_data} alt="Operation evidence" />}
            <button className="primary" onClick={() => saveRow('operations', operationForm)}>Save operator record</button>
          </div>

          <div className="card">
            <h3>Operator scoring guide</h3>
            <div className="alert-strip">
              <span className="page-pill green">90–100 Good</span>
              <span className="page-pill orange">70–89 Watch</span>
              <span className="page-pill red">Below 70 High risk</span>
            </div>
            <p className="muted">Use supervisor photos and notes to build a fair history of damages, abuse and rough working conditions.</p>
          </div>
        </div>

        <div className="card"><h3>Operating history</h3><MiniTable rows={data.operations} columns={[
          { key: 'operator_name', label: 'Operator' },
          { key: 'machine_group', label: 'Group' },
          { key: 'machine_type', label: 'Type' },
          { key: 'fleet_no', label: 'Fleet' },
          { key: 'score', label: 'Score' },
          { key: 'offence', label: 'Offence' },
          { key: 'supervisor', label: 'Supervisor' }
        ]} /></div>
      </div>
    )
  }

  function renderHoses() {
    return (
      <div className="page-grid hoses-page">
        <SectionTitle title="Hoses & Fittings" subtitle="Bulk hose/fitting requests, individual hose repairs, offsite repairs, admin approval and stock levels." />

        <div className="stats-grid">
          <StatCard title="Hose stock items" value={data.hoses.length} tone="blue" />
          <StatCard title="Open requests" value={metrics.hosePending.length} tone="orange" />
          <StatCard title="Awaiting approval" value={data.hose_requests.filter((h) => String(h.admin_approval || '') === 'Pending').length} tone="red" />
          <StatCard title="Offsite repairs" value={data.hose_requests.filter((h) => String(h.offsite_repair || '') === 'Yes').length} tone="purple" />
        </div>

        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Hose / fitting request</h3>
            <div className="form-grid">
              <Field label="Request date"><input type="date" value={hoseRequestForm.request_date} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, request_date: e.target.value })} /></Field>
              <Field label="Fleet number"><input value={hoseRequestForm.fleet_no} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, fleet_no: e.target.value })} /></Field>
              <Field label="Requested by"><input value={hoseRequestForm.requested_by} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, requested_by: e.target.value })} /></Field>
              <Field label="Hose type"><select value={hoseRequestForm.hose_type} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, hose_type: e.target.value })}>{hoseTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Hose no"><input value={hoseRequestForm.hose_no} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, hose_no: e.target.value })} /></Field>
              <Field label="Size"><input value={hoseRequestForm.size} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, size: e.target.value })} /></Field>
              <Field label="Length"><input value={hoseRequestForm.length} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, length: e.target.value })} /></Field>
              <Field label="Qty"><input type="number" value={hoseRequestForm.qty} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, qty: e.target.value })} /></Field>
              <Field label="Repair type"><select value={hoseRequestForm.repair_type} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, repair_type: e.target.value })}><option>New hose</option><option>Repair hose</option><option>Fittings only</option><option>Bulk order</option></select></Field>
              <Field label="Offsite repair"><select value={hoseRequestForm.offsite_repair} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, offsite_repair: e.target.value })}><option>No</option><option>Yes</option></select></Field>
              <Field label="Admin approval"><select value={hoseRequestForm.admin_approval} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, admin_approval: e.target.value })}><option>Pending</option><option>Approved</option><option>Rejected</option></select></Field>
              <Field label="Status"><select value={hoseRequestForm.status} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, status: e.target.value })}>{hoseStages.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="ETA"><input type="date" value={hoseRequestForm.eta} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, eta: e.target.value })} /></Field>
              <Field label="Fittings"><input value={hoseRequestForm.fittings} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, fittings: e.target.value })} /></Field>
            </div>
            <Field label="Reason / notes"><textarea rows={3} value={hoseRequestForm.reason} onChange={(e) => setHoseRequestForm({ ...hoseRequestForm, reason: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('hose_requests', hoseRequestForm)}>Save hose request</button>
          </div>

          <div className="card form-card">
            <h3>Hose stock</h3>
            <div className="form-grid">
              <Field label="Hose no"><input value={hoseStockForm.hose_no} onChange={(e) => setHoseStockForm({ ...hoseStockForm, hose_no: e.target.value })} /></Field>
              <Field label="Hose type"><select value={hoseStockForm.hose_type} onChange={(e) => setHoseStockForm({ ...hoseStockForm, hose_type: e.target.value })}>{hoseTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Size"><input value={hoseStockForm.size} onChange={(e) => setHoseStockForm({ ...hoseStockForm, size: e.target.value })} /></Field>
              <Field label="Length"><input value={hoseStockForm.length} onChange={(e) => setHoseStockForm({ ...hoseStockForm, length: e.target.value })} /></Field>
              <Field label="Fitting A"><select value={hoseStockForm.fitting_a} onChange={(e) => setHoseStockForm({ ...hoseStockForm, fitting_a: e.target.value })}>{fittingTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Fitting B"><select value={hoseStockForm.fitting_b} onChange={(e) => setHoseStockForm({ ...hoseStockForm, fitting_b: e.target.value })}>{fittingTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field label="Stock qty"><input type="number" value={hoseStockForm.stock_qty} onChange={(e) => setHoseStockForm({ ...hoseStockForm, stock_qty: e.target.value })} /></Field>
              <Field label="Min qty"><input type="number" value={hoseStockForm.min_qty} onChange={(e) => setHoseStockForm({ ...hoseStockForm, min_qty: e.target.value })} /></Field>
              <Field label="Shelf"><input value={hoseStockForm.shelf_location} onChange={(e) => setHoseStockForm({ ...hoseStockForm, shelf_location: e.target.value })} /></Field>
              <Field label="Status"><select value={hoseStockForm.status} onChange={(e) => setHoseStockForm({ ...hoseStockForm, status: e.target.value })}><option>In stock</option><option>Low stock</option><option>Out of stock</option></select></Field>
            </div>
            <button className="primary" onClick={() => saveRow('hoses', hoseStockForm)}>Save hose stock</button>
          </div>
        </div>

        <div className="grid-2">
          <div className="card"><h3>Hose requests</h3><MiniTable rows={data.hose_requests} columns={[
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'requested_by', label: 'By' },
            { key: 'hose_type', label: 'Type' },
            { key: 'size', label: 'Size' },
            { key: 'offsite_repair', label: 'Offsite' },
            { key: 'admin_approval', label: 'Approval', render: (r) => <Badge value={r.admin_approval} /> },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
          <div className="card"><h3>Hose/fitting stock</h3><MiniTable rows={data.hoses} columns={[
            { key: 'hose_no', label: 'Hose no' },
            { key: 'hose_type', label: 'Type' },
            { key: 'size', label: 'Size' },
            { key: 'stock_qty', label: 'Qty' },
            { key: 'min_qty', label: 'Min' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
        </div>
      </div>
    )
  }

  function renderTools() {
    return (
      <div className="page-grid tools-page">
        <SectionTitle title="Tools" subtitle="Tool inventory, quantities, brands, tool orders and employee-issued tool register." />

        <div className="stats-grid">
          <StatCard title="Tools in register" value={data.tools.length} tone="blue" />
          <StatCard title="Tool orders" value={metrics.toolOrders.length} tone="orange" />
          <StatCard title="Issued tools" value={data.employee_tools.filter((t) => String(t.status || '') === 'Issued').length} tone="purple" />
          <StatCard title="Returned tools" value={data.employee_tools.filter((t) => String(t.status || '') === 'Returned').length} tone="green" />
        </div>

        <div className="grid-3">
          <div className="card form-card accent-card">
            <h3>Tool inventory</h3>
            <Field label="Tool name"><input value={toolForm.tool_name} onChange={(e) => setToolForm({ ...toolForm, tool_name: e.target.value })} /></Field>
            <Field label="Category"><select value={toolForm.category} onChange={(e) => setToolForm({ ...toolForm, category: e.target.value })}>{toolCategories.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Brand"><input value={toolForm.brand} onChange={(e) => setToolForm({ ...toolForm, brand: e.target.value })} /></Field>
            <Field label="Serial no"><input value={toolForm.serial_no} onChange={(e) => setToolForm({ ...toolForm, serial_no: e.target.value })} /></Field>
            <Field label="Qty"><input type="number" value={toolForm.stock_qty} onChange={(e) => setToolForm({ ...toolForm, stock_qty: e.target.value })} /></Field>
            <Field label="Section"><select value={toolForm.workshop_section} onChange={(e) => setToolForm({ ...toolForm, workshop_section: e.target.value })}>{workshopSections.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Condition"><select value={toolForm.condition} onChange={(e) => setToolForm({ ...toolForm, condition: e.target.value })}><option>Good</option><option>Fair</option><option>Damaged</option><option>Lost</option></select></Field>
            <button className="primary" onClick={() => saveRow('tools', toolForm)}>Save tool</button>
          </div>

          <div className="card form-card">
            <h3>Tool order</h3>
            <Field label="Tool name"><input value={toolOrderForm.tool_name} onChange={(e) => setToolOrderForm({ ...toolOrderForm, tool_name: e.target.value })} /></Field>
            <Field label="Category"><select value={toolOrderForm.category} onChange={(e) => setToolOrderForm({ ...toolOrderForm, category: e.target.value })}>{toolCategories.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Brand"><input value={toolOrderForm.brand} onChange={(e) => setToolOrderForm({ ...toolOrderForm, brand: e.target.value })} /></Field>
            <Field label="Qty"><input type="number" value={toolOrderForm.qty} onChange={(e) => setToolOrderForm({ ...toolOrderForm, qty: e.target.value })} /></Field>
            <Field label="Requested by"><input value={toolOrderForm.requested_by} onChange={(e) => setToolOrderForm({ ...toolOrderForm, requested_by: e.target.value })} /></Field>
            <Field label="Status"><select value={toolOrderForm.status} onChange={(e) => setToolOrderForm({ ...toolOrderForm, status: e.target.value })}><option>Requested</option><option>Approved</option><option>Ordered</option><option>Delivered</option><option>Cancelled</option></select></Field>
            <Field label="Reason"><textarea rows={2} value={toolOrderForm.reason} onChange={(e) => setToolOrderForm({ ...toolOrderForm, reason: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('tool_orders', toolOrderForm)}>Save tool order</button>
          </div>

          <div className="card form-card">
            <h3>Employee issued tools</h3>
            <Field label="Employee"><input value={employeeToolForm.employee_name} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, employee_name: e.target.value })} /></Field>
            <Field label="Tool"><input value={employeeToolForm.tool_name} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, tool_name: e.target.value })} /></Field>
            <Field label="Brand"><input value={employeeToolForm.brand} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, brand: e.target.value })} /></Field>
            <Field label="Serial no"><input value={employeeToolForm.serial_no} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, serial_no: e.target.value })} /></Field>
            <Field label="Issue date"><input type="date" value={employeeToolForm.issue_date} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, issue_date: e.target.value })} /></Field>
            <Field label="Issued by"><input value={employeeToolForm.issued_by} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, issued_by: e.target.value })} /></Field>
            <Field label="Status"><select value={employeeToolForm.status} onChange={(e) => setEmployeeToolForm({ ...employeeToolForm, status: e.target.value })}><option>Issued</option><option>Returned</option><option>Lost</option><option>Damaged</option></select></Field>
            <button className="primary" onClick={() => saveRow('employee_tools', employeeToolForm)}>Save issued tool</button>
          </div>
        </div>

        <div className="grid-2">
          <div className="card"><h3>Tool inventory</h3><MiniTable rows={data.tools} columns={[
            { key: 'tool_name', label: 'Tool' },
            { key: 'category', label: 'Category' },
            { key: 'brand', label: 'Brand' },
            { key: 'stock_qty', label: 'Qty' },
            { key: 'workshop_section', label: 'Section' },
            { key: 'condition', label: 'Condition', render: (r) => <Badge value={r.condition} /> }
          ]} /></div>
          <div className="card"><h3>Employee issued tools</h3><MiniTable rows={data.employee_tools} columns={[
            { key: 'employee_name', label: 'Employee' },
            { key: 'tool_name', label: 'Tool' },
            { key: 'serial_no', label: 'Serial' },
            { key: 'issue_date', label: 'Issued' },
            { key: 'issued_by', label: 'By' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
        </div>
      </div>
    )
  }

  function renderFabrication() {
    return (
      <div className="page-grid fabrication-page">
        <SectionTitle title="Boiler Shop & Panel Beaters" subtitle="Material stock, machine-booked material requests and major ongoing department projects." />

        <div className="stats-grid">
          <StatCard title="Material stock" value={data.fabrication_stock.length} tone="blue" />
          <StatCard title="Material requests" value={data.fabrication_requests.length} tone="orange" />
          <StatCard title="Ongoing projects" value={metrics.fabProjects.length} tone="purple" />
          <StatCard title="Low material" value={data.fabrication_stock.filter((x) => n(x.stock_qty) <= n(x.min_qty)).length} tone="red" />
        </div>

        <div className="grid-3">
          <div className="card form-card accent-card">
            <h3>Material stock</h3>
            <Field label="Department"><select value={fabStockForm.department} onChange={(e) => setFabStockForm({ ...fabStockForm, department: e.target.value })}>{fabricationDepartments.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Material type"><select value={fabStockForm.material_type} onChange={(e) => setFabStockForm({ ...fabStockForm, material_type: e.target.value })}>{materialTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Description"><input value={fabStockForm.description} onChange={(e) => setFabStockForm({ ...fabStockForm, description: e.target.value })} /></Field>
            <Field label="Size/spec"><input value={fabStockForm.size_spec} onChange={(e) => setFabStockForm({ ...fabStockForm, size_spec: e.target.value })} /></Field>
            <Field label="Qty"><input type="number" value={fabStockForm.stock_qty} onChange={(e) => setFabStockForm({ ...fabStockForm, stock_qty: e.target.value })} /></Field>
            <Field label="Min qty"><input type="number" value={fabStockForm.min_qty} onChange={(e) => setFabStockForm({ ...fabStockForm, min_qty: e.target.value })} /></Field>
            <Field label="Unit"><input value={fabStockForm.unit} onChange={(e) => setFabStockForm({ ...fabStockForm, unit: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('fabrication_stock', fabStockForm)}>Save material</button>
          </div>

          <div className="card form-card">
            <h3>Material request</h3>
            <Field label="Department"><select value={fabRequestForm.department} onChange={(e) => setFabRequestForm({ ...fabRequestForm, department: e.target.value })}>{fabricationDepartments.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Fleet number"><input value={fabRequestForm.fleet_no} onChange={(e) => setFabRequestForm({ ...fabRequestForm, fleet_no: e.target.value })} /></Field>
            <Field label="Requested by"><input value={fabRequestForm.requested_by} onChange={(e) => setFabRequestForm({ ...fabRequestForm, requested_by: e.target.value })} /></Field>
            <Field label="Material type"><select value={fabRequestForm.material_type} onChange={(e) => setFabRequestForm({ ...fabRequestForm, material_type: e.target.value })}>{materialTypes.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Description"><input value={fabRequestForm.description} onChange={(e) => setFabRequestForm({ ...fabRequestForm, description: e.target.value })} /></Field>
            <Field label="Qty"><input type="number" value={fabRequestForm.qty} onChange={(e) => setFabRequestForm({ ...fabRequestForm, qty: e.target.value })} /></Field>
            <Field label="Project"><input value={fabRequestForm.project_name} onChange={(e) => setFabRequestForm({ ...fabRequestForm, project_name: e.target.value })} /></Field>
            <Field label="Status"><select value={fabRequestForm.status} onChange={(e) => setFabRequestForm({ ...fabRequestForm, status: e.target.value })}><option>Requested</option><option>Approved</option><option>Issued</option><option>Ordered</option><option>Completed</option></select></Field>
            <button className="primary" onClick={() => saveRow('fabrication_requests', fabRequestForm)}>Save request</button>
          </div>

          <div className="card form-card">
            <h3>Major project</h3>
            <Field label="Department"><select value={fabProjectForm.department} onChange={(e) => setFabProjectForm({ ...fabProjectForm, department: e.target.value })}>{fabricationDepartments.map((x) => <option key={x}>{x}</option>)}</select></Field>
            <Field label="Project name"><input value={fabProjectForm.project_name} onChange={(e) => setFabProjectForm({ ...fabProjectForm, project_name: e.target.value })} /></Field>
            <Field label="Fleet"><input value={fabProjectForm.fleet_no} onChange={(e) => setFabProjectForm({ ...fabProjectForm, fleet_no: e.target.value })} /></Field>
            <Field label="Supervisor"><input value={fabProjectForm.supervisor} onChange={(e) => setFabProjectForm({ ...fabProjectForm, supervisor: e.target.value })} /></Field>
            <Field label="Start"><input type="date" value={fabProjectForm.start_date} onChange={(e) => setFabProjectForm({ ...fabProjectForm, start_date: e.target.value })} /></Field>
            <Field label="Target"><input type="date" value={fabProjectForm.target_date} onChange={(e) => setFabProjectForm({ ...fabProjectForm, target_date: e.target.value })} /></Field>
            <Field label="Progress %"><input type="number" value={fabProjectForm.progress_percent} onChange={(e) => setFabProjectForm({ ...fabProjectForm, progress_percent: e.target.value })} /></Field>
            <button className="primary" onClick={() => saveRow('fabrication_projects', fabProjectForm)}>Save project</button>
          </div>
        </div>

        <div className="grid-3">
          <div className="card"><h3>Stock</h3><MiniTable rows={data.fabrication_stock} columns={[
            { key: 'department', label: 'Dept' },
            { key: 'material_type', label: 'Type' },
            { key: 'description', label: 'Description' },
            { key: 'stock_qty', label: 'Qty' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
          <div className="card"><h3>Requests</h3><MiniTable rows={data.fabrication_requests} columns={[
            { key: 'department', label: 'Dept' },
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'requested_by', label: 'By' },
            { key: 'description', label: 'Description' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
          <div className="card"><h3>Projects</h3><MiniTable rows={data.fabrication_projects} columns={[
            { key: 'department', label: 'Dept' },
            { key: 'project_name', label: 'Project' },
            { key: 'fleet_no', label: 'Fleet' },
            { key: 'progress_percent', label: 'Progress %' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
          ]} /></div>
        </div>
      </div>
    )
  }

  function renderPhotos() {
    return (
      <div className="page-grid">
        <SectionTitle title="Photos" subtitle="Upload machine, tyre, battery, breakdown, abuse and project photos." />
        <div className="grid-2">
          <div className="card form-card accent-card">
            <h3>Add photo</h3>
            <Field label="Fleet number"><input value={photoForm.fleet_no} onChange={(e) => setPhotoForm({ ...photoForm, fleet_no: e.target.value })} /></Field>
            <Field label="Linked type"><select value={photoForm.linked_type} onChange={(e) => setPhotoForm({ ...photoForm, linked_type: e.target.value })}><option>Breakdown</option><option>Repair</option><option>Tyre</option><option>Battery</option><option>Operations</option><option>Boiler/Panel</option></select></Field>
            <Field label="Caption"><input value={photoForm.caption} onChange={(e) => setPhotoForm({ ...photoForm, caption: e.target.value })} /></Field>
            <Field label="Photo"><input type="file" accept="image/*" onChange={(e) => attachPhoto(e.target.files?.[0], (v) => setPhotoForm({ ...photoForm, image_data: v }))} /></Field>
            {photoForm.image_data && <img className="thumb" src={photoForm.image_data} alt="Preview" />}
            <button className="primary" onClick={() => saveRow('photo_logs', photoForm)}>Save photo</button>
          </div>

          <div className="card">
            <h3>Photo register</h3>
            <div className="photo-grid">
              {data.photo_logs.map((p) => (
                <div className="photo-card" key={p.id}>
                  {p.image_data && <img src={p.image_data} alt={p.caption || 'Photo'} />}
                  <b>{p.fleet_no || 'No fleet'}</b>
                  <span>{p.linked_type} — {p.caption}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    )
  }

  function renderReports() {
    const deptCounts = Object.entries(data.fleet_machines.reduce((acc: Row, m) => {
      const key = m.department || 'Unknown'
      acc[key] = (acc[key] || 0) + 1
      return acc
    }, {})).map(([department, count]) => ({ id: department, department, count }))

    const typeCounts = Object.entries(data.fleet_machines.reduce((acc: Row, m) => {
      const key = m.machine_type || 'Unknown'
      acc[key] = (acc[key] || 0) + 1
      return acc
    }, {})).map(([machine_type, count]) => ({ id: machine_type, machine_type, count }))

    const fleetHistory = reportFleet
      ? [...data.breakdowns, ...data.repairs, ...data.services, ...data.spares_orders, ...data.tyres, ...data.batteries, ...data.hose_requests].filter((r) =>
          String(r.fleet_no || r.machine_fleet_no || '').toLowerCase() === reportFleet.toLowerCase()
        )
      : []

    return (
      <div className="page-grid reports-page">
        <SectionTitle title="Reports" subtitle="Machine availability, breakdowns per department, fleet type, service history, orders and fleet-specific history." right={<button onClick={() => window.print()}>Print / Save PDF</button>} />

        <div className="stats-grid">
          <StatCard title="Fleet" value={data.fleet_machines.length} tone="blue" />
          <StatCard title="Breakdowns" value={data.breakdowns.length} tone="red" />
          <StatCard title="Services" value={data.services.length} tone="purple" />
          <StatCard title="Orders" value={data.spares_orders.length + data.tyre_orders.length + data.battery_orders.length + data.hose_requests.length} tone="orange" />
        </div>

        <div className="grid-2">
          <div className="card"><h3>Fleet per department</h3><MiniTable rows={deptCounts} columns={[
            { key: 'department', label: 'Department' },
            { key: 'count', label: 'Machines' }
          ]} /></div>
          <div className="card"><h3>Fleet type breakdown</h3><MiniTable rows={typeCounts} columns={[
            { key: 'machine_type', label: 'Fleet type' },
            { key: 'count', label: 'Machines' }
          ]} /></div>
        </div>

        <div className="card">
          <h3>Fleet history lookup</h3>
          <div className="filters">
            <input placeholder="Enter fleet number..." value={reportFleet} onChange={(e) => setReportFleet(e.target.value)} />
          </div>
          <MiniTable rows={fleetHistory} empty="Enter a fleet number to see history." columns={[
            { key: 'fleet_no', label: 'Fleet', render: (r) => r.fleet_no || r.machine_fleet_no },
            { key: 'fault', label: 'Fault/service/order', render: (r) => r.fault || r.service_type || r.description || r.part_no || r.serial_no || r.hose_type || '-' },
            { key: 'status', label: 'Status', render: (r) => <Badge value={r.status || r.workflow_stage || r.stock_status} /> },
            { key: 'created_at', label: 'Date', render: (r) => dateOnly(r.created_at || r.request_date || r.due_date) }
          ]} />
        </div>
      </div>
    )
  }

  function renderAdmin() {
    return (
      <div className="page-grid">
        <SectionTitle title="Admin" subtitle="Sync, offline queue, data controls and quick checks." />
        <div className="stats-grid">
          <StatCard title="Online" value={online ? 'Yes' : 'No'} tone={online ? 'green' : 'red'} />
          <StatCard title="Queue" value={queueCount} tone="orange" />
          <StatCard title="Supabase" value={isSupabaseConfigured ? 'Set' : 'Missing'} tone={isSupabaseConfigured ? 'green' : 'red'} />
          <StatCard title="Tables" value={TABLES.length} tone="blue" />
        </div>
        <div className="card">
          <h3>Actions</h3>
          <div className="quick-actions">
            <button onClick={loadAll}>Reload from Supabase</button>
            <button onClick={syncOfflineQueue}>Sync offline queue</button>
            <button onClick={() => saveSnapshot(stripHeavyFileData(data))}>Save local snapshot</button>
            <button className="danger-btn" onClick={() => { localStorage.removeItem('turbo-workshop-snapshot'); setMessage('Local snapshot cleared.') }}>Clear local snapshot</button>
          </div>
        </div>
      </div>
    )
  }

  function renderTab() {
    if (tab === 'dashboard') return renderDashboard()
    if (tab === 'fleet') return renderFleet()
    if (tab === 'breakdowns') return renderBreakdowns()
    if (tab === 'repairs') return renderRepairs()
    if (tab === 'services') return renderServices()
    if (tab === 'spares') return renderSpares()
    if (tab === 'personnel') return renderPersonnel()
    if (tab === 'tyres') return renderTyres()
    if (tab === 'batteries') return renderBatteries()
    if (tab === 'operations') return renderOperations()
    if (tab === 'hoses') return renderHoses()
    if (tab === 'tools') return renderTools()
    if (tab === 'fabrication') return renderFabrication()
    if (tab === 'photos') return renderPhotos()
    if (tab === 'reports') return renderReports()
    if (tab === 'admin') return renderAdmin()
    return renderDashboard()
  }

  if (!session) {
    return (
      <main className="login-shell">
        <div className="login-card glass">
          <div className="brand-mark">TE</div>
          <h1>Turbo Energy Workshop Control</h1>
          <p className="muted">Fleet repairs, services, spares, personnel, tyres, batteries, hoses, tools and workshop projects.</p>
          {message && <div className="message">{message}</div>}
          <div className="form-card">
            <Field label="Username"><input value={loginName} onChange={(e) => setLoginName(e.target.value)} /></Field>
            <Field label="Password"><input type="password" value={loginPass} onChange={(e) => setLoginPass(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') handleLogin() }} /></Field>
            <button className="primary wide" onClick={handleLogin}>Login</button>
          </div>
          <div className="login-help">
            <b>Default logins</b><br />
            admin / admin123<br />
            admin workshop / workshop123<br />
            foreman / foreman123<br />
            user / user123
          </div>
        </div>
      </main>
    )
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="logo-block">
          <div className="brand-mark small">TE</div>
          <div>
            <b>Turbo Energy</b>
            <span>Workshop Control</span>
          </div>
        </div>

        <nav>
          {nav.map(([key, label]) => (
            <button key={key} className={tab === key ? 'active' : ''} onClick={() => setTab(key)}>
              {label}
            </button>
          ))}
        </nav>

        <div className="sync-box">
          <Badge value={online ? 'Online' : 'Offline'} />
          <span>Queue: {queueCount}</span>
          <span>Signed in: {session.username} ({session.role})</span>
          <button onClick={logout}>Logout</button>
        </div>
      </aside>

      <section className="workspace">
        <div className="topbar">
          <div>
            <h1>{nav.find(([key]) => key === tab)?.[1] || 'Dashboard'}</h1>
            <p>Navy/orange workshop system • laptop and phone friendly • online/offline records</p>
          </div>
          <div className="top-actions">
            <button onClick={() => setTab('dashboard')}>Dashboard</button>
            <button onClick={() => window.print()}>Print</button>
            <button onClick={syncOfflineQueue}>Sync</button>
          </div>
        </div>

        {message && <div className="message floating">{message}</div>}

        <datalist id="fleet-list">
          {data.fleet_machines.map((m) => <option key={m.id || m.fleet_no} value={m.fleet_no} />)}
        </datalist>

        {renderTab()}
      </section>

      {modal && <Modal title={modal.title} onClose={() => setModal(null)}>{modal.body}</Modal>}
    </main>
  )
}
