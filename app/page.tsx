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

function stripHeavyFileData(input: AppData): AppData {
  const next = { ...input } as AppData
  next.parts_books = (next.parts_books || []).map(({ file_data, object_url, ...book }) => book)
  return next
}

function safeFileName(name: string) {
  return String(name || 'parts-book.pdf').replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 120)
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
  const [partsBookForm, setPartsBookForm] = useState<Row>({ book_title: '', machine_group: 'Yellow Machine', machine_type: '', uploaded_by: '', file_name: '', mime_type: '', file_size: 0, file_data: '', file_url: '', storage_path: '', extracted_part_numbers: '', extracted_text: '', notes: '' })
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

  useEffect(() => { saveSnapshot(stripHeavyFileData(data)) }, [data])

  async function loadAll() {
    if (!supabase || !isSupabaseConfigured || (typeof navigator !== 'undefined' && !navigator.onLine)) return
    const next: AppData = { ...emptyData }
    for (const table of TABLES) {
      const orderCol = table === 'fleet_machines' ? 'fleet_no' : 'created_at'
      const { data: rows } = await supabase.from(table).select('*').order(orderCol, { ascending: false })
      if (rows) next[table] = rows as Row[]
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

  async function uploadGeneric(file: File | undefined, table: TableName, mapper: (row: Row) => Row, label: string) {
    if (!file) return
    const rows = await readExcelRows(file)
    const mapped = rows.map(mapper).filter((row) => Object.keys(row).length > 1)
    for (const row of mapped) await saveRow(table, row, 'insert')
    setMessage(`${mapped.length} ${label} records uploaded.`)
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
    let tempUrl = ''

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
        tempUrl = URL.createObjectURL(file)
      }
    } catch (err: any) {
      tempUrl = URL.createObjectURL(file)
      setMessage(`PDF loaded for this session only. Supabase Storage bucket may still need manual setup: ${err?.message || 'storage error'}`)
    }

    setPartsBookForm((prev) => ({
      ...prev,
      file_name: file.name,
      mime_type: file.type || 'application/pdf',
      file_size: file.size,
      file_url: fileUrl,
      storage_path: storagePath,
      object_url: tempUrl,
      extracted_part_numbers: detected.join('\n'),
      extracted_count: detected.length,
      extracted_text: rawText.slice(0, 7000)
    }))

    if (fileUrl) setMessage(`${detected.length} possible part numbers found. PDF uploaded to Supabase Storage.`)
  }

  function openPartsBook(book: Row) {
    const url = book.file_url || book.object_url
    const parts = String(book.extracted_part_numbers || '').split('\n').map((x) => x.trim()).filter(Boolean)

    setModal({
      title: book.book_title || book.file_name || 'Parts book',
      body: (
        <div className="pdf-modal-grid">
          <div className="pdf-frame-wrap">
            {url ? <iframe src={url} title="Parts book PDF" /> : <div className="empty">No PDF file URL saved for this book yet.</div>}
          </div>
          <div className="pdf-side">
            <h3>Detected part numbers</h3>
            <p className="muted">Click any part number to put it into the spares order sheet.</p>
            <div className="part-chip-box tall">
              {parts.length ? parts.map((part) => (
                <button key={part} className="part-chip" onClick={() => addPartToOrder(part)}>{part}</button>
              )) : <span className="muted">No part numbers detected. Type them manually in the order sheet.</span>}
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

  const metrics = useMemo(() => {
    const openBreakdowns = data.breakdowns.filter((r) => !['Closed', 'Completed'].includes(String(r.status || '')))
    const openRepairs = data.repairs.filter((r) => !['Closed', 'Completed'].includes(String(r.status || '')))
    const servicesDue = data.services.filter((r) => !['Completed', 'Closed'].includes(String(r.status || '')))
    const lowSpares = data.spares.filter((r) => n(r.stock_qty) <= n(r.min_qty))
    const peopleOff = data.personnel.filter((p) => String(p.employment_status || '').toLowerCase().includes('leave') || String(p.employment_status || '').toLowerCase().includes('off'))
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

  function listOptions(rows: Row[], key: string) {
    return Array.from(new Set(rows.map((r) => String(r[key] || '').trim()).filter(Boolean))).sort()
  }

  function renderDashboard() {
    return (
      <div className="page-grid">
        <SectionTitle
          title="Workshop Control Dashboard"
          subtitle="Live workshop overview for fleet, personnel, services, breakdowns, spares, hoses, tools and department projects."
          right={<div className="page-pill-row"><span className="page-pill blue">Current shift working</span><span className="page-pill orange">{today()}</span></div>}
        />

        <div className="stats-grid">
          <StatCard title="Total fleet" value={metrics.totalFleet} detail="Registered machines" tone="blue" onClick={() => setTab('fleet')} />
          <StatCard title="Breakdowns" value={metrics.openBreakdowns.length} detail="Open machine breakdowns" tone="red" onClick={() => setTab('breakdowns')} />
          <StatCard title="Being worked on" value={metrics.openRepairs.length} detail="Active repair jobs" tone="orange" onClick={() => setTab('repairs')} />
          <StatCard title="Service today / due" value={metrics.servicesDue.length} detail="Service schedule" tone="purple" onClick={() => setTab('services')} />
          <StatCard title="People off" value={metrics.peopleOff.length} detail="Leave/off schedule" tone="yellow" onClick={() => setTab('personnel')} />
          <StatCard title="Foremen on duty" value={metrics.foremen.length} detail="Current workshop supervision" tone="green" onClick={() => setTab('personnel')} />
          <StatCard title="Hose requests" value={metrics.hosePending.length} detail="Pending hose/fitting work" tone="blue" onClick={() => setTab('hoses')} />
          <StatCard title="Tool orders" value={metrics.toolOrders.length} detail="Tools requested" tone="orange" onClick={() => setTab('tools')} />
        </div>

        <div className="grid-2">
          <div className="card">
            <div className="card-head">
              <h3>Machines on breakdown</h3>
              <button onClick={() => setTab('breakdowns')}>Open list</button>
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
              <h3>Services due / today</h3>
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
              <h3>Alerts and requests</h3>
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
        <SectionTitle title="Fleet Register" subtitle="Upload or add machines. This becomes the master list used by services, repairs, tyres, batteries and orders." />
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
            <p className="small-note">Suggested columns: fleet_no, machine_type, make_model, department, location, hours, mileage, status.</p>
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
            <h3>Repeated breakdown hints</h3>
            <p className="muted">Use same fault/cause wording to build a useful history per machine. The reports page will group repeated faults by fleet number and fault type.</p>
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
            <p className="small-note">Suggested columns: fleet_no, service_type, due_date, scheduled_hours, job_card_no, supervisor, technician, status.</p>
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
        <SectionTitle title="Spares & Orders" subtitle="Dark user-friendly order board with requested, awaiting funding, funded, local/international orders, delivery ETA and PDF parts book viewer." />

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
            <p className="small-note">Saved books open below. If Supabase Storage bucket is missing, the PDF will open only during this session until we enable the bucket manually.</p>
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
            <p className="small-note">Suggested columns: employee_no, name, shift, section, role, employment_status, leave_balance_days, leave_start, leave_end, leave_type.</p>
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
            <h3>Repeated breakdown hints</h3>
            <p className="muted">Use same fault/cause wording to build a useful history per machine. The reports page will group repeated faults by fleet number and fault type.</p>
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
            <p className="small-note">Suggested columns: fleet_no, service_type, due_date, scheduled_hours, job_card_no, supervisor, technician, status.</p>
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
        <SectionTitle title="Spares & Orders" subtitle="Dark user-friendly order board with requested, awaiting funding, funded, local/international orders, delivery ETA and PDF parts book viewer." />

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
            <p className="small-note">Saved books open below. If Supabase Storage bucket is missing, the PDF will open only during this session until we enable the bucket manually.</p>
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
            <p className="small-note">Suggested columns: employee_no, name, shift, section, role, employment_status, leave_balance_days, leave_start, leave_end, leave_type.</p>
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
