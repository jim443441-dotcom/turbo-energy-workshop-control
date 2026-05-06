'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { isSupabaseConfigured, supabase } from '@/lib/supabase'

type Row = Record<string, any>

type TableName =
  | 'fleet_machines'
  | 'personnel'
  | 'leave_records'
  | 'services'
  | 'breakdowns'
  | 'repairs'
  | 'spares'
  | 'spares_orders'
  | 'tyres'
  | 'tyre_measurements'
  | 'tyre_orders'
  | 'batteries'
  | 'battery_orders'
  | 'operations'
  | 'hoses'
  | 'hose_requests'
  | 'tools'
  | 'tool_orders'
  | 'employee_tools'
  | 'fabrication_stock'
  | 'fabrication_requests'
  | 'fabrication_projects'
  | 'photo_logs'

type Register = {
  table: TableName
  label: string
  description: string
}

const REGISTERS: Register[] = [
  { table: 'fleet_machines', label: 'Fleet Register', description: 'Machines, fleet numbers, hours, mileage and status.' },
  { table: 'personnel', label: 'Personnel Register', description: 'Employees, shifts, trades, sections and leave balances.' },
  { table: 'leave_records', label: 'Leave / Off Schedule', description: 'Current, upcoming and past leave records.' },
  { table: 'services', label: 'Service Schedule', description: 'Services, job cards and supervisors.' },
  { table: 'breakdowns', label: 'Breakdowns', description: 'Breakdown records, faults, ETAs and fitters.' },
  { table: 'repairs', label: 'Repairs / Job Cards', description: 'Repairs, job cards and parts used.' },
  { table: 'spares', label: 'Spares Stock', description: 'Part numbers, stock, suppliers and shelves.' },
  { table: 'spares_orders', label: 'Spares Orders', description: 'Requests, funding, delivery and ETA.' },
  { table: 'tyres', label: 'Tyres Register', description: 'Tyre serials, fitment and reminders.' },
  { table: 'tyre_measurements', label: 'Tyre Measurements', description: 'Tread depth, pressure and checks.' },
  { table: 'tyre_orders', label: 'Tyre Orders', description: 'Tyres to quote or order.' },
  { table: 'batteries', label: 'Battery Register', description: 'Battery serials, fitment and maintenance.' },
  { table: 'battery_orders', label: 'Battery Orders', description: 'Battery requests against machines.' },
  { table: 'operations', label: 'Operations / Operators', description: 'Operator history, score and machine abuse records.' },
  { table: 'hoses', label: 'Hose Stock', description: 'Hose and fitting stock.' },
  { table: 'hose_requests', label: 'Hose Requests', description: 'Bulk hose requests and hose repairs.' },
  { table: 'tools', label: 'Tools Inventory', description: 'Workshop tools, quantities, brands and condition.' },
  { table: 'tool_orders', label: 'Tool Orders', description: 'Tool requests and ordering.' },
  { table: 'employee_tools', label: 'Employee Issued Tools', description: 'Tools issued to employees.' },
  { table: 'fabrication_stock', label: 'Boiler/Panel Stock', description: 'Boiler shop and panel beating stock.' },
  { table: 'fabrication_requests', label: 'Boiler/Panel Requests', description: 'Material requests booked to machines.' },
  { table: 'fabrication_projects', label: 'Boiler/Panel Projects', description: 'Major ongoing projects.' },
  { table: 'photo_logs', label: 'Photo Captions', description: 'Photo captions and fleet references.' }
]

const today = () => new Date().toISOString().slice(0, 10)
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`

function normal(value: any) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
}

function cleanRow(row: Row) {
  const out: Row = {}

  Object.entries(row).forEach(([key, value]) => {
    if (value === undefined || value === null) return
    if (typeof value === 'string' && value.trim() === '') return
    out[key] = value
  })

  return out
}

function makeSlug(value: any) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
}

function excelDateNumber(value: number) {
  const parsed = XLSX.SSF.parse_date_code(value)
  if (!parsed) return ''
  return `${String(parsed.y).padStart(4, '0')}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`
}

function dateOnly(value: any) {
  if (!value) return ''

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10)
  }

  if (typeof value === 'number' && value > 20000) {
    return excelDateNumber(value)
  }

  const text = String(value).trim()
  if (!text) return ''

  let m = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/)
  if (m) {
    const dd = m[1].padStart(2, '0')
    const mm = m[2].padStart(2, '0')
    const yyyy = m[3].length === 2 ? `20${m[3]}` : m[3]
    return `${yyyy}-${mm}-${dd}`
  }

  m = text.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/)
  if (m) {
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`
  }

  const d = new Date(text)
  if (!Number.isNaN(d.getTime())) {
    return d.toISOString().slice(0, 10)
  }

  return ''
}

function dateRange(value: any) {
  const text = String(value || '').trim()
  const matches = text.match(/\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/g) || []

  if (matches.length >= 2) {
    return [dateOnly(matches[0]), dateOnly(matches[1])]
  }

  if (matches.length === 1) {
    return [dateOnly(matches[0]), dateOnly(matches[0])]
  }

  return ['', '']
}

function daysInclusive(start: any, end: any) {
  const s = new Date(start)
  const e = new Date(end)

  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return 0

  return Math.max(0, Math.floor((e.getTime() - s.getTime()) / 86400000) + 1)
}

function getCell(row: Row, aliases: string[]) {
  const keys = Object.keys(row)

  for (const alias of aliases) {
    const found = keys.find((key) => normal(key) === normal(alias))
    if (found && String(row[found] ?? '').trim() !== '') return row[found]
  }

  for (const alias of aliases) {
    const wanted = normal(alias)

    const found = keys.find((key) => {
      const actual = normal(key)
      return actual && wanted && (actual.includes(wanted) || wanted.includes(actual))
    })

    if (found && String(row[found] ?? '').trim() !== '') return row[found]
  }

  return ''
}

function numberCell(row: Row, aliases: string[], fallback = 0) {
  const value = getCell(row, aliases)
  const num = Number(String(value || '').replace(/,/g, ''))
  return Number.isFinite(num) ? num : fallback
}

function fullName(row: Row) {
  const first = getCell(row, ['FirstName', 'First Name', 'Forename', 'FIRST NAME'])
  const surname = getCell(row, ['Surname', 'Last Name', 'SURNAME'])
  const normalName = getCell(row, ['Name', 'Employee Name', 'Full Name'])

  const joined = `${String(first || '').trim()} ${String(surname || '').trim()}`.trim()

  return joined || String(normalName || '').trim()
}

function headerScore(cells: any[]) {
  const words = [
    'fleet',
    'machine',
    'unit',
    'equipment',
    'firstname',
    'first name',
    'surname',
    'employee',
    'code',
    'occupation',
    'current job title',
    'department',
    'shift',
    'dates',
    'total',
    'leave type',
    'part no',
    'description',
    'qty',
    'serial',
    'status',
    'requested by',
    'supervisor',
    'foreman'
  ]

  let score = 0

  cells.forEach((cell) => {
    const text = normal(cell)
    words.forEach((word) => {
      if (text.includes(normal(word))) score += 1
    })
  })

  return score
}

async function readWorkbook(file: File) {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })

  const rows: Row[] = []
  const allHeadings = new Set<string>()
  const allCells: any[][] = []

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      raw: true
    }) as any[][]

    matrix.forEach((line) => allCells.push(line))

    if (!matrix.length) return

    let headerIndex = 0
    let best = -1

    matrix.forEach((line, index) => {
      const score = headerScore(line)
      if (score > best) {
        best = score
        headerIndex = index
      }
    })

    const headings = (matrix[headerIndex] || []).map((h, i) => {
      const clean = String(h || '').trim()
      return clean || `Column ${i + 1}`
    })

    headings.forEach((h) => allHeadings.add(h))

    matrix.slice(headerIndex + 1).forEach((line, offset) => {
      const obj: Row = {
        __sheet: sheetName,
        __row: headerIndex + offset + 2,
        __cells: line
      }

      let filled = 0

      headings.forEach((heading, index) => {
        const value = line[index]
        if (value !== undefined && value !== null && String(value).trim() !== '') {
          obj[heading] = value
          filled += 1
        }
      })

      if (filled > 0) rows.push(obj)
    })
  })

  return {
    rows,
    headings: Array.from(allHeadings),
    allCells
  }
}

function fleetFromCells(cells: any[]) {
  const text = cells.map((c) => String(c || '')).join(' ')
  const parts = text.split(/[\s,;:/|]+/).map((x) => x.trim()).filter(Boolean)

  const banned = new Set([
    'FLEET',
    'MACHINE',
    'EQUIPMENT',
    'TYPE',
    'DESCRIPTION',
    'DATE',
    'TOTAL',
    'STATUS',
    'DEPARTMENT',
    'WORKSHOP'
  ])

  for (const part of parts) {
    const clean = part.toUpperCase().replace(/[^A-Z0-9-]/g, '')
    if (!clean) continue
    if (banned.has(clean)) continue
    if (/^\d+$/.test(clean)) continue
    if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(clean)) continue

    if (/^[A-Z]{1,6}-?\d{1,6}[A-Z]?$/.test(clean)) {
      return clean.replace(/-/g, '')
    }
  }

  return ''
}

function inferMachineType(value: any, textValue = '') {
  const text = `${value} ${textValue}`.toUpperCase()

  if (text.includes('FEL') || text.includes('LOADER')) return 'FEL'
  if (text.includes('EXC') || text.includes('TEX') || text.includes('EXCAVATOR')) return 'Excavator'
  if (text.includes('DOZER') || text.includes('DZ')) return 'Dozer'
  if (text.includes('GRADER')) return 'Grader'
  if (text.includes('ROLLER')) return 'Roller'
  if (text.includes('ADT') || text.includes('HOWO') || text.includes('TRUCK') || text.includes('HT')) return 'Truck'
  if (text.includes('LDV') || text.includes('AFX') || text.includes('AG')) return 'LDV'

  return 'Machine'
}

function parsePersonnel(row: Row) {
  const name = fullName(row)
  const employeeNo = String(getCell(row, ['Code', 'Employee Code', 'Employee No', 'Clock No']) || '').trim()
  const role = String(getCell(row, ['Occupation', 'Current Job Title', 'Job Title', 'Trade', 'Role']) || 'Workshop Employee').trim()
  const section = String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim()

  if (!name) return null
  if (name.toLowerCase().includes('workshop department')) return null

  return cleanRow({
    id: employeeNo || makeSlug(name),
    employee_no: employeeNo,
    name,
    shift: String(getCell(row, ['Shift']) || 'Shift 1').trim(),
    section,
    role,
    staff_group: 'Workshop Staff',
    phone: String(getCell(row, ['Phone', 'Cell', 'Mobile']) || '').trim(),
    employment_status: String(getCell(row, ['Status', 'Employment Status']) || 'Active').trim(),
    leave_balance_days: numberCell(row, ['Leave Balance', 'Leave Days'], 30),
    leave_taken_days: numberCell(row, ['Leave Taken', 'Days Taken'], 0),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  })
}

function parseLeave(row: Row) {
  const name = fullName(row)
  const employeeNo = String(getCell(row, ['Code', 'Employee Code', 'Employee No', 'Clock No']) || '').trim()
  const role = String(getCell(row, ['Occupation', 'Current Job Title', 'Job Title', 'Trade', 'Role']) || '').trim()
  const section = String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim()

  const datesText = getCell(row, ['DATES', 'Dates', 'Date Range', 'Leave Dates'])
  const [rangeStart, rangeEnd] = dateRange(datesText)

  const start = dateOnly(getCell(row, ['Start Date', 'Leave Start', 'From'])) || rangeStart
  const end = dateOnly(getCell(row, ['End Date', 'Leave End', 'To'])) || rangeEnd

  if (!name && !employeeNo) return null
  if (!start && !end) return null

  const leaveType = String(getCell(row, ['LEAVE TYPE', 'Leave Type', 'Type']) || 'Annual Leave').trim()
  const days = numberCell(row, ['TOTAL', 'Total', 'Days', 'Leave Days'], daysInclusive(start, end))
  const id = `${employeeNo || makeSlug(name)}-${start}-${end}-${makeSlug(leaveType)}`

  return cleanRow({
    id,
    employee_no: employeeNo,
    person_name: name,
    shift: String(getCell(row, ['Shift']) || 'Shift 1').trim(),
    section,
    role,
    leave_type: leaveType,
    start_date: start,
    end_date: end,
    days,
    reason: String(getCell(row, ['Reason']) || '').trim(),
    status: String(getCell(row, ['Status']) || 'Approved').trim()
  })
}

function parseFleet(row: Row) {
  const cells = Array.isArray(row.__cells) ? row.__cells : Object.values(row)
  const cellText = cells.map((c) => String(c || '')).join(' ')

  const fleetNo =
    String(getCell(row, ['Fleet', 'Fleet No', 'Fleet Number', 'Machine', 'Machine No', 'Unit', 'Equipment', 'Plant No']) || '').trim() ||
    fleetFromCells(cells)

  if (!fleetNo) return null

  return cleanRow({
    id: fleetNo,
    fleet_no: fleetNo,
    machine_type:
      String(getCell(row, ['Machine Type', 'Equipment Type', 'Type']) || '').trim() ||
      inferMachineType(fleetNo, cellText),
    make_model: String(getCell(row, ['Make Model', 'Make', 'Model', 'Description']) || '').trim(),
    reg_no: String(getCell(row, ['Reg', 'Registration', 'Reg No']) || '').trim(),
    department: String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim(),
    location: String(getCell(row, ['Location', 'Site']) || '').trim(),
    hours: numberCell(row, ['Hours', 'Hrs', 'HM', 'Hour Meter'], 0),
    mileage: numberCell(row, ['Mileage', 'KM', 'Odometer'], 0),
    status: String(getCell(row, ['Status', 'State']) || 'Available').trim(),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  })
}

function parseGeneric(table: TableName, row: Row) {
  const fleetNo = String(getCell(row, ['Fleet', 'Fleet No', 'Machine', 'Unit']) || '').trim()
  const partNo = String(getCell(row, ['Part No', 'Part Number', 'Stock Code', 'Item No']) || '').trim()
  const description = String(getCell(row, ['Description', 'Item', 'Name']) || '').trim()
  const serial = String(getCell(row, ['Serial', 'Serial No', 'Serial Number']) || '').trim()
  const requestedBy = String(getCell(row, ['Requested By', 'Requester', 'Foreman']) || '').trim()

  const base: Row = {
    id: uid(),
    fleet_no: fleetNo,
    machine_fleet_no: fleetNo,
    part_no: partNo,
    description,
    serial_no: serial,
    requested_by: requestedBy,
    qty: numberCell(row, ['Qty', 'Quantity'], 1),
    stock_qty: numberCell(row, ['Stock Qty', 'Stock', 'Balance', 'Qty'], 0),
    status: String(getCell(row, ['Status', 'Stage']) || 'Imported').trim(),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  }

  if (table === 'spares') {
    return cleanRow({
      id: partNo || uid(),
      part_no: partNo,
      description,
      section: String(getCell(row, ['Section', 'Department']) || 'Stores').trim(),
      machine_type: String(getCell(row, ['Machine Type', 'Type']) || '').trim(),
      stock_qty: numberCell(row, ['Stock Qty', 'Stock', 'Balance', 'Qty'], 0),
      min_qty: numberCell(row, ['Min Qty', 'Minimum', 'Reorder Level'], 1),
      supplier: String(getCell(row, ['Supplier', 'Vendor']) || '').trim(),
      order_status: String(getCell(row, ['Status', 'Order Status']) || 'In stock').trim(),
      notes: base.notes
    })
  }

  if (table === 'spares_orders') {
    return cleanRow({
      id: uid(),
      request_date: dateOnly(getCell(row, ['Request Date', 'Date'])) || today(),
      machine_fleet_no: fleetNo,
      requested_by: requestedBy,
      part_no: partNo,
      description,
      qty: numberCell(row, ['Qty', 'Quantity'], 1),
      priority: String(getCell(row, ['Priority']) || 'Normal').trim(),
      workflow_stage: String(getCell(row, ['Stage', 'Workflow Stage', 'Status']) || 'Requested').trim(),
      status: String(getCell(row, ['Status', 'Stage']) || 'Requested').trim(),
      eta: dateOnly(getCell(row, ['ETA', 'Expected Date'])),
      notes: base.notes
    })
  }

  if (table === 'services') {
    return cleanRow({
      id: uid(),
      fleet_no: fleetNo,
      service_type: String(getCell(row, ['Service Type', 'Service', 'Type']) || 'Service').trim(),
      due_date: dateOnly(getCell(row, ['Due Date', 'Date'])),
      scheduled_hours: numberCell(row, ['Scheduled Hours', 'Due Hours', 'Hours', 'HM'], 0),
      completed_date: dateOnly(getCell(row, ['Completed Date', 'Done Date'])),
      completed_hours: numberCell(row, ['Completed Hours', 'Done Hours'], 0),
      job_card_no: String(getCell(row, ['Job Card', 'Job Card No']) || '').trim(),
      supervisor: String(getCell(row, ['Supervisor', 'Foreman']) || '').trim(),
      technician: String(getCell(row, ['Technician', 'Fitter', 'Mechanic']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Due').trim(),
      notes: base.notes
    })
  }

  if (table === 'breakdowns' || table === 'repairs') {
    return cleanRow({
      id: uid(),
      fleet_no: fleetNo,
      fault: String(getCell(row, ['Fault', 'Problem', 'Breakdown', 'Defect', 'Description']) || '').trim(),
      category: String(getCell(row, ['Category', 'Section']) || 'Mechanical').trim(),
      assigned_to: String(getCell(row, ['Assigned To', 'Fitter', 'Mechanic', 'Technician']) || '').trim(),
      reported_by: String(getCell(row, ['Reported By', 'Operator']) || '').trim(),
      job_card_no: String(getCell(row, ['Job Card', 'Job Card No']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Open').trim(),
      notes: base.notes
    })
  }

  if (table === 'tyres') {
    return cleanRow({
      id: serial || uid(),
      serial_no: serial,
      fleet_no: fleetNo,
      make: String(getCell(row, ['Make', 'Brand']) || '').trim(),
      company_no: String(getCell(row, ['Company No', 'Company Number']) || '').trim(),
      fitted_date: dateOnly(getCell(row, ['Fitted Date', 'Date Fitted', 'Fitment Date'])),
      fitted_by: String(getCell(row, ['Fitted By']) || '').trim(),
      position: String(getCell(row, ['Position']) || '').trim(),
      size: String(getCell(row, ['Size']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Fitted').trim(),
      notes: base.notes
    })
  }

  if (table === 'batteries') {
    return cleanRow({
      id: serial || uid(),
      serial_no: serial,
      fleet_no: fleetNo,
      make: String(getCell(row, ['Make', 'Brand']) || '').trim(),
      volts: String(getCell(row, ['Volts', 'Voltage']) || '').trim(),
      fitment_date: dateOnly(getCell(row, ['Fitment Date', 'Date Fitted'])),
      fitted_by: String(getCell(row, ['Fitted By']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Fitted').trim(),
      notes: base.notes
    })
  }

  if (table === 'tools') {
    return cleanRow({
      id: serial || description || uid(),
      tool_name: description || String(getCell(row, ['Tool', 'Tool Name']) || '').trim(),
      category: String(getCell(row, ['Category']) || 'Hand Tools').trim(),
      brand: String(getCell(row, ['Brand', 'Make']) || '').trim(),
      serial_no: serial,
      stock_qty: numberCell(row, ['Stock Qty', 'Qty', 'Quantity'], 0),
      condition: String(getCell(row, ['Condition']) || 'Good').trim(),
      status: String(getCell(row, ['Status']) || 'In stock').trim(),
      notes: base.notes
    })
  }

  const hasSomething = fleetNo || partNo || description || serial || requestedBy
  return hasSomething ? cleanRow(base) : null
}

function parseRows(table: TableName, rawRows: Row[], allCells: any[][]) {
  let rows: Row[] = []

  if (table === 'personnel') {
    rows = rawRows.map(parsePersonnel).filter(Boolean) as Row[]
  } else if (table === 'leave_records') {
    rows = rawRows.map(parseLeave).filter(Boolean) as Row[]
  } else if (table === 'fleet_machines') {
    rows = rawRows.map(parseFleet).filter(Boolean) as Row[]

    if (!rows.length) {
      const map = new Map<string, Row>()

      allCells.forEach((line) => {
        const fleetNo = fleetFromCells(line)
        if (!fleetNo) return

        const text = line.map((x) => String(x || '')).join(' ')

        map.set(fleetNo, {
          id: fleetNo,
          fleet_no: fleetNo,
          machine_type: inferMachineType(fleetNo, text),
          department: 'Workshop',
          status: 'Available',
          notes: text.slice(0, 160)
        })
      })

      rows = Array.from(map.values())
    }
  } else {
    rows = rawRows.map((row) => parseGeneric(table, row)).filter(Boolean) as Row[]
  }

  const map = new Map<string, Row>()

  rows.forEach((row) => {
    const key =
      table === 'personnel'
        ? String(row.employee_no || row.name || row.id)
        : table === 'fleet_machines'
          ? String(row.fleet_no || row.id)
          : String(row.id || uid())

    if (!key.trim()) return

    map.set(key, {
      ...row,
      id: key
    })
  })

  return Array.from(map.values())
}

function saveToLocalSnapshot(table: TableName, rows: Row[]) {
  try {
    const raw = localStorage.getItem('turbo-workshop-snapshot')
    const parsed = raw ? JSON.parse(raw) : { data: {} }
    const data = parsed.data || {}

    const existing: Row[] = Array.isArray(data[table]) ? data[table] : []
    const map = new Map<string, Row>()

    existing.forEach((row) => {
      const key = String(row.id || row.fleet_no || row.employee_no || JSON.stringify(row))
      map.set(key, row)
    })

    rows.forEach((row) => {
      const key = String(row.id || row.fleet_no || row.employee_no || uid())
      map.set(key, { ...map.get(key), ...row })
    })

    data[table] = Array.from(map.values())

    localStorage.setItem(
      'turbo-workshop-snapshot',
      JSON.stringify({
        updatedAt: new Date().toISOString(),
        data
      })
    )
  } catch {
    // local backup failure must not stop import
  }
}

function addToOfflineQueue(table: TableName, rows: Row[]) {
  try {
    const raw = localStorage.getItem('turbo-workshop-offline-queue')
    const queue = raw ? JSON.parse(raw) : []

    rows.forEach((row) => {
      queue.push({
        id: uid(),
        table,
        type: 'upsert',
        payload: row,
        createdAt: new Date().toISOString()
      })
    })

    localStorage.setItem('turbo-workshop-offline-queue', JSON.stringify(queue))
  } catch {
    // queue failure must not stop import
  }
}

function Badge({ value }: { value: string }) {
  const text = String(value || '')
  const low = text.toLowerCase()
  const cls = low.includes('fail') || low.includes('error') ? 'danger' : low.includes('uploaded') || low.includes('connected') ? 'good' : 'neutral'
  return <span className={`badge ${cls}`}>{text}</span>
}

export default function ImporterPage() {
  const [selected, setSelected] = useState<TableName>('personnel')
  const [search, setSearch] = useState('')
  const [status, setStatus] = useState('Ready')
  const [busy, setBusy] = useState(false)
  const [fileName, setFileName] = useState('')
  const [headings, setHeadings] = useState<string[]>([])
  const [preview, setPreview] = useState<Row[]>([])

  const selectedRegister = REGISTERS.find((r) => r.table === selected) || REGISTERS[0]

  const filtered = useMemo(() => {
    const q = search.toLowerCase()
    return REGISTERS.filter((r) => `${r.label} ${r.description} ${r.table}`.toLowerCase().includes(q))
  }, [search])

  async function saveRows(table: TableName, rows: Row[]) {
    saveToLocalSnapshot(table, rows)

    if (!supabase || !isSupabaseConfigured) {
      addToOfflineQueue(table, rows)
      setStatus(`Uploaded ${rows.length} records locally. Supabase is not configured.`)
      return
    }

    let uploaded = 0

    for (let i = 0; i < rows.length; i += 50) {
      const chunk = rows.slice(i, i + 50)

      const { error } = await supabase
        .from(table as any)
        .upsert(chunk, { onConflict: 'id' })

      if (error) {
        addToOfflineQueue(table, chunk)
        throw error
      }

      uploaded += chunk.length
      setStatus(`Uploaded ${uploaded}/${rows.length} records to ${selectedRegister.label}...`)
    }

    setStatus(`Uploaded ${rows.length} records to ${selectedRegister.label}.`)
  }

  async function handleUpload(file?: File) {
    if (!file) return

    setBusy(true)
    setFileName(file.name)
    setHeadings([])
    setPreview([])

    try {
      setStatus(`Reading ${file.name}...`)

      const parsed = await readWorkbook(file)
      setHeadings(parsed.headings)

      const rows = parseRows(selected, parsed.rows, parsed.allCells)
      setPreview(rows.slice(0, 12))

      if (!rows.length) {
        setStatus(`No usable rows found for ${selectedRegister.label}. Choose the correct register on the left.`)
        return
      }

      await saveRows(selected, rows)
    } catch (err: any) {
      setStatus(`Upload failed: ${err?.message || 'Unknown error'}`)
    } finally {
      setBusy(false)
    }
  }

  return (
    <main className="import-shell">
      <style>{`
        :root {
          --navy: #071a33;
          --navy2: #0b2b52;
          --orange: #ff7a1a;
          --blue: #9fe6ff;
          --card: rgba(9,31,58,.94);
          --line: rgba(159,230,255,.24);
          --text: #f4f8ff;
          --muted: #a9bad1;
        }

        body {
          margin: 0;
          background:
            radial-gradient(circle at top left, rgba(255,122,26,.22), transparent 34%),
            linear-gradient(135deg, #030912 0%, #071a33 48%, #0b2b52 100%);
          color: var(--text);
        }

        .import-shell {
          min-height: 100vh;
          padding: 24px;
          font-family: Arial, Helvetica, sans-serif;
        }

        .hero {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 16px;
          background: linear-gradient(135deg, var(--navy), var(--navy2));
          border: 1px solid var(--line);
          border-radius: 24px;
          padding: 22px;
          box-shadow: 0 18px 50px rgba(0,0,0,.35);
        }

        .brand {
          display: flex;
          align-items: center;
          gap: 14px;
        }

        .mark {
          width: 54px;
          height: 54px;
          border-radius: 16px;
          background: white;
          color: var(--navy);
          display: grid;
          place-items: center;
          font-weight: 900;
          font-size: 22px;
        }

        h1, h2, h3 { margin: 0; }
        p { color: var(--muted); line-height: 1.4; }

        a {
          color: var(--blue);
          font-weight: 900;
          text-decoration: none;
        }

        .grid {
          display: grid;
          grid-template-columns: 330px 1fr;
          gap: 18px;
          margin-top: 18px;
          align-items: start;
        }

        .card {
          background: var(--card);
          border: 1px solid var(--line);
          border-radius: 22px;
          padding: 18px;
          box-shadow: 0 14px 45px rgba(0,0,0,.28);
        }

        .module-search,
        .upload-box input {
          width: 100%;
          box-sizing: border-box;
          padding: 13px 14px;
          border-radius: 14px;
          border: 1px solid var(--line);
          background: rgba(255,255,255,.08);
          color: var(--text);
          outline: none;
          font-weight: 700;
        }

        .module-list {
          display: grid;
          gap: 9px;
          margin-top: 12px;
          max-height: 70vh;
          overflow: auto;
          padding-right: 4px;
        }

        .module-btn {
          text-align: left;
          border: 1px solid var(--line);
          background: rgba(255,255,255,.06);
          color: var(--text);
          border-radius: 16px;
          padding: 13px;
          cursor: pointer;
        }

        .module-btn.active {
          border-color: var(--orange);
          box-shadow: inset 4px 0 0 var(--orange);
          background: rgba(255,122,26,.16);
        }

        .module-btn b {
          display: block;
          margin-bottom: 4px;
        }

        .module-btn small {
          color: var(--muted);
        }

        .top-card {
          display: grid;
          grid-template-columns: 1fr 280px;
          gap: 16px;
          align-items: center;
        }

        .upload-box {
          border: 1px dashed rgba(159,230,255,.45);
          border-radius: 18px;
          padding: 18px;
          background: rgba(255,255,255,.06);
        }

        .btn {
          display: inline-flex;
          border: 0;
          border-radius: 14px;
          padding: 12px 16px;
          font-weight: 900;
          cursor: pointer;
          background: var(--orange);
          color: white;
        }

        .btn.secondary {
          background: rgba(255,255,255,.10);
          border: 1px solid var(--line);
        }

        .status {
          margin-top: 14px;
          padding: 14px;
          border-radius: 16px;
          background: rgba(159,230,255,.10);
          border: 1px solid var(--line);
          font-weight: 900;
        }

        .badge {
          display: inline-flex;
          align-items: center;
          border-radius: 999px;
          padding: 8px 12px;
          font-weight: 900;
          font-size: 12px;
        }

        .badge.good {
          background: rgba(34,197,94,.18);
          color: #8df7b1;
          border: 1px solid rgba(34,197,94,.38);
        }

        .badge.danger {
          background: rgba(239,68,68,.20);
          color: #ffb3b3;
          border: 1px solid rgba(239,68,68,.40);
        }

        .badge.neutral {
          background: rgba(159,230,255,.15);
          color: var(--blue);
          border: 1px solid var(--line);
        }

        .meta {
          display: grid;
          gap: 12px;
        }

        .meta div {
          background: rgba(255,255,255,.07);
          border: 1px solid var(--line);
          border-radius: 16px;
          padding: 14px;
        }

        .meta span {
          display: block;
          color: var(--muted);
          font-size: 12px;
          text-transform: uppercase;
          letter-spacing: .08em;
        }

        .meta b {
          display: block;
          margin-top: 6px;
          font-size: 16px;
          word-break: break-word;
        }

        .chips {
          display: flex;
          gap: 8px;
          flex-wrap: wrap;
          max-height: 140px;
          overflow: auto;
          margin-top: 10px;
        }

        .chip {
          background: rgba(159,230,255,.15);
          color: var(--blue);
          border: 1px solid var(--line);
          border-radius: 999px;
          padding: 8px 10px;
          font-weight: 900;
          font-size: 12px;
        }

        .table-wrap {
          overflow: auto;
          margin-top: 12px;
          border: 1px solid var(--line);
          border-radius: 16px;
        }

        table {
          width: 100%;
          border-collapse: collapse;
          min-width: 900px;
        }

        th, td {
          padding: 11px 12px;
          border-bottom: 1px solid rgba(159,230,255,.12);
          text-align: left;
          font-size: 13px;
          vertical-align: top;
        }

        th {
          color: var(--blue);
          text-transform: uppercase;
          letter-spacing: .06em;
          background: rgba(0,0,0,.24);
        }

        .empty {
          padding: 18px;
          border-radius: 16px;
          border: 1px dashed var(--line);
          color: var(--muted);
          margin-top: 10px;
        }

        .help-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 12px;
          margin-top: 18px;
        }

        .help-grid .card {
          padding: 14px;
        }

        .help-grid b {
          color: var(--blue);
        }

        @media (max-width: 980px) {
          .grid, .top-card, .help-grid {
            grid-template-columns: 1fr;
          }

          .import-shell {
            padding: 12px;
          }

          .hero {
            align-items: flex-start;
            flex-direction: column;
          }
        }
      `}</style>

      <section className="hero">
        <div className="brand">
          <div className="mark">TE</div>
          <div>
            <h1>Turbo Energy Smart Excel Importer</h1>
            <p>Fixed importer for personnel, leave, fleet, spares, tyres, batteries, hoses, tools and workshop sections.</p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <a className="btn secondary" href="/">Back to app</a>
          <Badge value={isSupabaseConfigured ? 'Supabase connected' : 'Supabase not configured'} />
        </div>
      </section>

      <section className="grid">
        <aside className="card">
          <h2>Choose register</h2>
          <p>Select what the Excel file is for, then upload.</p>

          <input
            className="module-search"
            placeholder="Search register..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />

          <div className="module-list">
            {filtered.map((item) => (
              <button
                type="button"
                key={item.table}
                className={`module-btn ${selected === item.table ? 'active' : ''}`}
                onClick={() => {
                  setSelected(item.table)
                  setStatus('Ready')
                  setFileName('')
                  setHeadings([])
                  setPreview([])
                }}
              >
                <b>{item.label}</b>
                <small>{item.description}</small>
              </button>
            ))}
          </div>
        </aside>

        <section>
          <div className="card top-card">
            <div>
              <h2>{selectedRegister.label}</h2>
              <p>{selectedRegister.description}</p>

              <div className="upload-box">
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busy}
                  onChange={(e) => {
                    handleUpload(e.target.files?.[0])
                    e.currentTarget.value = ''
                  }}
                />

                <p>
                  Personnel reads FIRST NAME, SURNAME, CURRENT JOB TITLE and Department.
                  Leave reads Code, FirstName, Surname, Occupation, Department, DATES, TOTAL and LEAVE TYPE.
                </p>
              </div>

              <div className="status">{status}</div>
            </div>

            <div className="meta">
              <div><span>Selected</span><b>{selectedRegister.label}</b></div>
              <div><span>Database table</span><b>{selectedRegister.table}</b></div>
              <div><span>File</span><b>{fileName || 'None'}</b></div>
              <div><span>Save mode</span><b>Upsert / update</b></div>
            </div>
          </div>

          <div className="help-grid">
            <div className="card">
              <b>Personnel sample</b>
              <p>Choose Personnel Register, then upload Workshop attendance.xlsx.</p>
            </div>

            <div className="card">
              <b>Leave sample</b>
              <p>Choose Leave / Off Schedule, then upload LEAVE SUMMARY.xlsx.</p>
            </div>

            <div className="card">
              <b>After upload</b>
              <p>When it says Uploaded, go back to the app and press CTRL + F5.</p>
            </div>
          </div>

          <div className="card" style={{ marginTop: 18 }}>
            <h3>Headings found</h3>

            {headings.length ? (
              <div className="chips">
                {headings.map((h) => <span className="chip" key={h}>{h}</span>)}
              </div>
            ) : (
              <div className="empty">No file loaded yet.</div>
            )}
          </div>

          <div className="card" style={{ marginTop: 18 }}>
            <h3>Preview before save</h3>

            {preview.length ? (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {Object.keys(preview[0]).slice(0, 12).map((key) => <th key={key}>{key}</th>)}
                    </tr>
                  </thead>

                  <tbody>
                    {preview.map((row) => (
                      <tr key={row.id || JSON.stringify(row)}>
                        {Object.keys(preview[0]).slice(0, 12).map((key) => (
                          <td key={key}>{String(row[key] ?? '')}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="empty">After upload, mapped records will show here.</div>
            )}
          </div>
        </section>
      </section>
    </main>
  )
}
