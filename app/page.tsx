'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { isSupabaseConfigured, supabase } from '@/lib/supabase'

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

type ImportTarget = {
  table: TableName
  label: string
  description: string
}

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
  ['importer', 'Excel Importer'],
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

const IMPORT_TARGETS: ImportTarget[] = [
  { table: 'fleet_machines', label: 'Fleet Register', description: 'Machine fleet numbers, type, hours, mileage, location and status.' },
  { table: 'personnel', label: 'Personnel Register', description: 'Workshop attendance, staff, shift, section, trade and leave balance.' },
  { table: 'leave_records', label: 'Leave / Off Schedule', description: 'LEAVE SUMMARY.xlsx, current leave, upcoming leave and past leave.' },
  { table: 'services', label: 'Service Schedule', description: 'Service due list, job cards, supervisors and service history.' },
  { table: 'breakdowns', label: 'Breakdowns', description: 'Breakdown sheets, faults, spares ETA, fitters and availability.' },
  { table: 'repairs', label: 'Repairs / Job Cards', description: 'Repair jobs, assigned people, parts used and job cards.' },
  { table: 'spares', label: 'Spares Stock', description: 'Parts stock, part numbers, stock quantity, supplier and shelf.' },
  { table: 'spares_orders', label: 'Spares Orders', description: 'Spares requests, awaiting funding, funded, local/international and ETA.' },
  { table: 'tyres', label: 'Tyres Register', description: 'Tyre make, serial, company number, fleet, fitment and measurements.' },
  { table: 'tyre_measurements', label: 'Tyre Measurements', description: 'Tread depth, tyre condition, pressure and measurement reminders.' },
  { table: 'tyre_orders', label: 'Tyre Orders', description: 'Tyre requests, quotes, stock and orders.' },
  { table: 'batteries', label: 'Battery Register', description: 'Battery make, serial, fleet, fitment, voltage and maintenance reminders.' },
  { table: 'battery_orders', label: 'Battery Orders', description: 'Battery requests and orders against machines.' },
  { table: 'operations', label: 'Operations / Operators', description: 'Operators, machines, scores, offences, damages and abuse pictures.' },
  { table: 'hoses', label: 'Hose Stock', description: 'Hose numbers, hose types, fittings, sizes and stock levels.' },
  { table: 'hose_requests', label: 'Hose Requests', description: 'Bulk hose requests and individual hose repair requests.' },
  { table: 'tools', label: 'Tools Inventory', description: 'Workshop tools, brands, quantities, serials and condition.' },
  { table: 'tool_orders', label: 'Tool Orders', description: 'Requests to order tools.' },
  { table: 'employee_tools', label: 'Employee Issued Tools', description: 'Tools issued to employees and return records.' },
  { table: 'fabrication_stock', label: 'Boiler/Panel Stock', description: 'Boiler shop and panel beater stock and materials.' },
  { table: 'fabrication_requests', label: 'Boiler/Panel Requests', description: 'Material requests booked to machines.' },
  { table: 'fabrication_projects', label: 'Boiler/Panel Projects', description: 'Major ongoing projects and material usage.' }
]

const shiftNames = ['Shift 1', 'Shift 2', 'Shift 3']
const personnelRoles = [
  'Manager',
  'Foreman',
  'Chargehand',
  'Diesel Plant Fitter',
  'Assistant',
  'Tyre Fitter',
  'Boiler Maker',
  'Auto Electrical',
  'Panel Beater',
  'Wash Bay Attendant',
  'Driver',
  'Diesel Mechanic'
]
const sections = ['Workshop', 'Mechanical', 'Auto Electrical', 'Tyre Bay', 'Boiler Shop', 'Panel Beating', 'Stores', 'Wash Bay', 'Operations']
const leaveTypes = ['Annual Leave', 'Sick Leave', 'Off Duty', 'Training', 'Unpaid Leave', 'AWOL']
const statuses = ['Available', 'Open', 'In progress', 'Requested', 'Awaiting Funding', 'Funded', 'Awaiting Delivery', 'Delivered', 'Completed', 'Approved', 'Active', 'Inactive']
const orderStages = ['Requested', 'Awaiting Funding', 'Funded', 'Awaiting Delivery', 'Delivered', 'Cancelled']
const priorities = ['Low', 'Normal', 'High', 'Critical']

const SNAPSHOT_KEY = 'turbo-workshop-snapshot'
const QUEUE_KEY = 'turbo-workshop-offline-queue'

const today = () => new Date().toISOString().slice(0, 10)
const nowIso = () => new Date().toISOString()
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
const num = (value: any, fallback = 0) => {
  const parsed = Number(String(value ?? '').replace(/,/g, '').trim())
  return Number.isFinite(parsed) ? parsed : fallback
}

function normal(value: any) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
}

function slug(value: any) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
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

function excelDateNumber(value: number) {
  const parsed = XLSX.SSF.parse_date_code(value)
  if (!parsed) return ''
  return `${String(parsed.y).padStart(4, '0')}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`
}

function dateOnly(value: any) {
  if (!value) return ''

  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  if (typeof value === 'number' && value > 20000) return excelDateNumber(value)

  const text = String(value).trim()
  if (!text) return ''

  let match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/)
  if (match) {
    const dd = match[1].padStart(2, '0')
    const mm = match[2].padStart(2, '0')
    const yyyy = match[3].length === 2 ? `20${match[3]}` : match[3]
    return `${yyyy}-${mm}-${dd}`
  }

  match = text.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/)
  if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`

  const parsed = new Date(text)
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10)

  return ''
}

function dateRange(value: any) {
  const text = String(value || '').trim()
  const matches = text.match(/\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/g) || []

  if (matches.length >= 2) return [dateOnly(matches[0]), dateOnly(matches[1])]
  if (matches.length === 1) return [dateOnly(matches[0]), dateOnly(matches[0])]

  return ['', '']
}

function daysInclusive(start: any, end: any) {
  const s = new Date(start)
  const e = new Date(end)

  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return 0

  return Math.max(0, Math.floor((e.getTime() - s.getTime()) / 86400000) + 1)
}

function isCurrentLeave(row: Row) {
  const start = row.start_date
  const end = row.end_date
  const t = new Date(today()).getTime()
  const s = new Date(start).getTime()
  const e = new Date(end).getTime()
  return !Number.isNaN(s) && !Number.isNaN(e) && t >= s && t <= e
}

function isUpcomingLeave(row: Row) {
  const start = new Date(row.start_date).getTime()
  const t = new Date(today()).getTime()
  return !Number.isNaN(start) && start > t
}

function isPastLeave(row: Row) {
  const end = new Date(row.end_date).getTime()
  const t = new Date(today()).getTime()
  return !Number.isNaN(end) && end < t
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

function fullName(row: Row) {
  const first = getCell(row, ['FIRST NAME', 'FirstName', 'First Name', 'Forename'])
  const surname = getCell(row, ['SURNAME', 'Surname', 'Last Name'])
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
  const headings = new Set<string>()
  const allCells: any[][] = []

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true }) as any[][]

    matrix.forEach((line) => allCells.push(line))

    if (!matrix.length) return

    let headerIndex = 0
    let bestScore = -1

    matrix.forEach((line, index) => {
      const score = headerScore(line)
      if (score > bestScore) {
        bestScore = score
        headerIndex = index
      }
    })

    const headerRow = matrix[headerIndex] || []
    const headerNames = headerRow.map((h, i) => {
      const clean = String(h || '').trim()
      return clean || `Column ${i + 1}`
    })

    headerNames.forEach((h) => headings.add(h))

    matrix.slice(headerIndex + 1).forEach((line, offset) => {
      const obj: Row = {
        __sheet: sheetName,
        __row: headerIndex + offset + 2,
        __cells: line
      }

      let filled = 0

      headerNames.forEach((heading, index) => {
        const value = line[index]
        if (value !== undefined && value !== null && String(value).trim() !== '') {
          obj[heading] = value
          filled += 1
        }
      })

      if (filled > 0) rows.push(obj)
    })
  })

  return { rows, headings: Array.from(headings), allCells }
}

function fleetFromCells(cells: any[]) {
  const text = cells.map((c) => String(c || '')).join(' ')
  const parts = text.split(/[\s,;:/|]+/).map((x) => x.trim()).filter(Boolean)

  const banned = new Set(['FLEET', 'MACHINE', 'EQUIPMENT', 'TYPE', 'DESCRIPTION', 'DATE', 'TOTAL', 'STATUS', 'DEPARTMENT', 'WORKSHOP'])

  for (const part of parts) {
    const clean = part.toUpperCase().replace(/[^A-Z0-9-]/g, '')
    if (!clean) continue
    if (banned.has(clean)) continue
    if (/^\d+$/.test(clean)) continue
    if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(clean)) continue

    if (/^[A-Z]{1,6}-?\d{1,6}[A-Z]?$/.test(clean)) return clean.replace(/-/g, '')
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
  const employeeNo = String(getCell(row, ['Code', 'Employee Code', 'Employee No', 'Clock No', 'Emp No']) || '').trim()
  const role = String(getCell(row, ['Occupation', 'CURRENT JOB TITLE', 'Current Job Title', 'Job Title', 'Trade', 'Role']) || 'Workshop Employee').trim()
  const section = String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim()

  if (!name) return null
  if (normal(name).includes('workshopdepartment')) return null
  if (normal(name).includes('firstname')) return null

  return cleanRow({
    id: employeeNo || slug(name),
    employee_no: employeeNo,
    name,
    shift: String(getCell(row, ['Shift']) || 'Shift 1').trim(),
    section,
    role,
    staff_group: 'Workshop Staff',
    phone: String(getCell(row, ['Phone', 'Cell', 'Mobile']) || '').trim(),
    employment_status: String(getCell(row, ['Status', 'Employment Status']) || 'Active').trim(),
    leave_balance_days: num(getCell(row, ['Leave Balance', 'Leave Days']), 30),
    leave_taken_days: num(getCell(row, ['Leave Taken', 'Days Taken']), 0),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  })
}

function parseLeave(row: Row) {
  const name = fullName(row)
  const employeeNo = String(getCell(row, ['Code', 'Employee Code', 'Employee No', 'Clock No', 'Emp No']) || '').trim()
  const role = String(getCell(row, ['Occupation', 'Current Job Title', 'Job Title', 'Trade', 'Role']) || '').trim()
  const section = String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim()

  const datesText = getCell(row, ['DATES', 'Dates', 'Date Range', 'Leave Dates'])
  const [rangeStart, rangeEnd] = dateRange(datesText)

  const start = dateOnly(getCell(row, ['Start Date', 'Leave Start', 'From'])) || rangeStart
  const end = dateOnly(getCell(row, ['End Date', 'Leave End', 'To'])) || rangeEnd

  if (!name && !employeeNo) return null
  if (!start && !end) return null

  const leaveType = String(getCell(row, ['LEAVE TYPE', 'Leave Type', 'Type']) || 'Annual Leave').trim()
  const leaveDays = num(getCell(row, ['TOTAL', 'Total', 'Days', 'Leave Days']), daysInclusive(start, end))
  const id = `${employeeNo || slug(name)}-${start}-${end}-${slug(leaveType)}`

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
    days: leaveDays,
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
    machine_type: String(getCell(row, ['Machine Type', 'Equipment Type', 'Type']) || '').trim() || inferMachineType(fleetNo, cellText),
    make_model: String(getCell(row, ['Make Model', 'Make', 'Model', 'Description']) || '').trim(),
    reg_no: String(getCell(row, ['Reg', 'Registration', 'Reg No']) || '').trim(),
    department: String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim(),
    location: String(getCell(row, ['Location', 'Site']) || '').trim(),
    hours: num(getCell(row, ['Hours', 'Hrs', 'HM', 'Hour Meter']), 0),
    mileage: num(getCell(row, ['Mileage', 'KM', 'Odometer']), 0),
    status: String(getCell(row, ['Status', 'State']) || 'Available').trim(),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  })
}

function parseGeneric(table: TableName, row: Row) {
  const fleetNo = String(getCell(row, ['Fleet', 'Fleet No', 'Machine', 'Unit']) || '').trim()
  const partNo = String(getCell(row, ['Part No', 'Part Number', 'Stock Code', 'Item No']) || '').trim()
  const description = String(getCell(row, ['Description', 'Item', 'Name', 'Tool Name', 'Tool']) || '').trim()
  const serial = String(getCell(row, ['Serial', 'Serial No', 'Serial Number']) || '').trim()
  const requestedBy = String(getCell(row, ['Requested By', 'Requester', 'Foreman']) || '').trim()

  if (table === 'services') {
    return cleanRow({
      id: uid(),
      fleet_no: fleetNo,
      service_type: String(getCell(row, ['Service Type', 'Service', 'Type']) || 'Service').trim(),
      due_date: dateOnly(getCell(row, ['Due Date', 'Date'])),
      scheduled_hours: num(getCell(row, ['Scheduled Hours', 'Due Hours', 'Hours', 'HM']), 0),
      completed_date: dateOnly(getCell(row, ['Completed Date', 'Done Date'])),
      completed_hours: num(getCell(row, ['Completed Hours', 'Done Hours']), 0),
      job_card_no: String(getCell(row, ['Job Card', 'Job Card No']) || '').trim(),
      supervisor: String(getCell(row, ['Supervisor', 'Foreman']) || '').trim(),
      technician: String(getCell(row, ['Technician', 'Fitter', 'Mechanic']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Due').trim(),
      notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
    })
  }

  if (table === 'breakdowns') {
    return cleanRow({
      id: uid(),
      fleet_no: fleetNo,
      fault: String(getCell(row, ['Fault', 'Problem', 'Breakdown', 'Defect', 'Description']) || '').trim(),
      category: String(getCell(row, ['Category', 'Section']) || 'Mechanical').trim(),
      reported_by: String(getCell(row, ['Reported By', 'Operator']) || '').trim(),
      assigned_to: String(getCell(row, ['Assigned To', 'Fitter', 'Mechanic', 'Technician']) || '').trim(),
      start_time: String(getCell(row, ['Start Time', 'Date', 'Date Reported']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Open').trim(),
      spare_eta: dateOnly(getCell(row, ['ETA', 'Spare ETA', 'Spares ETA'])),
      notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
    })
  }

  if (table === 'repairs') {
    return cleanRow({
      id: uid(),
      fleet_no: fleetNo,
      job_card_no: String(getCell(row, ['Job Card', 'Job Card No']) || '').trim(),
      fault: String(getCell(row, ['Fault', 'Problem', 'Defect', 'Description']) || '').trim(),
      assigned_to: String(getCell(row, ['Assigned To', 'Fitter', 'Mechanic', 'Technician']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'In progress').trim(),
      parts_used: String(getCell(row, ['Parts Used', 'Spares Used']) || '').trim(),
      notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
    })
  }

  if (table === 'spares') {
    return cleanRow({
      id: partNo || slug(description) || uid(),
      part_no: partNo,
      description,
      section: String(getCell(row, ['Section', 'Department']) || 'Stores').trim(),
      machine_type: String(getCell(row, ['Machine Type', 'Type']) || '').trim(),
      stock_qty: num(getCell(row, ['Stock Qty', 'Stock', 'Balance', 'Qty', 'Quantity']), 0),
      min_qty: num(getCell(row, ['Min Qty', 'Minimum', 'Reorder Level']), 1),
      supplier: String(getCell(row, ['Supplier', 'Vendor']) || '').trim(),
      shelf_location: String(getCell(row, ['Shelf', 'Bin', 'Location']) || '').trim(),
      order_status: String(getCell(row, ['Status', 'Order Status']) || 'In stock').trim()
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
      qty: num(getCell(row, ['Qty', 'Quantity']), 1),
      priority: String(getCell(row, ['Priority']) || 'Normal').trim(),
      workflow_stage: String(getCell(row, ['Stage', 'Workflow Stage', 'Status']) || 'Requested').trim(),
      status: String(getCell(row, ['Status', 'Stage']) || 'Requested').trim(),
      eta: dateOnly(getCell(row, ['ETA', 'Expected Date'])),
      notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
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
      tyre_type: String(getCell(row, ['Tyre Type', 'Type']) || '').trim(),
      size: String(getCell(row, ['Size']) || '').trim(),
      status: String(getCell(row, ['Status']) || 'Fitted').trim()
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
      status: String(getCell(row, ['Status']) || 'Fitted').trim()
    })
  }

  if (table === 'hoses' || table === 'tools' || table === 'fabrication_stock') {
    return cleanRow({
      id: serial || partNo || slug(description) || uid(),
      hose_no: String(getCell(row, ['Hose No', 'Hose Number']) || '').trim(),
      hose_type: String(getCell(row, ['Hose Type', 'Type']) || 'Hydraulic Hose').trim(),
      tool_name: description || String(getCell(row, ['Tool', 'Tool Name']) || '').trim(),
      description,
      category: String(getCell(row, ['Category']) || 'Hand Tools').trim(),
      brand: String(getCell(row, ['Brand', 'Make']) || '').trim(),
      serial_no: serial,
      stock_qty: num(getCell(row, ['Stock Qty', 'Qty', 'Quantity', 'Stock']), 0),
      min_qty: num(getCell(row, ['Min Qty', 'Minimum']), 1),
      condition: String(getCell(row, ['Condition']) || 'Good').trim(),
      status: String(getCell(row, ['Status']) || 'In stock').trim()
    })
  }

  const hasSomething = fleetNo || partNo || description || serial || requestedBy
  if (!hasSomething) return null

  return cleanRow({
    id: uid(),
    fleet_no: fleetNo,
    machine_fleet_no: fleetNo,
    part_no: partNo,
    description,
    serial_no: serial,
    requested_by: requestedBy,
    qty: num(getCell(row, ['Qty', 'Quantity']), 1),
    stock_qty: num(getCell(row, ['Stock Qty', 'Stock', 'Balance', 'Qty']), 0),
    status: String(getCell(row, ['Status', 'Stage']) || 'Imported').trim(),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  })
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

    map.set(key, { ...row, id: key })
  })

  return Array.from(map.values())
}

function readSnapshot(): AppData | null {
  try {
    const raw = localStorage.getItem(SNAPSHOT_KEY)
    if (!raw) return null
    return JSON.parse(raw).data as AppData
  } catch {
    return null
  }
}

function saveSnapshot(data: AppData) {
  try {
    localStorage.setItem(SNAPSHOT_KEY, JSON.stringify({ updatedAt: nowIso(), data }))
  } catch {
    // ignore localStorage limit
  }
}

function addToOfflineQueue(table: TableName, rows: Row[]) {
  try {
    const raw = localStorage.getItem(QUEUE_KEY)
    const queue = raw ? JSON.parse(raw) : []

    rows.forEach((row) => {
      queue.push({
        id: uid(),
        table,
        type: 'upsert',
        payload: row,
        createdAt: nowIso()
      })
    })

    localStorage.setItem(QUEUE_KEY, JSON.stringify(queue))
  } catch {
    // ignore queue failure
  }
}

function mergeRows(existing: Row[], incoming: Row[]) {
  const map = new Map<string, Row>()

  existing.forEach((row) => {
    const key = String(row.id || row.fleet_no || row.employee_no || JSON.stringify(row))
    map.set(key, row)
  })

  incoming.forEach((row) => {
    const key = String(row.id || row.fleet_no || row.employee_no || JSON.stringify(row))
    map.set(key, { ...map.get(key), ...row })
  })

  return Array.from(map.values())
}

function statusClass(value: any) {
  const text = String(value || '').toLowerCase()

  if (text.includes('open') || text.includes('break') || text.includes('awaiting') || text.includes('low') || text.includes('due') || text.includes('critical')) return 'danger'
  if (text.includes('request') || text.includes('fund') || text.includes('progress') || text.includes('normal')) return 'warn'
  if (text.includes('active') || text.includes('available') || text.includes('complete') || text.includes('delivered') || text.includes('approved') || text.includes('stock')) return 'good'

  return 'neutral'
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

function Input({ value, onChange, placeholder = '', type = 'text' }: { value: any; onChange: (value: any) => void; placeholder?: string; type?: string }) {
  return <input type={type} value={value ?? ''} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />
}

function Select({ value, onChange, options }: { value: any; onChange: (value: any) => void; options: string[] }) {
  return (
    <select value={value ?? ''} onChange={(e) => onChange(e.target.value)}>
      {options.map((option) => <option key={option}>{option}</option>)}
    </select>
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

function MiniTable({ rows, columns, empty = 'No records yet.' }: { rows: Row[]; columns: { key: string; label: string; render?: (row: Row) => any }[]; empty?: string }) {
  if (!rows.length) return <div className="empty">{empty}</div>

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {columns.map((col) => <th key={col.key}>{col.label}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id || JSON.stringify(row)}>
              {columns.map((col) => <td key={col.key}>{col.render ? col.render(row) : String(row[col.key] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function PageTitle({ title, subtitle, right }: { title: string; subtitle?: string; right?: any }) {
  return (
    <div className="page-title">
      <div>
        <h1>{title}</h1>
        {subtitle && <p>{subtitle}</p>}
      </div>
      {right && <div className="page-actions">{right}</div>}
    </div>
  )
}

export default function Home() {
  const [session, setSession] = useState<Session | null>(null)
  const [loginName, setLoginName] = useState('admin')
  const [loginPass, setLoginPass] = useState('admin123')
  const [tab, setTab] = useState('dashboard')
  const [data, setData] = useState<AppData>(emptyData)
  const [message, setMessage] = useState('')
  const [online, setOnline] = useState(true)
  const [selectedImport, setSelectedImport] = useState<TableName>('personnel')
  const [importSearch, setImportSearch] = useState('')
  const [importStatus, setImportStatus] = useState('Ready')
  const [importFile, setImportFile] = useState('')
  const [importHeadings, setImportHeadings] = useState<string[]>([])
  const [importPreview, setImportPreview] = useState<Row[]>([])
  const [busyImport, setBusyImport] = useState(false)

  const [machineForm, setMachineForm] = useState<Row>({ fleet_no: '', machine_type: '', make_model: '', department: 'Workshop', location: '', hours: 0, mileage: 0, status: 'Available' })
  const [personForm, setPersonForm] = useState<Row>({ employee_no: '', name: '', shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', employment_status: 'Active', leave_balance_days: 30, leave_taken_days: 0 })
  const [leaveForm, setLeaveForm] = useState<Row>({ employee_no: '', person_name: '', leave_type: 'Annual Leave', start_date: today(), end_date: today(), status: 'Approved' })
  const [quickForm, setQuickForm] = useState<Row>({})

  const selectedImportTarget = IMPORT_TARGETS.find((target) => target.table === selectedImport) || IMPORT_TARGETS[0]

  const filteredTargets = useMemo(() => {
    const q = importSearch.toLowerCase()
    return IMPORT_TARGETS.filter((target) => `${target.label} ${target.description} ${target.table}`.toLowerCase().includes(q))
  }, [importSearch])

  const dashboard = useMemo(() => {
    const totalFleet = data.fleet_machines.length
    const breakdowns = data.breakdowns.filter((row) => !String(row.status || '').toLowerCase().includes('complete'))
    const repairs = data.repairs.filter((row) => !String(row.status || '').toLowerCase().includes('complete'))
    const servicesDue = data.services.filter((row) => String(row.status || '').toLowerCase().includes('due'))
    const currentLeave = data.leave_records.filter(isCurrentLeave)
    const foremen = data.personnel.filter((row) => String(row.role || '').toLowerCase().includes('foreman') && !String(row.employment_status || '').toLowerCase().includes('inactive'))
    const lowSpares = data.spares.filter((row) => num(row.stock_qty) <= num(row.min_qty, 1))
    const hoseRequests = data.hose_requests.filter((row) => !String(row.status || '').toLowerCase().includes('complete'))
    const toolOrders = data.tool_orders.filter((row) => !String(row.status || '').toLowerCase().includes('delivered'))

    return { totalFleet, breakdowns, repairs, servicesDue, currentLeave, foremen, lowSpares, hoseRequests, toolOrders }
  }, [data])

  useEffect(() => {
    const savedSession = localStorage.getItem('turbo-workshop-session')
    if (savedSession) setSession(JSON.parse(savedSession))

    const snapshot = readSnapshot()
    if (snapshot) setData({ ...emptyData, ...snapshot })

    setOnline(typeof navigator === 'undefined' ? true : navigator.onLine)

    const up = () => setOnline(true)
    const down = () => setOnline(false)

    window.addEventListener('online', up)
    window.addEventListener('offline', down)

    loadAll()

    return () => {
      window.removeEventListener('online', up)
      window.removeEventListener('offline', down)
    }
  }, [])

  useEffect(() => {
    saveSnapshot(data)
  }, [data])

  function login() {
    const found = loginUsers.find((user) => user.username.toLowerCase() === loginName.toLowerCase().trim() && user.password === loginPass)

    if (!found) {
      setMessage('Wrong username or password.')
      return
    }

    const next = { username: found.username, role: found.role }
    setSession(next)
    localStorage.setItem('turbo-workshop-session', JSON.stringify(next))
    setMessage('Login successful.')
  }

  function logout() {
    setSession(null)
    localStorage.removeItem('turbo-workshop-session')
  }

  async function loadAll() {
    if (!supabase || !isSupabaseConfigured) return

    const current = readSnapshot() || emptyData
    const next: AppData = { ...emptyData, ...current }

    for (const table of TABLES) {
      try {
        const { data: rows } = await supabase.from(table as any).select('*').limit(10000)

        if (rows && Array.isArray(rows)) {
          next[table] = mergeRows(next[table] || [], rows as Row[])
        }
      } catch {
        // leave local data in place
      }
    }

    setData(next)
    saveSnapshot(next)
  }

  async function saveRows(table: TableName, rows: Row[]) {
    const stamped = rows.map((row) => cleanRow({ ...row, id: row.id || uid(), updated_at: nowIso() }))

    setData((prev) => {
      const next = { ...prev, [table]: mergeRows(prev[table] || [], stamped) }
      saveSnapshot(next)
      return next
    })

    if (!supabase || !isSupabaseConfigured || !navigator.onLine) {
      addToOfflineQueue(table, stamped)
      setMessage(`Saved ${stamped.length} record(s) locally. They will sync when online.`)
      return
    }

    let uploaded = 0

    for (let i = 0; i < stamped.length; i += 50) {
      const chunk = stamped.slice(i, i + 50)
      const { error } = await supabase.from(table as any).upsert(chunk, { onConflict: 'id' })

      if (error) {
        addToOfflineQueue(table, chunk)
        setMessage(`Saved locally because Supabase rejected the upload: ${error.message}`)
        return
      }

      uploaded += chunk.length
      setImportStatus(`Uploaded ${uploaded}/${stamped.length} records...`)
    }

    setMessage(`Saved ${stamped.length} record(s) online.`)
  }

  async function handleExcelUpload(table: TableName, file?: File) {
    if (!file) return

    setBusyImport(true)
    setImportFile(file.name)
    setImportStatus(`Reading ${file.name}...`)
    setImportHeadings([])
    setImportPreview([])

    try {
      const parsed = await readWorkbook(file)
      setImportHeadings(parsed.headings)

      const rows = parseRows(table, parsed.rows, parsed.allCells)
      setImportPreview(rows.slice(0, 15))

      if (!rows.length) {
        setImportStatus(`No usable rows found for ${IMPORT_TARGETS.find((t) => t.table === table)?.label || table}.`)
        return
      }

      await saveRows(table, rows)
      setImportStatus(`Uploaded ${rows.length} records to ${IMPORT_TARGETS.find((t) => t.table === table)?.label || table}.`)
    } catch (err: any) {
      setImportStatus(`Upload failed: ${err?.message || 'Unknown error'}`)
    } finally {
      setBusyImport(false)
    }
  }

  function renderLogin() {
    return (
      <main className="login-shell">
        <style>{baseStyles}</style>

        <div className="login-card">
          <div className="logo-box">TE</div>
          <h1>Turbo Energy Workshop Control</h1>
          <p>Workshop fleet repairs, services, spares, tyres, batteries, personnel and operations tracking.</p>

          <Field label="Username">
            <Input value={loginName} onChange={setLoginName} />
          </Field>

          <Field label="Password">
            <Input value={loginPass} onChange={setLoginPass} type="password" />
          </Field>

          <button className="primary full" onClick={login}>Login</button>

          {message && <div className="message">{message}</div>}
        </div>
      </main>
    )
  }

  function renderDashboard() {
    return (
      <div>
        <PageTitle
          title="Workshop Dashboard"
          subtitle="Current shift, foremen, breakdowns, services, spares, hoses, tools, tyres, batteries and projects in one live control room."
          right={<><Badge value="Current shift working" /><Badge value={today()} /></>}
        />

        <div className="stats">
          <StatCard title="Total Fleet" value={dashboard.totalFleet} detail="Registered machines" onClick={() => setTab('fleet')} />
          <StatCard title="Machines on Breakdown" value={dashboard.breakdowns.length} detail="Open breakdown records" tone="purple" onClick={() => setTab('breakdowns')} />
          <StatCard title="Machines Being Worked On" value={dashboard.repairs.length} detail="Active repairs" tone="brown" onClick={() => setTab('repairs')} />
          <StatCard title="Services Due" value={dashboard.servicesDue.length} detail="Service today/due" tone="indigo" onClick={() => setTab('services')} />
          <StatCard title="People Off" value={dashboard.currentLeave.length} detail="Leave/sick/off duty" tone="steel" onClick={() => setTab('personnel')} />
          <StatCard title="Foremen On Duty" value={dashboard.foremen.length} detail="Loaded foremen" tone="green" onClick={() => setTab('personnel')} />
          <StatCard title="Hose Requests" value={dashboard.hoseRequests.length} detail="Pending hoses/fittings" onClick={() => setTab('hoses')} />
          <StatCard title="Tool Orders" value={dashboard.toolOrders.length} detail="Requested tools" tone="brown" onClick={() => setTab('tools')} />
        </div>

        <div className="grid-2">
          <section className="card">
            <PageTitle title="Machines on breakdown / under repair" right={<button onClick={() => setTab('breakdowns')}>Open breakdowns</button>} />
            <MiniTable
              rows={dashboard.breakdowns.slice(0, 8)}
              empty="No open breakdowns."
              columns={[
                { key: 'fleet_no', label: 'Fleet' },
                { key: 'fault', label: 'Fault' },
                { key: 'assigned_to', label: 'Assigned' },
                { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
              ]}
            />
          </section>

          <section className="card">
            <PageTitle title="Machines on service today" right={<button onClick={() => setTab('services')}>Open services</button>} />
            <MiniTable
              rows={dashboard.servicesDue.slice(0, 8)}
              empty="No services due."
              columns={[
                { key: 'fleet_no', label: 'Fleet' },
                { key: 'service_type', label: 'Service' },
                { key: 'due_date', label: 'Due Date' },
                { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
              ]}
            />
          </section>

          <section className="card">
            <PageTitle title="Foremen on duty" right={<button onClick={() => setTab('personnel')}>Personnel</button>} />
            <MiniTable
              rows={dashboard.foremen.slice(0, 8)}
              empty="No foremen loaded yet."
              columns={[
                { key: 'name', label: 'Name' },
                { key: 'shift', label: 'Shift' },
                { key: 'section', label: 'Section' },
                { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }
              ]}
            />
          </section>

          <section className="card">
            <PageTitle title="Workshop alerts" right={<button onClick={() => setTab('reports')}>Reports</button>} />
            <div className="badge-row">
              <Badge value={`Low spares: ${dashboard.lowSpares.length}`} />
              <Badge value={`Tyres due: ${data.tyre_measurements.length}`} />
              <Badge value={`Batteries due: ${data.battery_orders.length}`} />
              <Badge value={`Projects: ${data.fabrication_projects.length}`} />
            </div>
          </section>
        </div>
      </div>
    )
  }

  function renderImporter() {
    return (
      <div>
        <PageTitle
          title="Excel Importer"
          subtitle="Upload Excel files here. This fixes personnel, leave, fleet and all workshop register uploads."
          right={<Badge value={isSupabaseConfigured ? 'Supabase connected' : 'Local mode'} />}
        />

        <div className="import-layout">
          <aside className="card">
            <h2>Choose register</h2>
            <p>Select what the Excel file is for.</p>

            <input placeholder="Search register..." value={importSearch} onChange={(e) => setImportSearch(e.target.value)} />

            <div className="import-list">
              {filteredTargets.map((target) => (
                <button
                  key={target.table}
                  className={selectedImport === target.table ? 'active import-btn' : 'import-btn'}
                  onClick={() => {
                    setSelectedImport(target.table)
                    setImportStatus('Ready')
                    setImportFile('')
                    setImportHeadings([])
                    setImportPreview([])
                  }}
                >
                  <b>{target.label}</b>
                  <small>{target.description}</small>
                </button>
              ))}
            </div>
          </aside>

          <section>
            <div className="card">
              <PageTitle
                title={selectedImportTarget.label}
                subtitle={selectedImportTarget.description}
                right={<Badge value={selectedImportTarget.table} />}
              />

              <div className="upload-box">
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  disabled={busyImport}
                  onChange={(e) => {
                    handleExcelUpload(selectedImport, e.target.files?.[0])
                    e.currentTarget.value = ''
                  }}
                />
                <p>
                  Personnel reads FIRST NAME, SURNAME, CURRENT JOB TITLE and Department.
                  Leave reads Code, FirstName, Surname, Occupation, Department, DATES, TOTAL and LEAVE TYPE.
                  Fleet scans for fleet numbers even when headings are poor.
                </p>
              </div>

              <div className="message">{importStatus}</div>

              <div className="meta-grid">
                <div><span>Selected</span><b>{selectedImportTarget.label}</b></div>
                <div><span>Table</span><b>{selectedImportTarget.table}</b></div>
                <div><span>File</span><b>{importFile || 'None'}</b></div>
                <div><span>Mode</span><b>Upsert / Update</b></div>
              </div>
            </div>

            <div className="card">
              <h2>Headings found</h2>
              {importHeadings.length ? (
                <div className="chip-row">
                  {importHeadings.map((h) => <span className="chip" key={h}>{h}</span>)}
                </div>
              ) : (
                <div className="empty">No file loaded yet.</div>
              )}
            </div>

            <div className="card">
              <h2>Preview before save</h2>
              {importPreview.length ? (
                <MiniTable
                  rows={importPreview}
                  columns={Object.keys(importPreview[0]).filter((k) => !k.startsWith('__')).slice(0, 10).map((key) => ({ key, label: key }))}
                />
              ) : (
                <div className="empty">After upload, mapped rows will show here.</div>
              )}
            </div>
          </section>
        </div>
      </div>
    )
  }

  function renderFleet() {
    return (
      <div>
        <PageTitle title="Fleet Register" subtitle="Machines, fleet numbers, machine types, hours, mileage and status." right={<ExcelButton table="fleet_machines" />} />

        <section className="card">
          <h2>Add machine</h2>
          <div className="form-grid">
            <Field label="Fleet number"><Input value={machineForm.fleet_no} onChange={(v) => setMachineForm({ ...machineForm, fleet_no: v })} /></Field>
            <Field label="Machine type"><Input value={machineForm.machine_type} onChange={(v) => setMachineForm({ ...machineForm, machine_type: v })} /></Field>
            <Field label="Make / Model"><Input value={machineForm.make_model} onChange={(v) => setMachineForm({ ...machineForm, make_model: v })} /></Field>
            <Field label="Department"><Select value={machineForm.department} onChange={(v) => setMachineForm({ ...machineForm, department: v })} options={sections} /></Field>
            <Field label="Location"><Input value={machineForm.location} onChange={(v) => setMachineForm({ ...machineForm, location: v })} /></Field>
            <Field label="Hours"><Input type="number" value={machineForm.hours} onChange={(v) => setMachineForm({ ...machineForm, hours: v })} /></Field>
            <Field label="Mileage"><Input type="number" value={machineForm.mileage} onChange={(v) => setMachineForm({ ...machineForm, mileage: v })} /></Field>
            <Field label="Status"><Select value={machineForm.status} onChange={(v) => setMachineForm({ ...machineForm, status: v })} options={statuses} /></Field>
          </div>
          <button className="primary" onClick={() => saveRows('fleet_machines', [{ ...machineForm, id: machineForm.fleet_no || uid() }])}>Save machine</button>
        </section>

        <section className="card">
          <MiniTable
            rows={data.fleet_machines}
            columns={[
              { key: 'fleet_no', label: 'Fleet' },
              { key: 'machine_type', label: 'Type' },
              { key: 'make_model', label: 'Make / Model' },
              { key: 'department', label: 'Department' },
              { key: 'hours', label: 'Hours' },
              { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
            ]}
          />
        </section>
      </div>
    )
  }

  function renderPersonnel() {
    const currentLeave = data.leave_records.filter(isCurrentLeave)
    const upcomingLeave = data.leave_records.filter(isUpcomingLeave)
    const pastLeave = data.leave_records.filter(isPastLeave)

    return (
      <div>
        <PageTitle title="Personnel, Shifts and Leave" subtitle="Shift 1, Shift 2, Shift 3, managers, foremen, trades, off schedule and leave history." right={<ExcelButton table="personnel" />} />

        <div className="stats">
          <StatCard title="Employees" value={data.personnel.length} detail="Loaded personnel" />
          <StatCard title="Current Leave / Off" value={currentLeave.length} detail="Away today" tone="brown" />
          <StatCard title="Upcoming Leave" value={upcomingLeave.length} detail="Future leave records" tone="indigo" />
          <StatCard title="Past Leave Records" value={pastLeave.length} detail="History" tone="green" />
        </div>

        <div className="grid-2">
          <section className="card">
            <h2>Add employee / slot</h2>
            <div className="form-grid">
              <Field label="Employee no"><Input value={personForm.employee_no} onChange={(v) => setPersonForm({ ...personForm, employee_no: v })} /></Field>
              <Field label="Name"><Input value={personForm.name} onChange={(v) => setPersonForm({ ...personForm, name: v })} /></Field>
              <Field label="Shift"><Select value={personForm.shift} onChange={(v) => setPersonForm({ ...personForm, shift: v })} options={shiftNames} /></Field>
              <Field label="Role"><Select value={personForm.role} onChange={(v) => setPersonForm({ ...personForm, role: v })} options={personnelRoles} /></Field>
              <Field label="Section"><Select value={personForm.section} onChange={(v) => setPersonForm({ ...personForm, section: v })} options={sections} /></Field>
              <Field label="Status"><Select value={personForm.employment_status} onChange={(v) => setPersonForm({ ...personForm, employment_status: v })} options={['Active', 'Inactive', 'On Leave', 'Off Duty']} /></Field>
              <Field label="Leave balance"><Input type="number" value={personForm.leave_balance_days} onChange={(v) => setPersonForm({ ...personForm, leave_balance_days: v })} /></Field>
              <Field label="Leave taken"><Input type="number" value={personForm.leave_taken_days} onChange={(v) => setPersonForm({ ...personForm, leave_taken_days: v })} /></Field>
            </div>
            <button className="primary" onClick={() => saveRows('personnel', [{ ...personForm, id: personForm.employee_no || slug(personForm.name) || uid() }])}>Save employee</button>
          </section>

          <section className="card">
            <h2>Upload personnel / leave Excel</h2>
            <div className="upload-stack">
              <label>
                <span>Personnel Excel</span>
                <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleExcelUpload('personnel', e.target.files?.[0])} />
              </label>
              <label>
                <span>Leave Summary Excel</span>
                <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleExcelUpload('leave_records', e.target.files?.[0])} />
              </label>
            </div>
          </section>
        </div>

        <section className="card">
          <h2>Leave / off schedule</h2>
          <div className="form-grid">
            <Field label="Employee"><Input value={leaveForm.person_name} onChange={(v) => setLeaveForm({ ...leaveForm, person_name: v })} /></Field>
            <Field label="Employee no"><Input value={leaveForm.employee_no} onChange={(v) => setLeaveForm({ ...leaveForm, employee_no: v })} /></Field>
            <Field label="Leave type"><Select value={leaveForm.leave_type} onChange={(v) => setLeaveForm({ ...leaveForm, leave_type: v })} options={leaveTypes} /></Field>
            <Field label="Start date"><Input type="date" value={leaveForm.start_date} onChange={(v) => setLeaveForm({ ...leaveForm, start_date: v })} /></Field>
            <Field label="End date"><Input type="date" value={leaveForm.end_date} onChange={(v) => setLeaveForm({ ...leaveForm, end_date: v })} /></Field>
            <Field label="Status"><Select value={leaveForm.status} onChange={(v) => setLeaveForm({ ...leaveForm, status: v })} options={['Approved', 'Pending', 'Rejected']} /></Field>
          </div>
          <button className="primary" onClick={() => saveRows('leave_records', [{ ...leaveForm, id: `${leaveForm.employee_no || slug(leaveForm.person_name)}-${leaveForm.start_date}-${leaveForm.end_date}`, days: daysInclusive(leaveForm.start_date, leaveForm.end_date) }])}>Save leave</button>
        </section>

        <div className="grid-2">
          <section className="card">
            <h2>Personnel Register</h2>
            <MiniTable
              rows={data.personnel}
              columns={[
                { key: 'employee_no', label: 'Emp No' },
                { key: 'name', label: 'Name' },
                { key: 'shift', label: 'Shift' },
                { key: 'role', label: 'Role' },
                { key: 'section', label: 'Section' },
                { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }
              ]}
            />
          </section>

          <section className="card">
            <h2>Current leave / off</h2>
            <MiniTable
              rows={currentLeave}
              columns={[
                { key: 'person_name', label: 'Name' },
                { key: 'leave_type', label: 'Type' },
                { key: 'start_date', label: 'Start' },
                { key: 'end_date', label: 'End' },
                { key: 'days', label: 'Days' }
              ]}
            />
          </section>

          <section className="card">
            <h2>Upcoming leave</h2>
            <MiniTable
              rows={upcomingLeave}
              columns={[
                { key: 'person_name', label: 'Name' },
                { key: 'leave_type', label: 'Type' },
                { key: 'start_date', label: 'Start' },
                { key: 'end_date', label: 'End' },
                { key: 'days', label: 'Days' }
              ]}
            />
          </section>

          <section className="card">
            <h2>Past leave records</h2>
            <MiniTable
              rows={pastLeave}
              columns={[
                { key: 'person_name', label: 'Name' },
                { key: 'leave_type', label: 'Type' },
                { key: 'start_date', label: 'Start' },
                { key: 'end_date', label: 'End' },
                { key: 'days', label: 'Days' }
              ]}
            />
          </section>
        </div>
      </div>
    )
  }

  function ExcelButton({ table }: { table: TableName }) {
    return (
      <label className="excel-button">
        Upload Excel
        <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleExcelUpload(table, e.target.files?.[0])} />
      </label>
    )
  }

  function renderGenericPage(title: string, subtitle: string, table: TableName, columns: { key: string; label: string; render?: (row: Row) => any }[]) {
    return (
      <div>
        <PageTitle title={title} subtitle={subtitle} right={<ExcelButton table={table} />} />

        <section className="card">
          <h2>Quick add</h2>
          <div className="form-grid">
            <Field label="Fleet / Machine"><Input value={quickForm.fleet_no || quickForm.machine_fleet_no || ''} onChange={(v) => setQuickForm({ ...quickForm, fleet_no: v, machine_fleet_no: v })} /></Field>
            <Field label="Description / Fault / Item"><Input value={quickForm.description || quickForm.fault || quickForm.tool_name || ''} onChange={(v) => setQuickForm({ ...quickForm, description: v, fault: v, tool_name: v })} /></Field>
            <Field label="Requested / Assigned by"><Input value={quickForm.requested_by || quickForm.assigned_to || ''} onChange={(v) => setQuickForm({ ...quickForm, requested_by: v, assigned_to: v })} /></Field>
            <Field label="Status"><Select value={quickForm.status || 'Requested'} onChange={(v) => setQuickForm({ ...quickForm, status: v })} options={statuses} /></Field>
            <Field label="Qty"><Input type="number" value={quickForm.qty || 1} onChange={(v) => setQuickForm({ ...quickForm, qty: v })} /></Field>
            <Field label="ETA / Date"><Input type="date" value={quickForm.eta || today()} onChange={(v) => setQuickForm({ ...quickForm, eta: v, request_date: v })} /></Field>
          </div>
          <button className="primary" onClick={() => saveRows(table, [{ ...quickForm, id: uid(), request_date: quickForm.request_date || today() }])}>Save record</button>
        </section>

        <section className="card">
          <MiniTable rows={data[table]} columns={columns} />
        </section>
      </div>
    )
  }

  function renderReports() {
    const byDepartment = data.fleet_machines.reduce((acc: Row, row) => {
      const key = row.department || 'Unknown'
      acc[key] = (acc[key] || 0) + 1
      return acc
    }, {})

    const byType = data.fleet_machines.reduce((acc: Row, row) => {
      const key = row.machine_type || 'Unknown'
      acc[key] = (acc[key] || 0) + 1
      return acc
    }, {})

    return (
      <div>
        <PageTitle title="Reports" subtitle="Machine availability, breakdowns by department, fleet type and workshop pressure points." right={<button onClick={() => window.print()}>Print</button>} />

        <div className="grid-2">
          <section className="card">
            <h2>Fleet by department</h2>
            <MiniTable rows={Object.entries(byDepartment).map(([department, count]) => ({ id: department, department, count }))} columns={[{ key: 'department', label: 'Department' }, { key: 'count', label: 'Machines' }]} />
          </section>

          <section className="card">
            <h2>Fleet by type</h2>
            <MiniTable rows={Object.entries(byType).map(([machine_type, count]) => ({ id: machine_type, machine_type, count }))} columns={[{ key: 'machine_type', label: 'Type' }, { key: 'count', label: 'Machines' }]} />
          </section>

          <section className="card">
            <h2>Open breakdowns</h2>
            <MiniTable rows={dashboard.breakdowns} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'fault', label: 'Fault' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
          </section>

          <section className="card">
            <h2>Low stock</h2>
            <MiniTable rows={dashboard.lowSpares} columns={[{ key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'min_qty', label: 'Min' }]} />
          </section>
        </div>
      </div>
    )
  }

  function renderContent() {
    if (tab === 'dashboard') return renderDashboard()
    if (tab === 'importer') return renderImporter()
    if (tab === 'fleet') return renderFleet()
    if (tab === 'personnel') return renderPersonnel()
    if (tab === 'reports') return renderReports()

    if (tab === 'breakdowns') return renderGenericPage('Breakdowns', 'Breakdown reports, faults, spares ETA and assigned fitters.', 'breakdowns', [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'fault', label: 'Fault' },
      { key: 'assigned_to', label: 'Assigned' },
      { key: 'spare_eta', label: 'Spares ETA' },
      { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
    ])

    if (tab === 'repairs') return renderGenericPage('Repairs / Job Cards', 'Repairs, job cards, parts used and fitters.', 'repairs', [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'job_card_no', label: 'Job Card' },
      { key: 'fault', label: 'Fault' },
      { key: 'assigned_to', label: 'Assigned' },
      { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
    ])

    if (tab === 'services') return renderGenericPage('Services', 'Past services, job cards, supervisors and service due history.', 'services', [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'service_type', label: 'Service' },
      { key: 'due_date', label: 'Due Date' },
      { key: 'supervisor', label: 'Supervisor' },
      { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
    ])

    if (tab === 'spares') {
      return (
        <div>
          <PageTitle title="Spares & Orders" subtitle="Stock, requests, awaiting funding, local/international orders and delivery ETA." right={<ExcelButton table="spares" />} />

          <div className="stats">
            {orderStages.map((stage) => <StatCard key={stage} title={stage} value={data.spares_orders.filter((r) => String(r.workflow_stage || r.status) === stage).length} detail="Spares orders" />)}
          </div>

          <div className="grid-2">
            <section className="card">
              <h2>Spares Stock</h2>
              <ExcelButton table="spares" />
              <MiniTable rows={data.spares} columns={[{ key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'min_qty', label: 'Min' }, { key: 'order_status', label: 'Status', render: (r) => <Badge value={r.order_status} /> }]} />
            </section>

            <section className="card">
              <h2>Spares Orders</h2>
              <ExcelButton table="spares_orders" />
              <MiniTable rows={data.spares_orders} columns={[{ key: 'machine_fleet_no', label: 'Fleet' }, { key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'qty', label: 'Qty' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>
          </div>
        </div>
      )
    }

    if (tab === 'tyres') return renderGenericPage('Tyres', 'Tyre make, serial number, company number, fitment, measurements and ordering.', 'tyres', [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'serial_no', label: 'Serial' },
      { key: 'make', label: 'Make' },
      { key: 'position', label: 'Position' },
      { key: 'size', label: 'Size' },
      { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
    ])

    if (tab === 'batteries') return renderGenericPage('Batteries', 'Battery make, serial number, date fitted, fleet, charging voltage and maintenance reminders.', 'batteries', [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'serial_no', label: 'Serial' },
      { key: 'make', label: 'Make' },
      { key: 'volts', label: 'Volts' },
      { key: 'fitment_date', label: 'Fitted' },
      { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }
    ])

    if (tab === 'operations') return renderGenericPage('Operations / Operators', 'Operators by machine type, history, score sheets, offences and damages.', 'operations', [
      { key: 'operator_name', label: 'Operator' },
      { key: 'machine_type', label: 'Machine Type' },
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'score', label: 'Score' },
      { key: 'offence', label: 'Offence' },
      { key: 'supervisor', label: 'Supervisor' }
    ])

    if (tab === 'hoses') {
      return (
        <div>
          <PageTitle title="Hoses" subtitle="Hose stock, fittings, bulk requests, individual hose repairs and offsite approval." right={<ExcelButton table="hoses" />} />

          <div className="grid-2">
            <section className="card">
              <h2>Hose Stock</h2>
              <ExcelButton table="hoses" />
              <MiniTable rows={data.hoses} columns={[{ key: 'hose_no', label: 'Hose No' }, { key: 'hose_type', label: 'Type' }, { key: 'size', label: 'Size' }, { key: 'stock_qty', label: 'Stock' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>

            <section className="card">
              <h2>Hose Requests</h2>
              <ExcelButton table="hose_requests" />
              <MiniTable rows={data.hose_requests} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'hose_type', label: 'Type' }, { key: 'requested_by', label: 'Requested By' }, { key: 'qty', label: 'Qty' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>
          </div>
        </div>
      )
    }

    if (tab === 'tools') {
      return (
        <div>
          <PageTitle title="Tools" subtitle="Inventory, orders and employee issued tools." right={<ExcelButton table="tools" />} />

          <div className="grid-2">
            <section className="card">
              <h2>Tools Inventory</h2>
              <ExcelButton table="tools" />
              <MiniTable rows={data.tools} columns={[{ key: 'tool_name', label: 'Tool' }, { key: 'brand', label: 'Brand' }, { key: 'serial_no', label: 'Serial' }, { key: 'stock_qty', label: 'Qty' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>

            <section className="card">
              <h2>Employee Issued Tools</h2>
              <ExcelButton table="employee_tools" />
              <MiniTable rows={data.employee_tools} columns={[{ key: 'employee_name', label: 'Employee' }, { key: 'tool_name', label: 'Tool' }, { key: 'issue_date', label: 'Issued' }, { key: 'issued_by', label: 'Issued By' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>
          </div>
        </div>
      )
    }

    if (tab === 'fabrication') {
      return (
        <div>
          <PageTitle title="Boiler Shop / Panel Beaters" subtitle="Stock, material requests, projects and material booked to machines." right={<ExcelButton table="fabrication_stock" />} />

          <div className="grid-2">
            <section className="card">
              <h2>Material Stock</h2>
              <ExcelButton table="fabrication_stock" />
              <MiniTable rows={data.fabrication_stock} columns={[{ key: 'department', label: 'Department' }, { key: 'material_type', label: 'Material' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>

            <section className="card">
              <h2>Major Projects</h2>
              <ExcelButton table="fabrication_projects" />
              <MiniTable rows={data.fabrication_projects} columns={[{ key: 'project_name', label: 'Project' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'supervisor', label: 'Supervisor' }, { key: 'progress_percent', label: 'Progress %' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
            </section>
          </div>
        </div>
      )
    }

    if (tab === 'photos') return renderGenericPage('Photos', 'Photo captions, machine damage, abuse and inspection references.', 'photo_logs', [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'linked_type', label: 'Type' },
      { key: 'caption', label: 'Caption' },
      { key: 'notes', label: 'Notes' }
    ])

    if (tab === 'admin') {
      return (
        <div>
          <PageTitle title="Admin" subtitle="Sync, local storage and system controls." />
          <section className="card">
            <div className="button-row">
              <button className="primary" onClick={loadAll}>Sync from Supabase</button>
              <button onClick={() => window.print()}>Print</button>
              <button onClick={() => {
                localStorage.removeItem(SNAPSHOT_KEY)
                setData(emptyData)
                setMessage('Local snapshot cleared.')
              }}>Clear local snapshot</button>
            </div>

            <div className="meta-grid">
              <div><span>Connection</span><b>{online ? 'Online' : 'Offline'}</b></div>
              <div><span>Supabase</span><b>{isSupabaseConfigured ? 'Configured' : 'Not configured'}</b></div>
              <div><span>User</span><b>{session?.username}</b></div>
              <div><span>Role</span><b>{session?.role}</b></div>
            </div>
          </section>
        </div>
      )
    }

    return renderDashboard()
  }

  if (!session) return renderLogin()

  return (
    <main className="app-shell">
      <style>{baseStyles}</style>

      <aside className="sidebar">
        <div className="brand">
          <div className="logo-box">TE</div>
          <div>
            <b>Turbo Energy</b>
            <small>Workshop Control</small>
          </div>
        </div>

        <nav>
          {nav.map(([key, label]) => (
            <button key={key} className={tab === key ? 'active' : ''} onClick={() => setTab(key)}>
              {label}
            </button>
          ))}
        </nav>
      </aside>

      <section className="main-area">
        <header className="topbar">
          <div>
            <h1>{nav.find(([key]) => key === tab)?.[1] || 'Dashboard'}</h1>
            <p>Navy/orange workshop system • laptop and phone friendly • online/offline records</p>
          </div>

          <div className="top-actions">
            <Badge value={online ? 'Online' : 'Offline'} />
            <span>{session.username} ({session.role})</span>
            <button onClick={loadAll}>Sync</button>
            <button onClick={logout}>Logout</button>
          </div>
        </header>

        {message && <div className="message">{message}</div>}

        {renderContent()}
      </section>
    </main>
  )
}

const baseStyles = `
  :root {
    --navy: #061a33;
    --navy2: #0b2b52;
    --orange: #ff7a1a;
    --blue: #9fe6ff;
    --line: rgba(159,230,255,.22);
    --card: rgba(8,29,55,.92);
    --text: #f5f9ff;
    --muted: #a9bad1;
    --danger: #ff8f8f;
    --good: #86efac;
    --warn: #ffd28a;
  }

  * {
    box-sizing: border-box;
  }

  body {
    margin: 0;
    background:
      radial-gradient(circle at top left, rgba(255,122,26,.20), transparent 30%),
      linear-gradient(135deg, #020814 0%, #061a33 48%, #0b2b52 100%);
    color: var(--text);
    font-family: Arial, Helvetica, sans-serif;
  }

  button, input, select, textarea {
    font: inherit;
  }

  button {
    border: 1px solid var(--line);
    background: rgba(255,255,255,.07);
    color: var(--text);
    border-radius: 14px;
    padding: 10px 14px;
    cursor: pointer;
    font-weight: 800;
  }

  button:hover {
    border-color: var(--blue);
  }

  .primary {
    background: linear-gradient(135deg, var(--orange), #ffae4d);
    color: #071a33;
    border: 0;
    font-weight: 900;
  }

  .full {
    width: 100%;
  }

  .login-shell {
    min-height: 100vh;
    display: grid;
    place-items: center;
    padding: 24px;
  }

  .login-card {
    width: min(460px, 100%);
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 28px;
    padding: 28px;
    box-shadow: 0 24px 80px rgba(0,0,0,.45);
  }

  .app-shell {
    min-height: 100vh;
    display: grid;
    grid-template-columns: 260px 1fr;
  }

  .sidebar {
    border-right: 1px solid var(--line);
    background: rgba(3,15,30,.92);
    padding: 18px 14px;
    position: sticky;
    top: 0;
    height: 100vh;
    overflow: auto;
  }

  .brand {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 22px;
  }

  .brand small {
    display: block;
    color: var(--muted);
  }

  .logo-box {
    width: 52px;
    height: 52px;
    border-radius: 14px;
    background: #ffffff;
    color: var(--navy);
    display: grid;
    place-items: center;
    font-weight: 900;
    font-size: 22px;
    box-shadow: 0 10px 30px rgba(0,0,0,.28);
  }

  nav {
    display: grid;
    gap: 8px;
  }

  nav button {
    text-align: left;
    border-radius: 14px;
  }

  nav button.active {
    background: rgba(255,122,26,.18);
    border-color: var(--orange);
    box-shadow: inset 4px 0 0 var(--orange);
  }

  .main-area {
    padding: 24px;
    overflow: auto;
  }

  .topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
    background: linear-gradient(135deg, var(--navy), var(--navy2));
    border: 1px solid var(--line);
    border-radius: 24px;
    padding: 22px;
    margin-bottom: 22px;
    box-shadow: 0 18px 55px rgba(0,0,0,.35);
  }

  .topbar h1,
  .page-title h1 {
    margin: 0;
    font-size: 28px;
  }

  .topbar p,
  .page-title p,
  p {
    color: var(--muted);
    margin: 6px 0 0;
    line-height: 1.4;
  }

  .top-actions {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
  }

  .page-title {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    margin: 16px 0;
  }

  .page-actions {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
  }

  .stats {
    display: grid;
    grid-template-columns: repeat(4, minmax(180px, 1fr));
    gap: 14px;
    margin-bottom: 18px;
  }

  .stat-card {
    text-align: left;
    min-height: 130px;
    border-radius: 20px;
    padding: 18px;
    background: linear-gradient(135deg, rgba(16,68,112,.95), rgba(11,43,82,.95));
    position: relative;
    overflow: hidden;
  }

  .stat-card::after {
    content: '';
    position: absolute;
    width: 90px;
    height: 90px;
    right: -18px;
    bottom: -18px;
    border-radius: 30px;
    background: rgba(255,255,255,.08);
  }

  .stat-card span {
    color: #b9d8ff;
    font-size: 13px;
    letter-spacing: .08em;
    text-transform: uppercase;
    font-weight: 900;
  }

  .stat-card b {
    display: block;
    font-size: 38px;
    margin-top: 10px;
  }

  .stat-card small {
    color: #d7eaff;
  }

  .stat-card.purple { background: linear-gradient(135deg, rgba(77,45,91,.95), rgba(32,42,82,.95)); }
  .stat-card.brown { background: linear-gradient(135deg, rgba(86,55,38,.95), rgba(34,45,60,.95)); }
  .stat-card.indigo { background: linear-gradient(135deg, rgba(38,51,111,.95), rgba(23,40,85,.95)); }
  .stat-card.steel { background: linear-gradient(135deg, rgba(45,67,72,.95), rgba(31,53,66,.95)); }
  .stat-card.green { background: linear-gradient(135deg, rgba(14,83,77,.95), rgba(12,56,71,.95)); }

  .grid-2 {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 18px;
    margin-top: 18px;
  }

  .import-layout {
    display: grid;
    grid-template-columns: 330px 1fr;
    gap: 18px;
    align-items: start;
  }

  .card {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 22px;
    padding: 18px;
    box-shadow: 0 14px 45px rgba(0,0,0,.28);
    margin-bottom: 18px;
  }

  .card h2 {
    margin: 0 0 12px;
  }

  .form-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 12px;
    margin-bottom: 14px;
  }

  .field {
    display: grid;
    gap: 6px;
  }

  .field span,
  .upload-stack span {
    color: #aee7ff;
    text-transform: uppercase;
    font-size: 12px;
    letter-spacing: .06em;
    font-weight: 900;
  }

  input,
  select,
  textarea {
    width: 100%;
    border: 1px solid var(--line);
    background: rgba(0,0,0,.28);
    color: var(--text);
    border-radius: 13px;
    padding: 12px;
    outline: none;
  }

  input:focus,
  select:focus,
  textarea:focus {
    border-color: var(--blue);
  }

  .upload-box,
  .upload-stack {
    border: 1px dashed rgba(159,230,255,.45);
    border-radius: 18px;
    padding: 16px;
    background: rgba(255,255,255,.06);
    display: grid;
    gap: 12px;
  }

  .excel-button {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: linear-gradient(135deg, var(--orange), #ffae4d);
    color: #071a33;
    border-radius: 14px;
    padding: 11px 14px;
    font-weight: 900;
    cursor: pointer;
  }

  .excel-button input {
    display: none;
  }

  .message {
    background: rgba(159,230,255,.12);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 13px 14px;
    margin-bottom: 16px;
    font-weight: 900;
  }

  .badge {
    display: inline-flex;
    align-items: center;
    border-radius: 999px;
    padding: 7px 11px;
    font-size: 12px;
    font-weight: 900;
    border: 1px solid var(--line);
  }

  .badge.good {
    background: rgba(34,197,94,.18);
    color: var(--good);
    border-color: rgba(34,197,94,.35);
  }

  .badge.danger {
    background: rgba(239,68,68,.18);
    color: var(--danger);
    border-color: rgba(239,68,68,.35);
  }

  .badge.warn {
    background: rgba(255,122,26,.18);
    color: var(--warn);
    border-color: rgba(255,122,26,.35);
  }

  .badge.neutral {
    background: rgba(159,230,255,.12);
    color: var(--blue);
  }

  .badge-row,
  .chip-row,
  .button-row {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
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

  .meta-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
    margin-top: 14px;
  }

  .meta-grid div {
    border: 1px solid var(--line);
    background: rgba(255,255,255,.06);
    border-radius: 16px;
    padding: 13px;
  }

  .meta-grid span {
    display: block;
    color: var(--muted);
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: .08em;
    font-weight: 900;
  }

  .meta-grid b {
    display: block;
    margin-top: 5px;
    word-break: break-word;
  }

  .table-wrap {
    overflow: auto;
    border: 1px solid var(--line);
    border-radius: 16px;
  }

  table {
    width: 100%;
    border-collapse: collapse;
    min-width: 760px;
  }

  th,
  td {
    padding: 11px 12px;
    border-bottom: 1px solid rgba(159,230,255,.12);
    text-align: left;
    font-size: 13px;
    vertical-align: top;
  }

  th {
    color: var(--blue);
    background: rgba(0,0,0,.24);
    text-transform: uppercase;
    letter-spacing: .06em;
  }

  .empty {
    padding: 18px;
    border: 1px dashed var(--line);
    border-radius: 16px;
    color: var(--muted);
  }

  .import-list {
    display: grid;
    gap: 9px;
    max-height: 70vh;
    overflow: auto;
    padding-right: 5px;
  }

  .import-btn {
    text-align: left;
    display: grid;
    gap: 4px;
  }

  .import-btn.active {
    border-color: var(--orange);
    box-shadow: inset 4px 0 0 var(--orange);
    background: rgba(255,122,26,.16);
  }

  .import-btn small {
    color: var(--muted);
  }

  @media (max-width: 1000px) {
    .app-shell {
      grid-template-columns: 1fr;
    }

    .sidebar {
      position: relative;
      height: auto;
    }

    nav {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .topbar,
    .page-title {
      align-items: flex-start;
      flex-direction: column;
    }

    .stats,
    .grid-2,
    .form-grid,
    .import-layout,
    .meta-grid {
      grid-template-columns: 1fr;
    }

    .main-area {
      padding: 12px;
    }
  }
`
