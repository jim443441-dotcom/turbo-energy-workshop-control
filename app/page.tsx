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
  { table: 'services', label: 'Service Schedule', description: 'Turbo Mining service sheet, service interval, hours till due and status.' },
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
const sections = ['WORKSHOPS', 'Workshop', 'Mechanical', 'Auto Electrical', 'Tyre Bay', 'Boiler Shop', 'Panel Beating', 'Stores', 'Wash Bay', 'Operations']
const leaveTypes = ['Annual Leave', 'Sick Leave', 'Off Duty', 'Training', 'Unpaid Leave', 'AWOL']
const orderStages = ['Requested', 'Awaiting Funding', 'Funded', 'Awaiting Delivery', 'Delivered', 'Cancelled']

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
  'Diesel Mechanic',
  'Workshop Clerk',
  'Workshop Pit Foreman',
  'Working Chargehand Light Vehicles',
  'DPF',
  'Vacant'
]

const statuses = [
  'Available',
  'Open',
  'In progress',
  'Requested',
  'Awaiting Funding',
  'Funded',
  'Awaiting Delivery',
  'Delivered',
  'Completed',
  'Approved',
  'Active',
  'Inactive',
  'Vacant',
  'On Leave',
  'Off Duty',
  'Scheduled',
  'Upcoming',
  'Due Soon',
  'Due',
  'Overdue'
]

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
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '')
}

function slug(value: any) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
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

function addDaysToToday(days: number) {
  if (!Number.isFinite(days)) return ''
  const d = new Date()
  d.setDate(d.getDate() + Math.ceil(days))
  return d.toISOString().slice(0, 10)
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
  const t = new Date(today()).getTime()
  const s = new Date(row.start_date).getTime()
  const e = new Date(row.end_date).getTime()
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
    'fleet', 'fleet no', 'machine', 'unit', 'equipment', 'description',
    'firstname', 'surname', 'employee', 'code', 'occupation', 'current job title', 'department', 'shift',
    'dates', 'total', 'leave type',
    'odo reading', 'last service', 'service interval', 'next service type', 'hours till service due', 'percent done',
    'part no', 'qty', 'serial', 'status', 'requested by', 'supervisor', 'foreman'
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

    const headerNames = (matrix[headerIndex] || []).map((h, i) => {
      const clean = String(h || '').trim()
      return clean || `Column ${i + 1}`
    })
    headerNames.forEach((h) => headings.add(h))

    matrix.slice(headerIndex + 1).forEach((line, offset) => {
      const obj: Row = { __sheet: sheetName, __row: headerIndex + offset + 2, __cells: line }
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
  const banned = new Set(['FLEET', 'MACHINE', 'EQUIPMENT', 'TYPE', 'DESCRIPTION', 'DATE', 'TOTAL', 'STATUS', 'DEPARTMENT', 'WORKSHOP', 'SERVICE', 'INTERVAL'])

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
  if (text.includes('LDV') || text.includes('AFX') || text.includes('AG') || text.includes('S-VEHICLE')) return 'LDV'
  return 'Machine'
}


const departmentHeaderNames = [
  'Mining',
  'Logistics',
  'Plant',
  'Workshop',
  'Engineering & Civils',
  'Stores & Procurement',
  'Admin'
]

function cleanFleetNo(value: any) {
  return String(value || '').trim().replace(/\s+/g, ' ').toUpperCase()
}

function isLikelyFleetNo(value: any) {
  const text = cleanFleetNo(value)
  if (!text) return false
  if (departmentHeaderNames.some((dept) => normal(dept) === normal(text))) return false
  if (['FLEET', 'MACHINE', 'DEPARTMENT', 'TOTAL', 'TYPE', 'STATUS', 'DATE'].includes(text)) return false
  if (/^\d+$/.test(text)) return false
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text)) return false
  return /^[A-Z]{1,6}\s?-?\d{1,6}[A-Z]?$/.test(text)
}

function departmentAllocationHeaderIndex(allCells: any[][]) {
  return allCells.findIndex((line) => {
    const matched = line.filter((cell) => departmentHeaderNames.some((dept) => normal(dept) === normal(cell))).length
    return matched >= 2
  })
}

function parseDepartmentAllocation(allCells: any[][]) {
  const headerIndex = departmentAllocationHeaderIndex(allCells)
  if (headerIndex < 0) return []

  const headers = allCells[headerIndex].map((cell) => String(cell || '').trim())
  const validColumns = headers
    .map((department, index) => ({ department, index }))
    .filter((item) => departmentHeaderNames.some((dept) => normal(dept) === normal(item.department)))

  const map = new Map<string, Row>()

  allCells.slice(headerIndex + 1).forEach((line) => {
    validColumns.forEach(({ department, index }) => {
      const fleetNo = cleanFleetNo(line[index])
      if (!isLikelyFleetNo(fleetNo)) return

      const key = normal(fleetNo)
      if (map.has(key)) return

      map.set(key, cleanRow({
        id: fleetNo,
        fleet_no: fleetNo,
        machine_type: inferMachineType(fleetNo),
        make_model: '',
        reg_no: fleetNo.startsWith('A') ? fleetNo : '',
        department,
        location: department,
        hours: 0,
        mileage: 0,
        status: 'Available',
        allocation_source: 'department_allocation',
        active_fleet_register: true,
        notes: 'Imported from department allocation Excel sheet'
      }))
    })
  })

  return Array.from(map.values())
}


function activeFleetRows(machines: Row[]) {
  const allocationRows = machines.filter((machine) => String(machine.allocation_source || '') === 'department_allocation')
  return allocationRows.length ? allocationRows : machines
}

function departmentSummaryRows(machines: Row[]) {
  const map = new Map<string, Row>()

  machines.forEach((machine) => {
    const department = String(machine.department || 'Unallocated').trim() || 'Unallocated'
    const existing = map.get(department) || { id: department, department, total: 0, available: 0, breakdown: 0, fleet_list: '' }
    existing.total += 1
    if (String(machine.status || '').toLowerCase().includes('available')) existing.available += 1
    if (String(machine.status || '').toLowerCase().includes('break')) existing.breakdown += 1
    existing.fleet_list = [existing.fleet_list, machine.fleet_no].filter(Boolean).join(', ')
    map.set(department, existing)
  })

  return Array.from(map.values()).sort((a, b) => String(a.department).localeCompare(String(b.department)))
}

function serviceStatus(hoursTillValue: any, percentDoneValue: any, hasServiceNumbers = true) {
  const hoursTill = Number(hoursTillValue)
  const percentDone = Number(percentDoneValue)

  if (!hasServiceNumbers || !Number.isFinite(hoursTill)) return 'Scheduled'
  if (hoursTill <= 0) return 'Due'
  if (hoursTill <= 50 || percentDone >= 0.98) return 'Due Soon'
  if (hoursTill <= 250 || percentDone >= 0.90) return 'Upcoming'
  return 'Scheduled'
}

function parsePersonnel(row: Row) {
  const name = fullName(row)
  const employeeNo = String(getCell(row, ['Code', 'Employee Code', 'Employee No', 'Clock No', 'Emp No']) || '').trim()
  const role = String(getCell(row, ['Occupation', 'CURRENT JOB TITLE', 'Current Job Title', 'Job Title', 'Trade', 'Role']) || 'Workshop Employee').trim()
  const section = String(getCell(row, ['Department', 'Dept', 'Section']) || 'Workshop').trim()
  const status = String(getCell(row, ['Status', 'Employment Status']) || 'Active').trim()
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
    employment_status: status,
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
  const [rangeStart, rangeEnd] = dateRange(getCell(row, ['DATES', 'Dates', 'Date Range', 'Leave Dates']))
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
  const rawFleet = getCell(row, ['Fleet', 'FLEET No', 'Fleet No', 'Fleet Number', 'Machine', 'Machine No', 'Unit', 'Equipment', 'Plant No', 'Asset', 'Registration', 'Reg No'])
  const fleetNo = cleanFleetNo(rawFleet) || fleetFromCells(cells)
  if (!fleetNo || !isLikelyFleetNo(fleetNo)) return null

  const department = String(getCell(row, ['Department', 'Dept', 'Section', 'Allocation', 'Allocated Department']) || 'Workshop').trim()

  return cleanRow({
    id: fleetNo,
    fleet_no: fleetNo,
    machine_type: String(getCell(row, ['Machine Type', 'Equipment Type', 'Type']) || '').trim() || inferMachineType(fleetNo, cellText),
    make_model: String(getCell(row, ['Make Model', 'Make', 'Model', 'Description', 'DESCRIPTION']) || '').trim(),
    reg_no: String(getCell(row, ['Reg', 'Registration', 'Reg No']) || (fleetNo.startsWith('A') ? fleetNo : '')).trim(),
    department,
    location: String(getCell(row, ['Location', 'Site']) || department).trim(),
    hours: num(getCell(row, ['Hours', 'Hrs', 'HM', 'Hour Meter', 'TOTAL HRS TO DATE', 'ODO READING CURRENT close']), 0),
    mileage: num(getCell(row, ['Mileage', 'KM', 'Odometer']), 0),
    status: String(getCell(row, ['Status', 'State']) || 'Available').trim(),
    notes: String(getCell(row, ['Notes', 'Remarks', 'Coments', 'Comments']) || '').trim()
  })
}

function parseService(row: Row) {
  const cells = Array.isArray(row.__cells) ? row.__cells : Object.values(row)
  const cellText = cells.map((c) => String(c || '')).join(' ')
  const fleetNo = String(getCell(row, ['FLEET No', 'Fleet', 'Fleet No', 'Fleet Number', 'Machine', 'Unit']) || '').trim() || fleetFromCells(cells)
  if (!fleetNo) return null
  if (normal(fleetNo).includes('fleetno')) return null

  const description = String(getCell(row, ['DESCRIPTION', 'Description', 'Machine Description']) || '').trim()
  const currentReading = num(getCell(row, ['CURRENT ODO READING (HOURS)', 'ODO READING CURRENT (close)', 'ODO READING         CURRENT (close)', 'Current Reading', 'Current Hours']), 0)
  const openingReading = num(getCell(row, ['ODO READING CURRENT (opening)', 'ODO READING         CURRENT (opening)', 'Opening Reading']), 0)
  const dailyTotal = num(getCell(row, ['ODO READING Total', 'ODO READING         Total', 'Daily Hours', 'Hours Today']), 0)
  const lastService = num(getCell(row, ['ODO READING AT LAST SERVICE', 'Last Service Hours', 'Last ODO Reading']), 0)
  const serviceInterval = num(getCell(row, ['SERVICE INTERVAL', 'Service Interval']), 0)
  const nextServiceType = String(getCell(row, ['NEXT SERVICE TYPE', 'Next Service Type', 'Service Type']) || 'Service').trim()
  const nextServiceHours = num(getCell(row, ['ODO READING AT SERVICE', 'ODO READING      AT SERVICE', 'Next Service Hours', 'Service Due Hours']), 0)
  const percentDoneRaw = num(getCell(row, ['PERCENT DONE', 'Percent Done']), 0)
  const percentDone = percentDoneRaw > 1 ? percentDoneRaw / 100 : percentDoneRaw
  const explicitHoursTill = getCell(row, ['HOURS TILL SERVICE DUE', 'Hours Till Service Due', 'Hours To Service'])
  const calculatedHoursTill = nextServiceHours && currentReading ? nextServiceHours - currentReading : undefined
  const hoursTill = num(explicitHoursTill, calculatedHoursTill ?? 999999)
  const estimateDays = num(getCell(row, ['ESTIMATE DAYS TILL SERVICE', 'Estimate Days Till Service', 'Days Till Service']), 0)
  const totalHours = num(getCell(row, ['TOTAL HRS TO DATE', 'Total Hrs To Date', 'Total Hours']), currentReading)
  const comments = String(getCell(row, ['COMMENTS', 'Comments', 'Coments', 'Coments/ Filter change']) || '').trim()
  const hasServiceNumbers = Boolean(fleetNo && (currentReading || nextServiceHours || explicitHoursTill !== ''))
  const status = serviceStatus(hoursTill, percentDone, hasServiceNumbers)

  return cleanRow({
    id: `${fleetNo}-${nextServiceHours || lastService || currentReading}-${nextServiceType || 'service'}`,
    fleet_no: fleetNo,
    machine_type: inferMachineType(fleetNo, `${description} ${cellText}`),
    description,
    service_type: nextServiceType ? `Service ${nextServiceType}` : 'Service',
    next_service_type: nextServiceType,
    current_hours: currentReading,
    opening_hours: openingReading,
    daily_hours: dailyTotal,
    last_service_hours: lastService,
    service_interval: serviceInterval,
    next_service_hours: nextServiceHours,
    scheduled_hours: nextServiceHours || lastService + serviceInterval,
    percent_done: Math.round(percentDone * 1000) / 10,
    hours_till_service_due: Math.round(hoursTill * 10) / 10,
    estimate_days_till_service: Math.round(estimateDays * 10) / 10,
    total_hours_to_date: totalHours,
    due_date: estimateDays ? addDaysToToday(estimateDays) : status === 'Due' ? today() : '',
    supervisor: '',
    technician: '',
    job_card_no: '',
    status,
    notes: comments
  })
}

function parseGeneric(table: TableName, row: Row) {
  const fleetNo = String(getCell(row, ['Fleet', 'Fleet No', 'Machine', 'Unit']) || '').trim()
  const partNo = String(getCell(row, ['Part No', 'Part Number', 'Stock Code', 'Item No']) || '').trim()
  const description = String(getCell(row, ['Description', 'Item', 'Name', 'Tool Name', 'Tool']) || '').trim()
  const serial = String(getCell(row, ['Serial', 'Serial No', 'Serial Number']) || '').trim()
  const requestedBy = String(getCell(row, ['Requested By', 'Requester', 'Foreman']) || '').trim()

  if (table === 'services') return parseService(row)

  if (table === 'breakdowns') {
    return cleanRow({
      id: uid(), fleet_no: fleetNo,
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
      id: uid(), fleet_no: fleetNo,
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
      id: partNo || slug(description) || uid(), part_no: partNo, description,
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
      id: uid(), request_date: dateOnly(getCell(row, ['Request Date', 'Date'])) || today(),
      machine_fleet_no: fleetNo, requested_by: requestedBy, part_no: partNo, description,
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
      id: serial || uid(), serial_no: serial, fleet_no: fleetNo,
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
      id: serial || uid(), serial_no: serial, fleet_no: fleetNo,
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
    id: uid(), fleet_no: fleetNo, machine_fleet_no: fleetNo, part_no: partNo, description, serial_no: serial,
    requested_by: requestedBy,
    qty: num(getCell(row, ['Qty', 'Quantity']), 1),
    stock_qty: num(getCell(row, ['Stock Qty', 'Stock', 'Balance', 'Qty']), 0),
    status: String(getCell(row, ['Status', 'Stage']) || 'Imported').trim(),
    notes: String(getCell(row, ['Notes', 'Remarks']) || '').trim()
  })
}

function parseRows(table: TableName, rawRows: Row[], allCells: any[][]) {
  let rows: Row[] = []
  if (table === 'personnel') rows = rawRows.map(parsePersonnel).filter(Boolean) as Row[]
  else if (table === 'leave_records') rows = rawRows.map(parseLeave).filter(Boolean) as Row[]
  else if (table === 'fleet_machines') {
    const allocationRows = parseDepartmentAllocation(allCells)
    rows = allocationRows.length ? allocationRows : rawRows.map(parseFleet).filter(Boolean) as Row[]

    if (!rows.length) {
      const map = new Map<string, Row>()
      allCells.forEach((line) => {
        const fleetNo = fleetFromCells(line)
        if (!fleetNo) return
        const text = line.map((x) => String(x || '')).join(' ')
        map.set(fleetNo, { id: fleetNo, fleet_no: fleetNo, machine_type: inferMachineType(fleetNo, text), department: 'Workshop', location: 'Workshop', status: 'Available', notes: text.slice(0, 160) })
      })
      rows = Array.from(map.values())
    }
  } else if (table === 'services') rows = rawRows.map(parseService).filter(Boolean) as Row[]
  else rows = rawRows.map((row) => parseGeneric(table, row)).filter(Boolean) as Row[]

  const map = new Map<string, Row>()
  rows.forEach((row) => {
    const key = table === 'personnel'
      ? String(row.employee_no || row.name || row.id)
      : table === 'fleet_machines'
        ? String(row.fleet_no || row.id)
        : table === 'services'
          ? String(row.id || `${row.fleet_no}-${row.next_service_hours || row.scheduled_hours || row.service_type}`)
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
    // browser storage can fill up when photos are saved
  }
}

function addToOfflineQueue(table: TableName, rows: Row[]) {
  try {
    const raw = localStorage.getItem(QUEUE_KEY)
    const queue = raw ? JSON.parse(raw) : []
    rows.forEach((row) => queue.push({ id: uid(), table, type: 'upsert', payload: row, createdAt: nowIso() }))
    localStorage.setItem(QUEUE_KEY, JSON.stringify(queue))
  } catch {
    // ignore queue failure
  }
}

function mergeRows(existing: Row[], incoming: Row[]) {
  const map = new Map<string, Row>()
  existing.forEach((row) => {
    const key = String(row.id || row.fleet_no || row.employee_no || row.name || JSON.stringify(row))
    map.set(key, row)
  })
  incoming.forEach((row) => {
    const key = String(row.id || row.fleet_no || row.employee_no || row.name || JSON.stringify(row))
    map.set(key, { ...map.get(key), ...row })
  })
  return Array.from(map.values())
}


function latestServiceRows(rows: Row[]) {
  const batches = rows
    .map((row) => String(row.import_batch || '').trim())
    .filter(Boolean)
    .sort()

  if (!batches.length) return rows

  const latest = batches[batches.length - 1]
  return rows.filter((row) => String(row.import_batch || '') === latest)
}

function sortedByHoursLeft(rows: Row[]) {
  return [...rows].sort((a, b) => num(a.hours_till_service_due, 999999) - num(b.hours_till_service_due, 999999))
}

function statusClass(value: any) {
  const text = String(value || '').toLowerCase()
  if (text.includes('overdue') || text === 'due' || text.includes('open') || text.includes('break') || text.includes('awaiting') || text.includes('low') || text.includes('critical')) return 'danger'
  if (text.includes('due soon') || text.includes('upcoming') || text.includes('request') || text.includes('fund') || text.includes('progress') || text.includes('normal') || text.includes('vacant')) return 'warn'
  if (text.includes('scheduled') || text.includes('active') || text.includes('available') || text.includes('complete') || text.includes('delivered') || text.includes('approved') || text.includes('stock')) return 'good'
  return 'neutral'
}

function Badge({ value }: { value: any }) {
  return <span className={`badge ${statusClass(value)}`}>{String(value || 'Not set')}</span>
}

function Field({ label, children }: { label: string; children: any }) {
  return <label className="field"><span>{label}</span>{children}</label>
}

function Input({ value, onChange, placeholder = '', type = 'text' }: { value: any; onChange: (value: any) => void; placeholder?: string; type?: string }) {
  return <input type={type} value={value ?? ''} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} />
}

function Select({ value, onChange, options }: { value: any; onChange: (value: any) => void; options: string[] }) {
  return <select value={value ?? ''} onChange={(e) => onChange(e.target.value)}>{options.map((option) => <option key={option}>{option}</option>)}</select>
}

function StatCard({ title, value, detail, tone = 'blue', onClick, active = false }: { title: string; value: any; detail?: string; tone?: string; onClick?: () => void; active?: boolean }) {
  return <button type="button" className={`stat-card ${tone} ${active ? 'selected-stat' : ''}`} onClick={onClick}><span>{title}</span><b>{value}</b>{detail && <small>{detail}</small>}</button>
}

function MiniTable({ rows, columns, empty = 'No records yet.', onRowClick, activeId }: { rows: Row[]; columns: { key: string; label: string; render?: (row: Row) => any }[]; empty?: string; onRowClick?: (row: Row) => void; activeId?: string }) {
  if (!rows.length) return <div className="empty">{empty}</div>
  return (
    <div className="table-wrap">
      <table>
        <thead><tr>{columns.map((col) => <th key={col.key}>{col.label}</th>)}</tr></thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id || JSON.stringify(row)} className={`${onRowClick ? 'click-row' : ''} ${activeId && String(row.id) === activeId ? 'active-row' : ''}`} onClick={() => onRowClick?.(row)}>
              {columns.map((col) => <td key={col.key}>{col.render ? col.render(row) : String(row[col.key] ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function PageTitle({ title, subtitle, right }: { title: string; subtitle?: string; right?: any }) {
  return <div className="page-title"><div><h1>{title}</h1>{subtitle && <p>{subtitle}</p>}</div>{right && <div className="page-actions">{right}</div>}</div>
}

async function filesToPhotoRows(files: FileList | null, linkedType: string, fleetNo = '', caption = '') {
  if (!files || !files.length) return []
  const selected = Array.from(files).filter((file) => file.type.startsWith('image/'))
  const rows = await Promise.all(selected.map((file) => new Promise<Row>((resolve) => {
    const reader = new FileReader()
    reader.onload = () => {
      resolve({
        id: uid(),
        fleet_no: fleetNo,
        linked_type: linkedType,
        caption: caption || file.name,
        notes: `${linkedType} photo uploaded ${new Date().toLocaleString()}`,
        file_name: file.name,
        image_data: String(reader.result || ''),
        created_at: nowIso()
      })
    }
    reader.readAsDataURL(file)
  })))
  return rows
}

export default function Home() {
  const [session, setSession] = useState<Session | null>(null)
  const [loginName, setLoginName] = useState('admin')
  const [loginPass, setLoginPass] = useState('admin123')
  const [tab, setTab] = useState('dashboard')
  const [data, setData] = useState<AppData>(emptyData)
  const [message, setMessage] = useState('')
  const [online, setOnline] = useState(true)

  const [selectedImport, setSelectedImport] = useState<TableName>('services')
  const [importSearch, setImportSearch] = useState('')
  const [importStatus, setImportStatus] = useState('Ready')
  const [importFile, setImportFile] = useState('')
  const [importHeadings, setImportHeadings] = useState<string[]>([])
  const [importPreview, setImportPreview] = useState<Row[]>([])
  const [busyImport, setBusyImport] = useState(false)

  const [machineForm, setMachineForm] = useState<Row>({ fleet_no: '', machine_type: '', make_model: '', department: 'Workshop', location: '', hours: 0, mileage: 0, status: 'Available' })
  const [personForm, setPersonForm] = useState<Row>({ id: '', employee_no: '', name: '', shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', employment_status: 'Active', leave_balance_days: 30, leave_taken_days: 0, phone: '', notes: '' })
  const [editingPersonId, setEditingPersonId] = useState('')
  const [leaveForm, setLeaveForm] = useState<Row>({ id: '', employee_no: '', person_name: '', leave_type: 'Annual Leave', start_date: today(), end_date: today(), status: 'Approved' })
  const [editingLeaveId, setEditingLeaveId] = useState('')
  const [leaveView, setLeaveView] = useState<'current' | 'upcoming' | 'past'>('current')
  const [quickForm, setQuickForm] = useState<Row>({})
  const [moduleForms, setModuleForms] = useState<Record<string, Row>>({})
  const [machineLookup, setMachineLookup] = useState('')
  const [selectedMachine, setSelectedMachine] = useState('')

  const selectedImportTarget = IMPORT_TARGETS.find((target) => target.table === selectedImport) || IMPORT_TARGETS[0]
  const filteredTargets = useMemo(() => {
    const q = importSearch.toLowerCase()
    return IMPORT_TARGETS.filter((target) => `${target.label} ${target.description} ${target.table}`.toLowerCase().includes(q))
  }, [importSearch])

  const dashboard = useMemo(() => {
    const activeFleet = activeFleetRows(data.fleet_machines)
    const totalFleet = activeFleet.length
    const breakdowns = data.breakdowns.filter((row) => !String(row.status || '').toLowerCase().includes('complete'))
    const repairs = data.repairs.filter((row) => !String(row.status || '').toLowerCase().includes('complete'))
    const activeServices = latestServiceRows(data.services)
    const servicesDue = activeServices.filter((row) => ['due', 'overdue'].includes(String(row.status || '').toLowerCase()))
    const serviceDueNow = servicesDue
    const currentLeave = data.leave_records.filter(isCurrentLeave)
    const foremen = data.personnel.filter((row) => String(row.role || '').toLowerCase().includes('foreman') && !String(row.employment_status || '').toLowerCase().includes('inactive'))
    const lowSpares = data.spares.filter((row) => num(row.stock_qty) <= num(row.min_qty, 1))
    const hoseRequests = data.hose_requests.filter((row) => !String(row.status || '').toLowerCase().includes('complete'))
    const toolOrders = data.tool_orders.filter((row) => !String(row.status || '').toLowerCase().includes('delivered'))
    const departmentSummary = departmentSummaryRows(activeFleet)
    return { totalFleet, breakdowns, repairs, servicesDue, serviceDueNow, currentLeave, foremen, lowSpares, hoseRequests, toolOrders, departmentSummary }
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

  useEffect(() => { saveSnapshot(data) }, [data])

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
        if (rows && Array.isArray(rows)) next[table] = mergeRows(next[table] || [], rows as Row[])
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

  async function savePhotoRows(rows: Row[]) {
    if (!rows.length) return
    setData((prev) => {
      const next = { ...prev, photo_logs: mergeRows(prev.photo_logs || [], rows) }
      saveSnapshot(next)
      return next
    })

    if (!supabase || !isSupabaseConfigured || !navigator.onLine) {
      setMessage(`Saved ${rows.length} photo(s) locally.`)
      return
    }

    const safeRows = rows.map((row) => cleanRow({
      id: row.id,
      fleet_no: row.fleet_no || '',
      linked_type: row.linked_type || 'Photo',
      caption: row.caption || '',
      notes: row.notes || ''
    }))

    try {
      await supabase.from('photo_logs' as any).upsert(safeRows, { onConflict: 'id' })
    } catch {
      // photo image data stays local so the app does not break when storage is not set up
    }
    setMessage(`Saved ${rows.length} photo(s).`)
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

  async function handlePhotoUpload(files: FileList | null, linkedType: string, fleetNo = '', caption = '') {
    const rows = await filesToPhotoRows(files, linkedType, fleetNo, caption)
    await savePhotoRows(rows)
  }

  function personNameFromLeave(row: Row) {
    const code = String(row.employee_no || '').trim()
    const rawName = String(row.person_name || '').trim()
    const matchByCode = data.personnel.find((p) => String(p.employee_no || '').trim() === code || String(p.employee_no || '').trim() === rawName)
    const matchByName = data.personnel.find((p) => normal(p.name) === normal(rawName))
    return matchByCode?.name || matchByName?.name || rawName || code || 'Unknown'
  }

  function clickLeaveCard(view: 'current' | 'upcoming' | 'past') {
    setLeaveView(view)
    setTimeout(() => document.getElementById('leave-list')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
  }

  function editPerson(row: Row) {
    setEditingPersonId(String(row.id || ''))
    setPersonForm({
      id: row.id || row.employee_no || slug(row.name) || uid(),
      employee_no: row.employee_no || '',
      name: row.name || '',
      shift: row.shift || 'Shift 1',
      section: row.section || 'Workshop',
      role: row.role || 'Diesel Plant Fitter',
      employment_status: row.employment_status || row.status || 'Active',
      leave_balance_days: row.leave_balance_days ?? 30,
      leave_taken_days: row.leave_taken_days ?? 0,
      phone: row.phone || '',
      notes: row.notes || ''
    })
    setTimeout(() => document.getElementById('employee-editor')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
  }

  function clearPersonForm() {
    setEditingPersonId('')
    setPersonForm({ id: '', employee_no: '', name: '', shift: 'Shift 1', section: 'Workshop', role: 'Diesel Plant Fitter', employment_status: 'Active', leave_balance_days: 30, leave_taken_days: 0, phone: '', notes: '' })
  }

  async function savePersonForm() {
    const id = personForm.id || personForm.employee_no || slug(personForm.name) || uid()
    await saveRows('personnel', [{
      ...personForm,
      id,
      employee_no: personForm.employee_no || '',
      name: personForm.name || 'Unnamed employee',
      shift: personForm.shift || 'Shift 1',
      section: personForm.section || 'Workshop',
      role: personForm.role || 'Workshop Employee',
      employment_status: personForm.employment_status || 'Active',
      leave_balance_days: num(personForm.leave_balance_days, 30),
      leave_taken_days: num(personForm.leave_taken_days, 0)
    }])
    setEditingPersonId(id)
  }

  function editLeave(row: Row) {
    setEditingLeaveId(String(row.id || ''))
    setLeaveForm({
      id: row.id || '',
      employee_no: row.employee_no || '',
      person_name: personNameFromLeave(row),
      leave_type: row.leave_type || 'Annual Leave',
      start_date: row.start_date || today(),
      end_date: row.end_date || today(),
      days: row.days || daysInclusive(row.start_date, row.end_date),
      reason: row.reason || '',
      status: row.status || 'Approved'
    })
    setTimeout(() => document.getElementById('leave-editor')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
  }

  async function saveLeaveForm() {
    const id = leaveForm.id || `${leaveForm.employee_no || slug(leaveForm.person_name)}-${leaveForm.start_date}-${leaveForm.end_date}-${slug(leaveForm.leave_type)}`
    await saveRows('leave_records', [{
      ...leaveForm,
      id,
      person_name: leaveForm.person_name || '',
      leave_type: leaveForm.leave_type || 'Annual Leave',
      start_date: leaveForm.start_date || today(),
      end_date: leaveForm.end_date || today(),
      days: daysInclusive(leaveForm.start_date, leaveForm.end_date),
      status: leaveForm.status || 'Approved'
    }])
    setEditingLeaveId(id)
  }

  function PhotoUploadBox({ linkedType, fleetNo = '' }: { linkedType: string; fleetNo?: string }) {
    return (
      <div className="photo-upload-mini">
        <span>Upload photos for {linkedType}</span>
        <input type="file" accept="image/*" multiple onChange={(e) => { handlePhotoUpload(e.target.files, linkedType, fleetNo); e.currentTarget.value = '' }} />
      </div>
    )
  }

  function PhotoGallery({ linkedType, fleetNo = '', limit = 8 }: { linkedType?: string; fleetNo?: string; limit?: number }) {
    const photos = data.photo_logs
      .filter((p) => (!linkedType || String(p.linked_type || '') === linkedType) && (!fleetNo || String(p.fleet_no || '') === fleetNo))
      .slice(-limit)
      .reverse()

    if (!photos.length) return <div className="empty">No photos uploaded yet.</div>
    return (
      <div className="photo-grid">
        {photos.map((photo) => (
          <div className="photo-card" key={photo.id}>
            {photo.image_data ? <img src={photo.image_data} alt={photo.caption || 'Workshop photo'} /> : <div className="photo-placeholder">Photo</div>}
            <b>{photo.caption || 'Photo'}</b>
            <small>{photo.fleet_no || photo.linked_type || ''}</small>
          </div>
        ))}
      </div>
    )
  }

  function renderLogin() {
    return (
      <main className="login-shell">
        <style>{baseStyles}</style>
        <div className="login-card">
          <div className="logo-box">TE</div>
          <h1>Turbo Energy Workshop Control</h1>
          <p>Workshop fleet repairs, services, spares, tyres, batteries, personnel and operations tracking.</p>
          <Field label="Username"><Input value={loginName} onChange={setLoginName} /></Field>
          <Field label="Password"><Input value={loginPass} onChange={setLoginPass} type="password" /></Field>
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
          <StatCard title="Services Due" value={dashboard.serviceDueNow.length} detail="Only due / overdue" tone="indigo" onClick={() => setTab('services')} />
          <StatCard title="People Off" value={dashboard.currentLeave.length} detail="Leave/sick/off duty" tone="steel" onClick={() => setTab('personnel')} />
          <StatCard title="Foremen On Duty" value={dashboard.foremen.length} detail="Loaded foremen" tone="green" onClick={() => setTab('personnel')} />
          <StatCard title="Hose Requests" value={dashboard.hoseRequests.length} detail="Pending hoses/fittings" onClick={() => setTab('hoses')} />
          <StatCard title="Tool Orders" value={dashboard.toolOrders.length} detail="Requested tools" tone="brown" onClick={() => setTab('tools')} />
        </div>

        <section className="card">
          <PageTitle title="Fleet by department" subtitle="Department allocation is read from the uploaded department allocation Excel sheet. Click Fleet Register to see the full machine list." right={<button onClick={() => setTab('fleet')}>Fleet Register</button>} />
          <MiniTable
            rows={dashboard.departmentSummary}
            empty="No department allocation loaded yet. Upload the department allocation Excel under Fleet Register."
            columns={[
              { key: 'department', label: 'Department' },
              { key: 'total', label: 'Machines' },
              { key: 'available', label: 'Available' },
              { key: 'breakdown', label: 'Breakdown' },
              { key: 'fleet_list', label: 'Fleet numbers' }
            ]}
          />
        </section>

        <div className="grid-2">
          <section className="card">
            <PageTitle title="Machines on breakdown / under repair" right={<button onClick={() => setTab('breakdowns')}>Open breakdowns</button>} />
            <MiniTable rows={dashboard.breakdowns.slice(0, 8)} empty="No open breakdowns." columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'fault', label: 'Fault' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
          </section>
          <section className="card">
            <PageTitle title="Machines due for service" right={<button onClick={() => setTab('services')}>Open services</button>} />
            <MiniTable rows={dashboard.servicesDue.slice(0, 8)} empty="No due services." columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'hours_till_service_due', label: 'Hours Left' }, { key: 'due_date', label: 'Due Date' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
          </section>
          <section className="card">
            <PageTitle title="Foremen on duty" right={<button onClick={() => setTab('personnel')}>Personnel</button>} />
            <MiniTable rows={dashboard.foremen.slice(0, 8)} empty="No foremen loaded yet." columns={[{ key: 'name', label: 'Name' }, { key: 'shift', label: 'Shift' }, { key: 'section', label: 'Section' }, { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }]} />
          </section>
          <section className="card">
            <PageTitle title="Workshop alerts" right={<button onClick={() => setTab('reports')}>Reports</button>} />
            <div className="badge-row"><Badge value={`Low spares: ${dashboard.lowSpares.length}`} /><Badge value={`Services due/soon: ${dashboard.servicesDue.length}`} /><Badge value={`Tyres due: ${data.tyre_measurements.length}`} /><Badge value={`Projects: ${data.fabrication_projects.length}`} /></div>
            <PhotoUploadBox linkedType="Dashboard" />
          </section>
        </div>
      </div>
    )
  }

  function renderImporter() {
    return (
      <div>
        <PageTitle title="Excel Importer" subtitle="Upload Excel files here. Service sheets now calculate due status from Hours Till Service Due instead of marking everything due." right={<Badge value={isSupabaseConfigured ? 'Supabase connected' : 'Local mode'} />} />
        <div className="import-layout">
          <aside className="card">
            <h2>Choose register</h2>
            <p>Select what the Excel file is for.</p>
            <input placeholder="Search register..." value={importSearch} onChange={(e) => setImportSearch(e.target.value)} />
            <div className="import-list">
              {filteredTargets.map((target) => (
                <button key={target.table} className={selectedImport === target.table ? 'active import-btn' : 'import-btn'} onClick={() => { setSelectedImport(target.table); setImportStatus('Ready'); setImportFile(''); setImportHeadings([]); setImportPreview([]) }}>
                  <b>{target.label}</b><small>{target.description}</small>
                </button>
              ))}
            </div>
          </aside>
          <section>
            <div className="card">
              <PageTitle title={selectedImportTarget.label} subtitle={selectedImportTarget.description} right={<Badge value={selectedImportTarget.table} />} />
              <div className="upload-box">
                <input type="file" accept=".xlsx,.xls,.csv" disabled={busyImport} onChange={(e) => { handleExcelUpload(selectedImport, e.target.files?.[0]); e.currentTarget.value = '' }} />
                <p>Service sheet reads FLEET No, ODO READING AT LAST SERVICE, SERVICE INTERVAL, NEXT SERVICE TYPE, ODO READING AT SERVICE, PERCENT DONE, HOURS TILL SERVICE DUE, ESTIMATE DAYS TILL SERVICE and comments.</p>
              </div>
              <div className="message">{importStatus}</div>
              <div className="meta-grid"><div><span>Selected</span><b>{selectedImportTarget.label}</b></div><div><span>Table</span><b>{selectedImportTarget.table}</b></div><div><span>File</span><b>{importFile || 'None'}</b></div><div><span>Mode</span><b>Upsert / Update</b></div></div>
            </div>
            <div className="card"><h2>Headings found</h2>{importHeadings.length ? <div className="chip-row">{importHeadings.map((h) => <span className="chip" key={h}>{h}</span>)}</div> : <div className="empty">No file loaded yet.</div>}</div>
            <div className="card"><h2>Preview before save</h2>{importPreview.length ? <MiniTable rows={importPreview} columns={Object.keys(importPreview[0]).filter((k) => !k.startsWith('__')).slice(0, 10).map((key) => ({ key, label: key }))} /> : <div className="empty">After upload, mapped rows will show here.</div>}</div>
          </section>
        </div>
      </div>
    )
  }

  function ExcelButton({ table }: { table: TableName }) {
    return <label className="excel-button">Upload Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { handleExcelUpload(table, e.target.files?.[0]); e.currentTarget.value = '' }} /></label>
  }

  function renderFleet() {
    return (
      <div>
        <PageTitle title="Fleet Register" subtitle="Machines, department allocation, fleet numbers, machine types, hours, mileage and status. Upload the department allocation Excel here." right={<><ExcelButton table="fleet_machines" /><PhotoUploadBox linkedType="Fleet" /></>} />
        <section className="card"><h2>Add machine</h2><div className="form-grid">
          <Field label="Fleet number"><Input value={machineForm.fleet_no} onChange={(v) => setMachineForm({ ...machineForm, fleet_no: v })} /></Field>
          <Field label="Machine type"><Input value={machineForm.machine_type} onChange={(v) => setMachineForm({ ...machineForm, machine_type: v })} /></Field>
          <Field label="Make / Model"><Input value={machineForm.make_model} onChange={(v) => setMachineForm({ ...machineForm, make_model: v })} /></Field>
          <Field label="Department"><Select value={machineForm.department} onChange={(v) => setMachineForm({ ...machineForm, department: v })} options={sections} /></Field>
          <Field label="Location"><Input value={machineForm.location} onChange={(v) => setMachineForm({ ...machineForm, location: v })} /></Field>
          <Field label="Hours"><Input type="number" value={machineForm.hours} onChange={(v) => setMachineForm({ ...machineForm, hours: v })} /></Field>
          <Field label="Mileage"><Input type="number" value={machineForm.mileage} onChange={(v) => setMachineForm({ ...machineForm, mileage: v })} /></Field>
          <Field label="Status"><Select value={machineForm.status} onChange={(v) => setMachineForm({ ...machineForm, status: v })} options={statuses} /></Field>
        </div><button className="primary" onClick={() => saveRows('fleet_machines', [{ ...machineForm, id: machineForm.fleet_no || uid() }])}>Save machine</button></section>
        <section className="card">
          <PageTitle title="Department breakdown" subtitle="This updates the dashboard fleet allocation cards." />
          <MiniTable rows={departmentSummaryRows(activeFleetRows(data.fleet_machines))} columns={[{ key: 'department', label: 'Department' }, { key: 'total', label: 'Machines' }, { key: 'available', label: 'Available' }, { key: 'breakdown', label: 'Breakdown' }, { key: 'fleet_list', label: 'Fleet numbers' }]} />
        </section>

        <section className="card">
          <PageTitle title="All machines" subtitle="Click a fleet number in Machine History to view service, repair, breakdown, spares, tyres, batteries, hoses and photos." />
          <MiniTable rows={activeFleetRows(data.fleet_machines)} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'machine_type', label: 'Type' }, { key: 'make_model', label: 'Make / Model' }, { key: 'department', label: 'Department' }, { key: 'hours', label: 'Hours' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
        </section>
        <section className="card"><h2>Fleet photos</h2><PhotoGallery linkedType="Fleet" /></section>
      </div>
    )
  }

  function renderPersonnel() {
    const currentLeave = data.leave_records.filter(isCurrentLeave)
    const upcomingLeave = data.leave_records.filter(isUpcomingLeave)
    const pastLeave = data.leave_records.filter(isPastLeave)
    const selectedLeaveRows = leaveView === 'current' ? currentLeave : leaveView === 'upcoming' ? upcomingLeave : pastLeave
    const selectedLeaveTitle = leaveView === 'current' ? 'Current leave / off list' : leaveView === 'upcoming' ? 'Upcoming leave list' : 'Past leave history'

    return (
      <div>
        <PageTitle title="Personnel, Shifts and Leave" subtitle="Click an employee to edit details or assign them to Shift 1, Shift 2 or Shift 3. Click leave cards to open the correct leave list." right={<><ExcelButton table="personnel" /><PhotoUploadBox linkedType="Personnel" /></>} />
        <div className="stats">
          <StatCard title="Employees" value={data.personnel.length} detail="Loaded personnel" onClick={() => document.getElementById('personnel-table')?.scrollIntoView({ behavior: 'smooth' })} />
          <StatCard title="Current Leave / Off" value={currentLeave.length} detail="Away today" tone="brown" active={leaveView === 'current'} onClick={() => clickLeaveCard('current')} />
          <StatCard title="Upcoming Leave" value={upcomingLeave.length} detail="Future leave records" tone="indigo" active={leaveView === 'upcoming'} onClick={() => clickLeaveCard('upcoming')} />
          <StatCard title="Past Leave Records" value={pastLeave.length} detail="History" tone="green" active={leaveView === 'past'} onClick={() => clickLeaveCard('past')} />
        </div>
        <section className="card" id="leave-list">
          <PageTitle title={selectedLeaveTitle} subtitle="Click any leave row to edit dates, type or status." right={<div className="button-row"><button className={leaveView === 'current' ? 'active-small' : ''} onClick={() => setLeaveView('current')}>Current</button><button className={leaveView === 'upcoming' ? 'active-small' : ''} onClick={() => setLeaveView('upcoming')}>Upcoming</button><button className={leaveView === 'past' ? 'active-small' : ''} onClick={() => setLeaveView('past')}>Past</button></div>} />
          <MiniTable rows={selectedLeaveRows} empty={`No ${leaveView} leave records.`} onRowClick={editLeave} activeId={editingLeaveId} columns={[{ key: 'person_name', label: 'Name', render: (r) => personNameFromLeave(r) }, { key: 'employee_no', label: 'Emp No' }, { key: 'leave_type', label: 'Type' }, { key: 'start_date', label: 'Start' }, { key: 'end_date', label: 'End' }, { key: 'days', label: 'Days' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} />
        </section>
        <div className="grid-2">
          <section className="card" id="employee-editor"><PageTitle title={editingPersonId ? 'Edit employee / assign shift' : 'Add employee / slot'} subtitle={editingPersonId ? 'You are editing the clicked employee. Change shift, role, section or status, then save.' : 'Add a new employee or vacant manpower slot.'} right={editingPersonId && <Badge value="Editing" />} />
            <div className="form-grid">
              <Field label="Employee no"><Input value={personForm.employee_no} onChange={(v) => setPersonForm({ ...personForm, employee_no: v })} /></Field>
              <Field label="Name"><Input value={personForm.name} onChange={(v) => setPersonForm({ ...personForm, name: v })} /></Field>
              <Field label="Shift"><Select value={personForm.shift} onChange={(v) => setPersonForm({ ...personForm, shift: v })} options={shiftNames} /></Field>
              <Field label="Role"><Select value={personForm.role} onChange={(v) => setPersonForm({ ...personForm, role: v })} options={personnelRoles} /></Field>
              <Field label="Section"><Select value={personForm.section} onChange={(v) => setPersonForm({ ...personForm, section: v })} options={sections} /></Field>
              <Field label="Status"><Select value={personForm.employment_status} onChange={(v) => setPersonForm({ ...personForm, employment_status: v })} options={['Active', 'Inactive', 'Vacant', 'On Leave', 'Off Duty']} /></Field>
              <Field label="Phone"><Input value={personForm.phone} onChange={(v) => setPersonForm({ ...personForm, phone: v })} /></Field>
              <Field label="Leave balance"><Input type="number" value={personForm.leave_balance_days} onChange={(v) => setPersonForm({ ...personForm, leave_balance_days: v })} /></Field>
              <Field label="Leave taken"><Input type="number" value={personForm.leave_taken_days} onChange={(v) => setPersonForm({ ...personForm, leave_taken_days: v })} /></Field>
              <Field label="Notes"><Input value={personForm.notes} onChange={(v) => setPersonForm({ ...personForm, notes: v })} /></Field>
            </div><div className="button-row"><button className="primary" onClick={savePersonForm}>{editingPersonId ? 'Save employee changes' : 'Save employee'}</button><button onClick={clearPersonForm}>Clear / new employee</button></div>
          </section>
          <section className="card"><h2>Upload personnel / leave Excel</h2><div className="upload-stack"><label><span>Personnel Excel</span><input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleExcelUpload('personnel', e.target.files?.[0])} /></label><label><span>Leave Summary Excel</span><input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleExcelUpload('leave_records', e.target.files?.[0])} /></label></div><PhotoUploadBox linkedType="Personnel" /></section>
        </div>
        <section className="card" id="leave-editor"><PageTitle title={editingLeaveId ? 'Edit leave / off record' : 'Leave / off schedule'} subtitle={editingLeaveId ? 'You are editing the clicked leave record.' : 'Create leave/off/sick records. The app calculates days automatically.'} right={editingLeaveId && <Badge value="Editing leave" />} />
          <div className="form-grid">
            <Field label="Employee"><Input value={leaveForm.person_name} onChange={(v) => setLeaveForm({ ...leaveForm, person_name: v })} /></Field>
            <Field label="Employee no"><Input value={leaveForm.employee_no} onChange={(v) => setLeaveForm({ ...leaveForm, employee_no: v })} /></Field>
            <Field label="Leave type"><Select value={leaveForm.leave_type} onChange={(v) => setLeaveForm({ ...leaveForm, leave_type: v })} options={leaveTypes} /></Field>
            <Field label="Start date"><Input type="date" value={leaveForm.start_date} onChange={(v) => setLeaveForm({ ...leaveForm, start_date: v })} /></Field>
            <Field label="End date"><Input type="date" value={leaveForm.end_date} onChange={(v) => setLeaveForm({ ...leaveForm, end_date: v })} /></Field>
            <Field label="Status"><Select value={leaveForm.status} onChange={(v) => setLeaveForm({ ...leaveForm, status: v })} options={['Approved', 'Pending', 'Rejected']} /></Field>
            <Field label="Reason"><Input value={leaveForm.reason || ''} onChange={(v) => setLeaveForm({ ...leaveForm, reason: v })} /></Field>
            <Field label="Calculated days"><Input value={daysInclusive(leaveForm.start_date, leaveForm.end_date)} onChange={() => {}} /></Field>
          </div><div className="button-row"><button className="primary" onClick={saveLeaveForm}>{editingLeaveId ? 'Save leave changes' : 'Save leave'}</button><button onClick={() => { setEditingLeaveId(''); setLeaveForm({ id: '', employee_no: '', person_name: '', leave_type: 'Annual Leave', start_date: today(), end_date: today(), status: 'Approved' }) }}>Clear / new leave</button></div>
        </section>
        <div className="grid-2"><section className="card" id="personnel-table"><PageTitle title="Personnel Register" subtitle="Click a person to load them into the editor above." /><MiniTable rows={data.personnel} onRowClick={editPerson} activeId={editingPersonId} columns={[{ key: 'employee_no', label: 'Emp No' }, { key: 'name', label: 'Name' }, { key: 'shift', label: 'Shift' }, { key: 'role', label: 'Role' }, { key: 'section', label: 'Section' }, { key: 'employment_status', label: 'Status', render: (r) => <Badge value={r.employment_status} /> }]} /></section><section className="card"><h2>Shift breakdown</h2><MiniTable rows={shiftNames.map((shift) => ({ id: shift, shift, employees: data.personnel.filter((p) => String(p.shift || '') === shift && String(p.employment_status || '').toLowerCase() !== 'vacant').length, vacancies: data.personnel.filter((p) => String(p.shift || '') === shift && String(p.employment_status || '').toLowerCase() === 'vacant').length }))} columns={[{ key: 'shift', label: 'Shift' }, { key: 'employees', label: 'Employees' }, { key: 'vacancies', label: 'Vacant slots' }]} /></section></div>
      </div>
    )
  }


  function getForm(key: string, defaults: Row = {}) {
    return { ...defaults, ...(moduleForms[key] || {}) }
  }

  function setFormValue(key: string, field: string, value: any) {
    setModuleForms((prev) => ({ ...prev, [key]: { ...(prev[key] || {}), [field]: value } }))
  }

  function clearForm(key: string) {
    setModuleForms((prev) => ({ ...prev, [key]: {} }))
  }

  function renderModuleForm(formKey: string, table: TableName, fields: { key: string; label: string; type?: string; options?: string[] }[], defaults: Row = {}) {
    const form = getForm(formKey, defaults)

    return (
      <>
        <div className="form-grid">
          {fields.map((field) => (
            <Field label={field.label} key={field.key}>
              {field.options
                ? <Select value={form[field.key] ?? defaults[field.key] ?? field.options[0]} onChange={(v) => setFormValue(formKey, field.key, v)} options={field.options} />
                : <Input type={field.type || 'text'} value={form[field.key] ?? defaults[field.key] ?? ''} onChange={(v) => setFormValue(formKey, field.key, v)} />}
            </Field>
          ))}
        </div>
        <div className="button-row">
          <button className="primary" onClick={() => {
            const payload: Row = { ...defaults, ...form, id: form.id || uid(), request_date: form.request_date || today(), created_at: nowIso() }
            if (table === 'services') {
              const current = num(payload.current_hours, 0)
              const last = num(payload.last_service_hours, 0)
              const interval = num(payload.service_interval, 250)
              const next = num(payload.next_service_hours, last + interval)
              const hoursLeft = next ? next - current : num(payload.hours_till_service_due, 999999)
              payload.next_service_hours = next
              payload.hours_till_service_due = Math.round(hoursLeft * 10) / 10
              payload.status = payload.status || serviceStatus(hoursLeft, interval ? (current - last) / interval : 0, true)
              payload.import_batch = 'manual-entry'
            }
            saveRows(table, [payload])
          }}>Save entry</button>
          <button onClick={() => clearForm(formKey)}>Clear</button>
        </div>
      </>
    )
  }

  function machineEvents(fleetNo: string) {
    const f = String(fleetNo || '').toLowerCase().trim()
    if (!f) return []

    const matchFleet = (row: Row) => [row.fleet_no, row.machine_fleet_no, row.fleet, row.machine, row.unit]
      .some((value) => String(value || '').toLowerCase().includes(f))

    return [
      ...latestServiceRows(data.services).filter(matchFleet).map((row) => ({ ...row, event_type: 'Service', event_text: row.service_type || row.next_service_type || 'Service', event_status: row.status })),
      ...data.breakdowns.filter(matchFleet).map((row) => ({ ...row, event_type: 'Breakdown', event_text: row.fault || row.description || 'Breakdown', event_status: row.status })),
      ...data.repairs.filter(matchFleet).map((row) => ({ ...row, event_type: 'Repair', event_text: row.fault || row.description || row.job_card_no || 'Repair', event_status: row.status })),
      ...data.spares_orders.filter(matchFleet).map((row) => ({ ...row, event_type: 'Spares order', event_text: row.description || row.part_no || 'Spares', event_status: row.status || row.workflow_stage })),
      ...data.tyres.filter(matchFleet).map((row) => ({ ...row, event_type: 'Tyre', event_text: row.serial_no || row.make || 'Tyre', event_status: row.status })),
      ...data.tyre_measurements.filter(matchFleet).map((row) => ({ ...row, event_type: 'Tyre measurement', event_text: row.position || row.tread_depth || 'Measurement', event_status: row.status })),
      ...data.batteries.filter(matchFleet).map((row) => ({ ...row, event_type: 'Battery', event_text: row.serial_no || row.make || 'Battery', event_status: row.status })),
      ...data.hose_requests.filter(matchFleet).map((row) => ({ ...row, event_type: 'Hose request', event_text: row.hose_type || row.description || 'Hose', event_status: row.status })),
      ...data.photo_logs.filter(matchFleet).map((row) => ({ ...row, event_type: 'Photo', event_text: row.caption || row.linked_type || 'Photo', event_status: 'Photo' }))
    ].slice(0, 200)
  }

  function renderMachineLookup(scope: string = 'all') {
    const q = machineLookup.toLowerCase().trim()
    const serviceRows = latestServiceRows(data.services)
    const fleetMatches = data.fleet_machines.filter((m) => `${m.fleet_no || ''} ${m.make_model || ''} ${m.machine_type || ''}`.toLowerCase().includes(q))
    const serviceMatches = serviceRows.filter((m) => String(m.fleet_no || '').toLowerCase().includes(q))
    const choices = (q ? [...fleetMatches, ...serviceMatches] : []).slice(0, 12)
    const target = selectedMachine || (choices[0]?.fleet_no || '')
    const events = machineEvents(target)

    return (
      <div className="machine-lookup">
        <div className="form-grid">
          <Field label="Search fleet / machine history">
            <Input value={machineLookup} onChange={(v) => { setMachineLookup(v); setSelectedMachine('') }} placeholder="Type fleet number e.g. FEL 1, AFQ 5794, HT 01" />
          </Field>
          <Field label="Selected machine">
            <Input value={target} onChange={setSelectedMachine} placeholder="Click a result below or type fleet number" />
          </Field>
        </div>
        {q && <div className="chip-row machine-results">
          {choices.length ? choices.map((row) => <button key={`${row.fleet_no}-${row.id}`} onClick={() => setSelectedMachine(String(row.fleet_no || ''))}>{row.fleet_no} {row.machine_type ? `• ${row.machine_type}` : ''}</button>) : <span className="chip">No machine found</span>}
        </div>}
        {target && <div className="grid-2">
          <section className="mini-panel"><h3>{target} history</h3><MiniTable rows={events} empty="No history for this machine yet." columns={[{ key: 'event_type', label: 'Type' }, { key: 'event_text', label: 'Details' }, { key: 'hours_till_service_due', label: 'Hours Left' }, { key: 'event_status', label: 'Status', render: (r) => <Badge value={r.event_status} /> }]} /></section>
          <section className="mini-panel"><h3>{target} photos / reminders</h3><PhotoUploadBox linkedType={scope === 'all' ? 'Machine History' : scope} fleetNo={target} /><PhotoGallery fleetNo={target} limit={6} /></section>
        </div>}
      </div>
    )
  }

  function renderSparesPage() {
    const lowStock = data.spares.filter((r) => num(r.stock_qty) <= num(r.min_qty, 1))
    return <div>
      <PageTitle title="Spares & Orders" subtitle="Stock, requests, awaiting funding, funded, local/international orders and delivery ETA." right={<><ExcelButton table="spares" /><PhotoUploadBox linkedType="Spares & Orders" /></>} />
      <div className="stats">{orderStages.map((stage) => <StatCard key={stage} title={stage} value={data.spares_orders.filter((r) => String(r.workflow_stage || r.status) === stage).length} detail="Spares orders" />)}</div>
      <div className="grid-2">
        <section className="card"><h2>Request spares / order</h2>{renderModuleForm('spares-order', 'spares_orders', [
          { key: 'machine_fleet_no', label: 'Fleet number' }, { key: 'request_date', label: 'Date', type: 'date' }, { key: 'requested_by', label: 'Requested by' }, { key: 'part_no', label: 'Part number' }, { key: 'description', label: 'Description' }, { key: 'qty', label: 'Qty', type: 'number' }, { key: 'priority', label: 'Priority', options: ['Low', 'Normal', 'High', 'Critical'] }, { key: 'workflow_stage', label: 'Workflow stage', options: orderStages }, { key: 'order_type', label: 'Order type', options: ['Local Order', 'International Order', 'Quote Only'] }, { key: 'eta', label: 'ETA', type: 'date' }, { key: 'notes', label: 'Notes' }
        ], { request_date: today(), qty: 1, workflow_stage: 'Requested', status: 'Requested', priority: 'Normal' })}</section>
        <section className="card"><h2>Add / update stock</h2>{renderModuleForm('spares-stock', 'spares', [
          { key: 'part_no', label: 'Part number' }, { key: 'description', label: 'Description' }, { key: 'machine_type', label: 'Machine type' }, { key: 'stock_qty', label: 'Stock qty', type: 'number' }, { key: 'min_qty', label: 'Minimum qty', type: 'number' }, { key: 'supplier', label: 'Supplier' }, { key: 'shelf_location', label: 'Shelf / bin' }, { key: 'order_status', label: 'Stock status', options: ['In stock', 'Low stock', 'Out of stock', 'Ordered'] }
        ], { stock_qty: 0, min_qty: 1, order_status: 'In stock' })}</section>
      </div>
      <section className="card"><h2>Low stock reminders</h2><MiniTable rows={lowStock} empty="No low stock reminders." columns={[{ key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'min_qty', label: 'Min' }, { key: 'supplier', label: 'Supplier' }]} /></section>
      <div className="grid-2"><section className="card"><h2>Spares stock</h2><ExcelButton table="spares" /><MiniTable rows={data.spares} columns={[{ key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'min_qty', label: 'Min' }, { key: 'order_status', label: 'Status', render: (r) => <Badge value={r.order_status} /> }]} /></section><section className="card"><h2>Spares orders</h2><ExcelButton table="spares_orders" /><MiniTable rows={data.spares_orders} columns={[{ key: 'machine_fleet_no', label: 'Fleet' }, { key: 'part_no', label: 'Part No' }, { key: 'description', label: 'Description' }, { key: 'qty', label: 'Qty' }, { key: 'workflow_stage', label: 'Stage', render: (r) => <Badge value={r.workflow_stage || r.status} /> }, { key: 'eta', label: 'ETA' }]} /></section></div>
      <section className="card"><h2>Spares photos</h2><PhotoGallery linkedType="Spares & Orders" /></section>
    </div>
  }

  function renderTyresPage() {
    return <div>
      <PageTitle title="Tyres" subtitle="Tyre register, repair requests, measurements, stock, reminders and ordering." right={<><ExcelButton table="tyres" /><PhotoUploadBox linkedType="Tyres" /></>} />
      <div className="grid-2">
        <section className="card"><h2>Fit / register tyre</h2>{renderModuleForm('tyre-register', 'tyres', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'position', label: 'Position' }, { key: 'make', label: 'Tyre make' }, { key: 'serial_no', label: 'Serial number' }, { key: 'company_no', label: 'Company number' }, { key: 'size', label: 'Size' }, { key: 'fitted_date', label: 'Date fitted', type: 'date' }, { key: 'fitted_by', label: 'Fitted by' }, { key: 'current_hours', label: 'Hours / mileage fitted', type: 'number' }, { key: 'status', label: 'Status', options: ['Fitted', 'In stock', 'Removed', 'Scrap', 'Repair'] }
        ], { fitted_date: today(), status: 'Fitted' })}</section>
        <section className="card"><h2>Tyre repair / measurement</h2>{renderModuleForm('tyre-measure', 'tyre_measurements', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'position', label: 'Position' }, { key: 'measurement_date', label: 'Date', type: 'date' }, { key: 'tread_depth', label: 'Tread depth', type: 'number' }, { key: 'pressure', label: 'Pressure', type: 'number' }, { key: 'reason', label: 'Repair / reason' }, { key: 'repaired_by', label: 'Checked / repaired by' }, { key: 'reminder_date', label: 'Next measurement reminder', type: 'date' }, { key: 'status', label: 'Status', options: ['Checked', 'Needs Repair', 'Repaired', 'Replace', 'Scrap'] }
        ], { measurement_date: today(), status: 'Checked' })}</section>
        <section className="card"><h2>Tyre order / quote</h2>{renderModuleForm('tyre-order', 'tyre_orders', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'requested_by', label: 'Requested by' }, { key: 'tyre_type', label: 'Tyre type' }, { key: 'size', label: 'Size' }, { key: 'qty', label: 'Qty', type: 'number' }, { key: 'reason', label: 'Reason' }, { key: 'status', label: 'Status', options: ['Requested', 'Quote Required', 'Awaiting Funding', 'Funded', 'Ordered', 'Delivered'] }, { key: 'eta', label: 'ETA', type: 'date' }
        ], { qty: 1, status: 'Requested' })}</section>
        <section className="card"><h2>Machine tyre history</h2>{renderMachineLookup('Tyres')}</section>
      </div>
      <section className="card"><h2>Tyre register</h2><ExcelButton table="tyres" /><MiniTable rows={data.tyres} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'position', label: 'Position' }, { key: 'serial_no', label: 'Serial' }, { key: 'make', label: 'Make' }, { key: 'size', label: 'Size' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section>
      <section className="card"><h2>Tyre photos</h2><PhotoGallery linkedType="Tyres" /></section>
    </div>
  }

  function renderBatteriesPage() {
    return <div>
      <PageTitle title="Batteries" subtitle="Battery register, stock, requests, maintenance reminders and fitment photos." right={<><ExcelButton table="batteries" /><PhotoUploadBox linkedType="Batteries" /></>} />
      <div className="grid-2">
        <section className="card"><h2>Fit / register battery</h2>{renderModuleForm('battery-register', 'batteries', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'make', label: 'Battery make' }, { key: 'serial_no', label: 'Serial number' }, { key: 'production_date', label: 'Production date', type: 'date' }, { key: 'fitment_date', label: 'Date fitted', type: 'date' }, { key: 'fitted_by', label: 'Fitted by' }, { key: 'volts', label: 'Charging system voltage' }, { key: 'hours_at_fitment', label: 'Hours / mileage fitted', type: 'number' }, { key: 'maintenance_due', label: 'Maintenance reminder', type: 'date' }, { key: 'status', label: 'Status', options: ['Fitted', 'In stock', 'Removed', 'Charging', 'Faulty', 'Scrap'] }
        ], { fitment_date: today(), status: 'Fitted' })}</section>
        <section className="card"><h2>Battery stock / order request</h2>{renderModuleForm('battery-order', 'battery_orders', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'requested_by', label: 'Requested by' }, { key: 'battery_type', label: 'Battery type' }, { key: 'volts', label: 'Voltage' }, { key: 'qty', label: 'Qty', type: 'number' }, { key: 'reason', label: 'Reason' }, { key: 'status', label: 'Status', options: ['Requested', 'Awaiting Funding', 'Funded', 'Ordered', 'Delivered', 'In stock'] }, { key: 'eta', label: 'ETA', type: 'date' }
        ], { qty: 1, status: 'Requested' })}</section>
      </div>
      <section className="card"><h2>Battery register</h2><ExcelButton table="batteries" /><MiniTable rows={data.batteries} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'serial_no', label: 'Serial' }, { key: 'make', label: 'Make' }, { key: 'volts', label: 'Volts' }, { key: 'fitment_date', label: 'Fitted' }, { key: 'maintenance_due', label: 'Maintenance' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section>
      <section className="card"><h2>Battery photos</h2><PhotoGallery linkedType="Batteries" /></section>
    </div>
  }

  function renderHosesPage() {
    return <div>
      <PageTitle title="Hoses" subtitle="Hose stock, fittings, bulk requests, individual hose repairs and offsite approval." right={<><ExcelButton table="hoses" /><PhotoUploadBox linkedType="Hoses" /></>} />
      <div className="grid-2">
        <section className="card"><h2>Hose stock entry</h2>{renderModuleForm('hose-stock', 'hoses', [
          { key: 'hose_no', label: 'Hose number' }, { key: 'hose_type', label: 'Hose type' }, { key: 'size', label: 'Size' }, { key: 'fitting_a', label: 'Fitting A' }, { key: 'fitting_b', label: 'Fitting B' }, { key: 'stock_qty', label: 'Stock qty', type: 'number' }, { key: 'min_qty', label: 'Minimum qty', type: 'number' }, { key: 'status', label: 'Status', options: ['In stock', 'Low stock', 'Out of stock', 'Ordered'] }
        ], { stock_qty: 0, min_qty: 1, status: 'In stock' })}</section>
        <section className="card"><h2>Hose request / repair</h2>{renderModuleForm('hose-request', 'hose_requests', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'request_type', label: 'Request type', options: ['Bulk Hose Request', 'Machine Hose Repair', 'Offsite Repair'] }, { key: 'hose_type', label: 'Hose type' }, { key: 'hose_no', label: 'Hose number' }, { key: 'qty', label: 'Qty', type: 'number' }, { key: 'requested_by', label: 'Requested by' }, { key: 'repair_reason', label: 'Reason / failure' }, { key: 'admin_approval', label: 'Admin approval', options: ['Pending Approval', 'Approved', 'Rejected'] }, { key: 'status', label: 'Status', options: ['Requested', 'Awaiting Approval', 'Approved', 'Offsite', 'Completed'] }
        ], { qty: 1, request_type: 'Machine Hose Repair', status: 'Requested', admin_approval: 'Pending Approval' })}</section>
      </div>
      <div className="grid-2"><section className="card"><h2>Hose Stock</h2><ExcelButton table="hoses" /><MiniTable rows={data.hoses} columns={[{ key: 'hose_no', label: 'Hose No' }, { key: 'hose_type', label: 'Type' }, { key: 'size', label: 'Size' }, { key: 'stock_qty', label: 'Stock' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section><section className="card"><h2>Hose Requests</h2><ExcelButton table="hose_requests" /><MiniTable rows={data.hose_requests} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'request_type', label: 'Type' }, { key: 'hose_type', label: 'Hose' }, { key: 'requested_by', label: 'Requested By' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section></div>
      <section className="card"><h2>Hose photos</h2><PhotoGallery linkedType="Hoses" /></section>
    </div>
  }

  function renderToolsPage() {
    return <div>
      <PageTitle title="Tools" subtitle="Inventory, orders and employee issued tools." right={<><ExcelButton table="tools" /><PhotoUploadBox linkedType="Tools" /></>} />
      <div className="grid-2">
        <section className="card"><h2>Add / update tool stock</h2>{renderModuleForm('tools-stock', 'tools', [
          { key: 'tool_name', label: 'Tool name' }, { key: 'brand', label: 'Brand' }, { key: 'serial_no', label: 'Serial no' }, { key: 'category', label: 'Category' }, { key: 'stock_qty', label: 'Quantity', type: 'number' }, { key: 'condition', label: 'Condition', options: ['New', 'Good', 'Needs Repair', 'Damaged', 'Scrap'] }, { key: 'status', label: 'Status', options: ['In stock', 'Issued', 'Ordered', 'Repair'] }
        ], { stock_qty: 1, condition: 'Good', status: 'In stock' })}</section>
        <section className="card"><h2>Issue tool to employee</h2>{renderModuleForm('employee-tool', 'employee_tools', [
          { key: 'employee_name', label: 'Employee name' }, { key: 'employee_no', label: 'Employee no' }, { key: 'tool_name', label: 'Tool name' }, { key: 'serial_no', label: 'Serial no' }, { key: 'issue_date', label: 'Issue date', type: 'date' }, { key: 'issued_by', label: 'Issued by' }, { key: 'return_date', label: 'Return date', type: 'date' }, { key: 'status', label: 'Status', options: ['Issued', 'Returned', 'Lost', 'Damaged'] }
        ], { issue_date: today(), status: 'Issued' })}</section>
        <section className="card"><h2>Tool order request</h2>{renderModuleForm('tool-order', 'tool_orders', [
          { key: 'tool_name', label: 'Tool name' }, { key: 'brand', label: 'Brand' }, { key: 'qty', label: 'Qty', type: 'number' }, { key: 'requested_by', label: 'Requested by' }, { key: 'reason', label: 'Reason' }, { key: 'status', label: 'Status', options: ['Requested', 'Awaiting Funding', 'Funded', 'Ordered', 'Delivered'] }, { key: 'eta', label: 'ETA', type: 'date' }
        ], { qty: 1, status: 'Requested' })}</section>
      </div>
      <div className="grid-2"><section className="card"><h2>Tools Inventory</h2><ExcelButton table="tools" /><MiniTable rows={data.tools} columns={[{ key: 'tool_name', label: 'Tool' }, { key: 'brand', label: 'Brand' }, { key: 'serial_no', label: 'Serial' }, { key: 'stock_qty', label: 'Qty' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section><section className="card"><h2>Employee Issued Tools</h2><ExcelButton table="employee_tools" /><MiniTable rows={data.employee_tools} columns={[{ key: 'employee_name', label: 'Employee' }, { key: 'tool_name', label: 'Tool' }, { key: 'issue_date', label: 'Issued' }, { key: 'issued_by', label: 'Issued By' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section></div>
      <section className="card"><h2>Tool photos</h2><PhotoGallery linkedType="Tools" /></section>
    </div>
  }

  function renderFabricationPage() {
    return <div>
      <PageTitle title="Boiler Shop / Panel Beaters" subtitle="Stock, material requests, projects and material booked to machines." right={<><ExcelButton table="fabrication_stock" /><PhotoUploadBox linkedType="Boiler/Panel" /></>} />
      <div className="grid-2">
        <section className="card"><h2>Material stock</h2>{renderModuleForm('fab-stock', 'fabrication_stock', [
          { key: 'department', label: 'Department', options: ['Boiler Shop', 'Panel Beating', 'Welding', 'Fabrication'] }, { key: 'material_type', label: 'Material type' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock qty', type: 'number' }, { key: 'min_qty', label: 'Minimum qty', type: 'number' }, { key: 'unit', label: 'Unit' }, { key: 'status', label: 'Status', options: ['In stock', 'Low stock', 'Out of stock', 'Ordered'] }
        ], { department: 'Boiler Shop', stock_qty: 0, min_qty: 1, status: 'In stock' })}</section>
        <section className="card"><h2>Request material / parts</h2>{renderModuleForm('fab-request', 'fabrication_requests', [
          { key: 'fleet_no', label: 'Fleet number' }, { key: 'department', label: 'Department', options: ['Boiler Shop', 'Panel Beating', 'Welding', 'Fabrication'] }, { key: 'requested_by', label: 'Requested by' }, { key: 'material_type', label: 'Material type' }, { key: 'description', label: 'Description' }, { key: 'qty', label: 'Qty', type: 'number' }, { key: 'status', label: 'Status', options: ['Requested', 'Awaiting Funding', 'Funded', 'Issued', 'Completed'] }
        ], { qty: 1, status: 'Requested' })}</section>
        <section className="card"><h2>Major project</h2>{renderModuleForm('fab-project', 'fabrication_projects', [
          { key: 'project_name', label: 'Project name' }, { key: 'fleet_no', label: 'Fleet number' }, { key: 'supervisor', label: 'Supervisor' }, { key: 'material_used', label: 'Material used' }, { key: 'progress_percent', label: 'Progress %', type: 'number' }, { key: 'target_date', label: 'Target date', type: 'date' }, { key: 'status', label: 'Status', options: ['Open', 'In progress', 'Awaiting Material', 'Completed'] }
        ], { progress_percent: 0, status: 'Open' })}</section>
      </div>
      <div className="grid-2"><section className="card"><h2>Material Stock</h2><ExcelButton table="fabrication_stock" /><MiniTable rows={data.fabrication_stock} columns={[{ key: 'department', label: 'Department' }, { key: 'material_type', label: 'Material' }, { key: 'description', label: 'Description' }, { key: 'stock_qty', label: 'Stock' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section><section className="card"><h2>Major Projects</h2><ExcelButton table="fabrication_projects" /><MiniTable rows={data.fabrication_projects} columns={[{ key: 'project_name', label: 'Project' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'supervisor', label: 'Supervisor' }, { key: 'progress_percent', label: 'Progress %' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section></div>
      <section className="card"><h2>Boiler/Panel photos</h2><PhotoGallery linkedType="Boiler/Panel" /></section>
    </div>
  }

  function renderServices() {
    const activeServices = latestServiceRows(data.services)
    const due = sortedByHoursLeft(activeServices.filter((r) => ['due', 'overdue'].includes(String(r.status || '').toLowerCase())))
    const dueSoon = sortedByHoursLeft(activeServices.filter((r) => String(r.status || '').toLowerCase() === 'due soon'))
    const upcoming = sortedByHoursLeft(activeServices.filter((r) => String(r.status || '').toLowerCase() === 'upcoming'))
    const scheduled = sortedByHoursLeft(activeServices.filter((r) => String(r.status || '').toLowerCase() === 'scheduled'))
    const columns = [
      { key: 'fleet_no', label: 'Fleet' },
      { key: 'machine_type', label: 'Type' },
      { key: 'service_type', label: 'Service' },
      { key: 'current_hours', label: 'Current' },
      { key: 'last_service_hours', label: 'Last Service' },
      { key: 'service_interval', label: 'Interval' },
      { key: 'next_service_hours', label: 'Next Service' },
      { key: 'hours_till_service_due', label: 'Hours Left' },
      { key: 'estimate_days_till_service', label: 'Days' },
      { key: 'status', label: 'Status', render: (r: Row) => <Badge value={r.status} /> }
    ]
    return (
      <div>
        <PageTitle title="Services" subtitle="Latest uploaded service sheet only. Due Now is true due/overdue only, not every machine." right={<><ExcelButton table="services" /><PhotoUploadBox linkedType="Services" /></>} />
        <div className="stats">
          <StatCard title="Due Now" value={due.length} detail="0 hours or overdue" tone="brown" />
          <StatCard title="Due Soon" value={dueSoon.length} detail="≤ 50 hours" tone="purple" />
          <StatCard title="Upcoming" value={upcoming.length} detail="≤ 250 hours" tone="indigo" />
          <StatCard title="Scheduled" value={scheduled.length} detail="Not due" tone="green" />
        </div>
        <section className="card"><h2>Search machine service history</h2>{renderMachineLookup('services')}</section>
        <section className="card"><h2>Quick service entry / reminder</h2>
          {renderModuleForm('service-entry', 'services', [
            { key: 'fleet_no', label: 'Fleet number' },
            { key: 'service_type', label: 'Service type', options: ['A Service', 'B Service', 'C Service', '250hr Service', '500hr Service', '1000hr Service', 'Repair Service'] },
            { key: 'current_hours', label: 'Current hours', type: 'number' },
            { key: 'last_service_hours', label: 'Last service hours', type: 'number' },
            { key: 'service_interval', label: 'Service interval', type: 'number' },
            { key: 'supervisor', label: 'Supervisor' },
            { key: 'job_card_no', label: 'Job card no' },
            { key: 'due_date', label: 'Reminder date', type: 'date' },
            { key: 'status', label: 'Status', options: ['Scheduled', 'Upcoming', 'Due Soon', 'Due', 'Completed'] },
            { key: 'notes', label: 'Notes' }
          ], { status: 'Scheduled', service_interval: 250 })}
        </section>
        <section className="card"><h2>Due / overdue service</h2><MiniTable rows={due} empty="No services due now." columns={columns} /></section>
        <section className="card"><h2>Due soon</h2><MiniTable rows={dueSoon} empty="No services due soon." columns={columns} /></section>
        <section className="card"><h2>Upcoming</h2><MiniTable rows={upcoming} empty="No upcoming services." columns={columns} /></section>
        <section className="card"><h2>All records from latest service sheet</h2><MiniTable rows={activeServices} columns={columns} /></section>
        <section className="card"><h2>Service photos</h2><PhotoGallery linkedType="Services" /></section>
      </div>
    )
  }

  function renderGenericPage(title: string, subtitle: string, table: TableName, columns: { key: string; label: string; render?: (row: Row) => any }[]) {
    return (
      <div>
        <PageTitle title={title} subtitle={subtitle} right={<><ExcelButton table={table} /><PhotoUploadBox linkedType={title} /></>} />
        <section className="card"><h2>Quick add</h2><div className="form-grid">
          <Field label="Fleet / Machine"><Input value={quickForm.fleet_no || quickForm.machine_fleet_no || ''} onChange={(v) => setQuickForm({ ...quickForm, fleet_no: v, machine_fleet_no: v })} /></Field>
          <Field label="Description / Fault / Item"><Input value={quickForm.description || quickForm.fault || quickForm.tool_name || ''} onChange={(v) => setQuickForm({ ...quickForm, description: v, fault: v, tool_name: v })} /></Field>
          <Field label="Requested / Assigned by"><Input value={quickForm.requested_by || quickForm.assigned_to || ''} onChange={(v) => setQuickForm({ ...quickForm, requested_by: v, assigned_to: v })} /></Field>
          <Field label="Status"><Select value={quickForm.status || 'Requested'} onChange={(v) => setQuickForm({ ...quickForm, status: v })} options={statuses} /></Field>
          <Field label="Qty"><Input type="number" value={quickForm.qty || 1} onChange={(v) => setQuickForm({ ...quickForm, qty: v })} /></Field>
          <Field label="ETA / Date"><Input type="date" value={quickForm.eta || today()} onChange={(v) => setQuickForm({ ...quickForm, eta: v, request_date: v })} /></Field>
        </div><button className="primary" onClick={() => saveRows(table, [{ ...quickForm, id: uid(), request_date: quickForm.request_date || today() }])}>Save record</button><PhotoUploadBox linkedType={title} fleetNo={quickForm.fleet_no || quickForm.machine_fleet_no || ''} /></section>
        <section className="card"><MiniTable rows={data[table]} columns={columns} /></section>
        <section className="card"><h2>{title} photos</h2><PhotoGallery linkedType={title} /></section>
      </div>
    )
  }

  function renderReports() {
    const reportFleet = activeFleetRows(data.fleet_machines)
    const byDepartment = reportFleet.reduce((acc: Row, row) => { const key = row.department || 'Unknown'; acc[key] = (acc[key] || 0) + 1; return acc }, {})
    const byType = reportFleet.reduce((acc: Row, row) => { const key = row.machine_type || 'Unknown'; acc[key] = (acc[key] || 0) + 1; return acc }, {})
    return (
      <div>
        <PageTitle title="Reports" subtitle="Machine availability, service due status, breakdowns by department, fleet type and workshop pressure points." right={<><button onClick={() => window.print()}>Print</button><PhotoUploadBox linkedType="Reports" /></>} />
        <div className="grid-2"><section className="card"><h2>Fleet by department</h2><MiniTable rows={Object.entries(byDepartment).map(([department, count]) => ({ id: department, department, count }))} columns={[{ key: 'department', label: 'Department' }, { key: 'count', label: 'Machines' }]} /></section><section className="card"><h2>Fleet by type</h2><MiniTable rows={Object.entries(byType).map(([machine_type, count]) => ({ id: machine_type, machine_type, count }))} columns={[{ key: 'machine_type', label: 'Type' }, { key: 'count', label: 'Machines' }]} /></section><section className="card"><h2>Open breakdowns</h2><MiniTable rows={dashboard.breakdowns} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'fault', label: 'Fault' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section><section className="card"><h2>Service due / soon</h2><MiniTable rows={dashboard.servicesDue} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'service_type', label: 'Service' }, { key: 'hours_till_service_due', label: 'Hours Left' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }]} /></section></div>
      </div>
    )
  }

  function renderPhotos() {
    return (
      <div>
        <PageTitle title="Photos" subtitle="Workshop photos uploaded from any page. Images preview locally; captions sync when Supabase allows." right={<PhotoUploadBox linkedType="General" />} />
        <section className="card"><PhotoGallery limit={200} /></section>
        <section className="card"><MiniTable rows={data.photo_logs} columns={[{ key: 'fleet_no', label: 'Fleet' }, { key: 'linked_type', label: 'Page' }, { key: 'caption', label: 'Caption' }, { key: 'notes', label: 'Notes' }]} /></section>
      </div>
    )
  }

  function renderContent() {
    if (tab === 'dashboard') return renderDashboard()
    if (tab === 'importer') return renderImporter()
    if (tab === 'fleet') return renderFleet()
    if (tab === 'personnel') return renderPersonnel()
    if (tab === 'services') return renderServices()
    if (tab === 'reports') return renderReports()
    if (tab === 'photos') return renderPhotos()

    if (tab === 'breakdowns') return renderGenericPage('Breakdowns', 'Breakdown reports, faults, spares ETA and assigned fitters.', 'breakdowns', [{ key: 'fleet_no', label: 'Fleet' }, { key: 'fault', label: 'Fault' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'spare_eta', label: 'Spares ETA' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }])
    if (tab === 'repairs') return renderGenericPage('Repairs / Job Cards', 'Repairs, job cards, parts used and fitters.', 'repairs', [{ key: 'fleet_no', label: 'Fleet' }, { key: 'job_card_no', label: 'Job Card' }, { key: 'fault', label: 'Fault' }, { key: 'assigned_to', label: 'Assigned' }, { key: 'status', label: 'Status', render: (r) => <Badge value={r.status} /> }])
    if (tab === 'tyres') return renderTyresPage()
    if (tab === 'batteries') return renderBatteriesPage()
    if (tab === 'operations') return renderGenericPage('Operations / Operators', 'Operators by machine type, history, score sheets, offences and damages.', 'operations', [{ key: 'operator_name', label: 'Operator' }, { key: 'machine_type', label: 'Machine Type' }, { key: 'fleet_no', label: 'Fleet' }, { key: 'score', label: 'Score' }, { key: 'offence', label: 'Offence' }, { key: 'supervisor', label: 'Supervisor' }])

    if (tab === 'spares') return renderSparesPage()

    if (tab === 'hoses') return renderHosesPage()

    if (tab === 'tools') return renderToolsPage()

    if (tab === 'fabrication') return renderFabricationPage()

    if (tab === 'admin') return <div><PageTitle title="Admin" subtitle="Sync, local storage and system controls." /><section className="card"><div className="button-row"><button className="primary" onClick={loadAll}>Sync from Supabase</button><button onClick={() => window.print()}>Print</button><button onClick={() => { localStorage.removeItem(SNAPSHOT_KEY); setData(emptyData); setMessage('Local snapshot cleared.') }}>Clear local snapshot</button></div><div className="meta-grid"><div><span>Connection</span><b>{online ? 'Online' : 'Offline'}</b></div><div><span>Supabase</span><b>{isSupabaseConfigured ? 'Configured' : 'Not configured'}</b></div><div><span>User</span><b>{session?.username}</b></div><div><span>Role</span><b>{session?.role}</b></div></div></section></div>

    return renderDashboard()
  }

  if (!session) return renderLogin()

  return (
    <main className="app-shell">
      <style>{baseStyles}</style>
      <aside className="sidebar"><div className="brand"><div className="logo-box">TE</div><div><b>Turbo Energy</b><small>Workshop Control</small></div></div><nav>{nav.map(([key, label]) => <button key={key} className={tab === key ? 'active' : ''} onClick={() => setTab(key)}>{label}</button>)}</nav></aside>
      <section className="main-area"><header className="topbar"><div><h1>{nav.find(([key]) => key === tab)?.[1] || 'Dashboard'}</h1><p>Navy/orange workshop system • laptop and phone friendly • online/offline records</p></div><div className="top-actions"><Badge value={online ? 'Online' : 'Offline'} /><span>{session.username} ({session.role})</span><button onClick={loadAll}>Sync</button><button onClick={logout}>Logout</button></div></header>{message && <div className="message">{message}</div>}{renderContent()}</section>
    </main>
  )
}

const baseStyles = `
  :root { --navy: #061a33; --navy2: #0b2b52; --orange: #ff7a1a; --blue: #9fe6ff; --line: rgba(159,230,255,.22); --card: rgba(8,29,55,.92); --text: #f5f9ff; --muted: #a9bad1; --danger: #ff8f8f; --good: #86efac; --warn: #ffd28a; }
  * { box-sizing: border-box; }
  body { margin: 0; background: radial-gradient(circle at top left, rgba(255,122,26,.20), transparent 30%), linear-gradient(135deg, #020814 0%, #061a33 48%, #0b2b52 100%); color: var(--text); font-family: Arial, Helvetica, sans-serif; }
  button, input, select, textarea { font: inherit; }
  button { border: 1px solid var(--line); background: rgba(255,255,255,.07); color: var(--text); border-radius: 14px; padding: 10px 14px; cursor: pointer; font-weight: 800; }
  button:hover { border-color: var(--blue); }
  .primary { background: linear-gradient(135deg, var(--orange), #ffae4d); color: #071a33; border: 0; font-weight: 900; }
  .full { width: 100%; }
  .login-shell { min-height: 100vh; display: grid; place-items: center; padding: 24px; }
  .login-card { width: min(460px, 100%); background: var(--card); border: 1px solid var(--line); border-radius: 28px; padding: 28px; box-shadow: 0 24px 80px rgba(0,0,0,.45); }
  .app-shell { min-height: 100vh; display: grid; grid-template-columns: 260px 1fr; }
  .sidebar { border-right: 1px solid var(--line); background: rgba(3,15,30,.92); padding: 18px 14px; position: sticky; top: 0; height: 100vh; overflow: auto; }
  .brand { display: flex; align-items: center; gap: 12px; margin-bottom: 22px; }
  .brand small { display: block; color: var(--muted); }
  .logo-box { width: 52px; height: 52px; border-radius: 14px; background: #ffffff; color: var(--navy); display: grid; place-items: center; font-weight: 900; font-size: 22px; box-shadow: 0 10px 30px rgba(0,0,0,.28); }
  nav { display: grid; gap: 8px; }
  nav button { text-align: left; border-radius: 14px; }
  nav button.active, .active-small { background: rgba(255,122,26,.18); border-color: var(--orange); box-shadow: inset 4px 0 0 var(--orange); }
  .main-area { padding: 24px; overflow: auto; }
  .topbar { display: flex; align-items: center; justify-content: space-between; gap: 18px; background: linear-gradient(135deg, var(--navy), var(--navy2)); border: 1px solid var(--line); border-radius: 24px; padding: 22px; margin-bottom: 22px; box-shadow: 0 18px 55px rgba(0,0,0,.35); }
  .topbar h1, .page-title h1 { margin: 0; font-size: 28px; }
  .topbar p, .page-title p, p { color: var(--muted); margin: 6px 0 0; line-height: 1.4; }
  .top-actions, .page-actions, .badge-row, .chip-row, .button-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
  .page-title { display: flex; align-items: center; justify-content: space-between; gap: 16px; margin: 16px 0; }
  .stats { display: grid; grid-template-columns: repeat(4, minmax(180px, 1fr)); gap: 14px; margin-bottom: 18px; }
  .stat-card { text-align: left; min-height: 130px; border-radius: 20px; padding: 18px; background: linear-gradient(135deg, rgba(16,68,112,.95), rgba(11,43,82,.95)); position: relative; overflow: hidden; }
  .stat-card.selected-stat { border-color: var(--orange); box-shadow: 0 0 0 2px rgba(255,122,26,.28), 0 14px 45px rgba(0,0,0,.28); }
  .stat-card::after { content: ''; position: absolute; width: 90px; height: 90px; right: -18px; bottom: -18px; border-radius: 30px; background: rgba(255,255,255,.08); }
  .stat-card span { color: #b9d8ff; font-size: 13px; letter-spacing: .08em; text-transform: uppercase; font-weight: 900; }
  .stat-card b { display: block; font-size: 38px; margin-top: 10px; }
  .stat-card small { color: #d7eaff; }
  .stat-card.purple { background: linear-gradient(135deg, rgba(77,45,91,.95), rgba(32,42,82,.95)); }
  .stat-card.brown { background: linear-gradient(135deg, rgba(86,55,38,.95), rgba(34,45,60,.95)); }
  .stat-card.indigo { background: linear-gradient(135deg, rgba(38,51,111,.95), rgba(23,40,85,.95)); }
  .stat-card.steel { background: linear-gradient(135deg, rgba(45,67,72,.95), rgba(31,53,66,.95)); }
  .stat-card.green { background: linear-gradient(135deg, rgba(14,83,77,.95), rgba(12,56,71,.95)); }
  .grid-2 { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; margin-top: 18px; }
  .import-layout { display: grid; grid-template-columns: 330px 1fr; gap: 18px; align-items: start; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 22px; padding: 18px; box-shadow: 0 14px 45px rgba(0,0,0,.28); margin-bottom: 18px; }
  .card h2 { margin: 0 0 12px; }
  .form-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; margin-bottom: 14px; }
  .field { display: grid; gap: 6px; }
  .field span, .upload-stack span, .photo-upload-mini span { color: #aee7ff; text-transform: uppercase; font-size: 12px; letter-spacing: .06em; font-weight: 900; }
  input, select, textarea { width: 100%; border: 1px solid var(--line); background: rgba(0,0,0,.28); color: var(--text); border-radius: 13px; padding: 12px; outline: none; }
  input:focus, select:focus, textarea:focus { border-color: var(--blue); }
  .upload-box, .upload-stack, .photo-upload-mini { border: 1px dashed rgba(159,230,255,.45); border-radius: 18px; padding: 16px; background: rgba(255,255,255,.06); display: grid; gap: 12px; margin-top: 12px; }
  .excel-button { display: inline-flex; align-items: center; gap: 8px; background: linear-gradient(135deg, var(--orange), #ffae4d); color: #071a33; border-radius: 14px; padding: 11px 14px; font-weight: 900; cursor: pointer; }
  .excel-button input { display: none; }
  .message { background: rgba(159,230,255,.12); border: 1px solid var(--line); border-radius: 16px; padding: 13px 14px; margin-bottom: 16px; font-weight: 900; }
  .badge { display: inline-flex; align-items: center; border-radius: 999px; padding: 7px 11px; font-size: 12px; font-weight: 900; border: 1px solid var(--line); }
  .badge.good { background: rgba(34,197,94,.18); color: var(--good); border-color: rgba(34,197,94,.35); }
  .badge.danger { background: rgba(239,68,68,.18); color: var(--danger); border-color: rgba(239,68,68,.35); }
  .badge.warn { background: rgba(255,122,26,.18); color: var(--warn); border-color: rgba(255,122,26,.35); }
  .badge.neutral { background: rgba(159,230,255,.12); color: var(--blue); }
  .chip { background: rgba(159,230,255,.15); color: var(--blue); border: 1px solid var(--line); border-radius: 999px; padding: 8px 10px; font-weight: 900; font-size: 12px; }
  .meta-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-top: 14px; }
  .meta-grid div { border: 1px solid var(--line); background: rgba(255,255,255,.06); border-radius: 16px; padding: 13px; }
  .meta-grid span { display: block; color: var(--muted); text-transform: uppercase; font-size: 11px; letter-spacing: .08em; font-weight: 900; }
  .meta-grid b { display: block; margin-top: 5px; word-break: break-word; }
  .table-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 16px; }
  table { width: 100%; border-collapse: collapse; min-width: 760px; }
  th, td { padding: 11px 12px; border-bottom: 1px solid rgba(159,230,255,.12); text-align: left; font-size: 13px; vertical-align: top; }
  th { color: var(--blue); background: rgba(0,0,0,.24); text-transform: uppercase; letter-spacing: .06em; }
  .click-row { cursor: pointer; }
  .click-row:hover { background: rgba(255,122,26,.10); }
  .active-row { background: rgba(255,122,26,.18); outline: 1px solid rgba(255,122,26,.45); }
  .empty { padding: 18px; border: 1px dashed var(--line); border-radius: 16px; color: var(--muted); }
  .import-list { display: grid; gap: 9px; max-height: 70vh; overflow: auto; padding-right: 5px; }
  .import-btn { text-align: left; display: grid; gap: 4px; }
  .import-btn.active { border-color: var(--orange); box-shadow: inset 4px 0 0 var(--orange); background: rgba(255,122,26,.16); }
  .import-btn small { color: var(--muted); }
  .photo-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 12px; }
  .photo-card { border: 1px solid var(--line); background: rgba(255,255,255,.06); border-radius: 16px; padding: 10px; display: grid; gap: 8px; }
  .photo-card img, .photo-placeholder { width: 100%; height: 120px; object-fit: cover; border-radius: 12px; background: rgba(0,0,0,.28); display: grid; place-items: center; color: var(--muted); }
  .photo-card small { color: var(--muted); }

  .machine-lookup {
    display: grid;
    gap: 14px;
  }

  .machine-results button {
    padding: 8px 11px;
    border-radius: 999px;
    color: var(--blue);
    background: rgba(159,230,255,.12);
  }

  .mini-panel {
    background: rgba(255,255,255,.04);
    border: 1px solid var(--line);
    border-radius: 18px;
    padding: 14px;
  }

  .mini-panel h3 {
    margin: 0 0 10px;
  }

  @media (max-width: 1000px) { .app-shell { grid-template-columns: 1fr; } .sidebar { position: relative; height: auto; } nav { grid-template-columns: repeat(2, minmax(0, 1fr)); } .topbar, .page-title { align-items: flex-start; flex-direction: column; } .stats, .grid-2, .form-grid, .import-layout, .meta-grid { grid-template-columns: 1fr; } .main-area { padding: 12px; } }
`

