'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { isSupabaseConfigured, supabase } from '@/lib/supabase'

type Row = Record<string, any>
type FieldType = 'text' | 'number' | 'date'

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

type FieldDef = {
  key: string
  aliases: string[]
  type?: FieldType
  fallback?: any
}

type ImportConfig = {
  table: TableName
  label: string
  description: string
  qualityKeys: string[]
  fields: FieldDef[]
}

const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
const today = () => new Date().toISOString().slice(0, 10)

const IMPORTS: ImportConfig[] = [
  {
    table: 'fleet_machines',
    label: 'Fleet Register',
    description: 'Machines, fleet numbers, hours, mileage, department and status.',
    qualityKeys: ['fleet_no'],
    fields: [
      f('fleet_no', ['fleet', 'fleet no', 'fleet number', 'machine', 'machine no', 'unit', 'unit no', 'equipment', 'equipment no']),
      f('machine_type', ['machine type', 'type', 'equipment type', 'fleet type']),
      f('make_model', ['make model', 'model', 'make', 'description']),
      f('reg_no', ['reg', 'registration', 'reg no']),
      f('department', ['department', 'dept', 'section', 'area'], 'text', 'Workshop'),
      f('location', ['location', 'site']),
      f('hours', ['hours', 'hrs', 'hm', 'hour meter', 'meter reading'], 'number', 0),
      f('mileage', ['mileage', 'km', 'odometer'], 'number', 0),
      f('status', ['status', 'state'], 'text', 'Available'),
      f('notes', ['notes', 'remarks'])
    ]
  },
  {
    table: 'personnel',
    label: 'Personnel Register',
    description: 'Employees, shifts, trades, sections, leave balances and status.',
    qualityKeys: ['name', 'employee_no'],
    fields: [
      f('employee_no', ['code', 'employee code', 'employee no', 'emp no', 'clock no', 'staff no']),
      f('name', ['name', 'employee name', 'full name', 'employee']),
      f('shift', ['shift'], 'text', 'Shift 1'),
      f('section', ['department', 'dept', 'section'], 'text', 'Workshop'),
      f('role', ['occupation', 'current job title', 'job title', 'trade', 'role', 'position'], 'text', 'Workshop Employee'),
      f('staff_group', ['staff group', 'group'], 'text', 'Workshop Staff'),
      f('phone', ['phone', 'cell', 'mobile']),
      f('employment_status', ['employment status', 'status'], 'text', 'Active'),
      f('leave_balance_days', ['leave balance', 'leave days', 'days leave'], 'number', 30),
      f('leave_taken_days', ['leave taken', 'days taken'], 'number', 0),
      f('notes', ['notes', 'remarks'])
    ]
  },
  {
    table: 'leave_records',
    label: 'Leave / Off Schedule',
    description: 'Current, upcoming and past leave records.',
    qualityKeys: ['person_name', 'employee_no'],
    fields: [
      f('employee_no', ['code', 'employee code', 'employee no', 'emp no', 'clock no']),
      f('person_name', ['name', 'employee name', 'full name', 'person name', 'employee']),
      f('shift', ['shift'], 'text', 'Shift 1'),
      f('section', ['department', 'dept', 'section']),
      f('role', ['occupation', 'current job title', 'job title', 'trade', 'role']),
      f('leave_type', ['leave type', 'type'], 'text', 'Annual Leave'),
      f('start_date', ['start date', 'leave start', 'from'], 'date'),
      f('end_date', ['end date', 'leave end', 'to'], 'date'),
      f('days', ['total', 'days', 'total days', 'leave days'], 'number', 0),
      f('reason', ['reason']),
      f('status', ['status'], 'text', 'Approved')
    ]
  },
  {
    table: 'services',
    label: 'Service Schedule',
    description: 'Services, service history, job cards and supervisors.',
    qualityKeys: ['fleet_no', 'service_type'],
    fields: [
      f('fleet_no', ['fleet', 'fleet no', 'machine', 'unit']),
      f('service_type', ['service type', 'service', 'type'], 'text', 'Service'),
      f('due_date', ['due date', 'service due date', 'date'], 'date'),
      f('scheduled_hours', ['scheduled hours', 'due hours', 'hours', 'hm'], 'number', 0),
      f('completed_date', ['completed date', 'date completed', 'done date'], 'date'),
      f('completed_hours', ['completed hours', 'done hours'], 'number', 0),
      f('job_card_no', ['job card', 'job card no']),
      f('supervisor', ['supervisor', 'foreman']),
      f('technician', ['technician', 'fitter', 'mechanic']),
      f('status', ['status'], 'text', 'Due'),
      f('notes', ['notes', 'remarks'])
    ]
  },
  {
    table: 'breakdowns',
    label: 'Breakdowns',
    description: 'Breakdown records, faults, ETAs and assigned fitters.',
    qualityKeys: ['fleet_no', 'fault'],
    fields: [
      f('fleet_no', ['fleet', 'fleet no', 'machine', 'unit']),
      f('category', ['category', 'area', 'section', 'department'], 'text', 'Mechanical'),
      f('fault', ['fault', 'breakdown', 'defect', 'problem', 'description']),
      f('reported_by', ['reported by', 'requester', 'raised by', 'operator']),
      f('assigned_to', ['assigned to', 'worked by', 'fitter', 'mechanic', 'technician']),
      f('start_time', ['start time', 'start date', 'date reported', 'date']),
      f('status', ['status'], 'text', 'Open'),
      f('spare_eta', ['spares eta', 'spare eta', 'eta'], 'date'),
      f('expected_available_date', ['expected available', 'available date', 'return date'], 'date'),
      f('cause', ['cause', 'root cause']),
      f('action_taken', ['action taken', 'repair done', 'work done']),
      f('preventative_recommendation', ['preventative', 'recommendation', 'preventative maintenance']),
      f('notes', ['notes', 'remarks'])
    ]
  },
  {
    table: 'repairs',
    label: 'Repairs / Job Cards',
    description: 'Repairs, job cards, parts used and fitters.',
    qualityKeys: ['fleet_no', 'fault', 'job_card_no'],
    fields: [
      f('fleet_no', ['fleet', 'fleet no', 'machine', 'unit']),
      f('job_card_no', ['job card', 'job card no', 'jc no', 'job no']),
      f('fault', ['fault', 'problem', 'defect', 'description']),
      f('assigned_to', ['assigned to', 'worked by', 'fitter', 'mechanic', 'technician']),
      f('start_time', ['start time', 'start date', 'date']),
      f('end_time', ['end time', 'end date', 'completed date']),
      f('status', ['status'], 'text', 'In progress'),
      f('parts_used', ['parts used', 'spares used', 'parts']),
      f('notes', ['notes', 'remarks'])
    ]
  },
  {
    table: 'spares',
    label: 'Spares Stock',
    description: 'Part numbers, stock quantities, suppliers and shelves.',
    qualityKeys: ['part_no', 'description'],
    fields: [
      f('part_no', ['part no', 'part number', 'spare no', 'stock code', 'item no']),
      f('description', ['description', 'item', 'spare', 'part name']),
      f('section', ['section', 'department'], 'text', 'Stores'),
      f('machine_group', ['machine group', 'group'], 'text', 'Yellow Machine'),
      f('machine_type', ['machine type', 'type']),
      f('stock_qty', ['stock qty', 'qty', 'quantity', 'stock', 'balance'], 'number', 0),
      f('min_qty', ['min qty', 'minimum', 'minimum stock', 'reorder level'], 'number', 1),
      f('supplier', ['supplier', 'vendor']),
      f('shelf_location', ['shelf', 'bin', 'location']),
      f('lead_time_days', ['lead time', 'lead time days'], 'number', 0),
      f('order_status', ['order status', 'status'], 'text', 'In stock')
    ]
  },
  {
    table: 'spares_orders',
    label: 'Spares Orders',
    description: 'Requests, awaiting funding, funded, local/international and ETA.',
    qualityKeys: ['machine_fleet_no', 'part_no', 'description', 'spares_items'],
    fields: [
      f('request_date', ['request date', 'date'], 'date'),
      f('machine_fleet_no', ['fleet', 'fleet no', 'machine', 'unit']),
      f('machine_group', ['machine group', 'group'], 'text', 'Yellow Machine'),
      f('workshop_section', ['workshop section', 'section', 'department'], 'text', 'Mechanical'),
      f('requested_by', ['requested by', 'requester', 'foreman', 'raised by']),
      f('part_no', ['part no', 'part number', 'spare no', 'stock code']),
      f('description', ['description', 'item', 'spare']),
      f('qty', ['qty', 'quantity', 'number'], 'number', 1),
      f('spares_items', ['spares items', 'items', 'list']),
      f('priority', ['priority'], 'text', 'Normal'),
      f('workflow_stage', ['workflow stage', 'stage', 'status'], 'text', 'Requested'),
      f('status', ['status', 'stage'], 'text', 'Requested'),
      f('order_type', ['order type', 'local international', 'local/international'], 'text', 'Not selected'),
      f('funding_status', ['funding status', 'funding'], 'text', 'Not funded'),
      f('eta', ['eta', 'delivery date', 'expected date'], 'date'),
      f('notes', ['notes', 'remarks'])
    ]
  },
  {
    table: 'tyres',
    label: 'Tyres Register',
    description: 'Tyre make, serial, company number, fitment and reminders.',
    qualityKeys: ['serial_no', 'company_no', 'fleet_no'],
    fields: [
      f('make', ['make', 'brand']),
      f('serial_no', ['serial', 'serial no', 'serial number', 'sn']),
      f('company_no', ['company no', 'company number']),
      f('production_date', ['production date', 'date of production'], 'date'),
      f('fitted_date', ['fitted date', 'date fitted', 'fitment date'], 'date'),
      f('fleet_no', ['fleet', 'fleet no', 'machine', 'unit']),
      f('fitted_by', ['fitted by']),
      f('position', ['position', 'wheel position'], 'text', 'FL'),
      f('tyre_type', ['tyre type', 'type']),
      f('size', ['size']),
      f('fitment_hours', ['fitment hours', 'hours', 'hm'], 'number', 0),
      f('fitment_mileage', ['fitment mileage', 'mileage', 'km'], 'number', 0),
      f('current_tread_mm', ['current tread', 'tread', 'tread depth'], 'number', 0),
      f('measurement_due_date', ['measurement due', 'reminder'], 'date'),
      f('stock_status', ['stock status'], 'text', 'Fitted'),
      f('status', ['status'], 'text', 'Fitted'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'tyre_measurements',
    label: 'Tyre Measurements',
    description: 'Tread depth, pressure, condition and checks.',
    qualityKeys: ['tyre_serial_no', 'fleet_no'],
    fields: [
      f('tyre_serial_no', ['tyre serial', 'serial no', 'serial']),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('position', ['position']),
      f('measurement_date', ['measurement date', 'date'], 'date'),
      f('measured_by', ['measured by', 'checked by']),
      f('tread_depth_mm', ['tread depth', 'tread'], 'number', 0),
      f('pressure_psi', ['pressure', 'psi'], 'number', 0),
      f('condition', ['condition'], 'text', 'Good'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'tyre_orders',
    label: 'Tyre Orders',
    description: 'Tyres to quote/order against machines.',
    qualityKeys: ['fleet_no', 'tyre_type', 'size'],
    fields: [
      f('request_date', ['request date', 'date'], 'date'),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('tyre_type', ['tyre type', 'type']),
      f('size', ['size']),
      f('qty', ['qty', 'quantity'], 'number', 1),
      f('requested_by', ['requested by', 'requester', 'foreman']),
      f('reason', ['reason']),
      f('priority', ['priority'], 'text', 'Normal'),
      f('status', ['status'], 'text', 'Requested'),
      f('eta', ['eta'], 'date'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'batteries',
    label: 'Battery Register',
    description: 'Battery make, serial, fleet, voltage and maintenance reminders.',
    qualityKeys: ['serial_no', 'fleet_no'],
    fields: [
      f('make', ['make', 'brand']),
      f('serial_no', ['serial', 'serial no', 'serial number']),
      f('production_date', ['production date', 'date of production'], 'date'),
      f('fitment_date', ['fitment date', 'date fitted', 'fitted date'], 'date'),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('fitted_by', ['fitted by']),
      f('charging_voltage', ['charging voltage', 'charging system voltage'], 'text', '24V'),
      f('volts', ['volts', 'voltage'], 'text', '12V'),
      f('fitment_hours', ['fitment hours', 'hours', 'hm'], 'number', 0),
      f('fitment_mileage', ['fitment mileage', 'mileage', 'km'], 'number', 0),
      f('maintenance_due_date', ['maintenance due', 'reminder'], 'date'),
      f('stock_status', ['stock status'], 'text', 'Fitted'),
      f('status', ['status'], 'text', 'Fitted'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'battery_orders',
    label: 'Battery Orders',
    description: 'Battery orders against machines.',
    qualityKeys: ['fleet_no', 'battery_type', 'voltage'],
    fields: [
      f('request_date', ['request date', 'date'], 'date'),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('battery_type', ['battery type', 'type']),
      f('voltage', ['voltage', 'volts'], 'text', '12V'),
      f('qty', ['qty', 'quantity'], 'number', 1),
      f('requested_by', ['requested by', 'requester', 'foreman']),
      f('reason', ['reason']),
      f('priority', ['priority'], 'text', 'Normal'),
      f('status', ['status'], 'text', 'Requested'),
      f('eta', ['eta'], 'date'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'operations',
    label: 'Operations / Operators',
    description: 'Operator history, scores, offences and damages.',
    qualityKeys: ['operator_name', 'fleet_no'],
    fields: [
      f('operator_name', ['operator name', 'operator', 'name']),
      f('employee_no', ['employee no', 'clock no', 'code']),
      f('machine_group', ['machine group', 'group'], 'text', 'Truck'),
      f('machine_type', ['machine type', 'type'], 'text', 'Truck'),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('shift', ['shift'], 'text', 'Shift 1'),
      f('score', ['score'], 'number', 100),
      f('offence', ['offence', 'offense']),
      f('damage_report', ['damage report', 'damage', 'abuse']),
      f('supervisor', ['supervisor', 'foreman']),
      f('event_date', ['event date', 'date'], 'date'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'hoses',
    label: 'Hose Stock',
    description: 'Hose and fitting stock.',
    qualityKeys: ['hose_no', 'hose_type', 'size'],
    fields: [
      f('hose_no', ['hose no', 'hose number']),
      f('hose_type', ['hose type', 'type'], 'text', 'Hydraulic Hose'),
      f('size', ['size']),
      f('length', ['length']),
      f('fitting_a', ['fitting a', 'fitting 1'], 'text', 'Straight'),
      f('fitting_b', ['fitting b', 'fitting 2'], 'text', 'Straight'),
      f('stock_qty', ['stock qty', 'qty', 'quantity'], 'number', 0),
      f('min_qty', ['min qty', 'minimum'], 'number', 1),
      f('shelf_location', ['shelf', 'location', 'bin']),
      f('status', ['status'], 'text', 'In stock'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'hose_requests',
    label: 'Hose Requests',
    description: 'Bulk hose requests and individual hose repairs.',
    qualityKeys: ['fleet_no', 'hose_type', 'requested_by'],
    fields: [
      f('request_date', ['request date', 'date'], 'date'),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('requested_by', ['requested by', 'requester', 'foreman']),
      f('hose_type', ['hose type', 'type'], 'text', 'Hydraulic Hose'),
      f('hose_no', ['hose no']),
      f('size', ['size']),
      f('length', ['length']),
      f('fittings', ['fittings']),
      f('qty', ['qty', 'quantity'], 'number', 1),
      f('repair_type', ['repair type'], 'text', 'New hose'),
      f('offsite_repair', ['offsite repair'], 'text', 'No'),
      f('admin_approval', ['admin approval', 'approval'], 'text', 'Pending'),
      f('status', ['status'], 'text', 'Requested'),
      f('eta', ['eta'], 'date'),
      f('reason', ['reason']),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'tools',
    label: 'Tools Inventory',
    description: 'Workshop tools, quantity, brand and condition.',
    qualityKeys: ['tool_name', 'serial_no'],
    fields: [
      f('tool_name', ['tool name', 'tool']),
      f('category', ['category'], 'text', 'Hand Tools'),
      f('brand', ['brand', 'make']),
      f('model', ['model']),
      f('serial_no', ['serial no', 'serial']),
      f('stock_qty', ['stock qty', 'qty', 'quantity'], 'number', 0),
      f('workshop_section', ['workshop section', 'section'], 'text', 'Mechanical'),
      f('condition', ['condition'], 'text', 'Good'),
      f('status', ['status'], 'text', 'In stock'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'tool_orders',
    label: 'Tool Orders',
    description: 'Requests to order tools.',
    qualityKeys: ['tool_name', 'requested_by'],
    fields: [
      f('request_date', ['request date', 'date'], 'date'),
      f('tool_name', ['tool name', 'tool']),
      f('category', ['category'], 'text', 'Hand Tools'),
      f('brand', ['brand']),
      f('qty', ['qty', 'quantity'], 'number', 1),
      f('requested_by', ['requested by', 'requester']),
      f('reason', ['reason']),
      f('priority', ['priority'], 'text', 'Normal'),
      f('status', ['status'], 'text', 'Requested'),
      f('eta', ['eta'], 'date'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'employee_tools',
    label: 'Employee Issued Tools',
    description: 'Tools issued to employees and return records.',
    qualityKeys: ['employee_name', 'tool_name'],
    fields: [
      f('employee_name', ['employee name', 'employee', 'name']),
      f('employee_no', ['employee no', 'clock no', 'code']),
      f('tool_name', ['tool name', 'tool']),
      f('brand', ['brand']),
      f('serial_no', ['serial no', 'serial']),
      f('issue_date', ['issue date', 'date issued'], 'date'),
      f('issued_by', ['issued by']),
      f('condition_issued', ['condition issued'], 'text', 'Good'),
      f('return_date', ['return date'], 'date'),
      f('condition_returned', ['condition returned']),
      f('status', ['status'], 'text', 'Issued'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'fabrication_stock',
    label: 'Boiler/Panel Stock',
    description: 'Material stock for boiler shop and panel beaters.',
    qualityKeys: ['description', 'material_type'],
    fields: [
      f('department', ['department'], 'text', 'Boiler Shop'),
      f('material_type', ['material type', 'material'], 'text', 'Steel Plate'),
      f('description', ['description', 'item']),
      f('size_spec', ['size spec', 'size']),
      f('stock_qty', ['stock qty', 'qty', 'quantity'], 'number', 0),
      f('min_qty', ['min qty', 'minimum'], 'number', 1),
      f('unit', ['unit'], 'text', 'pcs'),
      f('supplier', ['supplier']),
      f('location', ['location']),
      f('status', ['status'], 'text', 'In stock'),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'fabrication_requests',
    label: 'Boiler/Panel Requests',
    description: 'Material requests booked to machines and projects.',
    qualityKeys: ['description', 'fleet_no', 'project_name'],
    fields: [
      f('request_date', ['request date', 'date'], 'date'),
      f('department', ['department'], 'text', 'Boiler Shop'),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('requested_by', ['requested by', 'requester', 'foreman']),
      f('material_type', ['material type', 'material']),
      f('description', ['description', 'item']),
      f('qty', ['qty', 'quantity'], 'number', 1),
      f('project_name', ['project name', 'project']),
      f('status', ['status'], 'text', 'Requested'),
      f('eta', ['eta'], 'date'),
      f('reason', ['reason']),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'fabrication_projects',
    label: 'Boiler/Panel Projects',
    description: 'Major ongoing boiler shop and panel beating projects.',
    qualityKeys: ['project_name', 'fleet_no'],
    fields: [
      f('department', ['department'], 'text', 'Boiler Shop'),
      f('project_name', ['project name', 'project']),
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('supervisor', ['supervisor', 'foreman']),
      f('start_date', ['start date'], 'date'),
      f('target_date', ['target date'], 'date'),
      f('status', ['status'], 'text', 'Ongoing'),
      f('material_used', ['material used']),
      f('progress_percent', ['progress percent', 'progress'], 'number', 0),
      f('notes', ['notes'])
    ]
  },
  {
    table: 'photo_logs',
    label: 'Photo Captions Register',
    description: 'Excel captions only. Actual photos remain on Photos page.',
    qualityKeys: ['fleet_no', 'caption'],
    fields: [
      f('fleet_no', ['fleet', 'fleet no', 'machine']),
      f('linked_type', ['linked type', 'type'], 'text', 'Photo'),
      f('caption', ['caption', 'description']),
      f('notes', ['notes'])
    ]
  }
]

function f(key: string, aliases: string[], type: FieldType = 'text', fallback: any = ''): FieldDef {
  return { key, aliases, type, fallback }
}

function normalizeKey(value: any) {
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

function parseExcelDateNumber(value: number) {
  const parsed = XLSX.SSF.parse_date_code(value)
  if (!parsed) return ''

  const yyyy = String(parsed.y).padStart(4, '0')
  const mm = String(parsed.m).padStart(2, '0')
  const dd = String(parsed.d).padStart(2, '0')

  return `${yyyy}-${mm}-${dd}`
}

function dateOnly(value: any) {
  if (!value) return ''

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10)
  }

  if (typeof value === 'number' && value > 20000) {
    return parseExcelDateNumber(value)
  }

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
  if (match) {
    const yyyy = match[1]
    const mm = match[2].padStart(2, '0')
    const dd = match[3].padStart(2, '0')
    return `${yyyy}-${mm}-${dd}`
  }

  const parsed = new Date(text)
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10)
  }

  return ''
}

function daysInclusive(start: any, end: any) {
  const s = new Date(start)
  const e = new Date(end)

  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return 0

  return Math.max(0, Math.floor((e.getTime() - s.getTime()) / 86400000) + 1)
}

function exactCell(row: Row, aliases: string[]) {
  const keys = Object.keys(row)

  for (const alias of aliases) {
    const found = keys.find((key) => normalizeKey(key) === normalizeKey(alias))
    if (found && row[found] !== undefined && row[found] !== null && String(row[found]).trim() !== '') {
      return row[found]
    }
  }

  return ''
}

function fuzzyCell(row: Row, aliases: string[]) {
  const keys = Object.keys(row)

  for (const alias of aliases) {
    const wanted = normalizeKey(alias)

    const found = keys.find((key) => {
      const actual = normalizeKey(key)
      if (!actual || !wanted) return false
      return actual.includes(wanted) || wanted.includes(actual)
    })

    if (found && row[found] !== undefined && row[found] !== null && String(row[found]).trim() !== '') {
      return row[found]
    }
  }

  return ''
}

function getCell(row: Row, aliases: string[]) {
  const expanded = aliases.flatMap((alias) => {
    const n = normalizeKey(alias)
    const extra: string[] = []

    if (n.includes('fleet')) extra.push('fleet', 'fleet no', 'fleet number', 'machine', 'machine no', 'unit', 'unit no', 'equipment', 'asset')
    if (n.includes('employee')) extra.push('employee', 'employee name', 'emp no', 'employee no', 'clock no', 'staff no', 'code', 'name')
    if (n.includes('request')) extra.push('requested by', 'requester', 'foreman', 'supervisor', 'raised by', 'ordered by')
    if (n.includes('qty')) extra.push('qty', 'quantity', 'number', 'amount', 'stock', 'stock qty', 'balance')
    if (n.includes('part')) extra.push('part no', 'part number', 'spare no', 'spare number', 'item no', 'stock code')
    if (n.includes('serial')) extra.push('serial', 'serial no', 'serial number', 'sn', 's/n')
    if (n.includes('status')) extra.push('status', 'state', 'stage')
    if (n.includes('role')) extra.push('role', 'trade', 'position', 'job title', 'occupation', 'current job title')
    if (n.includes('section')) extra.push('section', 'department', 'dept', 'area')
    if (n.includes('hours')) extra.push('hours', 'hrs', 'hm', 'hour meter')
    if (n.includes('date')) extra.push('date', 'dates', 'request date', 'fitted date', 'fitment date', 'production date', 'due date')

    return [alias, ...extra]
  })

  return exactCell(row, expanded) || fuzzyCell(row, expanded)
}

function convertValue(raw: Row, field: FieldDef) {
  const found = getCell(raw, field.aliases)
  const value = found === '' || found === undefined || found === null ? field.fallback : found

  if (field.type === 'number') {
    const num = Number(String(value ?? '').replace(/,/g, ''))
    return Number.isFinite(num) ? num : field.fallback ?? 0
  }

  if (field.type === 'date') {
    return dateOnly(value)
  }

  return String(value ?? '').trim()
}

function makeId(table: TableName, mapped: Row) {
  if (table === 'fleet_machines') return mapped.fleet_no || uid()
  if (table === 'personnel') return mapped.employee_no || mapped.name || uid()
  return uid()
}

function parseDateRange(value: any) {
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

function fullNameFromRaw(raw: Row) {
  const firstName = getCell(raw, ['FirstName', 'First Name', 'Forename', 'Name'])
  const surname = getCell(raw, ['Surname', 'Last Name'])

  return `${String(firstName || '').trim()} ${String(surname || '').trim()}`.trim()
}

function polishMappedRow(table: TableName, raw: Row, mapped: Row) {
  const next = { ...mapped }

  if (table === 'personnel') {
    const fullName = fullNameFromRaw(raw)

    if (!next.name && fullName) next.name = fullName
    if (!next.employee_no) next.employee_no = getCell(raw, ['Code', 'Employee Code', 'Clock No', 'Emp No'])
    if (!next.role || next.role === 'Workshop Employee') {
      next.role = getCell(raw, ['Occupation', 'Current Job Title', 'Job Title', 'Trade']) || next.role
    }
    if (!next.section || next.section === 'Workshop') {
      next.section = getCell(raw, ['Department', 'Dept', 'Section']) || next.section
    }

    next.shift = next.shift || 'Shift 1'
    next.staff_group = next.staff_group || 'Workshop Staff'
    next.employment_status = next.employment_status || 'Active'
    next.id = next.employee_no || next.name || next.id || uid()
  }

  if (table === 'leave_records') {
    const fullName = fullNameFromRaw(raw)
    const dates = getCell(raw, ['DATES', 'Dates', 'Date Range', 'Leave Dates'])
    const [startDate, endDate] = parseDateRange(dates)

    if (!next.person_name && fullName) next.person_name = fullName
    if (!next.employee_no) next.employee_no = getCell(raw, ['Code', 'Employee Code', 'Clock No', 'Emp No'])
    if (!next.role) next.role = getCell(raw, ['Occupation', 'Current Job Title', 'Job Title', 'Trade'])
    if (!next.section) next.section = getCell(raw, ['Department', 'Dept', 'Section'])
    if (!next.leave_type || next.leave_type === 'Annual Leave') {
      next.leave_type = getCell(raw, ['LEAVE TYPE', 'Leave Type', 'Type']) || 'Annual Leave'
    }

    if (!next.start_date && startDate) next.start_date = startDate
    if (!next.end_date && endDate) next.end_date = endDate

    const totalDays = Number(getCell(raw, ['TOTAL', 'Total', 'Days', 'Leave Days']))
    if (!next.days || Number(next.days) === 0) {
      next.days = Number.isFinite(totalDays) && totalDays > 0 ? totalDays : daysInclusive(next.start_date, next.end_date)
    }

    next.shift = next.shift || 'Shift 1'
    next.status = next.status || 'Approved'
    next.id = next.id || uid()
  }

  return cleanRow(next)
}

function applyDefaults(table: TableName, row: Row) {
  const next = { ...row }

  const defaultTodayTables = [
    'spares_orders',
    'tyre_measurements',
    'tyre_orders',
    'battery_orders',
    'hose_requests',
    'tool_orders',
    'employee_tools',
    'fabrication_requests',
    'fabrication_projects',
    'operations'
  ]

  if (defaultTodayTables.includes(table)) {
    const key =
      table === 'tyre_measurements' ? 'measurement_date' :
      table === 'employee_tools' ? 'issue_date' :
      table === 'fabrication_projects' ? 'start_date' :
      table === 'operations' ? 'event_date' :
      'request_date'

    if (!next[key]) next[key] = today()
  }

  if (table === 'spares_orders') {
    next.workflow_stage = next.workflow_stage || 'Requested'
    next.status = next.status || next.workflow_stage
  }

  return cleanRow(next)
}

function mapRow(config: ImportConfig, raw: Row) {
  const mapped: Row = {}

  config.fields.forEach((field) => {
    mapped[field.key] = convertValue(raw, field)
  })

  mapped.id = makeId(config.table, mapped)

  return applyDefaults(config.table, polishMappedRow(config.table, raw, cleanRow(mapped)))
}

function hasUsefulData(config: ImportConfig, row: Row) {
  return config.qualityKeys.some((key) => {
    const value = row[key]
    return value !== undefined && value !== null && String(value).trim() !== ''
  })
}

function scoreHeaderRow(row: any[]) {
  const words = [
    'fleet', 'machine', 'unit', 'equipment',
    'firstname', 'first name', 'surname', 'employee', 'code',
    'occupation', 'current job title', 'department', 'shift',
    'dates', 'total', 'leave type',
    'part no', 'description', 'qty', 'serial', 'status',
    'requested by', 'supervisor', 'foreman'
  ]

  const cells = row.map((cell) => String(cell || '').trim()).filter(Boolean)
  let score = 0

  cells.forEach((cell) => {
    const normal = normalizeKey(cell)
    words.forEach((word) => {
      if (normal.includes(normalizeKey(word))) score += 1
    })
  })

  return score + Math.min(cells.length, 10) * 0.1
}

async function readWorkbook(file: File) {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
  const allRows: Row[] = []

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    const raw = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      raw: false
    }) as any[][]

    if (!raw.length) return

    let headerIndex = 0
    let bestScore = -1

    raw.forEach((row, index) => {
      const score = scoreHeaderRow(row)

      if (score > bestScore) {
        bestScore = score
        headerIndex = index
      }
    })

    const headers = (raw[headerIndex] || []).map((header, index) => {
      const clean = String(header || '').trim()
      return clean || `Column ${index + 1}`
    })

    raw.slice(headerIndex + 1).forEach((row) => {
      const obj: Row = {}
      let filled = 0

      headers.forEach((header, index) => {
        const value = row[index]

        if (value !== undefined && value !== null && String(value).trim() !== '') {
          obj[header] = value
          filled += 1
        }
      })

      if (filled > 0) {
        allRows.push({ ...obj, _sheet_name: sheetName })
      }
    })
  })

  return allRows
}

function Badge({ value }: { value: string }) {
  const v = String(value || 'Ready')
  const low = v.toLowerCase()

  const cls =
    low.includes('failed') || low.includes('error')
      ? 'danger'
      : low.includes('uploaded') || low.includes('connected')
        ? 'good'
        : 'neutral'

  return <span className={`badge ${cls}`}>{v}</span>
}

export default function ImporterPage() {
  const [selected, setSelected] = useState<TableName>('personnel')
  const [status, setStatus] = useState('Ready')
  const [busy, setBusy] = useState(false)
  const [previewRows, setPreviewRows] = useState<Row[]>([])
  const [rawHeadings, setRawHeadings] = useState<string[]>([])
  const [fileName, setFileName] = useState('')
  const [search, setSearch] = useState('')

  const selectedConfig = IMPORTS.find((item) => item.table === selected) || IMPORTS[0]

  const filteredImports = useMemo(() => {
    const q = search.toLowerCase()
    return IMPORTS.filter((item) => `${item.label} ${item.description} ${item.table}`.toLowerCase().includes(q))
  }, [search])

  async function insertRows(config: ImportConfig, rows: Row[]) {
    if (!supabase || !isSupabaseConfigured) {
      throw new Error('Supabase is not configured.')
    }

    let uploaded = 0
    const chunkSize = 100

    for (let i = 0; i < rows.length; i += chunkSize) {
      const chunk = rows.slice(i, i + chunkSize)

      const { error } = await supabase.from(config.table as any).upsert(chunk, { onConflict: 'id' })

      if (error) throw error

      uploaded += chunk.length
      setStatus(`Uploaded ${uploaded}/${rows.length} records to ${config.label}...`)
    }
  }

  async function handleUpload(file?: File) {
    if (!file) return

    setBusy(true)
    setFileName(file.name)
    setPreviewRows([])
    setRawHeadings([])

    try {
      setStatus(`Reading ${file.name}...`)

      const rawRows = await readWorkbook(file)
      const headings = Array.from(new Set(rawRows.flatMap((row) => Object.keys(row).filter((key) => key !== '_sheet_name'))))
      setRawHeadings(headings)

      const mapped = rawRows
        .map((row) => mapRow(selectedConfig, row))
        .filter((row) => hasUsefulData(selectedConfig, row))

      setPreviewRows(mapped.slice(0, 10))

      if (!mapped.length) {
        setStatus(`No usable rows found for ${selectedConfig.label}. Choose the correct register on the left.`)
        return
      }

      await insertRows(selectedConfig, mapped)
      setStatus(`Uploaded ${mapped.length} records to ${selectedConfig.label}.`)
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

        h1, h2, h3 {
          margin: 0;
        }

        p {
          color: var(--muted);
          line-height: 1.4;
        }

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

        .button-row {
          display: flex;
          gap: 10px;
          flex-wrap: wrap;
          margin-top: 14px;
        }

        .btn {
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
          grid-template-columns: 1fr;
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

        .headings {
          display: flex;
          gap: 8px;
          flex-wrap: wrap;
          max-height: 130px;
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
          .grid,
          .top-card,
          .help-grid {
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
            <p>Upload Excel/CSV for personnel, leave, fleet, spares, tyres, batteries, hoses, tools and workshop sections.</p>
          </div>
        </div>

        <div className="button-row">
          <a className="btn secondary" href="/">Back to app</a>
          <Badge value={isSupabaseConfigured ? 'Supabase connected' : 'Supabase not configured'} />
        </div>
      </section>

      <section className="grid">
        <aside className="card">
          <h2>Choose register</h2>
          <p>Select what your Excel file contains, then upload the file.</p>

          <input
            className="module-search"
            placeholder="Search register..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />

          <div className="module-list">
            {filteredImports.map((item) => (
              <button
                key={item.table}
                className={`module-btn ${selected === item.table ? 'active' : ''}`}
                onClick={() => {
                  setSelected(item.table)
                  setPreviewRows([])
                  setRawHeadings([])
                  setFileName('')
                  setStatus('Ready')
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
              <h2>{selectedConfig.label}</h2>
              <p>{selectedConfig.description}</p>

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
                  Accepts different headings. For your leave file it reads Code, FirstName, Surname,
                  Occupation, Department, DATES, TOTAL and LEAVE TYPE.
                </p>
              </div>

              <div className="status">{status}</div>
            </div>

            <div className="meta">
              <div>
                <span>Selected</span>
                <b>{selectedConfig.label}</b>
              </div>

              <div>
                <span>Database table</span>
                <b>{selectedConfig.table}</b>
              </div>

              <div>
                <span>Save mode</span>
                <b>Upsert / update</b>
              </div>

              <div>
                <span>File</span>
                <b>{fileName || 'None'}</b>
              </div>
            </div>
          </div>

          <div className="help-grid">
            <div className="card">
              <b>1. Choose register</b>
              <p>Choose Personnel Register for attendance/personnel files. Choose Leave / Off Schedule for leave files.</p>
            </div>

            <div className="card">
              <b>2. Upload Excel</b>
              <p>The importer finds the correct heading row even if headings are not on row 1.</p>
            </div>

            <div className="card">
              <b>3. Check app</b>
              <p>When it says Uploaded, go back to the normal app page and press CTRL + F5.</p>
            </div>
          </div>

          <div className="card" style={{ marginTop: 18 }}>
            <h3>Headings found in last file</h3>

            {rawHeadings.length ? (
              <div className="headings">
                {rawHeadings.map((h) => (
                  <span className="chip" key={h}>{h}</span>
                ))}
              </div>
            ) : (
              <div className="empty">No file loaded yet.</div>
            )}
          </div>

          <div className="card" style={{ marginTop: 18 }}>
            <h3>Preview of mapped rows</h3>

            {previewRows.length ? (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {Object.keys(previewRows[0]).slice(0, 12).map((key) => (
                        <th key={key}>{key}</th>
                      ))}
                    </tr>
                  </thead>

                  <tbody>
                    {previewRows.map((row) => (
                      <tr key={row.id || JSON.stringify(row)}>
                        {Object.keys(previewRows[0]).slice(0, 12).map((key) => (
                          <td key={key}>{String(row[key] ?? '')}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="empty">After upload, the first mapped rows will show here.</div>
            )}
          </div>
        </section>
      </section>
    </main>
  )
}
