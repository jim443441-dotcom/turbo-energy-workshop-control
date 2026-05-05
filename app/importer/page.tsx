'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { isSupabaseConfigured, supabase } from '@/lib/supabase'

type Row = Record<string, any>

type TableName =
  | 'fleet_machines'
  | 'breakdowns'
  | 'repairs'
  | 'services'
  | 'spares'
  | 'spares_orders'
  | 'personnel'
  | 'leave_records'
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

type FieldType = 'text' | 'number' | 'date'

type FieldDef = {
  key: string
  aliases: string[]
  type?: FieldType
  fallback?: any
}

type ImportItem = {
  table: TableName
  label: string
  description: string
  mode: 'insert' | 'upsert'
  conflictKey: string
  qualityKeys: string[]
}

const IMPORTS: ImportItem[] = [
  { table: 'fleet_machines', label: 'Fleet Register', description: 'Machines, departments, hours, mileage and status.', mode: 'upsert', conflictKey: 'fleet_no', qualityKeys: ['fleet_no'] },
  { table: 'personnel', label: 'Personnel Register', description: 'Employees, shifts, trades, sections and leave balances.', mode: 'upsert', conflictKey: 'id', qualityKeys: ['name', 'employee_no'] },
  { table: 'leave_records', label: 'Leave / Off Schedule', description: 'Current, upcoming and past leave records.', mode: 'insert', conflictKey: 'id', qualityKeys: ['person_name', 'employee_no'] },
  { table: 'services', label: 'Service Schedule', description: 'Due services, service history, supervisors and job cards.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'service_type'] },
  { table: 'breakdowns', label: 'Breakdowns', description: 'Breakdown reports, faults, spares ETA and assigned fitters.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'fault'] },
  { table: 'repairs', label: 'Repairs / Job Cards', description: 'Repair jobs, job cards, parts used and who worked on the machine.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'fault', 'job_card_no'] },
  { table: 'spares', label: 'Spares Stock', description: 'Part numbers, stock quantity, minimum stock, suppliers and shelves.', mode: 'insert', conflictKey: 'id', qualityKeys: ['part_no', 'description'] },
  { table: 'spares_orders', label: 'Spares Orders', description: 'Requests, awaiting funding, funded, local/international orders and ETA.', mode: 'insert', conflictKey: 'id', qualityKeys: ['machine_fleet_no', 'part_no', 'description', 'spares_items'] },
  { table: 'tyres', label: 'Tyres Register', description: 'Tyre make, serial, company number, fleet fitment and reminder dates.', mode: 'insert', conflictKey: 'id', qualityKeys: ['serial_no', 'company_no', 'fleet_no'] },
  { table: 'tyre_measurements', label: 'Tyre Measurements', description: 'Tread depth, pressure, condition and repair checks.', mode: 'insert', conflictKey: 'id', qualityKeys: ['tyre_serial_no', 'fleet_no'] },
  { table: 'tyre_orders', label: 'Tyre Orders', description: 'Tyres to quote/order against fleet numbers.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'tyre_type', 'size'] },
  { table: 'batteries', label: 'Battery Register', description: 'Battery make, serial, fleet, voltage, fitment and maintenance reminders.', mode: 'insert', conflictKey: 'id', qualityKeys: ['serial_no', 'fleet_no'] },
  { table: 'battery_orders', label: 'Battery Orders', description: 'Battery order requests against machines.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'battery_type', 'voltage'] },
  { table: 'operations', label: 'Operations / Operators', description: 'Operator history, machine type, score, offences and damages.', mode: 'insert', conflictKey: 'id', qualityKeys: ['operator_name', 'fleet_no'] },
  { table: 'hoses', label: 'Hose Stock', description: 'Hose and fitting stock levels.', mode: 'insert', conflictKey: 'id', qualityKeys: ['hose_no', 'hose_type', 'size'] },
  { table: 'hose_requests', label: 'Hose Requests', description: 'Bulk hose requests and individual hose repairs.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'hose_type', 'requested_by'] },
  { table: 'tools', label: 'Tools Inventory', description: 'Workshop tools, quantities, brands and condition.', mode: 'insert', conflictKey: 'id', qualityKeys: ['tool_name', 'serial_no'] },
  { table: 'tool_orders', label: 'Tool Orders', description: 'Requests to order tools.', mode: 'insert', conflictKey: 'id', qualityKeys: ['tool_name', 'requested_by'] },
  { table: 'employee_tools', label: 'Employee Issued Tools', description: 'Tools issued to employees and return records.', mode: 'insert', conflictKey: 'id', qualityKeys: ['employee_name', 'tool_name'] },
  { table: 'fabrication_stock', label: 'Boiler/Panel Stock', description: 'Material stock for boiler shop and panel beaters.', mode: 'insert', conflictKey: 'id', qualityKeys: ['description', 'material_type'] },
  { table: 'fabrication_requests', label: 'Boiler/Panel Requests', description: 'Material requests booked to machines and projects.', mode: 'insert', conflictKey: 'id', qualityKeys: ['description', 'fleet_no', 'project_name'] },
  { table: 'fabrication_projects', label: 'Boiler/Panel Projects', description: 'Major ongoing boiler shop and panel beating projects.', mode: 'insert', conflictKey: 'id', qualityKeys: ['project_name', 'fleet_no'] },
  { table: 'photo_logs', label: 'Photo Captions Register', description: 'Excel captions only. Actual photos still upload in the Photos page.', mode: 'insert', conflictKey: 'id', qualityKeys: ['fleet_no', 'caption'] }
]

const SCHEMAS: Record<TableName, FieldDef[]> = {
  fleet_machines: [
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'fleet number', 'machine', 'machine no', 'unit', 'unit no', 'equipment', 'equipment no', 'plant no'] },
    { key: 'machine_type', aliases: ['machine type', 'type', 'equipment type', 'fleet type'] },
    { key: 'make_model', aliases: ['make model', 'model', 'make', 'description'] },
    { key: 'reg_no', aliases: ['reg', 'registration', 'reg no', 'registration no'] },
    { key: 'department', aliases: ['department', 'dept', 'section', 'area'], fallback: 'Workshop' },
    { key: 'location', aliases: ['location', 'site'] },
    { key: 'hours', aliases: ['hours', 'hrs', 'hm', 'hour meter'], type: 'number', fallback: 0 },
    { key: 'mileage', aliases: ['mileage', 'km', 'odometer'], type: 'number', fallback: 0 },
    { key: 'status', aliases: ['status', 'state'], fallback: 'Available' },
    { key: 'service_interval_hours', aliases: ['service interval', 'service interval hours'], type: 'number', fallback: 250 },
    { key: 'last_service_hours', aliases: ['last service', 'last service hours'], type: 'number', fallback: 0 },
    { key: 'next_service_hours', aliases: ['next service', 'next service hours'], type: 'number', fallback: 0 },
    { key: 'notes', aliases: ['notes', 'remarks', 'comment'] }
  ],
  breakdowns: [
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'fleet number', 'machine', 'unit'] },
    { key: 'category', aliases: ['category', 'area', 'section', 'department'], fallback: 'Mechanical' },
    { key: 'fault', aliases: ['fault', 'breakdown', 'defect', 'problem', 'description', 'complaint'] },
    { key: 'reported_by', aliases: ['reported by', 'requester', 'raised by', 'operator'] },
    { key: 'assigned_to', aliases: ['assigned to', 'worked by', 'fitter', 'mechanic', 'technician'] },
    { key: 'start_time', aliases: ['start time', 'start date', 'date reported', 'date'] },
    { key: 'status', aliases: ['status', 'state'], fallback: 'Open' },
    { key: 'spare_eta', aliases: ['spares eta', 'spare eta', 'eta'], type: 'date' },
    { key: 'expected_available_date', aliases: ['expected available', 'available date', 'return date'], type: 'date' },
    { key: 'cause', aliases: ['cause', 'root cause'] },
    { key: 'action_taken', aliases: ['action taken', 'repair done', 'work done'] },
    { key: 'preventative_recommendation', aliases: ['preventative', 'recommendation', 'preventative maintenance'] },
    { key: 'notes', aliases: ['notes', 'remarks'] }
  ],
  repairs: [
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine', 'unit'] },
    { key: 'job_card_no', aliases: ['job card', 'job card no', 'jc no', 'job no'] },
    { key: 'fault', aliases: ['fault', 'problem', 'defect', 'description'] },
    { key: 'assigned_to', aliases: ['assigned to', 'worked by', 'fitter', 'mechanic', 'technician'] },
    { key: 'start_time', aliases: ['start time', 'start date', 'date'] },
    { key: 'end_time', aliases: ['end time', 'end date', 'completed date'] },
    { key: 'status', aliases: ['status'], fallback: 'In progress' },
    { key: 'parts_used', aliases: ['parts used', 'spares used', 'parts'] },
    { key: 'notes', aliases: ['notes', 'remarks'] }
  ],
  services: [
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine', 'unit'] },
    { key: 'service_type', aliases: ['service type', 'service', 'type'], fallback: '250hr service' },
    { key: 'due_date', aliases: ['due date', 'service due date', 'date'], type: 'date' },
    { key: 'scheduled_hours', aliases: ['scheduled hours', 'due hours', 'hours', 'hm'], type: 'number', fallback: 0 },
    { key: 'completed_date', aliases: ['completed date', 'date completed', 'done date'], type: 'date' },
    { key: 'completed_hours', aliases: ['completed hours', 'done hours'], type: 'number', fallback: 0 },
    { key: 'job_card_no', aliases: ['job card', 'job card no'] },
    { key: 'supervisor', aliases: ['supervisor', 'foreman'] },
    { key: 'technician', aliases: ['technician', 'fitter', 'mechanic'] },
    { key: 'status', aliases: ['status'], fallback: 'Due' },
    { key: 'notes', aliases: ['notes', 'remarks'] }
  ],
  spares: [
    { key: 'part_no', aliases: ['part no', 'part number', 'spare no', 'stock code', 'item no'] },
    { key: 'description', aliases: ['description', 'item', 'spare', 'part name'] },
    { key: 'section', aliases: ['section', 'department'], fallback: 'Stores' },
    { key: 'machine_group', aliases: ['machine group', 'group'], fallback: 'Yellow Machine' },
    { key: 'machine_type', aliases: ['machine type', 'type'] },
    { key: 'stock_qty', aliases: ['stock qty', 'qty', 'quantity', 'stock', 'balance'], type: 'number', fallback: 0 },
    { key: 'min_qty', aliases: ['min qty', 'minimum', 'minimum stock', 'reorder level'], type: 'number', fallback: 1 },
    { key: 'supplier', aliases: ['supplier', 'vendor'] },
    { key: 'shelf_location', aliases: ['shelf', 'bin', 'location'] },
    { key: 'lead_time_days', aliases: ['lead time', 'lead time days'], type: 'number', fallback: 0 },
    { key: 'order_status', aliases: ['order status', 'status'], fallback: 'In stock' }
  ],
  spares_orders: [
    { key: 'request_date', aliases: ['request date', 'date'], type: 'date' },
    { key: 'machine_fleet_no', aliases: ['fleet', 'fleet no', 'machine', 'unit'] },
    { key: 'machine_group', aliases: ['machine group', 'group'], fallback: 'Yellow Machine' },
    { key: 'workshop_section', aliases: ['workshop section', 'section', 'department'], fallback: 'Mechanical' },
    { key: 'requested_by', aliases: ['requested by', 'requester', 'foreman', 'raised by'] },
    { key: 'part_no', aliases: ['part no', 'part number', 'spare no', 'stock code'] },
    { key: 'description', aliases: ['description', 'item', 'spare'] },
    { key: 'qty', aliases: ['qty', 'quantity', 'number'], type: 'number', fallback: 1 },
    { key: 'spares_items', aliases: ['spares items', 'items', 'list'] },
    { key: 'priority', aliases: ['priority'], fallback: 'Normal' },
    { key: 'workflow_stage', aliases: ['workflow stage', 'stage', 'status'], fallback: 'Requested' },
    { key: 'status', aliases: ['status', 'stage'], fallback: 'Requested' },
    { key: 'order_type', aliases: ['order type', 'local international', 'local/international'], fallback: 'Not selected' },
    { key: 'funding_status', aliases: ['funding status', 'funding'], fallback: 'Not funded' },
    { key: 'eta', aliases: ['eta', 'delivery date', 'expected date'], type: 'date' },
    { key: 'notes', aliases: ['notes', 'remarks'] }
  ],
  personnel: [
    { key: 'employee_no', aliases: ['employee no', 'emp no', 'clock no', 'staff no', 'code', 'employee code'] },
    { key: 'name', aliases: ['name', 'employee name', 'employee', 'worker', 'person', 'full name'] },
    { key: 'shift', aliases: ['shift'], fallback: 'Shift 1' },
    { key: 'section', aliases: ['section', 'department', 'dept'], fallback: 'Workshop' },
    { key: 'role', aliases: ['role', 'trade', 'position', 'job title', 'occupation', 'current job title'], fallback: 'Diesel Plant Fitter' },
    { key: 'staff_group', aliases: ['staff group', 'group'], fallback: 'Workshop Staff' },
    { key: 'phone', aliases: ['phone', 'cell', 'mobile'] },
    { key: 'employment_status', aliases: ['employment status', 'status'], fallback: 'Active' },
    { key: 'leave_balance_days', aliases: ['leave balance', 'leave days', 'days leave'], type: 'number', fallback: 30 },
    { key: 'leave_taken_days', aliases: ['leave taken', 'days taken'], type: 'number', fallback: 0 },
    { key: 'notes', aliases: ['notes', 'remarks'] }
  ],
  leave_records: [
    { key: 'employee_no', aliases: ['employee no', 'emp no', 'clock no', 'code', 'employee code'] },
    { key: 'person_name', aliases: ['person name', 'employee name', 'name', 'employee', 'full name'] },
    { key: 'shift', aliases: ['shift'], fallback: 'Shift 1' },
    { key: 'section', aliases: ['section', 'department'] },
    { key: 'role', aliases: ['role', 'trade', 'position', 'occupation', 'current job title'] },
    { key: 'leave_type', aliases: ['leave type', 'type'], fallback: 'Annual Leave' },
    { key: 'start_date', aliases: ['start date', 'leave start', 'from'], type: 'date' },
    { key: 'end_date', aliases: ['end date', 'leave end', 'to'], type: 'date' },
    { key: 'days', aliases: ['days', 'total days', 'total'], type: 'number' },
    { key: 'reason', aliases: ['reason'] },
    { key: 'status', aliases: ['status'], fallback: 'Approved' }
  ],
  tyres: [
    { key: 'make', aliases: ['make', 'brand'] },
    { key: 'serial_no', aliases: ['serial', 'serial no', 'serial number', 'sn'] },
    { key: 'company_no', aliases: ['company no', 'company number'] },
    { key: 'production_date', aliases: ['production date', 'date of production'], type: 'date' },
    { key: 'fitted_date', aliases: ['fitted date', 'date fitted', 'fitment date'], type: 'date' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine', 'unit'] },
    { key: 'fitted_by', aliases: ['fitted by'] },
    { key: 'position', aliases: ['position', 'wheel position'], fallback: 'FL' },
    { key: 'tyre_type', aliases: ['tyre type', 'type'] },
    { key: 'size', aliases: ['size'] },
    { key: 'fitment_hours', aliases: ['fitment hours', 'hours', 'hm'], type: 'number', fallback: 0 },
    { key: 'fitment_mileage', aliases: ['fitment mileage', 'mileage', 'km'], type: 'number', fallback: 0 },
    { key: 'current_tread_mm', aliases: ['current tread', 'tread', 'tread depth'], type: 'number', fallback: 0 },
    { key: 'measurement_due_date', aliases: ['measurement due', 'reminder'], type: 'date' },
    { key: 'stock_status', aliases: ['stock status'], fallback: 'Fitted' },
    { key: 'status', aliases: ['status'], fallback: 'Fitted' },
    { key: 'notes', aliases: ['notes'] }
  ],
  tyre_measurements: [
    { key: 'tyre_serial_no', aliases: ['tyre serial', 'serial no', 'serial'] },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'position', aliases: ['position'] },
    { key: 'measurement_date', aliases: ['measurement date', 'date'], type: 'date' },
    { key: 'measured_by', aliases: ['measured by', 'checked by'] },
    { key: 'tread_depth_mm', aliases: ['tread depth', 'tread'], type: 'number', fallback: 0 },
    { key: 'pressure_psi', aliases: ['pressure', 'psi'], type: 'number', fallback: 0 },
    { key: 'condition', aliases: ['condition'], fallback: 'Good' },
    { key: 'notes', aliases: ['notes'] }
  ],
  tyre_orders: [
    { key: 'request_date', aliases: ['request date', 'date'], type: 'date' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'tyre_type', aliases: ['tyre type', 'type'] },
    { key: 'size', aliases: ['size'] },
    { key: 'qty', aliases: ['qty', 'quantity'], type: 'number', fallback: 1 },
    { key: 'requested_by', aliases: ['requested by', 'requester', 'foreman'] },
    { key: 'reason', aliases: ['reason'] },
    { key: 'priority', aliases: ['priority'], fallback: 'Normal' },
    { key: 'status', aliases: ['status'], fallback: 'Requested' },
    { key: 'eta', aliases: ['eta'], type: 'date' },
    { key: 'notes', aliases: ['notes'] }
  ],
  batteries: [
    { key: 'make', aliases: ['make', 'brand'] },
    { key: 'serial_no', aliases: ['serial', 'serial no', 'serial number'] },
    { key: 'production_date', aliases: ['production date', 'date of production'], type: 'date' },
    { key: 'fitment_date', aliases: ['fitment date', 'date fitted', 'fitted date'], type: 'date' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'fitted_by', aliases: ['fitted by'] },
    { key: 'charging_voltage', aliases: ['charging voltage', 'charging system voltage'], fallback: '24V' },
    { key: 'volts', aliases: ['volts', 'voltage'], fallback: '12V' },
    { key: 'fitment_hours', aliases: ['fitment hours', 'hours', 'hm'], type: 'number', fallback: 0 },
    { key: 'fitment_mileage', aliases: ['fitment mileage', 'mileage', 'km'], type: 'number', fallback: 0 },
    { key: 'maintenance_due_date', aliases: ['maintenance due', 'reminder'], type: 'date' },
    { key: 'stock_status', aliases: ['stock status'], fallback: 'Fitted' },
    { key: 'status', aliases: ['status'], fallback: 'Fitted' },
    { key: 'notes', aliases: ['notes'] }
  ],
  battery_orders: [
    { key: 'request_date', aliases: ['request date', 'date'], type: 'date' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'battery_type', aliases: ['battery type', 'type'] },
    { key: 'voltage', aliases: ['voltage', 'volts'], fallback: '12V' },
    { key: 'qty', aliases: ['qty', 'quantity'], type: 'number', fallback: 1 },
    { key: 'requested_by', aliases: ['requested by', 'requester', 'foreman'] },
    { key: 'reason', aliases: ['reason'] },
    { key: 'priority', aliases: ['priority'], fallback: 'Normal' },
    { key: 'status', aliases: ['status'], fallback: 'Requested' },
    { key: 'eta', aliases: ['eta'], type: 'date' },
    { key: 'notes', aliases: ['notes'] }
  ],
  operations: [
    { key: 'operator_name', aliases: ['operator name', 'operator', 'name'] },
    { key: 'employee_no', aliases: ['employee no', 'clock no'] },
    { key: 'machine_group', aliases: ['machine group', 'group'], fallback: 'Truck' },
    { key: 'machine_type', aliases: ['machine type', 'type'], fallback: 'Truck' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'shift', aliases: ['shift'], fallback: 'Shift 1' },
    { key: 'score', aliases: ['score'], type: 'number', fallback: 100 },
    { key: 'offence', aliases: ['offence', 'offense'] },
    { key: 'damage_report', aliases: ['damage report', 'damage', 'abuse'] },
    { key: 'supervisor', aliases: ['supervisor', 'foreman'] },
    { key: 'event_date', aliases: ['event date', 'date'], type: 'date' },
    { key: 'notes', aliases: ['notes'] }
  ],
  hoses: [
    { key: 'hose_no', aliases: ['hose no', 'hose number'] },
    { key: 'hose_type', aliases: ['hose type', 'type'], fallback: 'Hydraulic Hose' },
    { key: 'size', aliases: ['size'] },
    { key: 'length', aliases: ['length'] },
    { key: 'fitting_a', aliases: ['fitting a', 'fitting 1'], fallback: 'Straight' },
    { key: 'fitting_b', aliases: ['fitting b', 'fitting 2'], fallback: 'Straight' },
    { key: 'stock_qty', aliases: ['stock qty', 'qty', 'quantity'], type: 'number', fallback: 0 },
    { key: 'min_qty', aliases: ['min qty', 'minimum'], type: 'number', fallback: 1 },
    { key: 'shelf_location', aliases: ['shelf', 'location', 'bin'] },
    { key: 'status', aliases: ['status'], fallback: 'In stock' },
    { key: 'notes', aliases: ['notes'] }
  ],
  hose_requests: [
    { key: 'request_date', aliases: ['request date', 'date'], type: 'date' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'requested_by', aliases: ['requested by', 'requester', 'foreman'] },
    { key: 'hose_type', aliases: ['hose type', 'type'], fallback: 'Hydraulic Hose' },
    { key: 'hose_no', aliases: ['hose no'] },
    { key: 'size', aliases: ['size'] },
    { key: 'length', aliases: ['length'] },
    { key: 'fittings', aliases: ['fittings'] },
    { key: 'qty', aliases: ['qty', 'quantity'], type: 'number', fallback: 1 },
    { key: 'repair_type', aliases: ['repair type'], fallback: 'New hose' },
    { key: 'offsite_repair', aliases: ['offsite repair'], fallback: 'No' },
    { key: 'admin_approval', aliases: ['admin approval', 'approval'], fallback: 'Pending' },
    { key: 'status', aliases: ['status'], fallback: 'Requested' },
    { key: 'eta', aliases: ['eta'], type: 'date' },
    { key: 'reason', aliases: ['reason'] },
    { key: 'notes', aliases: ['notes'] }
  ],
  tools: [
    { key: 'tool_name', aliases: ['tool name', 'tool'] },
    { key: 'category', aliases: ['category'], fallback: 'Hand Tools' },
    { key: 'brand', aliases: ['brand', 'make'] },
    { key: 'model', aliases: ['model'] },
    { key: 'serial_no', aliases: ['serial no', 'serial'] },
    { key: 'stock_qty', aliases: ['stock qty', 'qty', 'quantity'], type: 'number', fallback: 0 },
    { key: 'workshop_section', aliases: ['workshop section', 'section'], fallback: 'Mechanical' },
    { key: 'condition', aliases: ['condition'], fallback: 'Good' },
    { key: 'status', aliases: ['status'], fallback: 'In stock' },
    { key: 'notes', aliases: ['notes'] }
  ],
  tool_orders: [
    { key: 'request_date', aliases: ['request date', 'date'], type: 'date' },
    { key: 'tool_name', aliases: ['tool name', 'tool'] },
    { key: 'category', aliases: ['category'], fallback: 'Hand Tools' },
    { key: 'brand', aliases: ['brand'] },
    { key: 'qty', aliases: ['qty', 'quantity'], type: 'number', fallback: 1 },
    { key: 'requested_by', aliases: ['requested by', 'requester'] },
    { key: 'reason', aliases: ['reason'] },
    { key: 'priority', aliases: ['priority'], fallback: 'Normal' },
    { key: 'status', aliases: ['status'], fallback: 'Requested' },
    { key: 'eta', aliases: ['eta'], type: 'date' },
    { key: 'notes', aliases: ['notes'] }
  ],
  employee_tools: [
    { key: 'employee_name', aliases: ['employee name', 'employee', 'name'] },
    { key: 'employee_no', aliases: ['employee no', 'clock no'] },
    { key: 'tool_name', aliases: ['tool name', 'tool'] },
    { key: 'brand', aliases: ['brand'] },
    { key: 'serial_no', aliases: ['serial no', 'serial'] },
    { key: 'issue_date', aliases: ['issue date', 'date issued'], type: 'date' },
    { key: 'issued_by', aliases: ['issued by'] },
    { key: 'condition_issued', aliases: ['condition issued'], fallback: 'Good' },
    { key: 'return_date', aliases: ['return date'], type: 'date' },
    { key: 'condition_returned', aliases: ['condition returned'] },
    { key: 'status', aliases: ['status'], fallback: 'Issued' },
    { key: 'notes', aliases: ['notes'] }
  ],
  fabrication_stock: [
    { key: 'department', aliases: ['department'], fallback: 'Boiler Shop' },
    { key: 'material_type', aliases: ['material type', 'material'], fallback: 'Steel Plate' },
    { key: 'description', aliases: ['description', 'item'] },
    { key: 'size_spec', aliases: ['size spec', 'size'] },
    { key: 'stock_qty', aliases: ['stock qty', 'qty', 'quantity'], type: 'number', fallback: 0 },
    { key: 'min_qty', aliases: ['min qty', 'minimum'], type: 'number', fallback: 1 },
    { key: 'unit', aliases: ['unit'], fallback: 'pcs' },
    { key: 'supplier', aliases: ['supplier'] },
    { key: 'location', aliases: ['location'] },
    { key: 'status', aliases: ['status'], fallback: 'In stock' },
    { key: 'notes', aliases: ['notes'] }
  ],
  fabrication_requests: [
    { key: 'request_date', aliases: ['request date', 'date'], type: 'date' },
    { key: 'department', aliases: ['department'], fallback: 'Boiler Shop' },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'requested_by', aliases: ['requested by', 'requester', 'foreman'] },
    { key: 'material_type', aliases: ['material type', 'material'] },
    { key: 'description', aliases: ['description', 'item'] },
    { key: 'qty', aliases: ['qty', 'quantity'], type: 'number', fallback: 1 },
    { key: 'project_name', aliases: ['project name', 'project'] },
    { key: 'status', aliases: ['status'], fallback: 'Requested' },
    { key: 'eta', aliases: ['eta'], type: 'date' },
    { key: 'reason', aliases: ['reason'] },
    { key: 'notes', aliases: ['notes'] }
  ],
  fabrication_projects: [
    { key: 'department', aliases: ['department'], fallback: 'Boiler Shop' },
    { key: 'project_name', aliases: ['project name', 'project'] },
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'supervisor', aliases: ['supervisor', 'foreman'] },
    { key: 'start_date', aliases: ['start date'], type: 'date' },
    { key: 'target_date', aliases: ['target date'], type: 'date' },
    { key: 'status', aliases: ['status'], fallback: 'Ongoing' },
    { key: 'material_used', aliases: ['material used'] },
    { key: 'progress_percent', aliases: ['progress percent', 'progress'], type: 'number', fallback: 0 },
    { key: 'notes', aliases: ['notes'] }
  ],
  photo_logs: [
    { key: 'fleet_no', aliases: ['fleet', 'fleet no', 'machine'] },
    { key: 'linked_type', aliases: ['linked type', 'type'], fallback: 'Photo' },
    { key: 'caption', aliases: ['caption', 'description'] },
    { key: 'notes', aliases: ['notes'] }
  ]
}

const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
const today = () => new Date().toISOString().slice(0, 10)
const nowIso = () => new Date().toISOString()

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

function dateOnly(value: any) {
  if (!value) return ''

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10)
  }

  if (typeof value === 'number' && value > 25000) {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) {
      const mm = String(parsed.m).padStart(2, '0')
      const dd = String(parsed.d).padStart(2, '0')
      return `${parsed.y}-${mm}-${dd}`
    }
  }

  const text = String(value).trim()

  const dotDate = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/)
  if (dotDate) {
    const dd = dotDate[1].padStart(2, '0')
    const mm = dotDate[2].padStart(2, '0')
    const yyyy = dotDate[3].length === 2 ? `20${dotDate[3]}` : dotDate[3]
    return `${yyyy}-${mm}-${dd}`
  }

  const d = new Date(text)
  if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10)

  return text.slice(0, 10)
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
  const extra = aliases.flatMap((alias) => {
    const a = normalizeKey(alias)
    const more: string[] = []

    if (a.includes('fleet')) more.push('fleet', 'fleet no', 'fleet number', 'machine', 'machine no', 'unit', 'unit no', 'equipment', 'asset', 'plant no')
    if (a.includes('machine')) more.push('machine', 'machine no', 'fleet', 'fleet no', 'unit', 'equipment', 'asset')
    if (a.includes('employee')) more.push('employee', 'employee name', 'emp no', 'clock no', 'staff no', 'worker', 'person', 'name')
    if (a.includes('request')) more.push('requested by', 'requester', 'foreman', 'supervisor', 'raised by', 'ordered by')
    if (a.includes('qty')) more.push('qty', 'quantity', 'number', 'amount', 'stock', 'stock qty', 'balance')
    if (a.includes('part')) more.push('part no', 'part number', 'spare no', 'spare number', 'item no', 'stock code', 'material code')
    if (a.includes('serial')) more.push('serial', 'serial no', 'serial number', 's/n', 'sn')
    if (a.includes('date')) more.push('date', 'request date', 'fitted date', 'fitment date', 'production date', 'due date', 'start date', 'end date')
    if (a.includes('status')) more.push('status', 'state', 'stage')
    if (a.includes('role')) more.push('role', 'trade', 'position', 'job title', 'designation')
    if (a.includes('section')) more.push('section', 'department', 'dept', 'area')
    if (a.includes('hours')) more.push('hours', 'hrs', 'hm', 'hour meter', 'meter reading')
    if (a.includes('mileage')) more.push('mileage', 'km', 'odometer')

    return [alias, ...more]
  })

  return exactCell(row, extra) || fuzzyCell(row, extra)
}

function convertValue(row: Row, field: FieldDef) {
  const raw = getCell(row, field.aliases)
  const value = raw === '' || raw === undefined || raw === null ? field.fallback : raw

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

function mapRow(table: TableName, raw: Row) {
  const schema = SCHEMAS[table]
  const mapped: Row = {}

  schema.forEach((field) => {
    mapped[field.key] = convertValue(raw, field)
  })

  if (table === 'leave_records' && !mapped.days) {
    mapped.days = daysInclusive(mapped.start_date, mapped.end_date)
  }

  if (table === 'spares_orders') {
    if (!mapped.request_date) mapped.request_date = today()
    if (!mapped.workflow_stage) mapped.workflow_stage = 'Requested'
    if (!mapped.status) mapped.status = mapped.workflow_stage
  }

  if (table === 'tyre_measurements' && !mapped.measurement_date) mapped.measurement_date = today()
  if (table === 'tyre_orders' && !mapped.request_date) mapped.request_date = today()
  if (table === 'battery_orders' && !mapped.request_date) mapped.request_date = today()
  if (table === 'hose_requests' && !mapped.request_date) mapped.request_date = today()
  if (table === 'tool_orders' && !mapped.request_date) mapped.request_date = today()
  if (table === 'employee_tools' && !mapped.issue_date) mapped.issue_date = today()
  if (table === 'fabrication_requests' && !mapped.request_date) mapped.request_date = today()
  if (table === 'fabrication_projects' && !mapped.start_date) mapped.start_date = today()
  if (table === 'operations' && !mapped.event_date) mapped.event_date = today()

  mapped.id = makeId(table, mapped)

  return cleanRow(mapped)
}

function scoreHeaderRow(row: any[]) {
  const headingWords = [
    'fleet', 'machine', 'unit', 'equipment',
    'firstname', 'first name', 'surname', 'employee', 'code',
    'occupation', 'current job title', 'department', 'shift',
    'dates', 'total', 'leave type',
    'part no', 'description', 'qty', 'serial', 'status',
    'requested by', 'supervisor', 'foreman'
  ]

  const cells = row.map((item) => String(item || '').trim()).filter(Boolean)

  let score = 0

  cells.forEach((item) => {
    const normal = normalizeKey(item)

    headingWords.forEach((word) => {
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
        allRows.push({
          ...obj,
          _sheet_name: sheetName
        })
      }
    })
  })

  return allRows
}

function parseDateRange(value: any) {
  const text = String(value || '').trim()
  const matches = text.match(/\d{1,4}[./-]\d{1,2}[./-]\d{1,4}/g) || []

  if (matches.length >= 2) {
    return [dateOnly(matches[0]), dateOnly(matches[1])]
  }

  return ['', '']
}

function fullNameFromRaw(raw: Row) {
  const firstName = getCell(raw, ['FirstName', 'First Name', 'first_name', 'Forename'])
  const surname = getCell(raw, ['Surname', 'Last Name', 'last_name'])

  return `${String(firstName || '').trim()} ${String(surname || '').trim()}`.trim()
}

function polishMappedRow(table: TableName, raw: Row, mapped: Row) {
  const next = { ...mapped }

  if (table === 'personnel') {
    const fullName = fullNameFromRaw(raw)

    if (!next.name && fullName) next.name = fullName
    if (!next.employee_no) next.employee_no = getCell(raw, ['Code', 'Employee Code', 'Clock No', 'Emp No'])
    if (!next.role || next.role === 'Diesel Plant Fitter') {
      next.role = getCell(raw, ['Occupation', 'Current Job Title', 'Job Title', 'Trade']) || next.role
    }
    if (!next.section || next.section === 'Workshop') {
      next.section = getCell(raw, ['Department', 'Dept', 'Section']) || next.section
    }
    if (!next.employment_status) next.employment_status = 'Active'
    if (!next.shift) next.shift = 'Shift 1'
    if (!next.staff_group) next.staff_group = 'Workshop Staff'

    next.id = next.employee_no || next.name || next.id || uid()
  }

  if (table === 'leave_records') {
    const fullName = fullNameFromRaw(raw)
    const range = getCell(raw, ['DATES', 'Dates', 'Date Range', 'Leave Dates'])
    const [startDate, endDate] = parseDateRange(range)

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

    if (!next.status) next.status = 'Approved'
    next.id = next.id || uid()
  }

  return cleanRow(next)
}

function hasUsefulData(item: ImportItem, row: Row) {
  return item.qualityKeys.some((key) => {
    const value = row[key]
    return value !== undefined && value !== null && String(value).trim() !== ''
  })
}

function Badge({ value }: { value: string }) {
  const v = String(value || 'Ready')
  const low = v.toLowerCase()
  const cls = low.includes('error') || low.includes('failed') ? 'danger' : low.includes('uploaded') || low.includes('success') ? 'good' : 'neutral'

  return <span className={`badge ${cls}`}>{v}</span>
}

export default function ImporterPage() {
  const [selected, setSelected] = useState<TableName>('fleet_machines')
  const [status, setStatus] = useState('Ready')
  const [busy, setBusy] = useState(false)
  const [previewRows, setPreviewRows] = useState<Row[]>([])
  const [rawHeadings, setRawHeadings] = useState<string[]>([])
  const [fileName, setFileName] = useState('')
  const [search, setSearch] = useState('')

  const selectedImport = IMPORTS.find((item) => item.table === selected) || IMPORTS[0]

  const filteredImports = useMemo(() => {
    const q = search.toLowerCase()
    return IMPORTS.filter((item) => `${item.label} ${item.description} ${item.table}`.toLowerCase().includes(q))
  }, [search])

  async function insertRows(table: TableName, rows: Row[], item: ImportItem) {
    if (!supabase || !isSupabaseConfigured) {
      throw new Error('Supabase is not configured. Check NEXT_PUBLIC_SUPABASE_URL and NEXT_PUBLIC_SUPABASE_ANON_KEY.')
    }

    let uploaded = 0
    const chunkSize = 100

    for (let i = 0; i < rows.length; i += chunkSize) {
      const chunk = rows.slice(i, i + chunkSize).map((row) => ({
        ...row,
        created_at: row.created_at || nowIso(),
        updated_at: nowIso()
      }))

      const request =
        item.mode === 'upsert'
          ? supabase.from(table as any).upsert(chunk, { onConflict: item.conflictKey })
          : supabase.from(table as any).insert(chunk)

      const { error } = await request

      if (error) throw error

      uploaded += chunk.length
      setStatus(`Uploaded ${uploaded}/${rows.length} records to ${item.label}...`)
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
        .map((row) => polishMappedRow(selected, row, mapRow(selected, row)))
        .filter((row) => hasUsefulData(selectedImport, row))

      setPreviewRows(mapped.slice(0, 10))

      if (!mapped.length) {
        setStatus(`No usable rows found for ${selectedImport.label}. Try another import type or check headings.`)
        return
      }

      await insertRows(selected, mapped, selectedImport)

      setStatus(`Uploaded ${mapped.length} records to ${selectedImport.label}.`)
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
          --grey: #d8e1eb;
          --card: rgba(9, 31, 58, .92);
          --line: rgba(159, 230, 255, .22);
          --text: #f4f8ff;
          --muted: #a9bad1;
        }

        body {
          margin: 0;
          background:
            radial-gradient(circle at top left, rgba(255,122,26,.20), transparent 34%),
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
          box-shadow: inset 0 0 0 3px rgba(255,122,26,.20);
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
          font-weight: 800;
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
        .upload-box input,
        select {
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

        select option {
          background: #071a33;
          color: white;
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
          background: rgba(255,122,26,.14);
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
          font-weight: 800;
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
          font-weight: 800;
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
          position: sticky;
          top: 0;
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
            <p>Universal Excel/CSV upload center for every workshop page. Works with different heading names, merged-top sheets and multiple sheets.</p>
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
              <h2>{selectedImport.label}</h2>
              <p>{selectedImport.description}</p>

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
                  Accepts headings like Fleet Number, Machine, Unit, Qty, Quantity, Requested By,
                  Foreman, FirstName, Surname, Code, Occupation, Department, Dates, Total and Leave Type.
                </p>
              </div>

              <div className="status">{status}</div>
            </div>

            <div className="meta">
              <div>
                <span>Selected</span>
                <b>{selectedImport.label}</b>
              </div>

              <div>
                <span>Database table</span>
                <b>{selectedImport.table}</b>
              </div>

              <div>
                <span>Mode</span>
                <b>{selectedImport.mode}</b>
              </div>

              <div>
                <span>File</span>
                <b>{fileName || 'None'}</b>
              </div>
            </div>
          </div>

          <div className="help-grid">
            <div className="card">
              <b>Finds headings automatically</b>
              <p>It checks the whole sheet and finds the best heading row, even when headings are on row 4.</p>
            </div>

            <div className="card">
              <b>Personnel names</b>
              <p>It joins FirstName + Surname and reads Code, Occupation and Department.</p>
            </div>

            <div className="card">
              <b>Leave dates</b>
              <p>It reads DATES like 23.04.2026 - 14.05.2026 and splits start/end automatically.</p>
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
