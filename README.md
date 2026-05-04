# Turbo Energy Workshop Control App

This is a complete starter workshop management app for:

- Fleet register upload from Excel/CSV
- Breakdown tracking
- Repairs and job cards
- Service schedule upload and reminders
- Spares stock availability
- Spares order requests
- Personnel sections, leave and hours worked
- Tyre serial number register and replacement reminders
- Battery serial number register and replacement reminders
- Photo uploads from phone or laptop
- Reports with print/save PDF
- Admin/user login gate
- Offline local saving with sync queue

Theme: navy blue, orange, grey, white and light blue.

---

## Starter logins

Use these after the app opens:

- `admin` / `admin123`
- `admin workshop` / `workshop123`
- `foreman` / `foreman123`
- `user` / `user123`

Admin can upload fleet and service Excel files. Users can add breakdowns, work logs, photos and workshop records.

---

## Supabase setup

1. Open Supabase.
2. Open your project.
3. Go to **SQL Editor**.
4. Click **New query**.
5. Open this project file:

   `database/schema.sql`

6. Copy everything from `schema.sql`.
7. Paste into Supabase SQL Editor.
8. Click **Run**.

This creates all required tables and prototype RLS policies.

---

## Environment variables for Vercel

In Supabase:

1. Go to **Project Settings**.
2. Go to **API**.
3. Copy:
   - Project URL
   - anon/public key

In Vercel:

1. Open your project.
2. Go to **Settings**.
3. Go to **Environment Variables**.
4. Add:

```env
NEXT_PUBLIC_SUPABASE_URL=https://your-project-ref.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=your-anon-public-key
```

5. Save.
6. Redeploy.

---

## GitHub upload steps

1. Extract this zip file.
2. Open GitHub.
3. Create a new repository, for example:

   `turbo-energy-workshop-control`

4. Click **Add file**.
5. Click **Upload files**.
6. Upload all files and folders from inside the extracted folder.
7. Commit message:

```text
Initial workshop fleet management app
```

8. Commit changes.

---

## Vercel deploy steps

1. Open Vercel.
2. Click **Add New Project**.
3. Import your GitHub repo.
4. Framework should show **Next.js**.
5. Add the Supabase environment variables.
6. Click **Deploy**.

---

## Excel upload columns

The app accepts flexible headings. These examples work.

### Fleet register

```csv
fleet_no,machine_type,make_model,reg_no,department,location,hours,mileage,status,service_interval_hours,last_service_hours,notes
```

Also accepted: `fleet`, `fleet no`, `machine no`, `unit`, `type`, `machine type`, `hmr`, `current hours`, `section`, `site`, `remarks`.

### Service schedule

```csv
fleet_no,service_type,scheduled_hours,due_date,status,technician,notes
```

Also accepted: `fleet`, `machine`, `service`, `due hours`, `next service hours`, `service date`, `fitter`, `mechanic`.

Sample templates are included in the `samples` folder.

---

## Offline behavior

- When internet is off, records are saved into the browser offline queue.
- When internet returns, click **Admin > Sync Offline Queue**.
- A local snapshot is kept on the device, so the app still opens on phone/laptop even when signal is poor.

Important: offline data is stored per device/browser until it syncs.

---

## Notes for the next upgrade

This version is designed to get you live quickly. Recommended next upgrades:

1. Supabase Auth for proper company security.
2. Supabase Storage for heavy photo uploading.
3. Edit buttons for all records.
4. WhatsApp notifications for breakdowns, services, low spares, tyre reminders and battery reminders.
5. Role permissions by Admin, Workshop Manager, Foreman, Chargehand and Stores.
