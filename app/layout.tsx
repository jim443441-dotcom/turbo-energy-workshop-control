import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Turbo Energy Workshop Control',
  description: 'Fleet repairs, services, spares, personnel, tyres, batteries and workshop records'
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
