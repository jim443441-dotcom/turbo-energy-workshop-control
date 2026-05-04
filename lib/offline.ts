export type OfflineAction = {
  id: string
  table: string
  type: 'insert' | 'upsert' | 'delete'
  payload: any
  createdAt: string
}

const QUEUE_KEY = 'turbo-workshop-offline-queue'
const SNAPSHOT_KEY = 'turbo-workshop-snapshot'

export function getOfflineQueue(): OfflineAction[] {
  if (typeof window === 'undefined') return []
  try {
    return JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]')
  } catch {
    return []
  }
}

export function setOfflineQueue(actions: OfflineAction[]) {
  if (typeof window === 'undefined') return
  localStorage.setItem(QUEUE_KEY, JSON.stringify(actions))
}

export function addOfflineAction(action: Omit<OfflineAction, 'id' | 'createdAt'>) {
  const queue = getOfflineQueue()
  queue.push({
    ...action,
    id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    createdAt: new Date().toISOString()
  })
  setOfflineQueue(queue)
}

export function saveSnapshot(data: any) {
  if (typeof window === 'undefined') return
  localStorage.setItem(SNAPSHOT_KEY, JSON.stringify({ updatedAt: new Date().toISOString(), data }))
}

export function readSnapshot<T>(): T | null {
  if (typeof window === 'undefined') return null
  try {
    const raw = localStorage.getItem(SNAPSHOT_KEY)
    if (!raw) return null
    return JSON.parse(raw).data as T
  } catch {
    return null
  }
}
