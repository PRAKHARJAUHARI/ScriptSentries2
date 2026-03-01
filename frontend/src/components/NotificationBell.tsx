// src/components/NotificationBell.tsx
import { useState, useEffect, useRef } from 'react'
import { Bell, Check, Loader2 } from 'lucide-react'
import { getNotifications, markAllRead, type NotificationResponse } from '../api/authApi'
import { useAuth } from '../context/AuthContext'

function formatTime(iso: string) {
  const diff = Math.floor((Date.now() - new Date(iso).getTime()) / 1000)
  if (diff < 60)    return 'just now'
  if (diff < 3600)  return `${Math.floor(diff / 60)}m ago`
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`
  return new Date(iso).toLocaleDateString()
}

interface Props {
  onNavigateToRisk?: (riskFlagId: number) => void
}

export function NotificationBell({ onNavigateToRisk }: Props) {
  const { user } = useAuth()
  const [open,          setOpen]          = useState(false)
  const [notifications, setNotifications] = useState<NotificationResponse[]>([])
  const [loading,       setLoading]       = useState(false)
  const panelRef = useRef<HTMLDivElement>(null)

  const unread = notifications.filter(n => !n.isRead).length

  const load = async () => {
    if (!user) return
    setLoading(true)
    try {
      setNotifications(await getNotifications(user.userId))
    } catch { /* ignore */ }
    finally { setLoading(false) }
  }

  // Initial load
  useEffect(() => { load() }, [user])

  // Reload when panel opens
  useEffect(() => { if (open) load() }, [open])

  // Poll every 30 s
  useEffect(() => {
    if (!user) return
    const id = setInterval(async () => {
      setNotifications(await getNotifications(user.userId).catch(() => []))
    }, 30_000)
    return () => clearInterval(id)
  }, [user])

  // Close on outside click
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (panelRef.current && !panelRef.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  const handleMarkRead = async () => {
    if (!user) return
    try {
      await markAllRead(user.userId)
      setNotifications(prev => prev.map(n => ({ ...n, isRead: true })))
    } catch { /* ignore */ }
  }

  if (!user) return null

  return (
    <div className="relative" ref={panelRef}>
      {/* Bell button */}
      <button
        onClick={() => setOpen(o => !o)}
        className="relative p-2 rounded-lg hover:bg-slate-100 transition-colors"
        aria-label="Notifications"
      >
        <Bell size={17} className={unread > 0 ? 'text-emerald-600' : 'text-slate-400'} />
        {unread > 0 && (
          <span className="absolute -top-0.5 -right-0.5 w-4 h-4 bg-red-500 rounded-full
                           text-[9px] font-bold text-white flex items-center justify-center
                           animate-pulse">
            {unread > 9 ? '9+' : unread}
          </span>
        )}
      </button>

      {/* Panel */}
      {open && (
        <div className="absolute right-0 top-full mt-2 w-80 bg-white border border-slate-200
                        rounded-2xl shadow-xl z-50 overflow-hidden animate-fade-in">

          {/* Header */}
          <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
            <div className="flex items-center gap-2">
              <Bell size={14} className="text-emerald-600" />
              <span className="text-sm font-semibold text-slate-800">Notifications</span>
              {unread > 0 && (
                <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-emerald-50
                                 text-emerald-700 border border-emerald-200 font-medium">
                  {unread} new
                </span>
              )}
            </div>
            {unread > 0 && (
              <button
                onClick={handleMarkRead}
                className="flex items-center gap-1 text-xs text-slate-400
                           hover:text-emerald-600 transition-colors font-medium"
              >
                <Check size={11} /> Mark all read
              </button>
            )}
          </div>

          {/* List */}
          <div className="max-h-72 overflow-y-auto">
            {loading ? (
              <div className="flex items-center justify-center py-10">
                <Loader2 size={16} className="animate-spin text-slate-300" />
              </div>
            ) : notifications.length === 0 ? (
              <div className="text-center py-10 text-slate-400 text-sm">
                <Bell size={24} className="mx-auto mb-2 text-slate-200" />
                No notifications yet
              </div>
            ) : (
              notifications.map(n => (
                <button
                  key={n.id}
                  onClick={() => {
                    if (n.riskFlagId && onNavigateToRisk) onNavigateToRisk(n.riskFlagId)
                    setOpen(false)
                  }}
                  className={`w-full text-left px-4 py-3 border-b border-slate-50
                              hover:bg-slate-50 transition-colors last:border-b-0 ${
                                !n.isRead ? 'bg-emerald-50/60' : ''
                              }`}
                >
                  <div className="flex items-start gap-2.5">
                    {/* Unread dot */}
                    <div className={`w-1.5 h-1.5 rounded-full mt-1.5 flex-shrink-0 ${
                      !n.isRead ? 'bg-emerald-500' : 'bg-transparent'
                    }`} />
                    <div className="flex-1 min-w-0">
                      <p className="text-xs text-slate-700 leading-relaxed">{n.message}</p>
                      <p className="text-[10px] text-slate-400 mt-1">{formatTime(n.createdAt)}</p>
                    </div>
                  </div>
                </button>
              ))
            )}
          </div>

          {notifications.length > 0 && (
            <div className="px-4 py-2.5 border-t border-slate-100 bg-slate-50">
              <p className="text-[10px] text-slate-400 text-center">
                Click a notification to jump to the risk flag
              </p>
            </div>
          )}
        </div>
      )}
    </div>
  )
}