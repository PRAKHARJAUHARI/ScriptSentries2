// src/components/Timeline.tsx
import { useState, useEffect } from 'react'
import {
  AlertTriangle, AlertCircle, Info, Clock,
  Loader2, Edit3, Check, X, Film, Trash2, User
} from 'lucide-react'
import {
  getProjectTimeline, renameVersion, deleteScript,
  type TimelineEntry, type ProjectTimeline
} from '../api/projectApi'
import { useAuth, canDeleteScript, canRename } from '../context/AuthContext'

interface Props {
  projectId: number
  onOpenScript: (scriptId: number) => void
}

function formatDate(iso: string) {
  return new Date(iso).toLocaleDateString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  })
}

function RiskBadge({ count, severity }: { count: number; severity: 'HIGH' | 'MEDIUM' | 'LOW' }) {
  if (count === 0) return null
  const cfg = {
    HIGH:   'bg-red-50 text-red-600 border-red-200',
    MEDIUM: 'bg-amber-50 text-amber-600 border-amber-200',
    LOW:    'bg-emerald-50 text-emerald-600 border-emerald-200',
  }
  const Icon = severity === 'HIGH' ? AlertTriangle : severity === 'MEDIUM' ? AlertCircle : Info
  return (
    <span className={`inline-flex items-center gap-1 text-[10px] px-2 py-0.5 rounded-full border font-medium ${cfg[severity]}`}>
      <Icon size={9} /> {count} {severity}
    </span>
  )
}

function VersionCard({
  entry, isLast, onOpenScript, onRenamed, onDeleted,
}: {
  entry: TimelineEntry
  isLast: boolean
  onOpenScript: (id: number) => void
  onRenamed: (scriptId: number, name: string) => void
  onDeleted: (scriptId: number) => void
}) {
  const { user }    = useAuth()
  const userCanEdit = user ? canRename(user.role) : false
  const userCanDel  = user ? canDeleteScript(user.role) : false

  const [editing,    setEditing]    = useState(false)
  const [draft,      setDraft]      = useState(entry.versionName)
  const [saving,     setSaving]     = useState(false)
  const [deleting,   setDeleting]   = useState(false)
  const [confirmDel, setConfirmDel] = useState(false)

  const dotColor =
    entry.status === 'COMPLETE' ? 'bg-emerald-500' :
    entry.status === 'FAILED'   ? 'bg-red-400'     : 'bg-amber-400'

  const handleSave = async () => {
    if (!draft.trim() || !user) return
    setSaving(true)
    try {
      await renameVersion(entry.scriptId, draft.trim(), user.userId)
      onRenamed(entry.scriptId, draft.trim())
      setEditing(false)
    } catch { /* ignore */ }
    finally { setSaving(false) }
  }

  const handleDelete = async () => {
    if (!user) return
    setDeleting(true)
    try {
      await deleteScript(entry.scriptId, user.userId)
      onDeleted(entry.scriptId)
    } catch { /* ignore */ }
    finally { setDeleting(false); setConfirmDel(false) }
  }

  return (
    <div className="flex gap-5">
      {/* Timeline spine */}
      <div className="flex flex-col items-center flex-shrink-0 w-8">
        <div className={`w-3 h-3 rounded-full border-2 border-white ring-1 ring-slate-200 ${dotColor} z-10 mt-1.5 flex-shrink-0`} />
        {!isLast && <div className="w-px flex-1 bg-slate-200 mt-1 min-h-[2rem]" />}
      </div>

      {/* Card */}
      <div 
        onClick={() => !editing && !confirmDel && onOpenScript(entry.scriptId)}
        className={`flex-1 mb-6 p-4 rounded-2xl border transition-all duration-200 bg-white group shadow-sm cursor-pointer ${
        entry.highCount > 0
          ? 'border-red-200 hover:border-red-300'
          : 'border-slate-200 hover:border-slate-300 hover:shadow-md'
      }`}>

        {/* Header */}
        <div className="flex items-start justify-between gap-3 mb-3">
          <div className="flex-1 min-w-0">
            {editing ? (
              <div className="flex items-center gap-2">
                <input
                  value={draft}
                  onChange={e => setDraft(e.target.value)}
                  onKeyDown={e => {
                    if (e.key === 'Enter') handleSave()
                    if (e.key === 'Escape') { setEditing(false); setDraft(entry.versionName) }
                  }}
                  autoFocus
                  className="flex-1 bg-white border border-emerald-400 text-slate-800 text-sm
                             rounded-lg px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-300"
                />
                <button onClick={handleSave} disabled={saving}
                  className="p-1.5 bg-emerald-600 hover:bg-emerald-500 rounded-lg transition-colors">
                  {saving
                    ? <Loader2 size={12} className="animate-spin text-white" />
                    : <Check size={12} className="text-white" />}
                </button>
                <button onClick={() => { setEditing(false); setDraft(entry.versionName) }}
                  className="p-1.5 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors">
                  <X size={12} className="text-slate-500" />
                </button>
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <h3 className="font-semibold text-slate-800 text-sm truncate">{draft}</h3>
                {userCanEdit && (
                  <button onClick={(e) => { e.stopPropagation(); setEditing(true) }}
                    className="p-1 hover:bg-slate-100 rounded-lg transition-colors
                               opacity-0 group-hover:opacity-100 text-slate-400 hover:text-slate-600">
                    <Edit3 size={11} />
                  </button>
                )}
              </div>
            )}
            <p className="text-[11px] text-slate-400 mt-0.5 truncate">{entry.filename}</p>
          </div>

          <div className="flex items-center gap-2 flex-shrink-0">
            <span className={`text-[10px] px-2 py-0.5 rounded-full font-medium uppercase tracking-wider ${
              entry.status === 'COMPLETE' ? 'bg-emerald-50 text-emerald-600 border border-emerald-200' :
              entry.status === 'FAILED'   ? 'bg-red-50 text-red-500 border border-red-200' :
                                            'bg-amber-50 text-amber-600 border border-amber-200'
            }`}>
              {entry.status}
            </span>

            {userCanDel && !confirmDel && (
              <button
                onClick={(e) => { e.stopPropagation(); setConfirmDel(true) }}
                className="p-1.5 opacity-0 group-hover:opacity-100 hover:bg-red-50 rounded-lg
                           transition-all text-slate-300 hover:text-red-500">
                <Trash2 size={13} />
              </button>
            )}
          </div>
        </div>

        {/* Delete confirmation */}
        {confirmDel && (
          <div className="flex items-center gap-3 bg-red-50 border border-red-200 rounded-xl px-3 py-2.5 mb-3">
            <AlertTriangle size={13} className="text-red-500 flex-shrink-0" />
            <p className="text-xs text-red-600 flex-1">Delete this version permanently?</p>
            <button onClick={(e) => { e.stopPropagation(); handleDelete() }} disabled={deleting}
              className="text-xs font-medium text-white px-2 py-1 bg-red-500 hover:bg-red-600
                         rounded-lg disabled:opacity-50 transition-colors">
              {deleting ? <Loader2 size={11} className="animate-spin" /> : 'Delete'}
            </button>
            <button onClick={(e) => { e.stopPropagation(); setConfirmDel(false) }}
              className="text-xs text-slate-500 hover:text-slate-700 px-2 py-1 transition-colors">
              Cancel
            </button>
          </div>
        )}

        {/* Risk badges */}
        <div className="flex items-center gap-2 flex-wrap mb-3">
          <RiskBadge count={entry.highCount}   severity="HIGH"   />
          <RiskBadge count={entry.mediumCount} severity="MEDIUM" />
          <RiskBadge count={entry.lowCount}    severity="LOW"    />
          {entry.totalRisks === 0 && (
            <span className="text-[10px] text-slate-400 italic">No risks flagged</span>
          )}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3 text-[11px] text-slate-400">
            <span className="flex items-center gap-1">
              <Clock size={9} /> {formatDate(entry.uploadedAt)}
            </span>
            <span>·</span>
            <span>{entry.totalPages} pages</span>
            {entry.uploadedBy && (
              <>
                <span>·</span>
                <span className="flex items-center gap-1">
                  <User size={9} /> @{entry.uploadedBy.username}
                </span>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

export function Timeline({ projectId, onOpenScript }: Props) {
  const [timeline, setTimeline] = useState<ProjectTimeline | null>(null)
  const [loading,  setLoading]  = useState(true)
  const [error,    setError]    = useState(false)

  useEffect(() => {
    let cancelled = false
    setLoading(true)
    setError(false)
    getProjectTimeline(projectId)
      .then(data => { if (!cancelled) setTimeline(data) })
      .catch(() => { if (!cancelled) setError(true) })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [projectId])

  const handleRenamed = (scriptId: number, newName: string) => {
    setTimeline(prev => prev ? ({
      ...prev,
      versions: prev.versions.map(v => v.scriptId === scriptId ? { ...v, versionName: newName } : v),
    }) : prev)
  }

  const handleDeleted = (scriptId: number) => {
    setTimeline(prev => prev ? ({
      ...prev,
      versions: prev.versions.filter(v => v.scriptId !== scriptId),
      totalVersions: (prev.totalVersions ?? 1) - 1,
    }) : prev)
  }

  if (loading) {
    return (
      <div className="flex items-center justify-center py-14">
        <Loader2 size={18} className="animate-spin text-slate-400" />
      </div>
    )
  }

  if (error) {
    return (
      <div className="text-center py-10 text-red-400 text-sm">
        Failed to load timeline. Check your connection and try again.
      </div>
    )
  }

  if (!timeline || timeline.versions.length === 0) {
    return (
      <div className="text-center py-16 border-2 border-dashed border-slate-200 rounded-2xl bg-white">
        <Film size={32} className="text-slate-300 mx-auto mb-3" />
        <p className="text-slate-500 font-medium text-sm">No script versions yet</p>
        <p className="text-slate-400 text-xs mt-1">
          Click <span className="font-medium text-emerald-600">Upload New Script</span> above to get started
        </p>
      </div>
    )
  }

  return (
    <div className="space-y-0">
      {timeline.versions.map((entry, idx) => (
        <VersionCard
          key={entry.scriptId}
          entry={entry}
          isLast={idx === timeline.versions.length - 1}
          onOpenScript={onOpenScript}
          onRenamed={handleRenamed}
          onDeleted={handleDeleted}
        />
      ))}
    </div>
  )
}
