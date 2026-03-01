// src/components/VersionLabelModal.tsx
import { useState } from 'react'
import { Tag, Loader2, AlertTriangle, AlertCircle, ChevronRight } from 'lucide-react'
import type { Script } from '../types'
import { assignScriptToProject, type ProjectResponse } from '../api/projectApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  script: Script
  projects: ProjectResponse[]
  defaultProjectId?: number
  onComplete: (versionName: string, projectId?: number) => void
  onSkip: () => void
}

const SUGGESTED_LABELS = ['Draft 1', 'Draft 2', 'Revised Draft', "Director's Cut", 'Final', 'Locked']

export function VersionLabelModal({ script, projects, defaultProjectId, onComplete, onSkip }: Props) {
  const { user } = useAuth()
  const [versionName, setVersionName] = useState('')
  const [projectId,   setProjectId]   = useState<number | undefined>(defaultProjectId)
  const [saving,      setSaving]      = useState(false)
  const [error,       setError]       = useState<string | null>(null)

  const highCount  = script.risks?.filter(r => r.severity === 'HIGH').length ?? 0
  const medCount   = script.risks?.filter(r => r.severity === 'MEDIUM').length ?? 0
  const totalRisks = script.risks?.length ?? script.riskCount ?? 0

  const handleSave = async () => {
    if (!versionName.trim()) return
    if (!user) return
    setSaving(true)
    setError(null)
    try {
      if (projectId !== undefined) {
        await assignScriptToProject(script.id, projectId, versionName.trim(), user.userId)
      }
      onComplete(versionName.trim(), projectId)
    } catch (e: unknown) {
      const msg = (e as { response?: { data?: { error?: string } } })?.response?.data?.error
      setError(msg ?? 'Failed to save version label.')
    } finally {
      setSaving(false)
    }
  }

  return (
    <>
      <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-[2px] z-40" />
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 pointer-events-none">
        <div className="w-full max-w-md bg-white border border-slate-200 rounded-2xl shadow-xl pointer-events-auto overflow-hidden">

          {/* Header */}
          <div className="p-5 border-b border-slate-100 bg-emerald-50/60">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 rounded-xl bg-emerald-100 border border-emerald-200 flex items-center justify-center">
                <Tag size={15} className="text-emerald-600" />
              </div>
              <div>
                <h2 className="font-display text-lg text-slate-900">Label This Version</h2>
                <p className="text-[11px] text-slate-500">Analysis complete — name this draft before saving</p>
              </div>
            </div>
            <div className="flex items-center gap-3 bg-white rounded-xl p-3 border border-slate-200">
              <div className="flex-1 min-w-0">
                <p className="text-xs text-slate-600 truncate font-medium">{script.filename}</p>
                <p className="text-[11px] text-slate-400 mt-0.5">{script.totalPages} pages analyzed</p>
              </div>
              <div className="flex items-center gap-2">
                {highCount > 0 && (
                  <span className="flex items-center gap-1 text-[10px] text-red-600 bg-red-50 border border-red-200 rounded-full px-2 py-0.5 font-medium">
                    <AlertTriangle size={9} /> {highCount} HIGH
                  </span>
                )}
                {medCount > 0 && (
                  <span className="flex items-center gap-1 text-[10px] text-amber-700 bg-amber-50 border border-amber-200 rounded-full px-2 py-0.5 font-medium">
                    <AlertCircle size={9} /> {medCount} MED
                  </span>
                )}
                <span className="text-[10px] text-slate-400">{totalRisks} total</span>
              </div>
            </div>
          </div>

          <div className="p-5 space-y-4">
            {/* Suggested labels */}
            <div>
              <p className="text-[10px] text-slate-400 mb-2 uppercase tracking-widest font-semibold">Quick Labels</p>
              <div className="flex flex-wrap gap-2">
                {SUGGESTED_LABELS.map(label => (
                  <button key={label} type="button" onClick={() => setVersionName(label)}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-all ${
                      versionName === label
                        ? 'bg-emerald-600 border-emerald-600 text-white font-medium'
                        : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300 hover:text-slate-700'
                    }`}>
                    {label}
                  </button>
                ))}
              </div>
            </div>

            {/* Custom name */}
            <div>
              <label className="block text-[10px] text-slate-500 mb-1.5 uppercase tracking-widest font-semibold">Version Name *</label>
              <input type="text" value={versionName} onChange={e => setVersionName(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleSave()}
                autoFocus
                placeholder="e.g. Draft 1, Director's Cut, Final…"
                className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm rounded-xl px-4 py-2.5
                           focus:outline-none focus:border-emerald-400 focus:ring-2 focus:ring-emerald-100
                           placeholder-slate-300 transition-all" />
            </div>

            {/* Project selector */}
            {projects.length > 0 && (
              <div>
                <label className="block text-[10px] text-slate-500 mb-1.5 uppercase tracking-widest font-semibold">Project</label>
                <select
                  value={projectId ?? ''}
                  onChange={e => setProjectId(e.target.value ? Number(e.target.value) : undefined)}
                  className="w-full bg-slate-50 border border-slate-200 text-slate-700 text-sm rounded-xl px-4 py-2.5
                             focus:outline-none focus:border-emerald-400 focus:ring-2 focus:ring-emerald-100 transition-all">
                  <option value="">— No project —</option>
                  {projects.map(p => (
                    <option key={p.id} value={p.id}>{p.name}{p.studioName ? ` · ${p.studioName}` : ''}</option>
                  ))}
                </select>
              </div>
            )}

            {error && (
              <div className="text-xs text-red-600 bg-red-50 border border-red-200 rounded-xl px-4 py-3">{error}</div>
            )}
          </div>

          {/* Footer */}
          <div className="p-4 border-t border-slate-100 bg-slate-50 flex items-center gap-3">
            <button onClick={handleSave} disabled={!versionName.trim() || saving}
              className="flex-1 flex items-center justify-center gap-2 py-2.5 bg-emerald-600 hover:bg-emerald-500
                         text-white text-sm font-medium rounded-xl transition-all disabled:opacity-40 active:scale-95">
              {saving ? <><Loader2 size={14} className="animate-spin" /> Saving…</> : <>Save Version <ChevronRight size={14} /></>}
            </button>
            <button onClick={onSkip}
              className="px-5 py-2.5 bg-white border border-slate-200 text-slate-500 text-sm rounded-xl hover:bg-slate-50 transition-all">
              Skip
            </button>
          </div>
        </div>
      </div>
    </>
  )
}
