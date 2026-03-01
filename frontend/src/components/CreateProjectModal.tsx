// src/components/CreateProjectModal.tsx
import { useState } from 'react'
import { X, Film, Loader2, AlertTriangle } from 'lucide-react'
import { createProject, type ProjectResponse } from '../api/projectApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  onClose: () => void
  onCreated: (project: ProjectResponse) => void
}

export function CreateProjectModal({ onClose, onCreated }: Props) {
  const { user } = useAuth()
  const [name,     setName]     = useState('')
  const [studio,   setStudio]   = useState('')
  const [director, setDirector] = useState('')
  const [genre,    setGenre]    = useState('')
  const [logline,  setLogline]  = useState('')
  const [saving,   setSaving]   = useState(false)
  const [error,    setError]    = useState<string | null>(null)

  const handleCreate = async () => {
    if (!name.trim()) { setError('Project name is required.'); return }
    if (!user) return
    setSaving(true)
    setError(null)
    try {
      const project = await createProject({
        name: name.trim(),
        studioName: studio.trim() || undefined,
        director:   director.trim() || undefined,
        genre:      genre.trim() || undefined,
        logline:    logline.trim() || undefined,
      }, user.userId)
      onCreated(project)
    } catch (e: unknown) {
      setError((e as { response?: { data?: { error?: string } } })?.response?.data?.error ?? 'Failed to create project.')
    } finally {
      setSaving(false)
    }
  }

  return (
    <>
      <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-[2px] z-40" onClick={onClose} />
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 pointer-events-none">
        <div className="w-full max-w-md bg-white border border-slate-200 rounded-2xl
                        shadow-xl pointer-events-auto overflow-hidden">

          <div className="flex items-center justify-between p-5 border-b border-slate-100">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-center">
                <Film size={16} className="text-emerald-600" />
              </div>
              <h2 className="font-display text-lg text-slate-900">New Project</h2>
            </div>
            <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-xl transition-colors">
              <X size={16} className="text-slate-400" />
            </button>
          </div>

          <div className="p-5 space-y-4">
            <div>
              <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">
                Project Name *
              </label>
              <input type="text" value={name} onChange={e => setName(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleCreate()}
                placeholder="e.g. Untitled Feature Film"
                autoFocus
                className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                           rounded-xl px-4 py-2.5 focus:outline-none focus:border-emerald-400
                           focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all" />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">Studio</label>
                <input type="text" value={studio} onChange={e => setStudio(e.target.value)}
                  placeholder="Studio name"
                  className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                             rounded-xl px-4 py-2.5 focus:outline-none focus:border-emerald-400
                             focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all" />
              </div>
              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">Genre</label>
                <input type="text" value={genre} onChange={e => setGenre(e.target.value)}
                  placeholder="e.g. Thriller"
                  className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                             rounded-xl px-4 py-2.5 focus:outline-none focus:border-emerald-400
                             focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all" />
              </div>
            </div>

            <div>
              <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">Director</label>
              <input type="text" value={director} onChange={e => setDirector(e.target.value)}
                placeholder="Director's name"
                className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                           rounded-xl px-4 py-2.5 focus:outline-none focus:border-emerald-400
                           focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all" />
            </div>

            <div>
              <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">Logline</label>
              <textarea value={logline} onChange={e => setLogline(e.target.value)}
                rows={2} placeholder="One-line description..."
                className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                           rounded-xl px-4 py-2.5 focus:outline-none focus:border-emerald-400
                           focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 resize-none transition-all" />
            </div>

            {error && (
              <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-4 py-3">
                <AlertTriangle size={13} className="text-red-500 flex-shrink-0" />
                <p className="text-xs text-red-600">{error}</p>
              </div>
            )}
          </div>

          <div className="p-4 border-t border-slate-100 bg-slate-50 flex gap-3">
            <button onClick={handleCreate} disabled={saving || !name.trim()}
              className="flex-1 inline-flex items-center justify-center gap-2 bg-emerald-600
                         hover:bg-emerald-500 text-white font-medium py-2.5 rounded-xl text-sm
                         transition-all disabled:opacity-40 active:scale-95">
              {saving ? <><Loader2 size={14} className="animate-spin" /> Creating…</> : <><Film size={14} /> Create Project</>}
            </button>
            <button onClick={onClose} className="px-5 py-2.5 text-sm text-slate-500 hover:text-slate-800
                                                  border border-slate-200 bg-white rounded-xl transition-colors">
              Cancel
            </button>
          </div>
        </div>
      </div>
    </>
  )
}