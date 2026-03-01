// src/components/RiskDrawer.tsx
import { useState, useEffect } from 'react'
import {
  X, AlertTriangle, AlertCircle, Info,
  FileText, Lightbulb, Lock, Save, Loader2
} from 'lucide-react'
import type { RiskFlag, ClearanceStatus } from '../types'
import { updateRisk } from '../api/api'
import { StatusSelect } from './StatusSelect'
import { CommentsPanel } from './CommentsPanel'

interface Props {
  risk: RiskFlag | null
  onClose: () => void
  onUpdated: (updated: RiskFlag) => void
  projectId?: number
}

const SEVERITY_COLORS = {
  HIGH:   { bg: 'bg-red-50',     border: 'border-red-200',    text: 'text-red-600',    icon: AlertTriangle },
  MEDIUM: { bg: 'bg-amber-50',   border: 'border-amber-200',  text: 'text-amber-600',  icon: AlertCircle   },
  LOW:    { bg: 'bg-emerald-50', border: 'border-emerald-200', text: 'text-emerald-600', icon: Info          },
}

export function RiskDrawer({ risk, onClose, onUpdated, projectId }: Props) {
  const [status,       setStatus]       = useState<ClearanceStatus>('PENDING')
  const [restrictions, setRestrictions] = useState('')
  const [saving,       setSaving]       = useState(false)
  const [saved,        setSaved]        = useState(false)

  useEffect(() => {
    if (risk) {
      setStatus(risk.status)
      setRestrictions(risk.restrictions ?? '')
      setSaved(false)
    }
  }, [risk])

  if (!risk) return null

  const sev  = SEVERITY_COLORS[risk.severity]
  const Icon = sev.icon

  const handleSave = async () => {
    setSaving(true)
    try {
      const updated = await updateRisk(risk.id, { status, restrictions })
      onUpdated(updated)
      setSaved(true)
      setTimeout(() => {
        setSaved(false)
        onClose()
      }, 500)
    } catch (e) {
      console.error('Save failed', e)
    } finally {
      setSaving(false)
    }
  }

  return (
    <>
      {/* Backdrop */}
      <div
        className="fixed inset-0 bg-black/50 backdrop-blur-sm z-40 animate-fade-in"
        onClick={onClose}
      />

      {/* Centered Modal */}
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 animate-fade-in">
        <div className="w-full max-w-2xl max-h-[90vh] bg-white rounded-2xl shadow-2xl border border-slate-200 
                        overflow-hidden flex flex-col">

          {/* Header */}
          <div className={`flex items-start justify-between p-6 border-b border-slate-100 ${sev.bg}`}>
            <div className="flex items-start gap-4 flex-1 min-w-0">
              <div className={`p-3 rounded-lg ${sev.bg} border ${sev.border} mt-0.5 flex-shrink-0`}>
                <Icon size={18} className={sev.text} />
              </div>
              <div className="min-w-0">
                <div className="flex items-center gap-2 flex-wrap mb-3">
                  <span className="bg-white border border-slate-300 text-slate-900 text-xs font-mono font-semibold px-3 py-1.5 rounded">
                    {[
                      `Pg ${risk.pageNumber}`,
                      risk.episodeNumber ? `Ep ${risk.episodeNumber}` : null,
                      risk.sceneNumber ? `Scene ${risk.sceneNumber}` : null,
                    ].filter(Boolean).join(' | ')}
                  </span>
                </div>
                <h2 className="font-display text-xl text-slate-900 mb-2">{risk.entityName}</h2>
                <div className="flex items-center gap-2 flex-wrap mb-1">
                  <span className={`px-3 py-1 text-xs font-semibold rounded ${sev.bg} ${sev.text}`}>
                    {risk.severity}
                  </span>
                  <span className="bg-slate-100 text-slate-600 text-xs font-medium px-3 py-1 rounded">
                    {risk.category.replace(/_/g, ' ')}
                  </span>
                </div>
                <p className="text-slate-400 text-sm mt-2">{risk.subCategory.replace(/_/g, ' ')}</p>
              </div>
            </div>
            <button
              onClick={onClose}
              className="p-2 hover:bg-white/50 rounded-lg transition-colors ml-4 flex-shrink-0">
              <X size={20} className="text-slate-600" />
            </button>
          </div>

          {/* Scrollable body */}
          <div className="flex-1 overflow-y-auto p-6 space-y-6">

          {/* Script Excerpt */}
          {risk.snippet && (
            <section>
              <h3 className="flex items-center gap-2 text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">
                <FileText size={12} /> Script Excerpt
              </h3>
              <blockquote className="bg-slate-50 border-l-2 border-emerald-400 px-4 py-3 rounded-r-lg
                                     font-mono text-sm text-slate-600 leading-relaxed">
                "{risk.snippet}"
              </blockquote>
            </section>
          )}

          {/* Legal Reason */}
          <section>
            <h3 className="flex items-center gap-2 text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">
              <AlertTriangle size={12} /> Legal Reason
            </h3>
            <p className="text-slate-700 text-sm leading-relaxed bg-slate-50 rounded-lg p-3 border border-slate-100">
              {risk.reason}
            </p>
          </section>

          {/* AI Suggestion */}
          {risk.suggestion && (
            <section>
              <h3 className="flex items-center gap-2 text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">
                <Lightbulb size={12} /> AI Suggestion
              </h3>
              <p className="text-slate-600 text-sm leading-relaxed bg-emerald-50 border border-emerald-100 rounded-lg p-3">
                {risk.suggestion}
              </p>
            </section>
          )}

          {/* Attorney workspace */}
          <div className="border-t border-slate-100 pt-5">
            <p className="text-xs text-slate-400 mb-4 uppercase tracking-wider font-semibold">
              Attorney Workspace
            </p>

            {/* Clearance Status */}
            <div className="mb-4">
              <label className="block text-xs text-slate-500 font-medium mb-1.5">Clearance Status</label>
              <StatusSelect value={status} onChange={setStatus} />
            </div>

            {/* Team Comments Section */}
            <div className="mb-4">
              <CommentsPanel riskFlagId={risk.id} projectId={projectId} />
            </div>

            {/* Restrictions */}
            <div className="mb-5">
              <label className="flex items-center gap-1.5 text-xs text-slate-500 font-medium mb-1.5">
                <Lock size={11} /> Restrictions
              </label>
              <textarea
                value={restrictions}
                onChange={e => setRestrictions(e.target.value)}
                rows={2}
                placeholder="e.g. 'Approved for US only', 'Must blur in theatrical release'…"
                className="w-full bg-white border border-slate-200 text-slate-700 text-sm
                           rounded-lg px-3 py-2.5 focus:outline-none focus:border-emerald-400
                           focus:ring-2 focus:ring-emerald-100 placeholder-slate-300
                           resize-none transition-colors"
              />
            </div>
          </div>
          </div>

          {/* Footer */}
          <div className="p-6 border-t border-slate-100 bg-slate-50 flex items-center gap-3 justify-end">
            <button onClick={onClose} className="btn-ghost px-6">
              Cancel
            </button>
            <button onClick={handleSave} disabled={saving} className="btn-primary px-6">
              {saving ? (
                <><Loader2 size={15} className="animate-spin" /> Saving…</>
              ) : saved ? (
                <><span>✓</span> Saved</>
              ) : (
                <><Save size={15} /> Save Changes</>
              )}
            </button>
          </div>
        </div>
      </div>
    </>
  )
}
