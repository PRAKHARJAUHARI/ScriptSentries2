// src/components/ManageMembersModal.tsx
import { useState } from 'react'
import { X, UserPlus, Trash2, Loader2, AlertTriangle, Search, Crown } from 'lucide-react'
import {
  addProjectMember, removeProjectMember,
  type ProjectResponse, type MemberResponse, type ProjectRole
} from '../api/projectApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  project: ProjectResponse
  onClose: () => void
  onUpdated: (project: ProjectResponse) => void
}

const ROLES: { value: ProjectRole; label: string; desc: string; color: string }[] = [
  { value: 'ATTORNEY',                label: 'Attorney',            desc: 'Full control',       color: 'text-violet-700 bg-violet-50 border-violet-200' },
  { value: 'ANALYST',                 label: 'Analyst',             desc: 'Upload & clear',     color: 'text-blue-700   bg-blue-50   border-blue-200'   },
  { value: 'MAIN_PRODUCTION_CONTACT', label: 'Production Contact',  desc: 'Monitor progress',   color: 'text-amber-700  bg-amber-50  border-amber-200'  },
  { value: 'PRODUCTION_ASSISTANT',    label: 'Prod. Assistant',     desc: 'View & comment',     color: 'text-slate-600  bg-slate-100 border-slate-200'  },
  { value: 'VIEWER',                  label: 'Viewer',              desc: 'Read-only',           color: 'text-zinc-500   bg-zinc-100  border-zinc-200'   },
]

function roleStyle(r: string) {
  return ROLES.find(x => x.value === r)?.color ?? 'text-slate-500 bg-slate-100 border-slate-200'
}
function roleLabel(r: string) {
  return ROLES.find(x => x.value === r)?.label ?? r.replace(/_/g, ' ')
}

export function ManageMembersModal({ project, onClose, onUpdated }: Props) {
  const { user } = useAuth()

  // Add form state
  const [userId,   setUserId]   = useState('')
  const [newRole,  setNewRole]  = useState<ProjectRole>('ANALYST')
  const [adding,   setAdding]   = useState(false)
  const [addError, setAddError] = useState<string | null>(null)

  // Remove state
  const [removingId, setRemovingId] = useState<number | null>(null)
  const [removeError, setRemoveError] = useState<string | null>(null)

  const handleAdd = async () => {
    const uid = parseInt(userId.trim(), 10)
    if (!uid || isNaN(uid)) { setAddError('Enter a valid numeric User ID.'); return }
    if (!user) return

    setAdding(true)
    setAddError(null)
    try {
      const member = await addProjectMember(project.id, { userId: uid, projectRole: newRole }, user.userId)
      const updated: ProjectResponse = {
        ...project,
        members: [...project.members, member],
      }
      onUpdated(updated)
      setUserId('')
    } catch (e: unknown) {
      setAddError((e as { response?: { data?: { error?: string } } })?.response?.data?.error ?? 'Failed to add member.')
    } finally {
      setAdding(false)
    }
  }

  const handleRemove = async (member: MemberResponse) => {
    if (!user) return
    setRemovingId(member.user.id)
    setRemoveError(null)
    try {
      await removeProjectMember(project.id, member.user.id, user.userId)
      const updated: ProjectResponse = {
        ...project,
        members: project.members.filter(m => m.user.id !== member.user.id),
      }
      onUpdated(updated)
    } catch (e: unknown) {
      setRemoveError((e as { response?: { data?: { error?: string } } })?.response?.data?.error ?? 'Failed to remove member.')
    } finally {
      setRemovingId(null)
    }
  }

  const isCreator = (m: MemberResponse) =>
    project.createdBy && m.user.id === project.createdBy.id

  return (
    <>
      <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-[2px] z-40" onClick={onClose} />
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 pointer-events-none">
        <div className="w-full max-w-lg bg-white border border-slate-200 rounded-2xl shadow-xl
                        pointer-events-auto flex flex-col max-h-[90vh]">

          {/* Header */}
          <div className="flex items-center justify-between p-5 border-b border-slate-100 flex-shrink-0">
            <div>
              <h2 className="font-display text-lg text-slate-900">Manage Team</h2>
              <p className="text-[11px] text-slate-400 mt-0.5 truncate max-w-[280px]">{project.name}</p>
            </div>
            <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-xl transition-colors">
              <X size={16} className="text-slate-400" />
            </button>
          </div>

          <div className="overflow-y-auto flex-1">

            {/* Add member */}
            <div className="p-5 border-b border-slate-100">
              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-3 flex items-center gap-1.5">
                <UserPlus size={11} /> Add Member
              </p>
              <div className="flex gap-2 mb-3">
                <div className="relative flex-1">
                  <Search size={13} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  <input
                    type="number"
                    value={userId}
                    onChange={e => setUserId(e.target.value)}
                    onKeyDown={e => e.key === 'Enter' && handleAdd()}
                    placeholder="User ID (numeric)"
                    className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                               rounded-xl pl-9 pr-4 py-2.5 focus:outline-none focus:border-emerald-400
                               focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all"
                  />
                </div>
                <button onClick={handleAdd} disabled={adding || !userId.trim()}
                  className="inline-flex items-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-500
                             text-white text-sm font-medium rounded-xl transition-all disabled:opacity-40">
                  {adding ? <Loader2 size={13} className="animate-spin" /> : <UserPlus size={13} />}
                  Add
                </button>
              </div>

              {/* Role selector */}
              <div className="grid grid-cols-5 gap-1.5">
                {ROLES.map(r => (
                  <button key={r.value} type="button" onClick={() => setNewRole(r.value)}
                    className={`text-center p-2 rounded-xl border text-left transition-all ${
                      newRole === r.value
                        ? r.color + ' ring-2 ring-offset-1 ring-emerald-300'
                        : 'bg-white border-slate-200 hover:border-slate-300'
                    }`}>
                    <p className="text-[10px] font-bold text-slate-700 leading-tight">{r.label}</p>
                    <p className="text-[9px] text-slate-400 mt-0.5">{r.desc}</p>
                  </button>
                ))}
              </div>

              {addError && (
                <div className="flex items-center gap-2 mt-3 bg-red-50 border border-red-200 rounded-xl px-3 py-2">
                  <AlertTriangle size={12} className="text-red-500 flex-shrink-0" />
                  <p className="text-xs text-red-600">{addError}</p>
                </div>
              )}
            </div>

            {/* Current members */}
            <div className="p-5">
              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-3">
                Current Members ({project.members.length})
              </p>

              {removeError && (
                <div className="flex items-center gap-2 mb-3 bg-red-50 border border-red-200 rounded-xl px-3 py-2">
                  <AlertTriangle size={12} className="text-red-500 flex-shrink-0" />
                  <p className="text-xs text-red-600">{removeError}</p>
                </div>
              )}

              <div className="space-y-2">
                {project.members.map(m => {
                  const creator = isCreator(m)
                  const isSelf  = user?.userId === m.user.id
                  const canRemove = !creator && !isSelf

                  return (
                    <div key={m.id}
                      className="flex items-center gap-3 bg-slate-50 border border-slate-100 rounded-xl px-3 py-2.5">
                      <div className="w-8 h-8 rounded-full bg-emerald-100 border border-emerald-200
                                      flex items-center justify-center flex-shrink-0">
                        <span className="text-[11px] font-bold text-emerald-700">
                          {m.user.username.charAt(0).toUpperCase()}
                        </span>
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-1.5">
                          <p className="text-xs font-semibold text-slate-700 truncate">@{m.user.username}</p>
                          {creator && <Crown size={10} className="text-amber-500 flex-shrink-0" aria-label="Project creator" />}
                          {isSelf && <span className="text-[9px] text-slate-400">(you)</span>}
                        </div>
                        <p className="text-[10px] text-slate-400 truncate">{m.user.email}</p>
                      </div>
                      <span className={`text-[9px] px-1.5 py-0.5 rounded border font-semibold flex-shrink-0 ${roleStyle(m.projectRole)}`}>
                        {roleLabel(m.projectRole)}
                      </span>
                      {canRemove ? (
                        <button
                          onClick={() => handleRemove(m)}
                          disabled={removingId === m.user.id}
                          className="p-1.5 hover:bg-red-50 hover:text-red-500 text-slate-300
                                     rounded-lg transition-colors disabled:opacity-40 flex-shrink-0">
                          {removingId === m.user.id
                            ? <Loader2 size={13} className="animate-spin" />
                            : <Trash2 size={13} />
                          }
                        </button>
                      ) : (
                        <div className="w-7 flex-shrink-0" />
                      )}
                    </div>
                  )
                })}
              </div>
            </div>
          </div>

          <div className="p-4 border-t border-slate-100 bg-slate-50 flex-shrink-0">
            <button onClick={onClose} className="w-full py-2 text-sm text-slate-500 hover:text-slate-800 transition-colors">
              Done
            </button>
          </div>
        </div>
      </div>
    </>
  )
}