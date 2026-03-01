// src/components/ManageTeamModal.tsx
import { useState, useRef, useEffect } from 'react'
import { X, Loader2, Plus, Trash2, Users } from 'lucide-react'
import { addProjectMember, removeProjectMember, searchUsers, type MemberResponse, type ProjectRole, type UserSummary } from '../api/projectApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  projectId: number
  currentMembers: MemberResponse[]
  createdBy: UserSummary | null
  onClose: () => void
  onMemberAdded: (member: MemberResponse) => void
  onMemberRemoved: (userId: number) => void
}

const ROLE_COLORS: Record<string, string> = {
  ATTORNEY:                'text-violet-700 bg-violet-50 border-violet-200',
  ANALYST:                 'text-blue-700   bg-blue-50   border-blue-200',
  MAIN_PRODUCTION_CONTACT: 'text-amber-700  bg-amber-50  border-amber-200',
  PRODUCTION_ASSISTANT:    'text-slate-600  bg-slate-100 border-slate-200',
  VIEWER:                  'text-zinc-500   bg-zinc-100  border-zinc-200',
}

const PROJECT_ROLES: ProjectRole[] = [
  'ATTORNEY',
  'ANALYST',
  'MAIN_PRODUCTION_CONTACT',
  'PRODUCTION_ASSISTANT',
  'VIEWER',
]

export function ManageTeamModal({
  projectId,
  currentMembers,
  createdBy,
  onClose,
  onMemberAdded,
  onMemberRemoved,
}: Props) {
  const { user } = useAuth()
  const [searchQuery, setSearchQuery] = useState('')
  const [searchResults, setSearchResults] = useState<UserSummary[]>([])
  const [selectedUser, setSelectedUser] = useState<UserSummary | null>(null)
  const [selectedRole, setSelectedRole] = useState<ProjectRole>('ANALYST')
  const [loading, setLoading] = useState(false)
  const [adding, setAdding] = useState(false)
  const [removing, setRemoving] = useState<number | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [showConfirm, setShowConfirm] = useState<number | null>(null)
  const searchTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Search users as user types
  useEffect(() => {
    if (!searchQuery.trim()) {
      setSearchResults([])
      return
    }
    if (searchTimer.current) clearTimeout(searchTimer.current)
    setLoading(true)
    searchTimer.current = setTimeout(async () => {
      try {
        const results = await searchUsers(searchQuery)
        // Filter out already-members
        const memberIds = new Set(currentMembers.map(m => m.user.id))
        setSearchResults(results.filter(r => !memberIds.has(r.id)))
      } catch {
        setError('Failed to search users')
        setSearchResults([])
      } finally {
        setLoading(false)
      }
    }, 300)
  }, [searchQuery, currentMembers])

  const handleAddMember = async () => {
    if (!selectedUser || !user) return
    setAdding(true)
    setError(null)
    try {
      const member = await addProjectMember(
        projectId,
        { userId: selectedUser.id, projectRole: selectedRole },
        user.userId,
      )
      onMemberAdded(member)
      setSearchQuery('')
      setSelectedUser(null)
      setSelectedRole('ANALYST')
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to add member')
    } finally {
      setAdding(false)
    }
  }

  const handleRemoveMember = async (memberId: number) => {
    if (!user) return
    setRemoving(memberId)
    setError(null)
    try {
      await removeProjectMember(projectId, memberId, user.userId)
      onMemberRemoved(memberId)
      setShowConfirm(null)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to remove member')
    } finally {
      setRemoving(null)
    }
  }

  const isCreator = (memberId: number) => createdBy?.id === memberId
  const currentMemberIds = new Set(currentMembers.map(m => m.user.id))

  return (
    <>
      {/* Backdrop */}
      <div
        className="fixed inset-0 bg-slate-900/30 backdrop-blur-[1px] z-40"
        onClick={onClose}
      />

      {/* Modal */}
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 pointer-events-none">
        <div className="w-full max-w-2xl bg-white rounded-2xl shadow-2xl border border-slate-200
                        pointer-events-auto flex flex-col overflow-hidden max-h-[90vh]">

          {/* Header */}
          <div className="flex items-center justify-between p-5 border-b border-slate-100 bg-slate-50">
            <div className="flex items-center gap-2">
              <Users size={16} className="text-emerald-600" />
              <h2 className="font-display text-lg text-slate-900">Manage Team</h2>
            </div>
            <button
              onClick={onClose}
              className="p-2 hover:bg-white rounded-lg transition-colors text-slate-400 hover:text-slate-600">
              <X size={18} />
            </button>
          </div>

          {/* Scrollable body */}
          <div className="flex-1 overflow-y-auto p-5 space-y-5">

            {/* Add Member Section */}
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4">
              <h3 className="font-semibold text-sm text-emerald-900 mb-3">Add Team Member</h3>

              {/* Search input */}
              <div className="mb-3">
                <label className="block text-xs text-slate-600 font-medium mb-1.5">Search User</label>
                <input
                  type="text"
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                  placeholder="Type name or email to search…"
                  className="w-full bg-white border border-emerald-300 text-slate-800 text-sm
                             rounded-lg px-3 py-2 focus:outline-none focus:border-emerald-500
                             focus:ring-2 focus:ring-emerald-100 placeholder-slate-400"
                />
                {loading && <Loader2 size={14} className="animate-spin text-slate-400 mt-1" />}
              </div>

              {/* Search results dropdown */}
              {searchResults.length > 0 && !selectedUser && (
                <div className="mb-3 bg-white border border-emerald-200 rounded-lg overflow-hidden">
                  {searchResults.map(u => (
                    <button
                      key={u.id}
                      onClick={() => setSelectedUser(u)}
                      className="w-full text-left px-3 py-2 border-b border-emerald-100 last:border-b-0
                                 hover:bg-emerald-50 transition-colors"
                    >
                      <div className="font-medium text-sm text-slate-800">@{u.username}</div>
                      <div className="text-xs text-slate-400">{u.email}</div>
                    </button>
                  ))}
                </div>
              )}

              {/* Selected user + role */}
              {selectedUser && !currentMemberIds.has(selectedUser.id) && (
                <div className="mb-3 p-3 bg-white border border-emerald-200 rounded-lg">
                  <div className="flex items-center justify-between mb-3">
                    <div>
                      <div className="font-medium text-sm text-slate-800">@{selectedUser.username}</div>
                      <div className="text-xs text-slate-400">{selectedUser.email}</div>
                    </div>
                    <button
                      onClick={() => setSelectedUser(null)}
                      className="p-1 hover:bg-slate-100 rounded transition-colors">
                      <X size={14} className="text-slate-400" />
                    </button>
                  </div>

                  {/* Role select */}
                  <label className="block text-xs text-slate-600 font-medium mb-1.5">Project Role</label>
                  <select
                    value={selectedRole}
                    onChange={e => setSelectedRole(e.target.value as ProjectRole)}
                    className="w-full bg-white border border-slate-200 text-slate-800 text-sm
                               rounded-lg px-3 py-2 focus:outline-none focus:border-emerald-400
                               focus:ring-2 focus:ring-emerald-100"
                  >
                    {PROJECT_ROLES.map(role => (
                      <option key={role} value={role}>
                        {role.replace(/_/g, ' ')}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {/* Error */}
              {error && (
                <div className="mb-3 px-3 py-2 bg-red-50 border border-red-200 rounded-lg text-xs text-red-600">
                  {error}
                </div>
              )}

              {/* Add button */}
              <button
                onClick={handleAddMember}
                disabled={!selectedUser || adding}
                className="w-full flex items-center justify-center gap-2 px-4 py-2 rounded-lg
                           font-medium text-sm transition-all
                           disabled:opacity-50 disabled:cursor-not-allowed
                           bg-emerald-600 hover:bg-emerald-500 text-white active:scale-95"
              >
                {adding ? (
                  <><Loader2 size={14} className="animate-spin" /> Adding…</>
                ) : (
                  <><Plus size={14} /> Add Member</>
                )}
              </button>
            </div>

            {/* Current Members */}
            <div>
              <h3 className="font-semibold text-sm text-slate-900 mb-3">
                Team Members ({currentMembers.length})
              </h3>
              <div className="space-y-2">
                {currentMembers.map(member => (
                  <div
                    key={member.id}
                    className="flex items-center justify-between p-3 bg-slate-50 border border-slate-100 rounded-lg"
                  >
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-medium text-sm text-slate-800">@{member.user.username}</span>
                        <span className={`text-[9px] px-1.5 py-0.5 rounded border font-semibold ${
                          ROLE_COLORS[member.projectRole] ?? 'text-slate-500 bg-slate-100 border-slate-200'
                        }`}>
                          {member.projectRole.replace(/_/g, ' ')}
                        </span>
                        {isCreator(member.user.id) && (
                          <span className="text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-200 font-semibold">
                            Creator
                          </span>
                        )}
                      </div>
                      <div className="text-xs text-slate-400 mt-0.5">{member.user.email}</div>
                    </div>

                    {/* Remove button (disabled for creator) */}
                    {!isCreator(member.user.id) && (
                      <>
                        {showConfirm === member.user.id ? (
                          <div className="flex items-center gap-2 ml-2">
                            <button
                              onClick={() => handleRemoveMember(member.user.id)}
                              disabled={removing === member.user.id}
                              className="px-2 py-1 text-xs font-medium rounded bg-red-100 text-red-700 hover:bg-red-200
                                       transition-colors disabled:opacity-50"
                            >
                              {removing === member.user.id ? 'Removing…' : 'Confirm'}
                            </button>
                            <button
                              onClick={() => setShowConfirm(null)}
                              disabled={removing === member.user.id}
                              className="px-2 py-1 text-xs font-medium rounded bg-slate-200 text-slate-700 hover:bg-slate-300
                                       transition-colors disabled:opacity-50"
                            >
                              Cancel
                            </button>
                          </div>
                        ) : (
                          <button
                            onClick={() => setShowConfirm(member.user.id)}
                            className="p-2 hover:bg-red-50 rounded-lg transition-colors text-slate-400
                                     hover:text-red-600 ml-2 flex-shrink-0"
                          >
                            <Trash2 size={14} />
                          </button>
                        )}
                      </>
                    )}
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Footer */}
          <div className="p-4 border-t border-slate-100 bg-slate-50 flex justify-end">
            <button
              onClick={onClose}
              className="px-4 py-2 rounded-lg font-medium text-sm
                       border border-slate-200 text-slate-700 hover:bg-slate-100
                       transition-all active:scale-95"
            >
              Close
            </button>
          </div>
        </div>
      </div>
    </>
  )
}
