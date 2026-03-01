// src/components/MentionInput.tsx
import { useState, useRef, useEffect, useCallback } from 'react'
import { searchUsers, type UserSummary } from '../api/authApi'
import { AtSign, Loader2 } from 'lucide-react'

interface Props {
  value: string
  onChange: (value: string) => void
  placeholder?: string
  rows?: number
  disabled?: boolean
  projectId?: number
}

const ROLE_COLORS: Record<string, string> = {
  ATTORNEY:                'text-violet-700 bg-violet-50',
  ANALYST:                 'text-blue-700   bg-blue-50',
  MAIN_PRODUCTION_CONTACT: 'text-amber-700  bg-amber-50',
  PRODUCTION_ASSISTANT:    'text-slate-600  bg-slate-100',
  VIEWER:                  'text-zinc-500   bg-zinc-100',
}

export function MentionInput({ value, onChange, placeholder, rows = 3, disabled, projectId }: Props) {
  const [showDropdown, setShowDropdown] = useState(false)
  const [mentionQuery, setMentionQuery] = useState('')
  const [users,        setUsers]        = useState<UserSummary[]>([])
  const [loading,      setLoading]      = useState(false)
  const [activeIndex,  setActiveIndex]  = useState(0)
  const [mentionStart, setMentionStart] = useState(-1)

  const textareaRef  = useRef<HTMLTextAreaElement>(null)
  const dropdownRef  = useRef<HTMLDivElement>(null)
  const searchTimer  = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Debounced search when mentionQuery changes
  useEffect(() => {
    if (!mentionQuery) { setUsers([]); return }
    if (searchTimer.current) clearTimeout(searchTimer.current)
    setLoading(true)
    searchTimer.current = setTimeout(async () => {
      try {
        setUsers(await searchUsers(mentionQuery, projectId))
        setActiveIndex(0)
      } catch { setUsers([]) }
      finally { setLoading(false) }
    }, 200)
  }, [mentionQuery, projectId])

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const text      = e.target.value
    const cursorPos = e.target.selectionStart
    onChange(text)

    const beforeCursor = text.slice(0, cursorPos)
    const atIdx        = beforeCursor.lastIndexOf('@')

    if (atIdx !== -1) {
      const afterAt = beforeCursor.slice(atIdx + 1)
      // Only trigger if no space/newline between @ and cursor
      if (!afterAt.includes(' ') && !afterAt.includes('\n')) {
        setMentionStart(atIdx)
        setMentionQuery(afterAt)
        setShowDropdown(true)
        return
      }
    }

    setShowDropdown(false)
    setMentionQuery('')
    setMentionStart(-1)
  }

  const selectUser = useCallback((u: UserSummary) => {
    if (mentionStart === -1) return
    const before  = value.slice(0, mentionStart)
    const cursorPos = textareaRef.current?.selectionStart ?? value.length
    const after   = value.slice(cursorPos)
    const newText = `${before}@${u.username} ${after}`
    onChange(newText)

    setShowDropdown(false)
    setMentionQuery('')
    setMentionStart(-1)
    setUsers([])

    // Move cursor to just after the inserted mention
    setTimeout(() => {
      const ta = textareaRef.current
      if (!ta) return
      const pos = before.length + u.username.length + 2
      ta.focus()
      ta.setSelectionRange(pos, pos)
    }, 0)
  }, [value, mentionStart, onChange])

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (!showDropdown || users.length === 0) return
    if (e.key === 'ArrowDown') { e.preventDefault(); setActiveIndex(i => (i + 1) % users.length) }
    else if (e.key === 'ArrowUp')  { e.preventDefault(); setActiveIndex(i => (i - 1 + users.length) % users.length) }
    else if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); selectUser(users[activeIndex]) }
    else if (e.key === 'Escape')   { setShowDropdown(false) }
  }

  // Close on outside click
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (
        dropdownRef.current  && !dropdownRef.current.contains(e.target as Node) &&
        textareaRef.current  && !textareaRef.current.contains(e.target as Node)
      ) { setShowDropdown(false) }
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  return (
    <div className="relative">
      <div className="relative">
        <textarea
          ref={textareaRef}
          value={value}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          rows={rows}
          disabled={disabled}
          placeholder={placeholder ?? 'Add a comment… type @ to mention a colleague'}
          className="w-full bg-slate-50 border border-slate-200 text-slate-800 text-sm
                     rounded-xl px-4 py-2.5 pr-9 focus:outline-none focus:border-emerald-400
                     focus:ring-2 focus:ring-emerald-100 placeholder-slate-300
                     resize-none transition-all disabled:opacity-50"
        />
        <AtSign
          size={14}
          className={`absolute right-3 top-3 pointer-events-none transition-colors ${
            showDropdown ? 'text-emerald-500' : 'text-slate-300'
          }`}
        />
      </div>

      {/* Mention dropdown */}
      {showDropdown && (
        <div
          ref={dropdownRef}
          className="absolute bottom-full left-0 right-0 mb-1.5 z-50
                     bg-white border border-slate-200 rounded-xl shadow-xl
                     overflow-hidden"
        >
          {/* Header */}
          <div className="flex items-center gap-2 px-3 py-2 border-b border-slate-100 bg-slate-50">
            <AtSign size={11} className="text-emerald-500" />
            <span className="text-xs text-slate-500">
              {loading ? 'Searching…' : 'Mention a teammate'}
            </span>
            {loading && <Loader2 size={11} className="animate-spin text-slate-400 ml-auto" />}
          </div>

          {/* Results */}
          {users.length > 0 ? (
            <ul className="max-h-48 overflow-y-auto">
              {users.map((u, idx) => (
                <li
                  key={u.id}
                  onMouseDown={e => { e.preventDefault(); selectUser(u) }}
                  onMouseEnter={() => setActiveIndex(idx)}
                  className={`flex items-center gap-3 px-3 py-2.5 cursor-pointer transition-colors ${
                    idx === activeIndex ? 'bg-emerald-50' : 'hover:bg-slate-50'
                  }`}
                >
                  <div className="w-7 h-7 rounded-full bg-emerald-100 border border-emerald-200
                                  flex items-center justify-center flex-shrink-0">
                    <span className="text-[11px] font-bold text-emerald-700">
                      {u.username.charAt(0).toUpperCase()}
                    </span>
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="text-sm text-slate-800 font-medium">@{u.username}</span>
                      <span className={`text-[9px] px-1.5 py-0.5 rounded font-semibold ${
                        ROLE_COLORS[u.role] ?? 'text-slate-500 bg-slate-100'
                      }`}>
                        {u.role.replace(/_/g, ' ')}
                      </span>
                    </div>
                    <span className="text-[11px] text-slate-400 truncate block">{u.email}</span>
                  </div>
                  {idx === activeIndex && (
                    <span className="text-[10px] text-slate-400 flex-shrink-0">↵</span>
                  )}
                </li>
              ))}
            </ul>
          ) : !loading ? (
            <div className="px-4 py-5 text-center text-slate-400 text-sm">
              {mentionQuery
                ? `No users matching "@${mentionQuery}"`
                : 'Type a username to search'}
            </div>
          ) : null}

          <div className="px-3 py-1.5 border-t border-slate-100 bg-slate-50 flex gap-3 text-[10px] text-slate-400">
            <span>↑↓ navigate</span>
            <span>↵ / Tab select</span>
            <span>Esc close</span>
          </div>
        </div>
      )}
    </div>
  )
}