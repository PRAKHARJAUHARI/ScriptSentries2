// src/components/GlobalSearchModal.tsx
import { useState, useEffect, useRef, useCallback } from 'react'
import { Search, X, Loader2, AlertTriangle, AlertCircle, Info,
         FolderOpen, FileText, ChevronRight, Database } from 'lucide-react'
import { globalSearch, type SearchHit, type SearchResponse } from '../api/globalSearchApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  onClose: () => void
  /** Called when user clicks a result — navigate to that project + risk flag */
  onNavigate: (projectId: number, scriptId: number, riskFlagId: number) => void
}

const SEVERITY_CONFIG = {
  HIGH:   { icon: AlertTriangle, color: 'text-red-500',    bg: 'bg-red-50    border-red-200'    },
  MEDIUM: { icon: AlertCircle,   color: 'text-amber-500',  bg: 'bg-amber-50  border-amber-200'  },
  LOW:    { icon: Info,          color: 'text-emerald-500',bg: 'bg-emerald-50 border-emerald-200'},
}

const STATUS_COLORS: Record<string, string> = {
  PENDING:                'bg-slate-100  text-slate-600',
  CLEARED:                'bg-emerald-100 text-emerald-700',
  NOT_CLEAR:              'bg-red-100    text-red-700',
  NEGOTIATED_BY_ATTORNEY: 'bg-violet-100 text-violet-700',
  BRANDED_INTEGRATION:    'bg-blue-100   text-blue-700',
  NO_CLEARANCE_NECESSARY: 'bg-teal-100   text-teal-700',
  PERMISSIBLE:            'bg-lime-100   text-lime-700',
}

function statusLabel(s: string) {
  return s.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, c => c.toUpperCase())
}

/** Group hits by project name for the UI */
function groupByProject(hits: SearchHit[]) {
  const groups = new Map<string, { projectId: number; studioName: string | null; hits: SearchHit[] }>()
  for (const hit of hits) {
    const key = hit.projectName
    if (!groups.has(key)) {
      groups.set(key, { projectId: hit.projectId, studioName: hit.studioName, hits: [] })
    }
    groups.get(key)!.hits.push(hit)
  }
  return Array.from(groups.entries()).map(([name, v]) => ({ projectName: name, ...v }))
}

export function GlobalSearchModal({ onClose, onNavigate }: Props) {
  const { user } = useAuth()
  const [query,   setQuery]   = useState('')
  const [result,  setResult]  = useState<SearchResponse | null>(null)
  const [loading, setLoading] = useState(false)
  const [error,   setError]   = useState<string | null>(null)
  const inputRef   = useRef<HTMLInputElement>(null)
  const searchTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Auto-focus input on open
  useEffect(() => { inputRef.current?.focus() }, [])

  // Debounced search — fires 400ms after user stops typing
  useEffect(() => {
    if (query.trim().length < 2) {
      setResult(null)
      return
    }
    if (searchTimer.current) clearTimeout(searchTimer.current)
    setLoading(true)
    setError(null)
    searchTimer.current = setTimeout(async () => {
      try {
        setResult(await globalSearch(query.trim(), user!.userId))
      } catch {
        setError('Search failed. Please try again.')
      } finally {
        setLoading(false)
      }
    }, 400)
  }, [query])

  // Close on Escape
  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', handler)
    return () => document.removeEventListener('keydown', handler)
  }, [onClose])

  const handleNavigate = useCallback((hit: SearchHit) => {
    onNavigate(hit.projectId, hit.scriptId, hit.riskFlagId)
    onClose()
  }, [onNavigate, onClose])

  const groups = result ? groupByProject(result.hits) : []

  return (
    <>
      {/* Backdrop */}
      <div className="fixed inset-0 bg-black/60 backdrop-blur-[2px] z-50" onClick={onClose} />

      {/* Modal */}
      <div className="fixed inset-0 z-50 flex items-start justify-center pt-[10vh] px-4 pointer-events-none">
        <div className="w-full max-w-2xl bg-white rounded-2xl shadow-2xl border border-slate-200
                        pointer-events-auto flex flex-col overflow-hidden max-h-[78vh]">

          {/* Search input */}
          <div className="flex items-center gap-3 px-4 py-3 border-b border-slate-100">
            <Search size={18} className="text-slate-400 flex-shrink-0" />
            <input
              ref={inputRef}
              type="text"
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder='Search across all projects — e.g. "John Wick", "Nike", "Empire State"'
              className="flex-1 text-sm text-slate-800 placeholder-slate-300
                         focus:outline-none bg-transparent"
            />
            {loading && <Loader2 size={15} className="animate-spin text-slate-400 flex-shrink-0" />}
            <button onClick={onClose}
              className="p-1 hover:bg-slate-100 rounded-lg transition-colors flex-shrink-0">
              <X size={15} className="text-slate-400" />
            </button>
          </div>

          {/* Body */}
          <div className="flex-1 overflow-y-auto">

            {/* Empty state */}
            {!query.trim() && (
              <div className="flex flex-col items-center justify-center py-16 text-slate-400">
                <Database size={36} className="mb-3 text-slate-200" />
                <p className="text-sm font-medium">Search across all your projects</p>
                <p className="text-xs mt-1 text-slate-300">Find any entity name in any script, any version</p>
              </div>
            )}

            {/* Too short */}
            {query.trim().length === 1 && (
              <div className="text-center py-10 text-xs text-slate-400">
                Type at least 2 characters to search
              </div>
            )}

            {/* Error */}
            {error && (
              <div className="m-4 px-4 py-3 bg-red-50 border border-red-200 rounded-xl text-xs text-red-600">
                {error}
              </div>
            )}

            {/* No results */}
            {result && result.totalHits === 0 && (
              <div className="flex flex-col items-center justify-center py-16 text-slate-400">
                <Search size={32} className="mb-3 text-slate-200" />
                <p className="text-sm font-medium">No results for "{result.query}"</p>
                <p className="text-xs mt-1 text-slate-300">
                  This entity hasn't appeared in any script across your projects
                </p>
              </div>
            )}

            {/* Results */}
            {result && result.totalHits > 0 && (
              <div className="p-3 space-y-4">

                {/* Summary */}
                <div className="flex items-center justify-between px-1">
                  <span className="text-xs text-slate-500">
                    <span className="font-semibold text-slate-800">{result.totalHits}</span> result{result.totalHits !== 1 ? 's' : ''} for "
                    <span className="font-semibold text-emerald-600">{result.query}</span>"
                    {result.totalHits === 200 && (
                      <span className="text-amber-500 ml-1">— showing first 200</span>
                    )}
                  </span>
                  <span className="text-[10px] text-slate-400">
                    {groups.length} project{groups.length !== 1 ? 's' : ''}
                  </span>
                </div>

                {/* Grouped by project */}
                {groups.map(group => (
                  <div key={group.projectId}
                    className="border border-slate-100 rounded-xl overflow-hidden">

                    {/* Project header */}
                    <div className="flex items-center gap-2 px-3 py-2 bg-slate-50 border-b border-slate-100">
                      <FolderOpen size={13} className="text-emerald-600 flex-shrink-0" />
                      <span className="text-xs font-bold text-slate-800">{group.projectName}</span>
                      {group.studioName && (
                        <span className="text-[10px] text-slate-400">· {group.studioName}</span>
                      )}
                      <span className="ml-auto text-[10px] text-slate-400 bg-white border border-slate-200 px-1.5 py-0.5 rounded-full">
                        {group.hits.length} hit{group.hits.length !== 1 ? 's' : ''}
                      </span>
                    </div>

                    {/* Risk hits */}
                    {group.hits.map(hit => {
                      const sev = SEVERITY_CONFIG[hit.severity as keyof typeof SEVERITY_CONFIG]
                      const SevIcon = sev.icon
                      return (
                        <button
                          key={hit.riskFlagId}
                          onClick={() => handleNavigate(hit)}
                          className="w-full text-left px-3 py-2.5 border-b border-slate-50 last:border-b-0
                                     hover:bg-emerald-50/60 transition-colors group"
                        >
                          <div className="flex items-start gap-2.5">
                            {/* Severity icon */}
                            <div className={`p-1 rounded-lg border mt-0.5 flex-shrink-0 ${sev.bg}`}>
                              <SevIcon size={11} className={sev.color} />
                            </div>

                            <div className="flex-1 min-w-0">
                              {/* Top row: entity name + version + page */}
                              <div className="flex items-center gap-2 flex-wrap">
                                <span className="text-sm font-semibold text-slate-800">
                                  {hit.entityName}
                                </span>
                                <span className={`text-[9px] px-1.5 py-0.5 rounded font-semibold ${
                                  STATUS_COLORS[hit.status] ?? 'bg-slate-100 text-slate-500'
                                }`}>
                                  {statusLabel(hit.status)}
                                </span>
                              </div>

                              {/* Middle row: breadcrumb */}
                              <div className="flex items-center gap-1 mt-0.5 text-[10px] text-slate-400">
                                <FileText size={9} />
                                <span className="truncate max-w-[140px]">{hit.scriptFilename}</span>
                                {hit.versionName && (
                                  <>
                                    <ChevronRight size={9} />
                                    <span>{hit.versionName}</span>
                                  </>
                                )}
                                <ChevronRight size={9} />
                                <span>pg.{hit.pageNumber}</span>
                                <span className="mx-1">·</span>
                                <span className="text-slate-300">{hit.category.replace(/_/g, ' ')}</span>
                              </div>

                              {/* Snippet */}
                              {hit.snippet && (
                                <p className="text-[10px] text-slate-400 mt-0.5 line-clamp-1 italic">
                                  "{hit.snippet}"
                                </p>
                              )}
                            </div>

                            {/* Arrow — appears on hover */}
                            <ChevronRight
                              size={13}
                              className="text-slate-300 group-hover:text-emerald-500
                                         flex-shrink-0 mt-1 transition-colors"
                            />
                          </div>
                        </button>
                      )
                    })}
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Footer hint */}
          <div className="px-4 py-2 border-t border-slate-100 bg-slate-50 flex justify-between">
            <p className="text-[10px] text-slate-400">
              Searches entity names across all scripts in all your projects
            </p>
            <p className="text-[10px] text-slate-400">Esc to close</p>
          </div>
        </div>
      </div>
    </>
  )
}