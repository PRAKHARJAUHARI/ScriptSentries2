// src/components/RiskTable.tsx
import { useState } from 'react'
import { AlertTriangle, AlertCircle, Info, FileWarning } from 'lucide-react'
import type { RiskFlag, RiskSeverity, RiskCategory } from '../types'
import { StatusBadge } from './StatusSelect'

interface Props {
  risks: RiskFlag[]
  onSelectRisk: (risk: RiskFlag) => void
  selectedId?: number
}

const SEVERITY_CONFIG: Record<RiskSeverity, { icon: typeof AlertTriangle; color: string; dot: string }> = {
  HIGH:   { icon: AlertTriangle, color: 'text-red-500',    dot: 'bg-red-500'    },
  MEDIUM: { icon: AlertCircle,   color: 'text-amber-500',  dot: 'bg-amber-200'  },
  LOW:    { icon: Info,          color: 'text-emerald-600', dot: 'bg-emerald-500' },
}

const CATEGORY_LABELS: Partial<Record<RiskCategory, string>> = {
  FACT_BASED_ISSUES:     'Fact Issues',
  GOVERNMENT:            'Government',
  LIKENESS:              'Likeness',
  LOCATIONS:             'Locations',
  MARKETING_ADDED_VALUE: 'Marketing / MAV',
  MUSIC_CHOREOGRAPHY:    'Music / Choreo',
  NAMES_NUMBERS:         'Names / Numbers',
  PLAYBACK:              'Playback',
  PRODUCT_MISUSE:        'Product Misuse',
  PROPS_SET_DRESSING:    'Props / Set',
  REFERENCES:            'References',
  VEHICLES:              'Vehicles',
  WARDROBE:              'Wardrobe',
  OTHER:                 'Other',
}

export function RiskTable({ risks, onSelectRisk, selectedId }: Props) {
  const [filter,         setFilter]         = useState<RiskSeverity | 'ALL'>('ALL')
  const [categoryFilter, setCategoryFilter] = useState<string>('ALL')

  const categories = Array.from(new Set(risks.map(r => r.category)))

  // Sort by severity: HIGH -> MEDIUM -> LOW
  const severitySortOrder: Record<RiskSeverity, number> = { HIGH: 0, MEDIUM: 1, LOW: 2 }

  const filtered = risks
    .filter(r => {
      const sevMatch = filter === 'ALL' || r.severity === filter
      const catMatch = categoryFilter === 'ALL' || r.category === categoryFilter
      return sevMatch && catMatch
    })
    .sort((a, b) => severitySortOrder[a.severity] - severitySortOrder[b.severity])

  const counts = {
    HIGH:   risks.filter(r => r.severity === 'HIGH').length,
    MEDIUM: risks.filter(r => r.severity === 'MEDIUM').length,
    LOW:    risks.filter(r => r.severity === 'LOW').length,
  }

  if (risks.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-24 text-slate-400 animate-fade-in">
        <FileWarning size={48} className="mb-4 opacity-30" />
        <p className="font-display text-lg text-slate-500">No risks flagged</p>
        <p className="text-sm mt-1">This script is clear</p>
      </div>
    )
  }

  return (
    <div className="flex flex-col gap-4 animate-slide-up">

      {/* Filter bar */}
      <div className="flex items-center gap-2 flex-wrap">
        <button
          onClick={() => setFilter('ALL')}
          className={`badge border transition-all ${
            filter === 'ALL'
              ? 'bg-slate-800 text-white border-slate-700'
              : 'bg-white text-slate-500 border-slate-200 hover:border-slate-300'
          }`}>
          All {risks.length}
        </button>
        <button
          onClick={() => setFilter('HIGH')}
          className={`badge transition-all ${
            filter === 'HIGH'
              ? 'bg-red-500 text-white border border-red-500'
              : 'bg-red-50 text-red-500 border border-red-200 hover:bg-red-100'
          }`}>
          <span className="w-1.5 h-1.5 rounded-full bg-current inline-block" />
          High {counts.HIGH}
        </button>
        <button
          onClick={() => setFilter('MEDIUM')}
          className={`badge transition-all ${
            filter === 'MEDIUM'
              ? 'bg-amber-500 text-white border border-amber-500'
              : 'bg-amber-50 text-amber-600 border border-amber-200 hover:bg-amber-100'
          }`}>
          <span className="w-1.5 h-1.5 rounded-full bg-current inline-block" />
          Medium {counts.MEDIUM}
        </button>
        <button
          onClick={() => setFilter('LOW')}
          className={`badge transition-all ${
            filter === 'LOW'
              ? 'bg-emerald-600 text-white border border-emerald-600'
              : 'bg-emerald-50 text-emerald-600 border border-emerald-200 hover:bg-emerald-100'
          }`}>
          <span className="w-1.5 h-1.5 rounded-full bg-current inline-block" />
          Low {counts.LOW}
        </button>

        <select
          value={categoryFilter}
          onChange={e => setCategoryFilter(e.target.value)}
          className="ml-auto bg-white border border-slate-200 text-slate-600
                     rounded-lg px-3 py-1.5 text-xs focus:outline-none focus:border-emerald-400
                     transition-colors shadow-sm">
          <option value="ALL">All Categories</option>
          {categories.map(c => (
            <option key={c} value={c}>{CATEGORY_LABELS[c] ?? c}</option>
          ))}
        </select>
      </div>

      {/* Table */}
      <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100 bg-slate-50">
              <th className="text-left py-3 px-4 text-slate-500 font-semibold text-xs uppercase tracking-wider w-16">Episode<br/>Scene Page</th>
              <th className="text-left py-3 px-4 text-slate-500 font-semibold text-xs uppercase tracking-wider w-20">Severity</th>
              <th className="text-left py-3 px-4 text-slate-500 font-semibold text-xs uppercase tracking-wider flex-1">Category</th>
              <th className="text-left py-3 px-4 text-slate-500 font-semibold text-xs uppercase tracking-wider flex-1">Entity / Item</th>
              <th className="text-left py-3 px-4 text-slate-500 font-semibold text-xs uppercase tracking-wider w-20">Status</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((risk) => {
              const sev      = SEVERITY_CONFIG[risk.severity]
              const Icon     = sev.icon
              const isSelected = risk.id === selectedId

              return (
                <tr
                  key={risk.id}
                  onClick={() => onSelectRisk(risk)}
                  className={`border-b border-slate-100 cursor-pointer transition-colors
                              hover:bg-slate-50 ${
                    isSelected
                      ? 'bg-emerald-50 border-l-2 border-l-emerald-500'
                      : ''
                  }`}>
                  <td className="py-3 px-4">
                    <span className="bg-white border border-slate-300 text-slate-900 text-[10px] font-mono font-semibold px-2 py-1 rounded inline-block tracking-wide whitespace-nowrap">
                      {[
                        `Pg ${risk.pageNumber}`,
                        risk.episodeNumber ? `Ep ${risk.episodeNumber}` : null,
                        risk.sceneNumber ? `Scene ${risk.sceneNumber}` : null,
                      ].filter(Boolean).join(' | ')}
                    </span>
                  </td>
                  <td className="py-3 px-4">
                    <span className={`flex items-center gap-1 ${sev.color} font-semibold text-xs whitespace-nowrap`}>
                      <Icon size={13} />
                      {risk.severity}
                    </span>
                  </td>
                  <td className="py-3 px-4">
                    <div className="flex flex-col min-w-0">
                      <span className="text-slate-600 text-xs font-medium truncate">
                        {CATEGORY_LABELS[risk.category] ?? risk.category}
                      </span>
                      {risk.subCategory && (
                        <span className="text-slate-400 text-[11px] truncate mt-0.5">
                          {risk.subCategory}
                        </span>
                      )}
                    </div>
                  </td>
                  <td className="py-3 px-4">
                    <div className="flex flex-col min-w-0">
                      <span className="font-medium text-slate-800 text-sm truncate">
                        {risk.entityName}
                      </span>
                      {risk.snippet && (
                        <span className="text-slate-400 text-xs truncate max-w-xs mt-0.5">
                          {risk.snippet.slice(0, 60)}…
                        </span>
                      )}
                    </div>
                  </td>
                  <td className="py-3 px-4">
                    <StatusBadge status={risk.status} />
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>

        {filtered.length === 0 && (
          <div className="text-center py-12 text-slate-400 text-sm">
            No risks match the current filter
          </div>
        )}
      </div>
    </div>
  )
}
