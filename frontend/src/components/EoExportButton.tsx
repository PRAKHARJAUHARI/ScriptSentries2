// src/components/EoExportButton.tsx
import { useState } from 'react'
import { FileCheck, Loader2, AlertTriangle } from 'lucide-react'
import { downloadEoExport } from '../api/globalSearchApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  scriptId: number
  /** The user's role on this project — only ATTORNEY and ANALYST can generate E&O */
  projectRole: string
}

export function EoExportButton({ scriptId, projectRole }: Props) {
  const { user } = useAuth()
  const [loading, setLoading] = useState(false)
  const [error,   setError]   = useState<string | null>(null)

  const canExport = projectRole === 'ATTORNEY' || projectRole === 'ANALYST'

  const handleExport = async () => {
    if (!user || !canExport) return
    setLoading(true)
    setError(null)
    try {
      await downloadEoExport(scriptId, user.userId)
    } catch {
      setError('Export failed. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  if (!canExport) return null

  return (
    <div className="inline-flex flex-col items-start gap-1">
      <button
        onClick={handleExport}
        disabled={loading}
        className="inline-flex items-center gap-2 px-4 py-2 bg-violet-600 hover:bg-violet-500
                   text-white text-xs font-medium rounded-xl transition-all
                   disabled:opacity-50 active:scale-95 border border-violet-700"
        title="Download E&O Insurance Submission Workbook"
      >
        {loading
          ? <><Loader2 size={13} className="animate-spin" /> Generating E&O…</>
          : <><FileCheck size={13} /> Export E&O Insurance</>
        }
      </button>
      {error && (
        <div className="flex items-center gap-1 text-[10px] text-red-500">
          <AlertTriangle size={10} />
          {error}
        </div>
      )}
    </div>
  )
}
