// src/api/globalSearchApi.ts
import axios from 'axios'

const api = axios.create({ baseURL: '/api' })

api.interceptors.request.use((config) => {
  try {
    const stored = localStorage.getItem('ss_user')
    if (stored) {
      const { token } = JSON.parse(stored)
      if (token) config.headers.Authorization = `Bearer ${token}`
    }
  } catch { /* ignore */ }
  return config
})

// ── Types ─────────────────────────────────────────────────────────────────────

export interface SearchHit {
  riskFlagId: number
  entityName: string
  category: string
  subCategory: string
  severity: 'HIGH' | 'MEDIUM' | 'LOW'
  status: string
  snippet: string | null
  pageNumber: number
  scriptId: number
  scriptFilename: string
  versionName: string | null
  projectId: number
  projectName: string
  studioName: string | null
}

export interface SearchResponse {
  query: string
  totalHits: number
  hits: SearchHit[]
}

// ── API calls ─────────────────────────────────────────────────────────────────

export async function globalSearch(q: string, userId: number): Promise<SearchResponse> {
  const { data } = await api.get<SearchResponse>('/search', { params: { q, userId } })
  return data
}

/** Triggers a browser download of the E&O insurance Excel workbook */
export async function downloadEoExport(scriptId: number, userId: number): Promise<void> {
  const response = await api.get(`/scripts/${scriptId}/export/eo`, {
    params: { userId },
    responseType: 'blob',
  })

  const url      = URL.createObjectURL(new Blob([response.data]))
  const link     = document.createElement('a')
  const filename = response.headers['content-disposition']
    ?.match(/filename="?([^"]+)"?/)?.[1]
    ?? `EO_Submission_${scriptId}.xlsx`

  link.href = url
  link.setAttribute('download', filename)
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}
