// src/api/api.ts
import axios from 'axios'
import type { Script, RiskFlag, RiskUpdatePayload } from '../types'

const api = axios.create({ 
  baseURL: (import.meta.env.VITE_API_BASE_URL as string) || '/api' 
})

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

// ─── Scripts ──────────────────────────────────────────────────────────────────

export async function listScripts(): Promise<Script[]> {
  const { data } = await api.get<Script[]>('/scripts')
  return data
}

export async function getScript(id: number): Promise<Script> {
  const { data } = await api.get<Script>(`/scripts/${id}`)
  return data
}

/**
 * Upload a PDF for zero-retention AI analysis.
 * projectId MUST be appended to FormData — backend @RequestParam("projectId") is required.
 * Do NOT pass onProgress as the second arg — that was the original bug.
 */
export async function scanScript(
  file: File,
  projectId: number,
  onProgress?: (pct: number) => void,
): Promise<Script> {
  const form = new FormData()
  form.append('file', file)
  form.append('projectId', String(projectId))  // required by backend

  const { data } = await api.post<Script>('/scripts/scan', form, {
    // Do NOT set Content-Type — axios sets multipart/form-data with boundary
    onUploadProgress: (e) => {
      if (onProgress && e.total) {
        onProgress(Math.round((e.loaded / e.total) * 100))
      }
    },
  })

  return data
}

// ─── Risk Flags ───────────────────────────────────────────────────────────────

export async function updateRisk(
  id: number,
  payload: RiskUpdatePayload,
): Promise<RiskFlag> {
  const { data } = await api.patch<RiskFlag>(`/risks/${id}`, payload)
  return data
}

// ─── Export ───────────────────────────────────────────────────────────────────

export async function exportScript(scriptId: number): Promise<void> {
  const response = await api.get(`/scripts/${scriptId}/export`, {
    responseType: 'blob',
  })

  const contentDisposition = response.headers['content-disposition'] ?? ''
  const match = contentDisposition.match(/filename="?([^";\n]+)"?/)
  const filename = match?.[1] ?? `ScriptSentries_Report_${scriptId}.xlsx`

  const url = URL.createObjectURL(new Blob([response.data]))
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}
