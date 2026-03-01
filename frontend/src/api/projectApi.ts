// src/api/projectApi.ts
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

export type ProjectRole =
  | 'ATTORNEY' | 'ANALYST' | 'MAIN_PRODUCTION_CONTACT'
  | 'PRODUCTION_ASSISTANT' | 'VIEWER'

export interface UserSummary    { id: number; username: string; email: string; role: string }
export interface MemberResponse { id: number; user: UserSummary; projectRole: ProjectRole; joinedAt: string }

export interface ProjectResponse {
  id: number
  name: string
  studioName: string | null
  director: string | null
  producer: string | null
  productionEmail: string | null
  productionPhone: string | null
  genre: string | null
  logline: string | null
  expectedRelease: string | null
  imdbLink: string | null
  notes: string | null
  createdAt: string
  createdBy: UserSummary | null
  members: MemberResponse[]
  totalScripts: number
  totalRisks: number
}

export interface TimelineEntry {
  scriptId: number
  filename: string
  versionName: string
  uploadedAt: string
  deletedAt: string | null
  status: string
  totalPages: number
  highCount: number
  mediumCount: number
  lowCount: number
  totalRisks: number
  uploadedBy: UserSummary | null
}

export interface ProjectTimeline {
  projectId: number
  projectName: string
  studioName: string | null
  versions: TimelineEntry[]
  totalVersions: number
  totalHighRisks: number
}

export interface CreateProjectPayload {
  name: string
  studioName?: string
  director?: string
  producer?: string
  productionEmail?: string
  productionPhone?: string
  genre?: string
  logline?: string
  expectedRelease?: string
  imdbLink?: string
  notes?: string
}

export type UpdateProjectPayload = Partial<Omit<CreateProjectPayload, 'members'>>

// ── Projects ──────────────────────────────────────────────────────────────────

export async function createProject(payload: CreateProjectPayload, userId: number): Promise<ProjectResponse> {
  const { data } = await api.post<ProjectResponse>('/projects', payload, { params: { userId } })
  return data
}

export async function updateProject(id: number, payload: UpdateProjectPayload, userId: number): Promise<ProjectResponse> {
  const { data } = await api.patch<ProjectResponse>(`/projects/${id}`, payload, { params: { userId } })
  return data
}

export async function getProjects(userId: number): Promise<ProjectResponse[]> {
  const { data } = await api.get<ProjectResponse[]>('/projects', { params: { userId } })
  return data
}

export async function getProject(id: number): Promise<ProjectResponse> {
  const { data } = await api.get<ProjectResponse>(`/projects/${id}`)
  return data
}

export async function deleteProject(id: number, userId: number): Promise<void> {
  await api.delete(`/projects/${id}`, { params: { userId } })
}

// ── Timeline ──────────────────────────────────────────────────────────────────

export async function getProjectTimeline(projectId: number): Promise<ProjectTimeline> {
  const { data } = await api.get<ProjectTimeline>(`/projects/${projectId}/timeline`)
  return data
}

// ── Members ───────────────────────────────────────────────────────────────────

export async function addProjectMember(
  projectId: number,
  invite: { userId: number; projectRole: ProjectRole },
  requestingUserId: number,
): Promise<MemberResponse> {
  const { data } = await api.post<MemberResponse>(
    `/projects/${projectId}/members`,
    invite,
    { params: { requestingUserId } },
  )
  return data
}

export async function removeProjectMember(
  projectId: number,
  targetUserId: number,
  requestingUserId: number,
): Promise<void> {
  await api.delete(`/projects/${projectId}/members/${targetUserId}`, {
    params: { requestingUserId },
  })
}

// ── Scripts ───────────────────────────────────────────────────────────────────

export async function assignScriptToProject(
  scriptId: number,
  projectId: number,
  versionName: string | undefined,
  userId: number,
): Promise<void> {
  await api.post(`/projects/scripts/${scriptId}/assign`, null, {
    params: { projectId, versionName, userId },
  })
}

export async function renameVersion(scriptId: number, versionName: string, userId: number): Promise<void> {
  await api.patch(`/projects/scripts/${scriptId}/rename`, { versionName }, { params: { userId } })
}

export async function deleteScript(scriptId: number, userId: number): Promise<void> {
  await api.delete(`/projects/scripts/${scriptId}`, { params: { userId } })
}

// ── User search ───────────────────────────────────────────────────────────────

export async function searchUsers(query: string): Promise<UserSummary[]> {
  const { data } = await api.get<UserSummary[]>('/auth/users/search', { params: { q: query } })
  return data
}