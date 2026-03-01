// src/api/authApi.ts
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

// ─── Auth ─────────────────────────────────────────────────────────────────────

export interface RegisterPayload {
  username: string
  email: string
  password: string
  role?: string
}

export interface LoginPayload {
  email: string
  password: string
}

export interface AuthResponse {
  token: string
  userId: number
  username: string
  email: string
  role: string
}

export async function register(payload: RegisterPayload): Promise<AuthResponse> {
  const { data } = await api.post<AuthResponse>('/auth/register', payload)
  return data
}

export async function login(payload: LoginPayload): Promise<AuthResponse> {
  const { data } = await api.post<AuthResponse>('/auth/login', payload)
  return data
}

// ─── User Search ──────────────────────────────────────────────────────────────

export interface UserSummary {
  id: number
  username: string
  email: string
  role: string
}

/**
 * Search users by username or email prefix.
 * Used for @mention autocomplete AND member invite search.
 * Backend: GET /api/collab/users/search?q=<query>&projectId=<projectId>
 * If projectId provided, returns only users already in that project.
 */
export async function searchUsers(q: string, projectId?: number): Promise<UserSummary[]> {
  const params: { q: string; projectId?: number } = { q }
  if (projectId) params.projectId = projectId
  const { data } = await api.get<UserSummary[]>('/collab/users/search', { params })
  return data
}

// ─── Comments ─────────────────────────────────────────────────────────────────

export interface CommentResponse {
  id: number
  text: string
  author: UserSummary
  riskFlagId: number
  type?: string
  createdAt: string
}

export async function getComments(riskFlagId: number): Promise<CommentResponse[]> {
  const { data } = await api.get<CommentResponse[]>(`/collab/comments/${riskFlagId}`)
  return data
}

export async function postComment(payload: {
  text: string
  riskFlagId: number
  authorId: number
  type?: string
}): Promise<CommentResponse> {
  const { data } = await api.post<CommentResponse>('/collab/comments', payload)
  return data
}

// ─── Notifications ────────────────────────────────────────────────────────────

export interface NotificationResponse {
  id: number
  message: string
  isRead: boolean
  riskFlagId: number | null
  createdAt: string
}

export async function getNotifications(userId: number): Promise<NotificationResponse[]> {
  const { data } = await api.get<NotificationResponse[]>(`/collab/notifications/${userId}`)
  return data
}

export async function getUnreadCount(userId: number): Promise<number> {
  const { data } = await api.get<number>(`/collab/notifications/${userId}/unread-count`)
  return data
}

export async function markAllRead(userId: number): Promise<void> {
  await api.post(`/collab/notifications/${userId}/mark-read`)
}