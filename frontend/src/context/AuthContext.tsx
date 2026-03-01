import { createContext, useContext, useState, ReactNode } from 'react'

export type ProjectRole =
  | 'ATTORNEY'
  | 'ANALYST'
  | 'MAIN_PRODUCTION_CONTACT'
  | 'PRODUCTION_ASSISTANT'
  | 'VIEWER'

export interface AuthUser {
  userId: number
  username: string
  email: string
  role: ProjectRole
  token: string
}

interface AuthContextType {
  user: AuthUser | null
  login: (user: AuthUser) => void
  logout: () => void
  isAuthenticated: boolean
}

const AuthContext = createContext<AuthContextType | null>(null)

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<AuthUser | null>(() => {
    try {
      const stored = localStorage.getItem('ss_user')
      return stored ? JSON.parse(stored) : null
    } catch { return null }
  })

  const login = (userData: AuthUser) => {
    localStorage.setItem('ss_user', JSON.stringify(userData))
    setUser(userData)
  }

  const logout = () => {
    localStorage.removeItem('ss_user')
    setUser(null)
  }

  return (
    <AuthContext.Provider value={{ user, login, logout, isAuthenticated: !!user }}>
      {children}
    </AuthContext.Provider>
  )
}

export function useAuth() {
  const ctx = useContext(AuthContext)
  if (!ctx) throw new Error('useAuth must be used within AuthProvider')
  return ctx
}

// Permission helpers — mirrors ProjectRole.java
export const canUpload        = (r: ProjectRole) => r !== 'VIEWER'
export const canEdit          = (r: ProjectRole) => r === 'ATTORNEY' || r === 'ANALYST'
export const canFinalize      = (r: ProjectRole) => r === 'ATTORNEY'
export const canRename        = (r: ProjectRole) => r === 'ATTORNEY' || r === 'MAIN_PRODUCTION_CONTACT'
export const canDeleteScript  = (r: ProjectRole) => r === 'ATTORNEY' || r === 'ANALYST'
export const canDeleteProject = (r: ProjectRole) => r === 'ATTORNEY'
export const canManageMembers = (r: ProjectRole) => r === 'ATTORNEY' || r === 'ANALYST'
export const canAddViewer     = (r: ProjectRole) => r === 'ATTORNEY' || r === 'ANALYST'
export const isReadOnly       = (r: ProjectRole) => r === 'VIEWER' || r === 'PRODUCTION_ASSISTANT'