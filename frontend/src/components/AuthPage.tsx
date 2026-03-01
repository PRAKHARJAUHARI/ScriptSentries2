// src/components/AuthPage.tsx
import { useState, FormEvent } from 'react'
import { ShieldCheck, Mail, Lock, User, Loader2, Scale, Search, Phone, Clapperboard, Eye } from 'lucide-react'
import { login, register } from '../api/authApi'
import { useAuth, type ProjectRole } from '../context/AuthContext'

const ROLES: { value: ProjectRole; label: string; desc: string; icon: React.ElementType }[] = [
  { value: 'ATTORNEY',                label: 'Attorney',        desc: 'Full access + finalize',  icon: Scale        },
  { value: 'ANALYST',                 label: 'Analyst',         desc: 'Edit risks + comment',    icon: Search       },
  { value: 'MAIN_PRODUCTION_CONTACT', label: 'Prod. Contact',   desc: 'View + rename versions',  icon: Phone        },
  { value: 'PRODUCTION_ASSISTANT',    label: 'Prod. Assistant', desc: 'View only',               icon: Clapperboard },
  { value: 'VIEWER',                  label: 'Viewer',          desc: 'Read-only contact list',  icon: Eye          },
]

export function AuthPage() {
  const { login: ctxLogin } = useAuth()
  const [tab,      setTab]      = useState<'login' | 'register'>('login')
  const [email,    setEmail]    = useState('')
  const [password, setPassword] = useState('')
  const [username, setUsername] = useState('')
  const [role,     setRole]     = useState<ProjectRole>('ATTORNEY')
  const [loading,  setLoading]  = useState(false)
  const [error,    setError]    = useState<string | null>(null)

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault()
    setLoading(true)
    setError(null)
    try {
      const res = tab === 'login'
        ? await login({ email, password })
        : await register({ username, email, password, role })
      ctxLogin({
        userId: res.userId, username: res.username,
        email: res.email, role: res.role as ProjectRole, token: res.token,
      })
    } catch (e: unknown) {
      setError(
        (e as { response?: { data?: { message?: string } } })?.response?.data?.message
        ?? 'Authentication failed'
      )
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="w-full max-w-md">

        {/* Logo */}
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-2xl
                          border border-emerald-300 bg-emerald-50 mb-4">
            <ShieldCheck size={22} className="text-emerald-600" />
          </div>
          <h1 className="font-display text-2xl text-slate-900">
            Script<span className="text-emerald-600">Sentries</span>
          </h1>
          <p className="text-slate-400 text-sm mt-1">Zero-Retention Script Clearance</p>
        </div>

        {/* Card */}
        <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">

          {/* Tabs */}
          <div className="flex border-b border-slate-100">
            {(['login', 'register'] as const).map(t => (
              <button
                key={t}
                onClick={() => { setTab(t); setError(null) }}
                className={`flex-1 py-3.5 text-sm font-medium capitalize transition-colors ${
                  tab === t
                    ? 'text-emerald-600 border-b-2 border-emerald-500 bg-emerald-50/50'
                    : 'text-slate-400 hover:text-slate-600'
                }`}>
                {t === 'login' ? 'Sign In' : 'Register'}
              </button>
            ))}
          </div>

          <form onSubmit={handleSubmit} className="p-6 space-y-4">

            {tab === 'register' && (
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1.5">Username</label>
                <div className="relative">
                  <User size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text" value={username} onChange={e => setUsername(e.target.value)}
                    placeholder="yourhandle"
                    required
                    className="w-full pl-9 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-lg
                               text-slate-800 text-sm focus:outline-none focus:border-emerald-400
                               focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all"
                  />
                </div>
              </div>
            )}

            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1.5">Email</label>
              <div className="relative">
                <Mail size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="email" value={email} onChange={e => setEmail(e.target.value)}
                  placeholder="you@studio.com"
                  required
                  className="w-full pl-9 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-lg
                             text-slate-800 text-sm focus:outline-none focus:border-emerald-400
                             focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-500 mb-1.5">Password</label>
              <div className="relative">
                <Lock size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="password" value={password} onChange={e => setPassword(e.target.value)}
                  placeholder="••••••••"
                  required
                  className="w-full pl-9 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-lg
                             text-slate-800 text-sm focus:outline-none focus:border-emerald-400
                             focus:ring-2 focus:ring-emerald-100 placeholder-slate-300 transition-all"
                />
              </div>
            </div>

            {/* Role selector */}
            {tab === 'register' && (
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-2">Your Role</label>
                <div className="grid gap-2">
                  {ROLES.map(r => {
                    const IconComp = r.icon
                    const selected = role === r.value
                    return (
                      <button
                        key={r.value}
                        type="button"
                        onClick={() => setRole(r.value)}
                        className={`flex items-center gap-3 px-3 py-2.5 rounded-xl border text-left
                                   transition-all ${
                          selected
                            ? 'border-emerald-400 bg-emerald-50 ring-2 ring-emerald-100'
                            : 'border-slate-200 hover:border-slate-300 bg-white'
                        }`}>
                        <div className={`p-1.5 rounded-lg ${selected ? 'bg-emerald-100' : 'bg-slate-100'}`}>
                          <IconComp size={13} className={selected ? 'text-emerald-600' : 'text-slate-400'} />
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className={`text-xs font-medium ${selected ? 'text-emerald-700' : 'text-slate-700'}`}>
                            {r.label}
                          </p>
                          <p className="text-[10px] text-slate-400 truncate">{r.desc}</p>
                        </div>
                        {selected && (
                          <div className="w-2 h-2 rounded-full bg-emerald-500 flex-shrink-0" />
                        )}
                      </button>
                    )
                  })}
                </div>
              </div>
            )}

            {error && (
              <p className="text-xs text-red-500 bg-red-50 border border-red-200 rounded-lg px-3 py-2">
                {error}
              </p>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full btn-primary justify-center py-2.5 mt-2">
              {loading ? (
                <><Loader2 size={15} className="animate-spin" /> {tab === 'login' ? 'Signing in…' : 'Creating account…'}</>
              ) : (
                tab === 'login' ? 'Sign In' : 'Create Account'
              )}
            </button>
          </form>
        </div>

        <p className="text-center text-xs text-slate-400 mt-6">
          Your scripts are never stored. Zero retention, always.
        </p>
      </div>
    </div>
  )
}
