// src/components/CommentsPanel.tsx
import { useState, useEffect } from 'react'
import { MessageSquare, Send, Loader2 } from 'lucide-react'
import { MentionInput } from './MentionInput'
import { getComments, postComment, type CommentResponse } from '../api/authApi'
import { useAuth } from '../context/AuthContext'

interface Props {
  riskFlagId: number
  projectId?: number
}

const COMMENT_TYPES = ['Main', 'Legal', 'Notes', 'To Research', 'To Production', 'Contacts', 'Other'] as const
type CommentType = typeof COMMENT_TYPES[number]

const COMMENT_TYPE_COLORS: Record<CommentType, string> = {
  'Main': 'bg-blue-100 text-blue-700 border-blue-200',
  'Legal': 'bg-violet-100 text-violet-700 border-violet-200',
  'Notes': 'bg-slate-100 text-slate-700 border-slate-200',
  'To Research': 'bg-amber-100 text-amber-700 border-amber-200',
  'To Production': 'bg-emerald-100 text-emerald-700 border-emerald-200',
  'Contacts': 'bg-pink-100 text-pink-700 border-pink-200',
  'Other': 'bg-gray-100 text-gray-700 border-gray-200',
}

function formatTime(iso: string) {
  return new Date(iso).toLocaleString('en-US', {
    month: 'short', day: 'numeric',
    hour: '2-digit', minute: '2-digit',
  })
}

/** Render @username mentions in blue */
function highlightMentions(text: string) {
  return text.split(/(@\w+)/g).map((part, i) =>
    part.startsWith('@')
      ? <span key={i} className="text-emerald-600 font-semibold">{part}</span>
      : <span key={i}>{part}</span>
  )
}

const ROLE_COLORS: Record<string, string> = {
  ATTORNEY:                'text-violet-700 bg-violet-50 border-violet-200',
  ANALYST:                 'text-blue-700   bg-blue-50   border-blue-200',
  MAIN_PRODUCTION_CONTACT: 'text-amber-700  bg-amber-50  border-amber-200',
  PRODUCTION_ASSISTANT:    'text-slate-600  bg-slate-100 border-slate-200',
  VIEWER:                  'text-zinc-500   bg-zinc-100  border-zinc-200',
}

export function CommentsPanel({ riskFlagId, projectId }: Props) {
  const { user } = useAuth()
  const [comments,    setComments]    = useState<CommentResponse[]>([])
  const [text,        setText]        = useState('')
  const [commentType, setCommentType] = useState<CommentType>('Main')
  const [loading,     setLoading]     = useState(true)
  const [submitting,  setSubmitting]  = useState(false)
  const [error,       setError]       = useState<string | null>(null)

  useEffect(() => {
    setLoading(true)
    getComments(riskFlagId)
      .then(setComments)
      .catch(() => setComments([]))
      .finally(() => setLoading(false))
  }, [riskFlagId])

  const handleSubmit = async () => {
    if (!text.trim() || !user) return
    setSubmitting(true)
    setError(null)
    try {
      const comment = await postComment({
        text: text.trim(),
        riskFlagId,
        authorId: user.userId,
        type: commentType,
      })
      setComments(prev => [...prev, comment])
      setText('')
    } catch {
      setError('Failed to post comment. Please try again.')
    } finally {
      setSubmitting(false)
    }
  }

  return (
    <div className="space-y-4">
      <h3 className="flex items-center gap-2 text-[10px] font-bold text-slate-400 uppercase tracking-widest">
        <MessageSquare size={11} />
        Team Comments
        {comments.length > 0 && (
          <span className="text-[9px] px-1.5 py-0.5 rounded-full bg-slate-100 text-slate-500 font-medium normal-case">
            {comments.length}
          </span>
        )}
      </h3>

      {/* Comment list */}
      <div className="space-y-2.5 max-h-56 overflow-y-auto pr-0.5">
        {loading ? (
          <div className="flex items-center justify-center py-8">
            <Loader2 size={16} className="animate-spin text-slate-300" />
          </div>
        ) : comments.length === 0 ? (
          <div className="text-center py-6 text-slate-400 text-xs bg-slate-50 rounded-xl border border-slate-100">
            No comments yet — be the first to comment
          </div>
        ) : (
          comments.map(comment => (
            <div key={comment.id}
              className="bg-slate-50 border border-slate-100 rounded-xl p-3">
              <div className="flex items-center gap-2 mb-1.5 flex-wrap">
                {/* Avatar */}
                <div className="w-6 h-6 rounded-full bg-emerald-100 border border-emerald-200
                                flex items-center justify-center flex-shrink-0">
                  <span className="text-[10px] font-bold text-emerald-700">
                    {comment.author.username.charAt(0).toUpperCase()}
                  </span>
                </div>
                <span className="text-xs font-semibold text-slate-800">@{comment.author.username}</span>
                <span className={`text-[9px] px-1.5 py-0.5 rounded border font-semibold ${
                  ROLE_COLORS[comment.author.role] ?? 'text-slate-500 bg-slate-100 border-slate-200'
                }`}>
                  {comment.author.role.replace(/_/g, ' ')}
                </span>
                {comment.type && (
                  <span className={`text-[9px] px-1.5 py-0.5 rounded border font-medium ${
                    COMMENT_TYPE_COLORS[comment.type as CommentType] ?? 'bg-gray-100 text-gray-700 border-gray-200'
                  }`}>
                    {comment.type}
                  </span>
                )}
                <span className="text-[10px] text-slate-400 ml-auto">{formatTime(comment.createdAt)}</span>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed pl-8">
                {highlightMentions(comment.text)}
              </p>
            </div>
          ))
        )}
      </div>

      {/* Input */}
      {user ? (
        <div className="space-y-3">
          <div className="flex items-center gap-3 flex-wrap">
            <label className="text-xs font-semibold text-slate-600 uppercase tracking-wider">Type:</label>
            <select
              value={commentType}
              onChange={(e) => setCommentType(e.target.value as CommentType)}
              className="text-xs border border-slate-200 rounded px-2.5 py-1 bg-white text-slate-700 focus:outline-none focus:border-emerald-400 focus:ring-1 focus:ring-emerald-300"
            >
              {COMMENT_TYPES.map((type) => (
                <option key={type} value={type}>{type}</option>
              ))}
            </select>
            <span className={`text-[10px] px-2.5 py-1 rounded border font-medium ${COMMENT_TYPE_COLORS[commentType]}`}>
              {commentType}
            </span>
          </div>
          <MentionInput
            value={text}
            onChange={setText}
            placeholder="Comment… type @ to mention a colleague"
            rows={2}
            disabled={submitting}
            projectId={projectId}
          />
          {error && (
            <p className="text-xs text-red-500">{error}</p>
          )}
          <div className="flex justify-end">
            <button
              onClick={handleSubmit}
              disabled={!text.trim() || submitting}
              className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500
                         text-white text-xs font-medium rounded-xl transition-all
                         disabled:opacity-40 active:scale-95"
            >
              {submitting
                ? <><Loader2 size={12} className="animate-spin" /> Posting…</>
                : <><Send size={12} /> Post Comment</>
              }
            </button>
          </div>
        </div>
      ) : (
        <p className="text-xs text-slate-400 text-center py-3 bg-slate-50 rounded-xl border border-slate-100">
          Sign in to comment
        </p>
      )}
    </div>
  )
}