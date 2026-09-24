import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import logoLockup from './assets/genoffice-logo.svg'
import iconDocx from './assets/file-docx.svg'
import iconXlsx from './assets/file-xlsx.svg'
import iconPptx from './assets/file-pptx.svg'
import iconPdf from './assets/file-pdf.svg'
import iconMd from './assets/file-md.svg'
import type {
  AccountStatus,
  CloudProjectKind,
  CloudProjectsSnapshot,
  HomeApi,
  ProjectHomeApi,
  ProjectSummaryEntry,
  RecentEntry,
} from '../../shared/home-api'
import { useDismissablePopover } from '@genoffice/ui'
import { fileCountKey } from './counts'
import { listFilterExt, nextSelectAll } from './home-list'
import { useI18n } from './locale'
import type { I18n, StringKey } from './locale'
import { SettingsModal } from './SettingsModal'

declare global {
  interface Window {
    aiOffice: HomeApi
    aiOfficeProject?: ProjectHomeApi
  }
}

/** page size of the home list */
const PAGE_SIZE = 10

const FILE_ICONS: Record<string, string> = {
  docx: iconDocx,
  xlsx: iconXlsx,
  xlsm: iconXlsx,
  pptx: iconPptx,
  pdf: iconPdf,
  md: iconMd,
  markdown: iconMd,
}

/* Formats the open-local card advertises. Too long for the card at any window
   width, so it ellipsizes and a hover ScreenTip carries the full list. Keep in
   sync with the main-process open-dialog filter (OPEN_DIALOG_EXTENSIONS). */
const OPEN_LOCAL_EXTENSIONS = '.docx / .xlsx / .xlsm / .xls / .csv / .pptx / .pdf / .md / .code / .txt / .log / .py / .js / .ts'


function FileBadge({ ext, size }: { ext: string; size: number }) {
  const icon = FILE_ICONS[ext]
  if (icon) {
    return <img src={icon} width={size} height={size} alt="" aria-hidden="true" />
  }
  // Bitable / multi-dim table — indigo T (must not collide with sheet green X)
  const isBitable = ext === 'table' || ext === 'bitable' || ext === 'bitable.json'
  const isDraw = ext === 'board' || ext === 'excalidraw' || ext === 'excalidraw.json'
  const isCode =
    ext === 'code' ||
    ['py', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'go', 'rs', 'java', 'vue', 'css', 'sh', 'sql'].includes(ext)
  const label = isBitable ? 'T' : isDraw ? 'D' : isCode ? '</>' : ext ? ext[0].toUpperCase() : '?'
  const fill = isBitable ? '#4F46E5' : isDraw ? '#EA580C' : isCode ? '#0F766E' : ext === 'txt' ? '#0F766E' : '#98a2b3'
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect width="32" height="32" rx="7.5" fill={fill} />
      <text
        x="16"
        y="16.5"
        textAnchor="middle"
        dominantBaseline="central"
        fill="#fff"
        fontSize={isCode ? 11 : 17}
        fontWeight="700"
        fontFamily="system-ui, -apple-system, 'Segoe UI', sans-serif"
      >
        {label}
      </text>
    </svg>
  )
}

function formatModified(mtimeMs: number, i18n: I18n): string {
  const date = new Date(mtimeMs)
  const now = new Date()
  const startOfDay = (d: Date) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime()
  const days = Math.round((startOfDay(now) - startOfDay(date)) / 86400000)
  if (days <= 0) {
    return `${i18n.t('today')} · ${date.toLocaleTimeString(i18n.dateLocale, { hour: '2-digit', minute: '2-digit' })}`
  }
  if (days === 1) return i18n.t('yesterday')
  return date.toLocaleDateString(i18n.dateLocale, { month: 'short', day: 'numeric' })
}

function formatSize(bytes: number): string {
  if (!bytes || bytes < 0) return '—'
  if (bytes >= 1048576) return `${(bytes / 1048576).toFixed(1)} MB`
  if (bytes < 1024) return `${bytes} B`
  return `${Math.round(bytes / 1024)} KB`
}

function parentDir(path: string): string {
  const parts = path.split(/[\\/]/).filter(Boolean)
  const parent = parts[parts.length - 2] ?? ''
  // moreai-artifact://{id}/name — never show bare numeric id as "location"
  if (/^moreai-artifact:/i.test(path) || /^\d+$/.test(parent)) return ''
  return parent
}

function decodePathSegment(segment: string): string {
  try {
    return decodeURIComponent(segment)
  } catch {
    return segment
  }
}

function fileName(path: string): string {
  const raw = path.split(/[\\/]/).pop() ?? path
  return decodePathSegment(raw)
}

function baseName(entry: RecentEntry): string {
  return entry.ext ? entry.name.slice(0, -(entry.ext.length + 1)) : entry.name
}

// ── Project hooks ─────────────────────────────────────────

/** whether we are inside the shell (aiOfficeProject API available) */
function hasProjectApi(): boolean {
  return typeof window.aiOfficeProject !== 'undefined'
}

const FILTERS: { key: string; label: StringKey }[] = [
  { key: 'all', label: 'filterAll' },
  { key: 'docx', label: 'filterDocs' },
  { key: 'xlsx', label: 'filterSheets' },
  { key: 'pptx', label: 'filterSlides' },
  { key: 'pdf', label: 'filterPdf' },
  { key: 'md', label: 'filterMd' },
  { key: 'table', label: 'filterBitable' },
  { key: 'code', label: 'filterCode' },
]

/** Check glyph marking the selected sort option; invisible on the others so labels stay aligned */
function SortCheck({ visible }: { visible: boolean }): ReactElement {
  return (
    <svg
      className="cloud-sort-check"
      width="12"
      height="12"
      viewBox="0 0 16 16"
      fill="none"
      aria-hidden="true"
      style={visible ? undefined : { visibility: 'hidden' }}
    >
      <path
        d="M3 8.5L6.5 12L13 4.5"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

// ── AI artifacts sidebar (MoreAI embed) ──────────────────

type AiSourceId = 'myTasks' | 'personalAi' | 'chat' | 'automation'

const AI_SOURCE_ITEMS: { id: AiSourceId; labelKey: 'aiSourceMyTasks' | 'aiSourcePersonalAi' | 'aiSourceChat' | 'aiSourceAutomation' }[] =
  [
    { id: 'myTasks', labelKey: 'aiSourceMyTasks' },
    { id: 'personalAi', labelKey: 'aiSourcePersonalAi' },
    { id: 'chat', labelKey: 'aiSourceChat' },
    { id: 'automation', labelKey: 'aiSourceAutomation' },
  ]

function AiOutputsPanel({
  selectedId,
  selectedSessionKey,
  onSelect,
  onSelectSession,
}: {
  selectedId: AiSourceId | null
  selectedSessionKey: string | null
  onSelect: (id: AiSourceId | null) => void
  onSelectSession: (sourceId: AiSourceId, sessionKey: string) => void
}) {
  const { t } = useI18n()
  const [sessions, setSessions] = useState<
    Array<{ sessionKey: string; title: string; fileCount: number }>
  >([])

  useEffect(() => {
    if (!selectedId || !window.aiOffice.listAiSessions) {
      setSessions([])
      return
    }
    let cancelled = false
    void window.aiOffice.listAiSessions(selectedId).then((rows) => {
      if (!cancelled) setSessions(rows || [])
    })
    const onReady = () => {
      void window.aiOffice.listAiSessions?.(selectedId).then((rows) => {
        if (!cancelled) setSessions(rows || [])
      })
    }
    window.addEventListener('moreai-ai-home-ready', onReady)
    return () => {
      cancelled = true
      window.removeEventListener('moreai-ai-home-ready', onReady)
    }
  }, [selectedId])

  return (
    <div className="proj-panel ai-outputs-panel">
      <div className="proj-panel-head">
        <span className="proj-panel-title">{t('aiOutputs')}</span>
      </div>
      <ul className="proj-list">
        {AI_SOURCE_ITEMS.map((item) => {
          const isActive = selectedId === item.id
          return (
            <li
              key={item.id}
              className={`proj-item${isActive ? ' active' : ''}${isActive ? ' ai-source-open' : ''}`}
            >
              <div
                className="proj-item-main"
                role="button"
                tabIndex={0}
                onClick={() => onSelect(isActive ? null : item.id)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault()
                    onSelect(isActive ? null : item.id)
                  }
                }}
              >
                <span className="proj-item-icon" aria-hidden="true">
                  {item.id === 'myTasks' ? (
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                      <path
                        d="M2.8 2.8h10.4v10.4H2.8zM5.6 8.1l1.7 1.7 3.5-3.6"
                        stroke="currentColor"
                        strokeWidth="1.3"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                  ) : item.id === 'personalAi' ? (
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                      <circle cx="8" cy="5.2" r="2.6" stroke="currentColor" strokeWidth="1.3" />
                      <path
                        d="M2.8 13.2c.8-2.5 2.6-3.7 5.2-3.7s4.4 1.2 5.2 3.7"
                        stroke="currentColor"
                        strokeWidth="1.3"
                        strokeLinecap="round"
                      />
                    </svg>
                  ) : item.id === 'chat' ? (
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                      <path
                        d="M2.5 4A2 2 0 0 1 4.5 2h7A2 2 0 0 1 13.5 4v4.2a2 2 0 0 1-2 2H7.1L4.2 13.2V10.2H4.5a2 2 0 0 1-2-2V4z"
                        stroke="currentColor"
                        strokeWidth="1.3"
                        strokeLinejoin="round"
                      />
                    </svg>
                  ) : (
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                      <path
                        d="M2.5 3.8h4v4h-4zM9.5 3.8h4v4h-4zM5.5 8.5v2h5v-2M3.8 12.2h8.4"
                        stroke="currentColor"
                        strokeWidth="1.3"
                        strokeLinejoin="round"
                      />
                    </svg>
                  )}
                </span>
                <span className="proj-item-name">{t(item.labelKey)}</span>
              </div>
              {isActive ? (
                <div className="ai-session-list">
                  {sessions.length === 0 ? (
                    <div className="ai-session-empty">暂无会话</div>
                  ) : (
                    sessions.map((s) => {
                      const selected = selectedSessionKey === s.sessionKey
                      return (
                        <button
                          key={s.sessionKey}
                          type="button"
                          className={`ai-session-item${selected ? ' active' : ''}`}
                          title={s.title}
                          onClick={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                            onSelectSession(item.id, s.sessionKey)
                          }}
                        >
                          <span className="ai-session-icon" aria-hidden="true">
                            <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                              <path
                                d="M2.5 4.5h4l1 1.5h6v7h-11v-8.5z"
                                stroke="currentColor"
                                strokeWidth="1.2"
                                strokeLinejoin="round"
                              />
                            </svg>
                          </span>
                          <span className="ai-session-name">{s.title}</span>
                          {s.fileCount > 0 ? (
                            <span className="ai-session-count">{s.fileCount}</span>
                          ) : null}
                        </button>
                      )
                    })
                  )}
                </div>
              ) : null}
            </li>
          )
        })}
      </ul>
    </div>
  )
}

// ── Project sidebar component ────────────────────────────

interface ProjectPanelProps {
  projects: ProjectSummaryEntry[]
  selectedId: string | null
  onSelect: (id: string | null) => void
  onRefresh: () => void
}

function ProjectPanel({ projects, selectedId, onSelect, onRefresh }: ProjectPanelProps) {
  const { t } = useI18n()
  const [creating, setCreating] = useState(false)
  const [newName, setNewName] = useState('')
  // open menu id + fixed-position anchor (viewport coords), so the popup can
  // escape the scrollable project list without the list losing overflow-y
  const [projMenu, setProjMenu] = useState<{ id: string; top: number; right: number } | null>(null)
  const [renaming, setRenaming] = useState<{ id: string; value: string } | null>(null)
  const newInputRef = useRef<HTMLInputElement>(null)
  // wrap (… button + popup) of the row whose menu is open — the dismissal guard root
  const projMenuWrapRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (creating && newInputRef.current) newInputRef.current.focus()
  }, [creating])

  // unified dismissal: outside press, window blur, chrome press (tab strip / window drag)
  useDismissablePopover(projMenu !== null, () => setProjMenu(null), {
    inside: () => [projMenuWrapRef.current],
  })

  // also close on any scroll (the fixed-position popup would otherwise detach
  // from its row while the list scrolls)
  useEffect(() => {
    if (!projMenu) return
    const close = () => setProjMenu(null)
    window.addEventListener('scroll', close, true)
    return () => window.removeEventListener('scroll', close, true)
  }, [projMenu])

  const commitCreate = async () => {
    const name = newName.trim()
    setCreating(false)
    setNewName('')
    if (!name) return
    try {
      await window.aiOfficeProject?.createProject(name)
    } catch (error) {
      window.alert(error instanceof Error ? error.message : String(error))
      return
    }
    onRefresh()
  }

  const commitRename = async () => {
    if (!renaming) return
    const name = renaming.value.trim()
    const id = renaming.id
    setRenaming(null)
    if (!name) return
    try {
      await window.aiOfficeProject?.renameProject(id, name)
    } catch (error) {
      window.alert(error instanceof Error ? error.message : String(error))
      return
    }
    onRefresh()
  }

  // in-app confirm dialog (same style as the delete-files modal), not window.confirm
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null)

  const doDelete = (id: string) => {
    setProjMenu(null)
    setConfirmDeleteId(id)
  }

  const confirmDeleteNow = async () => {
    const id = confirmDeleteId
    setConfirmDeleteId(null)
    if (!id) return
    try {
      await window.aiOfficeProject?.deleteProject(id)
    } catch (error) {
      window.alert(error instanceof Error ? error.message : String(error))
      return
    }
    if (selectedId === id) onSelect(null)
    onRefresh()
  }

  useEffect(() => {
    if (!confirmDeleteId) return
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setConfirmDeleteId(null)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [confirmDeleteId])

  return (
    <div className="proj-panel">
      <div className="proj-panel-head">
        <span className="proj-panel-title">{t('projects')}</span>
        <button
          className="proj-add-btn"
          data-tip={t('newProject')}
          onClick={() => setCreating(true)}
          aria-label={t('newProject')}
        >
          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <path
              d="M7 1v12M1 7h12"
              stroke="currentColor"
              strokeWidth="1.7"
              strokeLinecap="round"
            />
          </svg>
        </button>
      </div>

      {creating && (
        <div className="proj-new-row">
          <input
            ref={newInputRef}
            className="proj-rename-input"
            placeholder={t('projectName')}
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onBlur={() => void commitCreate()}
            onKeyDown={(e) => {
              // IME (e.g. pinyin): Enter/Escape during composition only affects
              // the composition, it must not commit or cancel the field
              if (e.nativeEvent.isComposing) return
              if (e.key === 'Enter') void commitCreate()
              if (e.key === 'Escape') {
                setCreating(false)
                setNewName('')
              }
            }}
          />
        </div>
      )}

      <ul className="proj-list">
        {projects.map((proj) => {
          const isActive = selectedId === proj.id
          const isRenaming = renaming?.id === proj.id
          return (
            <li key={proj.id} className={`proj-item${isActive ? ' active' : ''}`}>
              <div
                className="proj-item-main"
                role="button"
                tabIndex={0}
                onClick={() => onSelect(isActive ? null : proj.id)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') onSelect(isActive ? null : proj.id)
                }}
              >
                <span className="proj-item-icon" aria-hidden="true">
                  <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                    <path
                      d="M1.5 4A1.5 1.5 0 0 1 3 2.5h3.1c.44 0 .85.19 1.13.52L8.4 4.4H13A1.5 1.5 0 0 1 14.5 5.9v5.6A1.5 1.5 0 0 1 13 13H3a1.5 1.5 0 0 1-1.5-1.5V4z"
                      stroke="currentColor"
                      strokeWidth="1.2"
                      strokeLinejoin="round"
                    />
                  </svg>
                </span>
                {isRenaming ? (
                  <input
                    className="proj-rename-input inline"
                    value={renaming.value}
                    autoFocus
                    onFocus={(e) => e.target.select()}
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => setRenaming({ id: proj.id, value: e.target.value })}
                    onBlur={() => void commitRename()}
                    onKeyDown={(e) => {
                      e.stopPropagation()
                      if (e.nativeEvent.isComposing) return
                      if (e.key === 'Enter') void commitRename()
                      if (e.key === 'Escape') setRenaming(null)
                    }}
                  />
                ) : (
                  <span className="proj-item-name">
                    {proj.isDefault ? t('defaultProject') : proj.name}
                  </span>
                )}
                <span className="proj-item-meta">
                  <span className="proj-item-count">{proj.fileCount}</span>
                </span>
              </div>

              {!proj.isDefault && (
                <div
                  className="proj-menu-wrap"
                  ref={projMenu?.id === proj.id ? projMenuWrapRef : undefined}
                >
                  <button
                    className="proj-more-btn"
                    aria-label={t('projMoreActions', { name: proj.name })}
                    aria-expanded={projMenu?.id === proj.id}
                    onClick={(e) => {
                      e.stopPropagation()
                      if (projMenu?.id === proj.id) {
                        setProjMenu(null)
                        return
                      }
                      const rect = e.currentTarget.getBoundingClientRect()
                      setProjMenu({
                        id: proj.id,
                        top: rect.bottom + 4,
                        right: window.innerWidth - rect.right,
                      })
                    }}
                  >
                    <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
                      <circle cx="3.2" cy="8" r="1.35" fill="currentColor" />
                      <circle cx="8" cy="8" r="1.35" fill="currentColor" />
                      <circle cx="12.8" cy="8" r="1.35" fill="currentColor" />
                    </svg>
                  </button>
                  {projMenu?.id === proj.id && (
                    <div
                      className="proj-menu"
                      role="menu"
                      style={{ top: projMenu.top, right: projMenu.right }}
                    >
                      <button
                        role="menuitem"
                        onClick={(e) => {
                          e.stopPropagation()
                          setProjMenu(null)
                          setRenaming({ id: proj.id, value: proj.name })
                        }}
                      >
                        {t('rename')}
                      </button>
                      <div className="row-menu-divider" />
                      <button
                        role="menuitem"
                        className="danger"
                        onClick={(e) => {
                          e.stopPropagation()
                          doDelete(proj.id)
                        }}
                      >
                        {t('deleteProject')}
                      </button>
                    </div>
                  )}
                </div>
              )}
            </li>
          )
        })}
      </ul>

      {confirmDeleteId &&
        (() => {
          // locale string is "title?\nbody" — split it across the dialog
          const [confirmTitle, ...confirmBody] = t('deleteProjectConfirm').split('\n')
          return (
            <div className="modal-overlay" onClick={() => setConfirmDeleteId(null)}>
              <div
                className="modal"
                role="dialog"
                aria-modal="true"
                aria-label={confirmTitle}
                onClick={(event) => event.stopPropagation()}
              >
                <h3>{confirmTitle}</h3>
                <p>{confirmBody.join('\n')}</p>
                <div className="modal-buttons">
                  <button
                    className="btn btn-secondary"
                    autoFocus
                    onClick={() => setConfirmDeleteId(null)}
                  >
                    {t('cancel')}
                  </button>
                  <button className="btn btn-danger" onClick={() => void confirmDeleteNow()}>
                    {t('delete')}
                  </button>
                </div>
              </div>
            </div>
          )
        })()}
    </div>
  )
}

// ── Account entry (bottom-left) ──────────────────────────
// Currently the Genspark (gsk) login entry; to be upgraded to a signup/account system later.
// Clicking it opens the settings modal directly (SettingsModal.tsx), which hosts
// login/logout plus preferences (language, theme, save location, update channel).

const LOGIN_POLL_MS = 2500
/** fallback deadline when the CLI does not report expires_in (device codes live ~300s) */
const LOGIN_MAX_WAIT_MS = 300_000

function AccountEntry({
  onStatusChange,
}: {
  onStatusChange?: (status: AccountStatus | null) => void
}) {
  const { t } = useI18n()
  const [status, setStatus] = useState<AccountStatus | null>(null)

  useEffect(() => {
    onStatusChange?.(status)
  }, [status, onStatusChange])
  const [waiting, setWaiting] = useState(false)
  // incremented on login retry, resetting the polling timer
  const [loginNonce, setLoginNonce] = useState(0)
  const [loginError, setLoginError] = useState<
    'timeout' | 'launch' | 'network' | 'expired' | 'failed' | null
  >(null)
  // auth URL reported by the login CLI — rescue entry when the browser did not open
  const [authUrl, setAuthUrl] = useState<string | null>(null)
  const [urlCopied, setUrlCopied] = useState(false)
  const loginDeadline = useRef(0)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [loggingOut, setLoggingOut] = useState(false)
  // bumped on logout so an in-flight status refresh (which can still
  // report logged-in) is discarded instead of resurrecting the UI
  const statusSeq = useRef(0)

  // query login state once on mount
  useEffect(() => {
    let alive = true
    void window.aiOffice.accountStatus?.().then((s) => {
      if (alive) setStatus(s)
    })
    return () => {
      alive = false
    }
  }, [])

  // login progress pushed from main (gsk login CLI output)
  useEffect(() => {
    const off = window.aiOffice.onAccountLogin?.((ev) => {
      if (ev.phase === 'url') {
        if (ev.url) setAuthUrl(ev.url)
        if (ev.expiresInSec) loginDeadline.current = Date.now() + ev.expiresInSec * 1000
      } else if (ev.phase === 'success') {
        void window.aiOffice.accountStatus().then((s) => {
          if (s.loggedIn) {
            setStatus(s)
            setWaiting(false)
            setAuthUrl(null)
          }
        })
      } else if (ev.phase === 'error') {
        setWaiting(false)
        setAuthUrl(null)
        setLoginError(
          ev.error === 'network' ? 'network' : ev.error === 'expired' ? 'expired' : 'failed',
        )
      }
    })
    return off
  }, [])

  // config-file polling stays as the fallback success path (works even if progress events are lost)
  useEffect(() => {
    if (!waiting) return
    const timer = setInterval(() => {
      void window.aiOffice.accountStatus().then((s) => {
        if (s.loggedIn) {
          setStatus(s)
          setWaiting(false)
          setAuthUrl(null)
        } else if (Date.now() > loginDeadline.current) {
          setWaiting(false)
          setAuthUrl(null)
          setLoginError('timeout')
        }
      })
    }, LOGIN_POLL_MS)
    return () => clearInterval(timer)
  }, [waiting, loginNonce])

  const loggedIn = status?.loggedIn ?? false
  const email = status?.email ?? ''
  const initial = email ? email[0].toUpperCase() : loggedIn ? 'G' : '?'
  const errorText = loginError
    ? {
        timeout: t('loginTimeout'),
        launch: t('loginLaunchFailed'),
        network: t('loginNetworkError'),
        expired: t('loginExpired'),
        failed: t('loginFailed'),
      }[loginError]
    : null

  const doLogout = () => {
    setLoggingOut(true)
    statusSeq.current++
    void window.aiOffice.accountLogout().then(() => {
      setLoggingOut(false)
      setStatus({ loggedIn: false })
    })
  }

  const startLogin = () => {
    // clicking again while waiting = relaunch the login (main kills the stale CLI, so the new device code is the live one)
    setLoginError(null)
    setWaiting(true)
    setAuthUrl(null)
    setUrlCopied(false)
    loginDeadline.current = Date.now() + LOGIN_MAX_WAIT_MS
    setLoginNonce((n) => n + 1)
    void window.aiOffice.accountLogin().then((launched) => {
      if (!launched) {
        setWaiting(false)
        setLoginError('launch')
      }
    })
  }

  const openLoginUrl = () => void window.aiOffice.openLoginUrl?.()

  const copyLoginUrl = () => {
    if (!authUrl) return
    void navigator.clipboard.writeText(authUrl).then(() => {
      setUrlCopied(true)
      window.setTimeout(() => setUrlCopied(false), 2000)
    })
  }

  const handleClick = () => {
    // refresh the login state / credit balance; drop the response
    // when a logout happened while it was in flight
    const seq = statusSeq.current
    void window.aiOffice.accountStatus?.().then((s) => {
      if (seq === statusSeq.current) setStatus(s)
    })
    setSettingsOpen(true)
  }

  return (
    <div className="account-entry">
      {settingsOpen && (
        <SettingsModal
          status={status}
          loggingOut={loggingOut}
          loginWaiting={waiting}
          loginUrl={authUrl}
          urlCopied={urlCopied}
          onOpenLoginUrl={openLoginUrl}
          onCopyLoginUrl={copyLoginUrl}
          onClose={() => setSettingsOpen(false)}
          onLogin={() => {
            setSettingsOpen(false)
            startLogin()
          }}
          onLogout={doLogout}
        />
      )}
      {!settingsOpen && waiting && authUrl && (
        <div className="login-hint" role="status">
          <button className="login-hint-open" onClick={openLoginUrl}>
            {t('loginOpenShort')}
          </button>
          <button
            className={`login-hint-copy${urlCopied ? ' copied' : ''}`}
            onClick={copyLoginUrl}
            // static tip: screentips are suppressed from pointerdown until the pointer
            // leaves the control, so a swapped-in "copied" tip would never show — the
            // check-mark icon is the visible feedback
            data-tip={t('loginCopyUrl')}
            aria-label={urlCopied ? t('loginCopied') : t('loginCopyUrl')}
          >
            {urlCopied ? (
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <path
                  d="m3.5 8.5 3 3 6-7"
                  stroke="currentColor"
                  strokeWidth="1.6"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            ) : (
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <rect
                  x="5.5"
                  y="5.5"
                  width="7"
                  height="7"
                  rx="1.5"
                  stroke="currentColor"
                  strokeWidth="1.3"
                />
                <path
                  d="M3.5 10.5V5a1.5 1.5 0 0 1 1.5-1.5h5.5"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinecap="round"
                />
              </svg>
            )}
          </button>
        </div>
      )}
      <button
        className="account-btn"
        onClick={handleClick}
        aria-haspopup="dialog"
        aria-expanded={settingsOpen}
        data-tip={
          loggedIn
            ? email || t('loggedInGenspark')
            : waiting
              ? t('waitingLogin')
              : (errorText ?? t('loginGenspark'))
        }
        aria-label={t('settings')}
      >
        <span
          className={`account-avatar${loggedIn ? ' logged-in' : ''}${waiting ? ' waiting' : ''}`}
        >
          {waiting ? (
            <svg
              className="account-spinner"
              width="14"
              height="14"
              viewBox="0 0 16 16"
              aria-hidden="true"
            >
              <circle
                cx="8"
                cy="8"
                r="6"
                stroke="currentColor"
                strokeWidth="1.8"
                fill="none"
                strokeDasharray="26"
                strokeDashoffset="18"
                strokeLinecap="round"
              />
            </svg>
          ) : (
            initial
          )}
        </span>
        <span className="account-text">
          <span className="account-name">
            {loggedIn
              ? email
                ? email.split('@')[0]
                : t('loggedIn')
              : waiting
                ? t('waitingShort')
                : t('login')}
          </span>
          {!loggedIn && !waiting && errorText && (
            <span className="account-sub error">{errorText}</span>
          )}
        </span>
        <svg
          className="account-chevron"
          width="14"
          height="14"
          viewBox="0 0 16 16"
          fill="none"
          aria-hidden="true"
        >
          <path
            d="M5 6.2 8 3.4l3 2.8M5 9.8l3 2.8 3-2.8"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
    </div>
  )
}

function isMoreaiHomeEmbed(): boolean {
  return typeof document !== 'undefined' && document.body.classList.contains('moreai-home-embed')
}

type ChatModelOption = {
  id: string
  label: string
  factory: string
  iconKey?: string
  iconSvg?: string
}

function MoreaiFactoryIcon({
  factory,
  iconKey,
  iconSvg,
  size = 16,
}: {
  factory: string
  iconKey?: string
  iconSvg?: string
  size?: number
}) {
  const [svgFailed, setSvgFailed] = useState(false)
  if (iconSvg && !svgFailed) {
    return (
      <img
        className="moreai-model-icon"
        src={iconSvg}
        width={size}
        height={size}
        alt=""
        onError={() => setSvgFailed(true)}
      />
    )
  }
  if (iconKey) {
    return (
      <svg className="moreai-model-icon" width={size} height={size} aria-hidden="true">
        <use href={`#icon-${iconKey}`} xlinkHref={`#icon-${iconKey}`} />
      </svg>
    )
  }
  return (
    <span className="moreai-model-letter" aria-hidden="true">
      {(factory || '?').slice(0, 1).toUpperCase()}
    </span>
  )
}

/** MoreAI embed: model picker (account lives in MoreAI sidebar). */
function MoreaiAccountModelEntry({
  onStatusChange,
}: {
  onStatusChange?: (status: AccountStatus | null) => void
}) {
  const { t } = useI18n()
  const [status, setStatus] = useState<AccountStatus | null>(null)
  const [models, setModels] = useState<ChatModelOption[]>([])
  const [modelId, setModelId] = useState<string>('')
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    onStatusChange?.(status)
  }, [status, onStatusChange])

  useEffect(() => {
    let alive = true
    void window.aiOffice.accountStatus?.().then((s) => {
      if (alive) setStatus(s)
    })
    const api = window.aiOffice as HomeApi & {
      listChatModels?: () => Promise<ChatModelOption[]>
      getChatModel?: () => Promise<string | null>
      setChatModel?: (id: string | null) => Promise<void>
    }
    void Promise.all([api.listChatModels?.() ?? Promise.resolve([]), api.getChatModel?.() ?? Promise.resolve(null)])
      .then(([list, selected]) => {
        if (!alive) return
        setModels(list)
        setModelId(selected || list[0]?.id || '')
      })
      .catch(() => {
        if (!alive) return
        setModels([])
        setModelId('')
      })
    return () => {
      alive = false
    }
  }, [])

  useEffect(() => {
    if (!open) return
    const onDoc = (e: MouseEvent) => {
      if (!wrapRef.current?.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', onDoc)
    return () => document.removeEventListener('mousedown', onDoc)
  }, [open])

  const selected = models.find((m) => m.id === modelId) || models[0]
  const modelLabel = selected?.label || '选择模型'

  const pickModel = (id: string) => {
    setModelId(id)
    setOpen(false)
    const api = window.aiOffice as HomeApi & {
      setChatModel?: (id: string | null) => Promise<void>
    }
    void api.setChatModel?.(id)
  }

  return (
    <div className="account-entry moreai-model-entry" ref={wrapRef}>
      <div className="proj-panel-head moreai-model-head">
        <span className="proj-panel-title">{t('officeAiModelTitle')}</span>
      </div>
      <button
        type="button"
        className={`nav-item moreai-model-btn${open ? ' active' : ''}`}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={t('officeAiModelTitle')}
        onClick={() => setOpen((v) => !v)}
      >
        <MoreaiFactoryIcon
          factory={selected?.factory || ''}
          iconKey={selected?.iconKey}
          iconSvg={selected?.iconSvg}
          size={16}
        />
        <span className="nav-label">{modelLabel}</span>
        <svg
          className="moreai-model-chevron"
          width="12"
          height="12"
          viewBox="0 0 16 16"
          fill="none"
          aria-hidden="true"
        >
          <path
            d={open ? 'M4 10.5 8 6.5l4 4' : 'M4 6.5 8 10.5l4-4'}
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
      <p className="moreai-model-desc">{t('officeAiModelHint')}</p>
      {open && (
        <div className="moreai-model-menu" role="listbox">
          {models.length === 0 ? (
            <div className="moreai-model-empty">模型列表加载失败，请刷新后重试</div>
          ) : (
            models.map((m) => (
              <button
                key={m.id}
                type="button"
                className={m.id === modelId ? 'active' : undefined}
                role="option"
                aria-selected={m.id === modelId}
                onClick={() => pickModel(m.id)}
              >
                <MoreaiFactoryIcon factory={m.factory} iconKey={m.iconKey} iconSvg={m.iconSvg} size={16} />
                <span className="moreai-model-label">{m.label}</span>
                <span className="moreai-model-factory">{m.factory}</span>
              </button>
            ))
          )}
        </div>
      )}
    </div>
  )
}

// ── Cloud (Genspark web) projects view ──────────────────

/** kind filter segments; labels shared with the recents type filter */
const CLOUD_FILTERS = [
  { key: 'all', label: 'filterAll' },
  { key: 'docs', label: 'filterDocs' },
  { key: 'sheets', label: 'filterSheets' },
  { key: 'slides', label: 'filterSlides' },
] as const satisfies readonly { key: 'all' | CloudProjectKind; label: StringKey }[]

/** module kind → file icon extension */
const CLOUD_KIND_EXT: Record<string, string> = { docs: 'docx', sheets: 'xlsx', slides: 'pptx' }

/** rows revealed per "load more" step; purely client-side over the local snapshot */
const CLOUD_REVEAL_STEP = 100

function CloudProjectsView() {
  const i18n = useI18n()
  const { t } = i18n
  const [snapshot, setSnapshot] = useState<CloudProjectsSnapshot | null>(null)
  const [loading, setLoading] = useState(true)
  const [syncing, setSyncing] = useState(false)
  const [loginWaiting, setLoginWaiting] = useState(false)
  const [kind, setKind] = useState<'all' | CloudProjectKind>('all')
  const [query, setQuery] = useState('')
  const [sort, setSort] = useState<'recent' | 'oldest'>('recent')
  const [sortMenuOpen, setSortMenuOpen] = useState(false)
  const [revealed, setRevealed] = useState(CLOUD_REVEAL_STEP)
  const sortRef = useRef<HTMLDivElement>(null)

  // the local store paints instantly; a background sync replaces it when done.
  // a failed sync keeps whatever is shown; with nothing shown the
  // !snapshot && !loading branch below renders the retry state
  const startSync = () => {
    setSyncing(true)
    void window.aiOffice.cloudProjectsSync?.().then((synced) => {
      setSyncing(false)
      setLoading(false)
      if (synced) setSnapshot(synced)
    })
  }
  const startSyncRef = useRef(startSync)
  startSyncRef.current = startSync

  useEffect(() => {
    let cancelled = false
    void window.aiOffice.cloudProjectsCached?.().then((stored) => {
      if (cancelled || !stored) return
      setSnapshot((prev) => prev ?? stored)
      setLoading(false)
    })
    startSyncRef.current()
    return () => {
      cancelled = true
    }
  }, [])

  // the sign-in button reuses the account login flow; sync once it lands
  useEffect(() => {
    const off = window.aiOffice.onAccountLogin?.((ev) => {
      if (ev.phase === 'success') {
        setLoginWaiting(false)
        startSyncRef.current()
      } else if (ev.phase === 'error') {
        setLoginWaiting(false)
      }
    })
    return off
  }, [])

  // unified dismissal: outside press, window blur, chrome press (tab strip / window drag)
  useDismissablePopover(sortMenuOpen, () => setSortMenuOpen(false), {
    inside: () => [sortRef.current],
  })

  const startLogin = () => {
    setLoginWaiting(true)
    void window.aiOffice.accountLogin?.().then((ok) => {
      if (!ok) setLoginWaiting(false)
    })
  }

  const changeKind = (k: 'all' | CloudProjectKind) => {
    if (k === kind) return
    setKind(k)
    setRevealed(CLOUD_REVEAL_STEP)
  }

  const openProject = (projectUrl: string) => {
    void window.aiOffice.openCloudProject?.(projectUrl)
  }

  // filter / search / sort are all local over the snapshot — no requests
  const q = query.trim().toLowerCase()
  let list = snapshot?.projects.filter((proj) => kind === 'all' || proj.kind === kind) ?? []
  if (q) list = list.filter((proj) => proj.title.toLowerCase().includes(q))
  if (sort === 'oldest') list = [...list].reverse()
  const visible = list.slice(0, revealed)

  const renderRows = () => {
    const items: ReactElement[] = []
    for (const proj of visible) {
      items.push(
        <li key={proj.projectId}>
          <button
            className="cloud-row"
            data-tip={t('cloudOpenInBrowser')}
            data-tip-anchor=".cloud-row-external"
            data-tip-place="right"
            onClick={() => openProject(proj.projectUrl)}
          >
            <FileBadge ext={CLOUD_KIND_EXT[proj.kind] ?? ''} size={24} />
            <span className="cloud-row-main">
              <span className="cloud-row-title">{proj.title || t('untitled')}</span>
              <svg
                className="cloud-row-external"
                width="13"
                height="13"
                viewBox="0 0 16 16"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.5 3.5H4a1.5 1.5 0 0 0-1.5 1.5v7A1.5 1.5 0 0 0 4 13.5h7A1.5 1.5 0 0 0 12.5 12V9.5M9.5 2.5h4v4M13 3l-5.5 5.5"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </span>
            <span className="cloud-row-time">
              {proj.ctimeMs ? formatModified(proj.ctimeMs, i18n) : ''}
            </span>
          </button>
        </li>,
      )
    }
    return items
  }

  const renderBody = () => {
    if (snapshot && !snapshot.available) {
      return (
        <p className="empty proj-empty">
          <span className="empty-hint">{t('cloudLoginHint')}</span>
          <button className="btn btn-secondary" disabled={loginWaiting} onClick={startLogin}>
            {loginWaiting ? t('waitingShort') : t('loginGenspark')}
          </button>
        </p>
      )
    }
    if (!snapshot) {
      if (loading || syncing) {
        return (
          <div className="load-more" aria-hidden="true">
            <span className="load-more-spinner" />
          </div>
        )
      }
      return (
        <p className="empty proj-empty">
          <span className="empty-hint">{t('cloudError')}</span>
          <button className="btn btn-secondary" onClick={() => startSync()}>
            {t('cloudRetry')}
          </button>
        </p>
      )
    }
    if (list.length === 0) {
      return (
        <p className="empty proj-empty">
          <span className="empty-hint">
            {t(q ? 'cloudNoResults' : kind === 'all' ? 'cloudEmpty' : 'emptyFiltered')}
          </span>
        </p>
      )
    }
    return (
      <div className="cloud-scroll">
        <div className="cloud-table">
          <div className="cloud-columns">
            <span className="col-name">{t('colName')}</span>
            <div className="cloud-col-sort" ref={sortRef}>
              <button
                className="cloud-col-sort-btn"
                aria-haspopup="menu"
                aria-expanded={sortMenuOpen}
                onClick={() => setSortMenuOpen((o) => !o)}
              >
                {t('colModified')}
                <svg
                  width="12"
                  height="12"
                  viewBox="0 0 16 16"
                  fill="none"
                  aria-hidden="true"
                  style={sort === 'oldest' ? { transform: 'rotate(180deg)' } : undefined}
                >
                  <path
                    d="M8 3v10M4.5 9.5L8 13l3.5-3.5"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
              {sortMenuOpen && (
                <div className="cloud-sort-menu" role="menu">
                  {(['recent', 'oldest'] as const).map((key) => (
                    <button
                      key={key}
                      className={sort === key ? 'active' : ''}
                      role="menuitemradio"
                      aria-checked={sort === key}
                      onClick={() => {
                        setSort(key)
                        setSortMenuOpen(false)
                        setRevealed(CLOUD_REVEAL_STEP)
                      }}
                    >
                      <SortCheck visible={sort === key} />
                      {t(key === 'recent' ? 'cloudSortRecent' : 'cloudSortOldest')}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
          <ul className="cloud-list">{renderRows()}</ul>
        </div>
        {list.length > revealed && (
          <div className="load-more">
            <button
              className="btn btn-secondary"
              onClick={() => setRevealed((n) => n + CLOUD_REVEAL_STEP)}
            >
              {t('cloudLoadMore')}
            </button>
          </div>
        )}
      </div>
    )
  }

  return (
    <main className="content">
      <section className="cloud-projects" aria-label={t('navCloud')}>
        <header className="cloud-hero">
          <div className="cloud-hero-top">
            <h1 className="cloud-title">{t('navCloud')}</h1>
          </div>
          <p className="cloud-subtitle">{t('cloudSubtitle')}</p>
          {snapshot?.available && (
            <div className="cloud-controls">
              <div className="cloud-seg" role="tablist" aria-label={t('filterAria')}>
                {CLOUD_FILTERS.map((f) => (
                  <button
                    key={f.key}
                    className={kind === f.key ? 'active' : ''}
                    role="tab"
                    aria-selected={kind === f.key}
                    onClick={() => changeKind(f.key)}
                  >
                    {t(f.label)}
                  </button>
                ))}
              </div>
              <button
                className={`cloud-refresh-btn${syncing ? ' syncing' : ''}`}
                data-tip={t('cloudRefresh')}
                aria-label={t('cloudRefresh')}
                disabled={syncing}
                onClick={() => startSync()}
              >
                <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <path
                    d="M13.6 8a5.6 5.6 0 1 1-1.64-3.96M13.6 2.4v3.2h-3.2"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
              <div className="cloud-search">
                <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <circle cx="7" cy="7" r="4.6" stroke="currentColor" strokeWidth="1.4" />
                  <path
                    d="M10.5 10.5L14 14"
                    stroke="currentColor"
                    strokeWidth="1.4"
                    strokeLinecap="round"
                  />
                </svg>
                <input
                  value={query}
                  placeholder={t('cloudSearchPlaceholder', { n: snapshot.projects.length })}
                  onChange={(e) => {
                    setQuery(e.target.value)
                    setRevealed(CLOUD_REVEAL_STEP)
                  }}
                />
              </div>
            </div>
          )}
        </header>
        {renderBody()}
      </section>
    </main>
  )
}

// ── Drop-to-open overlay ────────────────────────────────

/**
 * Full-window affordance while OS files hover over Home. Purely visual — the
 * actual open is owned by the preload drop bridge (installDropOpenBridge), so
 * this overlay stays pointer-events:none and never handles events itself.
 * Visibility tracks a dragenter/dragleave depth counter: `dragover` stops
 * being delivered while the cursor is stationary (macOS), so a debounce would
 * hide the overlay mid-drag. Enter fires before the matching leave when
 * moving between elements, so the depth never dips to zero inside the window.
 */
function DropToOpenOverlay(): ReactElement | null {
  const [visible, setVisible] = useState(false)
  useEffect(() => {
    let depth = 0
    const hasFiles = (ev: DragEvent): boolean => ev.dataTransfer?.types.includes('Files') ?? false
    // NB: the preload drop bridge also listens here and cancels file drags, so
    // defaultPrevented can't discriminate anything at this layer — only zones
    // that stopPropagation (none on Home) would keep us out entirely.
    const onDragEnter = (ev: DragEvent) => {
      if (!hasFiles(ev)) return
      depth += 1
      setVisible(true)
    }
    const onDragLeave = (ev: DragEvent) => {
      if (!hasFiles(ev)) return
      depth = Math.max(0, depth - 1)
      if (depth === 0) setVisible(false)
    }
    // drop/blur reset the depth outright: leaving the window mid-drag can eat
    // a dragleave, and a stuck overlay would be worse than a re-shown one
    const onHide = () => {
      depth = 0
      setVisible(false)
    }
    window.addEventListener('dragenter', onDragEnter)
    window.addEventListener('dragleave', onDragLeave)
    window.addEventListener('drop', onHide)
    window.addEventListener('blur', onHide)
    return () => {
      window.removeEventListener('dragenter', onDragEnter)
      window.removeEventListener('dragleave', onDragLeave)
      window.removeEventListener('drop', onHide)
      window.removeEventListener('blur', onHide)
    }
  }, [])
  const { t } = useI18n()
  if (!visible) return null
  return (
    <div className="home-drop-overlay" aria-hidden="true">
      <div className="home-drop-card">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path
            d="M12 3.5v11M7.5 10.5l4.5 4.5 4.5-4.5"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
          <path
            d="M4 16.5v2A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5v-2"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
          />
        </svg>
        <h2>{t('dropToOpenTitle')}</h2>
        <p>{OPEN_LOCAL_EXTENSIONS}</p>
      </div>
    </div>
  )
}

// ── Main component ──────────────────────────────────────

export function Home() {
  const i18n = useI18n()
  const { t, lang } = i18n
  // ── Paged list state (rows loaded for the current view + filter) ──
  const [entries, setEntries] = useState<RecentEntry[]>([])
  /** total count under the current view + filter (not just the loaded rows) */
  const [listTotal, setListTotal] = useState(0)
  const [listPage, setListPage] = useState(1)
  const [jumpDraft, setJumpDraft] = useState('1')
  /** Sidebar Recent / Starred / Shared totals (not tied to the type filter chips). */
  const [navCounts, setNavCounts] = useState({ recent: 0, starred: 0, shared: 0 })
  const [view, setView] = useState<'recent' | 'starred' | 'shared'>('recent')
  // Genspark web projects take over the content area (like a selected project)
  const [cloudMode, setCloudMode] = useState(false)
  const [filter, setFilter] = useState('all')
  // modified-column sort (WPS-style header popover), shared by the global and project tables
  const [fileSort, setFileSort] = useState<'recent' | 'oldest'>('recent')
  const [fileSortMenuOpen, setFileSortMenuOpen] = useState(false)
  const fileSortRef = useRef<HTMLDivElement>(null)
  const [rowMenu, setRowMenu] = useState<string | null>(null)
  // actions cell (… button + menu) of the row whose menu is open — the dismissal guard root
  const rowMenuWrapRef = useRef<HTMLSpanElement>(null)
  const [selected, setSelected] = useState<ReadonlySet<string>>(new Set())
  const [renaming, setRenaming] = useState<{ path: string; value: string } | null>(null)
  const [confirmDelete, setConfirmDelete] = useState<string[] | null>(null)
  // delayed so fast duplicates don't flash a spinner
  const [duplicating, setDuplicating] = useState(false)
  const duplicatingShowTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  useEffect(() => {
    return () => {
      if (duplicatingShowTimer.current) clearTimeout(duplicatingShowTimer.current)
    }
  }, [])
  // Genspark Projects is web-account data, so its nav entry only shows when logged in
  const [loggedIn, setLoggedIn] = useState(false)
  const handleAccountStatus = useCallback((s: AccountStatus | null) => {
    const on = s?.loggedIn ?? false
    setLoggedIn(on)
    if (!on) setCloudMode(false)
  }, [])
  // ── Project state ──
  const [projects, setProjects] = useState<ProjectSummaryEntry[]>([])
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(null)
  /** MoreAI: filter Home file list to one AI 产物 source (null = 全部 mixed) */
  const [selectedAiSource, setSelectedAiSource] = useState<AiSourceId | null>(null)
  const [selectedAiSessionKey, setSelectedAiSessionKey] = useState<string | null>(null)
  // Refs so reload()/docs-changed see cleared filters immediately (setState is async).
  const selectedAiSourceRef = useRef(selectedAiSource)
  const selectedAiSessionKeyRef = useRef(selectedAiSessionKey)
  const selectedProjectIdRef = useRef(selectedProjectId)
  selectedAiSourceRef.current = selectedAiSource
  selectedAiSessionKeyRef.current = selectedAiSessionKey
  selectedProjectIdRef.current = selectedProjectId

  const projectMode = hasProjectApi()

  // ── Paged loading ──
  const requestSeq = useRef(0)
  const listPageRef = useRef(1)
  listPageRef.current = listPage
  const viewRef = useRef(view)
  viewRef.current = view
  const filterRef = useRef(filter)
  filterRef.current = filter
  const fileSortOrderRef = useRef(fileSort)
  fileSortOrderRef.current = fileSort

  const reloadNavCounts = (seq: number) => {
    void window.aiOffice.recents({ offset: 0, limit: 0 }).then((page) => {
      if (seq !== requestSeq.current) return
      setNavCounts((prev) => ({ ...prev, recent: page.totalAll || page.total || 0 }))
    })
    void window.aiOffice.starred({ offset: 0, limit: 0 }).then((page) => {
      if (seq !== requestSeq.current) return
      setNavCounts((prev) => ({ ...prev, starred: page.totalAll || page.total || 0 }))
    })
    void (window.aiOffice.shared?.({ offset: 0, limit: 0 }) ?? Promise.resolve(null)).then((page) => {
      if (seq !== requestSeq.current || !page) return
      setNavCounts((prev) => ({ ...prev, shared: page.totalAll || page.total || 0 }))
    })
  }

  const reload = (opts?: { resetPage?: boolean }) => {
    const seq = ++requestSeq.current
    if (opts?.resetPage) {
      listPageRef.current = 1
      setListPage(1)
      setJumpDraft('1')
    }
    const page = Math.max(1, listPageRef.current)
    const aiSource = selectedAiSourceRef.current || undefined
    const sessionKey = selectedAiSessionKeyRef.current || undefined
    const projectId = selectedProjectIdRef.current || undefined
    const ext = listFilterExt(filterRef.current)
    const listApi =
      projectId
        ? window.aiOffice.recents
        : viewRef.current === 'starred'
        ? window.aiOffice.starred
        : viewRef.current === 'shared'
          ? (window.aiOffice.shared ?? (async () => ({ entries: [], total: 0, totalAll: 0 })))
          : window.aiOffice.recents
    void listApi({
      offset: (page - 1) * PAGE_SIZE,
      limit: PAGE_SIZE,
      ext,
      aiSource,
      sessionKey,
      projectId,
      order: fileSortOrderRef.current,
    }).then((result) => {
      if (seq !== requestSeq.current) return
      setEntries(result.entries)
      setListTotal(result.total)
      const pages = Math.max(1, Math.ceil((result.total || 0) / PAGE_SIZE) || 1)
      if (page > pages) {
        listPageRef.current = pages
        setListPage(pages)
      }
    })
    reloadNavCounts(seq)
  }
  const reloadRef = useRef(reload)
  reloadRef.current = reload

  const [projectTick, setProjectTick] = useState(0)

  const refresh = () => {
    reloadRef.current()
    setProjectTick((n) => n + 1)
  }

  useEffect(() => {
    reloadRef.current()
  }, [view, selectedAiSource, selectedAiSessionKey, selectedProjectId, filter, listPage, fileSort])

  useEffect(() => {
    if (!projectMode) return
    void window.aiOfficeProject!.listProjects().then(setProjects)
  }, [projectMode, projectTick])

  useEffect(() => {
    let timer: number | null = null
    const onAiReady = () => {
      if (timer != null) window.clearTimeout(timer)
      timer = window.setTimeout(() => {
        timer = null
        reloadRef.current()
      }, 80)
    }
    const onDocsChanged = () => {
      selectedAiSourceRef.current = null
      selectedAiSessionKeyRef.current = null
      setSelectedAiSource(null)
      setSelectedAiSessionKey(null)
      setSelectedProjectId(null)
      setCloudMode(false)
      void window.aiOffice.openAiSource?.(null)
      reloadRef.current({ resetPage: true })
      setProjectTick((n) => n + 1)
    }
    window.addEventListener('moreai-ai-home-ready', onAiReady)
    window.addEventListener('moreai-home-docs-changed', onDocsChanged)
    return () => {
      if (timer != null) window.clearTimeout(timer)
      window.removeEventListener('moreai-ai-home-ready', onAiReady)
      window.removeEventListener('moreai-home-docs-changed', onDocsChanged)
    }
  }, [])

  useEffect(() => {
    setListPage(1)
    setJumpDraft('1')
  }, [selectedAiSource, selectedAiSessionKey, selectedProjectId])

  useEffect(() => {
    const onFocus = () => {
      reloadRef.current()
    }
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [])

  useEffect(() => {
    setJumpDraft(String(listPage))
  }, [listPage])

  const [aiSessionTitle, setAiSessionTitle] = useState<string | null>(null)
  useEffect(() => {
    if (!selectedAiSource || !selectedAiSessionKey || !window.aiOffice.listAiSessions) {
      setAiSessionTitle(null)
      return
    }
    let cancelled = false
    void window.aiOffice.listAiSessions(selectedAiSource).then((rows) => {
      if (cancelled) return
      const hit = (rows || []).find((s) => s.sessionKey === selectedAiSessionKey)
      setAiSessionTitle(hit?.title || null)
    })
    return () => {
      cancelled = true
    }
  }, [selectedAiSource, selectedAiSessionKey])

  useDismissablePopover(fileSortMenuOpen, () => setFileSortMenuOpen(false), {
    inside: () => [fileSortRef.current],
  })

  // unified dismissal: outside press, window blur, chrome press (tab strip / window drag)
  useDismissablePopover(rowMenu !== null, () => setRowMenu(null), {
    inside: () => [rowMenuWrapRef.current],
  })

  // Escape closes the row menu and the delete-confirm dialog
  useEffect(() => {
    if (rowMenu === null && confirmDelete === null) return
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setRowMenu(null)
        setConfirmDelete(null)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [rowMenu, confirmDelete])

  // ── Project files state ────────────────────────────────

  const [moveFileMenu, setMoveFileMenu] = useState<string | null>(null)
  // submenu opens rightward by default; flips left when the window edge is too close
  const [moveMenuFlip, setMoveMenuFlip] = useState(false)
  // hover-open/close delays: avoid flashing the submenu while the pointer passes
  // through, and keep it open while crossing the 4px gap into it
  const moveMenuTimers = useRef<{ open: number | null; close: number | null }>({
    open: null,
    close: null,
  })
  // wrap (trigger + submenu) of the row whose move submenu is open — the dismissal guard root
  const moveMenuWrapRef = useRef<HTMLDivElement>(null)

  const openMoveMenu = (path: string) => {
    setMoveMenuFlip(false)
    setMoveFileMenu(path)
  }

  // ref runs pre-paint, so measuring the real width (long project names exceed
  // the min-width) and flipping never flashes; once flipped the check no longer hits
  const measureSubmenu = (el: HTMLDivElement | null) => {
    if (el && el.getBoundingClientRect().right > document.documentElement.clientWidth - 8) {
      setMoveMenuFlip(true)
    }
  }

  const clearMoveMenuTimer = (kind: 'open' | 'close') => {
    const timers = moveMenuTimers.current
    if (timers[kind] !== null) {
      window.clearTimeout(timers[kind])
      timers[kind] = null
    }
  }
  const [bulkMoveMenu, setBulkMoveMenu] = useState(false)
  // selection-bar wrap (trigger + menu) of the bulk move menu — the dismissal guard root
  const bulkMoveWrapRef = useRef<HTMLSpanElement>(null)

  // the submenu lives inside the row menu: when that closes, drop the stale
  // submenu state and any pending hover timers so it doesn't reopen expanded
  useEffect(() => {
    if (rowMenu === null) {
      clearMoveMenuTimer('open')
      clearMoveMenuTimer('close')
      setMoveFileMenu(null)
    }
  }, [rowMenu])

  // move-file submenu: unified dismissal (outside press, window blur, chrome press)
  useDismissablePopover(moveFileMenu !== null, () => setMoveFileMenu(null), {
    inside: () => [moveMenuWrapRef.current],
  })

  // bulk move-to-project menu in the selection bar: unified dismissal, plus Escape
  useDismissablePopover(bulkMoveMenu, () => setBulkMoveMenu(false), {
    inside: () => [bulkMoveWrapRef.current],
  })
  useEffect(() => {
    if (!bulkMoveMenu) return
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setBulkMoveMenu(false)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [bulkMoveMenu])

  // WPS-style sortable "modified" column header, shared by both file tables
  const renderModifiedHeader = () => (
    <div className="cloud-col-sort" ref={fileSortRef}>
      <button
        className="cloud-col-sort-btn"
        aria-haspopup="menu"
        aria-expanded={fileSortMenuOpen}
        onClick={() => setFileSortMenuOpen((o) => !o)}
      >
        {t('colModified')}
        <svg
          width="12"
          height="12"
          viewBox="0 0 16 16"
          fill="none"
          aria-hidden="true"
          style={fileSort === 'oldest' ? { transform: 'rotate(180deg)' } : undefined}
        >
          <path
            d="M8 3v10M4.5 9.5L8 13l3.5-3.5"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
      {fileSortMenuOpen && (
        <div className="cloud-sort-menu" role="menu">
          {(['recent', 'oldest'] as const).map((key) => (
            <button
              key={key}
              className={fileSort === key ? 'active' : ''}
              role="menuitemradio"
              aria-checked={fileSort === key}
              onClick={() => {
                setFileSort(key)
                setFileSortMenuOpen(false)
              }}
            >
              <SortCheck visible={fileSort === key} />
              {t(key === 'recent' ? 'cloudSortRecent' : 'cloudSortOldest')}
            </button>
          ))}
        </div>
      )}
    </div>
  )

  // ── Plain view (no project selected): `entries` is the loaded page ──
  const allListPaths = entries.map((e) => e.path)
  const selectedPaths = allListPaths.filter((p) => selected.has(p))
  const allSelected = allListPaths.length > 0 && selectedPaths.length === allListPaths.length
  const pagerTotal = listTotal
  const totalPages = Math.max(1, Math.ceil(pagerTotal / PAGE_SIZE) || 1)

  const goListPage = (next: number) => {
    const page = Math.min(totalPages, Math.max(1, next))
    setListPage(page)
    setJumpDraft(String(page))
  }

  const projSelectedPaths = selectedPaths
  const projAllSelected = allSelected

  const changeView = (next: 'recent' | 'starred' | 'shared') => {
    setView(next)
    setListPage(1)
    setJumpDraft('1')
    setSelected(new Set())
    setRowMenu(null)
  }

  const changeFilter = (key: string) => {
    setFilter(key)
    setListPage(1)
    setJumpDraft('1')
    setSelected(new Set())
    setRowMenu(null)
  }

  const toggleSelect = (path: string, on: boolean) => {
    setSelected((prev) => {
      const next = new Set(prev)
      if (on) next.add(path)
      else next.delete(path)
      return next
    })
  }

  const toggleSelectAll = () => {
    setSelected(nextSelectAll(allListPaths, selected))
  }

  const toggleStar = (path: string) => {
    void window.aiOffice.toggleStar(path).then(refresh)
  }

  const deleteFiles = (paths: string[]) => {
    setRowMenu(null)
    setConfirmDelete(paths)
  }

  const confirmDeleteNow = () => {
    const paths = confirmDelete ?? []
    setConfirmDelete(null)
    setSelected(new Set())
    void window.aiOffice.deleteFiles(paths).then(refresh)
  }

  const duplicateFile = (path: string) => {
    setRowMenu(null)
    if (duplicatingShowTimer.current) clearTimeout(duplicatingShowTimer.current)
    // PDF / large artifacts can take several seconds (fetch + re-upload); show a
    // blocking busy dialog after a short delay so the UI doesn't look stuck.
    duplicatingShowTimer.current = setTimeout(() => setDuplicating(true), 350)
    void window.aiOffice
      .duplicateFile(path)
      .then(refresh)
      .finally(() => {
        if (duplicatingShowTimer.current) {
          clearTimeout(duplicatingShowTimer.current)
          duplicatingShowTimer.current = null
        }
        setDuplicating(false)
      })
  }

  const exportFile = (path: string) => {
    setRowMenu(null)
    void window.aiOffice.exportFile(path).then((result) => {
      if (result && !result.ok) window.alert(result.error ?? t('exportFailed'))
    })
  }

  const startRename = (entry: RecentEntry) => {
    setRowMenu(null)
    setRenaming({ path: entry.path, value: baseName(entry) })
  }

  const commitRename = (entry: RecentEntry) => {
    const value = renaming?.value.trim() ?? ''
    setRenaming(null)
    if (!value || value === baseName(entry)) return
    const newName = entry.ext ? `${value}.${entry.ext}` : value
    void window.aiOffice.renameFile(entry.path, newName).then((result) => {
      if (!result.ok) window.alert(result.error ?? t('renameFailed'))
      refresh()
    })
  }

  const moveFileTo = async (filePath: string, targetProjectId: string) => {
    setMoveFileMenu(null)
    setRowMenu(null)
    try {
      await window.aiOfficeProject?.moveFile(filePath, targetProjectId)
    } catch (error) {
      window.alert(error instanceof Error ? error.message : String(error))
      return
    }
    refresh()
  }

  const moveFilesTo = async (paths: string[], targetProjectId: string) => {
    setBulkMoveMenu(false)
    setSelected(new Set())
    try {
      for (const path of paths) {
        await window.aiOfficeProject?.moveFile(path, targetProjectId)
      }
    } catch (error) {
      window.alert(error instanceof Error ? error.message : String(error))
    } finally {
      // A bulk move can fail after earlier paths succeeded; reload to restore
      // unmoved rows while keeping successfully moved rows out of this project.
      refresh()
    }
  }

  // ── New file (passes projectId when a project is selected) ──
  const handleNewDoc = () => {
    void window.aiOffice.newDoc(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const handleNewSheet = () => {
    void window.aiOffice.newSheet(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const handleNewSlide = () => {
    void window.aiOffice.newSlide(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const handleNewMarkdown = () => {
    void window.aiOffice.newMarkdown(
      selectedProjectId ? { projectId: selectedProjectId } : undefined,
    )
  }

  const handleNewPdf = () => {
    void window.aiOffice.newPdf(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const handleNewBitable = () => {
    void window.aiOffice.newBitable?.(
      selectedProjectId ? { projectId: selectedProjectId } : undefined,
    )
  }

  const handleNewCode = () => {
    void window.aiOffice.newCode?.(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const NEW_ITEMS = [
    { ext: 'docx', title: t('newDoc'), sub: '.docx', action: handleNewDoc },
    { ext: 'xlsx', title: t('newSheet'), sub: '.xlsx', action: handleNewSheet },
    { ext: 'pptx', title: t('newSlide'), sub: '.pptx', action: handleNewSlide },
    { ext: 'md', title: t('newMarkdown'), sub: '.md', action: handleNewMarkdown },
    { ext: 'pdf', title: t('newPdf'), sub: '.pdf', action: handleNewPdf },
    { ext: 'table', title: t('newBitable'), sub: '.table', action: handleNewBitable },
    { ext: 'txt', title: t('newCode'), sub: '.txt', action: handleNewCode },
  ]

  function renderQuickCards() {
    return (
      <div className="quick-cards">
        {NEW_ITEMS.map((item) => (
          <button key={item.ext} className="quick-card" onClick={() => void item.action()}>
            <FileBadge ext={item.ext} size={30} />
            <span className="quick-text">
              <span className="quick-title-row">
                <span className="quick-title">{item.title}</span>
                <span className="ai-chip">AI</span>
              </span>
              <span className="quick-sub">{item.sub}</span>
            </span>
          </button>
        ))}
        <button
          className="quick-card"
          onClick={() => void window.aiOffice.browse()}
          data-tip={OPEN_LOCAL_EXTENSIONS}
        >
          <span className="quick-folder">
            <svg width="18" height="18" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <path
                d="M1.5 4A1.5 1.5 0 0 1 3 2.5h3.1c.44 0 .85.19 1.13.52L8.4 4.4H13A1.5 1.5 0 0 1 14.5 5.9v5.6A1.5 1.5 0 0 1 13 13H3a1.5 1.5 0 0 1-1.5-1.5V4z"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinejoin="round"
              />
            </svg>
          </span>
          <span className="quick-text">
            <span className="quick-title-row">
              <span className="quick-title">{t('openLocal')}</span>
            </span>
            <span className="quick-sub">{OPEN_LOCAL_EXTENSIONS}</span>
          </span>
        </button>
      </div>
    )
  }

  // ── File row rendering (shared by the plain view and the project files view) ──

  function renderFileRow(entry: RecentEntry, context: 'global' | 'project' | 'plain') {
    const isRenaming = renaming?.path === entry.path
    const otherProjects = projects.filter(
      (p) => p.id !== (context === 'project' ? selectedProjectId : undefined),
    )
    return (
      <li className="recent-row" key={entry.path}>
        <div
          className="recent-item"
          role="button"
          tabIndex={0}
          onClick={() => {
            if (!isRenaming) void window.aiOffice.openPath(entry.path)
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter' && event.target === event.currentTarget) {
              void window.aiOffice.openPath(entry.path)
            }
          }}
        >
          <span className="col-check" onClick={(event) => event.stopPropagation()}>
            <input
              type="checkbox"
              className="row-check"
              checked={selected.has(entry.path)}
              onChange={(event) => toggleSelect(entry.path, event.target.checked)}
              aria-label={t('selectFile', { name: entry.name })}
            />
          </span>
          <span className="recent-icon">
            <FileBadge ext={entry.ext} size={24} />
          </span>
          {isRenaming ? (
            <input
              className="rename-input"
              value={renaming.value}
              autoFocus
              onFocus={(event) => event.target.select()}
              onClick={(event) => event.stopPropagation()}
              onChange={(event) => setRenaming({ path: entry.path, value: event.target.value })}
              onBlur={() => commitRename(entry)}
              onKeyDown={(event) => {
                event.stopPropagation()
                if (event.nativeEvent.isComposing) return
                if (event.key === 'Enter') commitRename(entry)
                if (event.key === 'Escape') setRenaming(null)
              }}
            />
          ) : (
            <span className="recent-name">{entry.name}</span>
          )}
          <span className="recent-path">{entry.location ?? parentDir(entry.path)}</span>
          <span className="recent-time">{formatModified(entry.mtimeMs, i18n)}</span>
          <span className="recent-size">{formatSize(entry.sizeBytes)}</span>
          <button
            className={`star-btn${entry.starred ? ' starred' : ''}`}
            aria-label={entry.starred ? t('unstar') : t('star')}
            onClick={(event) => {
              event.stopPropagation()
              toggleStar(entry.path)
            }}
          >
            <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">
              <path
                d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                fill={entry.starred ? '#f5a623' : 'none'}
                stroke={entry.starred ? '#f5a623' : 'currentColor'}
                strokeWidth="1.2"
                strokeLinejoin="round"
              />
            </svg>
          </button>
          <span
            className="recent-actions"
            ref={rowMenu === entry.path ? rowMenuWrapRef : undefined}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              className="more-btn"
              aria-label={t('moreActions')}
              aria-expanded={rowMenu === entry.path}
              onClick={() => setRowMenu(rowMenu === entry.path ? null : entry.path)}
            >
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
                <circle cx="3.2" cy="8" r="1.4" fill="currentColor" />
                <circle cx="8" cy="8" r="1.4" fill="currentColor" />
                <circle cx="12.8" cy="8" r="1.4" fill="currentColor" />
              </svg>
            </button>
            {rowMenu === entry.path && (
              <div className="row-menu" role="menu">
                <button
                  role="menuitem"
                  onClick={() => {
                    setRowMenu(null)
                    void window.aiOffice.openPath(entry.path)
                  }}
                >
                  {t('open')}
                </button>
                {!isMoreaiHomeEmbed() && (
                  <>
                    <button
                      role="menuitem"
                      onClick={() => {
                        setRowMenu(null)
                        void window.aiOffice.revealPath(entry.path)
                      }}
                    >
                      {t('revealInFolder')}
                    </button>
                    <button
                      role="menuitem"
                      onClick={() => {
                        setRowMenu(null)
                        void navigator.clipboard.writeText(entry.path)
                      }}
                    >
                      {t('copyPath')}
                    </button>
                  </>
                )}
                {projectMode && otherProjects.length > 0 && (
                  <>
                    <div className="row-menu-divider" />
                    <div
                      className="move-menu-wrap"
                      ref={moveFileMenu === entry.path ? moveMenuWrapRef : undefined}
                      onMouseEnter={() => {
                        clearMoveMenuTimer('close')
                        if (moveFileMenu === entry.path) return
                        clearMoveMenuTimer('open')
                        moveMenuTimers.current.open = window.setTimeout(
                          () => openMoveMenu(entry.path),
                          160,
                        )
                      }}
                      onMouseLeave={() => {
                        clearMoveMenuTimer('open')
                        clearMoveMenuTimer('close')
                        moveMenuTimers.current.close = window.setTimeout(
                          () => setMoveFileMenu(null),
                          140,
                        )
                      }}
                    >
                      <button
                        role="menuitem"
                        className="submenu-trigger"
                        onClick={(e) => {
                          e.stopPropagation()
                          clearMoveMenuTimer('open')
                          clearMoveMenuTimer('close')
                          if (moveFileMenu === entry.path) setMoveFileMenu(null)
                          else openMoveMenu(entry.path)
                        }}
                      >
                        {t('moveToProject')}
                        <svg
                          width="11"
                          height="11"
                          viewBox="0 0 12 12"
                          aria-hidden="true"
                          style={{ marginLeft: 'auto' }}
                        >
                          <path
                            d="M4.5 2.5l4 3.5-4 3.5"
                            stroke="currentColor"
                            strokeWidth="1.3"
                            strokeLinecap="round"
                            fill="none"
                          />
                        </svg>
                      </button>
                      {moveFileMenu === entry.path && (
                        <div
                          className={`submenu${moveMenuFlip ? ' submenu-left' : ''}`}
                          role="menu"
                          ref={measureSubmenu}
                        >
                          {otherProjects.map((p) => (
                            <button
                              key={p.id}
                              role="menuitem"
                              onClick={() => void moveFileTo(entry.path, p.id)}
                            >
                              {p.isDefault ? t('defaultProject') : p.name}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  </>
                )}
                <div className="row-menu-divider" />
                <button role="menuitem" onClick={() => startRename(entry)}>
                  {t('rename')}
                </button>
                <button role="menuitem" onClick={() => duplicateFile(entry.path)}>
                  {t('duplicate')}
                </button>
                <button role="menuitem" onClick={() => exportFile(entry.path)}>
                  {t('exportToLocal')}
                </button>
                {context === 'global' && selectedPaths.length === 0 && (
                  <>
                    <div className="row-menu-divider" />
                    <button
                      role="menuitem"
                      className="danger"
                      onClick={() => deleteFiles([entry.path])}
                    >
                      {t('deleteFiles')}
                    </button>
                  </>
                )}
              </div>
            )}
          </span>
        </div>
      </li>
    )
  }

  function renderFilterPills(): ReactElement {
    return (
      <div className="filter-pills" role="tablist" aria-label={t('filterAria')}>
        {FILTERS.map((f) => (
          <button
            key={f.key}
            className={`filter-pill${filter === f.key ? ' active' : ''}`}
            onClick={() => changeFilter(f.key)}
          >
            {t(f.label)}
          </button>
        ))}
      </div>
    )
  }

  function renderCountHeading(label: string, count: number): ReactElement {
    return (
      <div className="recents-heading">
        <span className="section-label" title={label}>
          {label}
        </span>
        <span className="file-count">{t(fileCountKey(count), { n: count })}</span>
      </div>
    )
  }

  // ── Project files view ────────────────────────────────

  function renderProjectContent() {
    const proj = projects.find((p) => p.id === selectedProjectId)
    if (!proj) return null
    const otherProjects = projects.filter((p) => p.id !== proj.id)
    const projectLabel = proj.isDefault ? t('defaultProject') : proj.name

    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="section-head">
            <span className="section-label">{t('secQuickStart')}</span>
          </div>
          {renderQuickCards()}
        </section>

        <section className="recents" aria-label={projectLabel}>
          <div className="recents-toolbar">
            {projSelectedPaths.length > 0 ? (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: projSelectedPaths.length })}
                </span>
                {otherProjects.length > 0 && (
                  <span className="selection-move-wrap" ref={bulkMoveWrapRef}>
                    <button
                      className="selection-action"
                      aria-expanded={bulkMoveMenu}
                      onClick={() => setBulkMoveMenu((open) => !open)}
                    >
                      {t('moveToProject')}
                    </button>
                    {bulkMoveMenu && (
                      <div className="selection-move-menu" role="menu">
                        {otherProjects.map((p) => (
                          <button
                            key={p.id}
                            role="menuitem"
                            onClick={() => void moveFilesTo(projSelectedPaths, p.id)}
                          >
                            {p.isDefault ? t('defaultProject') : p.name}
                          </button>
                        ))}
                      </div>
                    )}
                  </span>
                )}
                <button
                  className="selection-action danger"
                  onClick={() => deleteFiles(projSelectedPaths)}
                >
                  {t('deleteFiles')}
                </button>
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            ) : (
              renderFilterPills()
            )}
            {renderCountHeading(projectLabel, listTotal)}
          </div>

          {entries.length === 0 ? (
            <p className="empty proj-empty">
              <svg
                className="proj-empty-icon"
                width="48"
                height="48"
                viewBox="0 0 24 24"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.29297 3.75H14.1729C14.4927 3.75 14.7979 3.88392 15.0146 4.11914L18.5566 7.96387C18.7512 8.17512 18.8593 8.45208 18.8594 8.73926V19.1055C18.8593 19.7376 18.346 20.25 17.7139 20.25H6.29297C5.66091 20.2499 5.14855 19.7375 5.14844 19.1055V4.89453C5.14855 4.26247 5.66091 3.75011 6.29297 3.75Z"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
                <path
                  d="M13.8984 4V7.11C13.8984 8.15382 14.7446 9 15.7884 9H18.8984"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
              </svg>
              <span className="empty-hint">
                {filter !== 'all' ? t('emptyFiltered') : t('projEmptyHint')}
              </span>
            </p>
          ) : (
            <div className={`recent-table${projSelectedPaths.length > 0 ? ' has-selection' : ''}`}>
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={projAllSelected}
                    onChange={toggleSelectAll}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span>{t('colLocation')}</span>
                {renderModifiedHeader()}
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {entries.map((entry) => renderFileRow(entry, 'project'))}
              </ul>
              {renderFilePager()}
            </div>
          )}
        </section>
      </main>
    )
  }

  /** AI 产物 browse — same chrome as 项目: 快速开始 + 文件列表 (no greeting jump). */
  function renderAiBrowseContent() {
    const sourceLabelKey = AI_SOURCE_ITEMS.find((i) => i.id === selectedAiSource)?.labelKey
    const aiLabel = aiSessionTitle || (sourceLabelKey ? t(sourceLabelKey) : t('secAiFiles'))
    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="section-head">
            <span className="section-label">{t('secQuickStart')}</span>
          </div>
          {renderQuickCards()}
        </section>

        <section className="recents" aria-label={aiLabel}>
          <div className="recents-toolbar">
            {selectedPaths.length > 0 ? (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: selectedPaths.length })}
                </span>
                <button
                  className="selection-action danger"
                  onClick={() => deleteFiles(selectedPaths)}
                >
                  {t('deleteFiles')}
                </button>
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            ) : (
              renderFilterPills()
            )}
            {renderCountHeading(aiLabel, listTotal)}
          </div>

          {entries.length === 0 ? (
            <p className="empty proj-empty">
              <svg
                className="proj-empty-icon"
                width="48"
                height="48"
                viewBox="0 0 24 24"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.29297 3.75H14.1729C14.4927 3.75 14.7979 3.88392 15.0146 4.11914L18.5566 7.96387C18.7512 8.17512 18.8593 8.45208 18.8594 8.73926V19.1055C18.8593 19.7376 18.346 20.25 17.7139 20.25H6.29297C5.66091 20.2499 5.14855 19.7375 5.14844 19.1055V4.89453C5.14855 4.26247 5.66091 3.75011 6.29297 3.75Z"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
                <path
                  d="M13.8984 4V7.11C13.8984 8.15382 14.7446 9 15.7884 9H18.8984"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
              </svg>
              <span className="empty-hint">
                {filter !== 'all' ? t('emptyFiltered') : t('projEmptyHint')}
              </span>
            </p>
          ) : (
            <div className={`recent-table${selectedPaths.length > 0 ? ' has-selection' : ''}`}>
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={allSelected}
                    onChange={toggleSelectAll}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span>{t('colLocation')}</span>
                {renderModifiedHeader()}
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {entries.map((entry) => renderFileRow(entry, 'plain'))}
              </ul>
              {renderFilePager()}
            </div>
          )}
        </section>
      </main>
    )
  }

  function renderFilePager(): ReactElement | null {
    if (pagerTotal <= 0) return null
    const submitJump = () => {
      const n = Number.parseInt(jumpDraft, 10)
      if (Number.isFinite(n)) goListPage(n)
    }
    return (
      <nav className="file-pager" aria-label={t('pagerAria')}>
        <button
          type="button"
          className="file-pager-btn"
          disabled={listPage <= 1}
          onClick={() => goListPage(listPage - 1)}
        >
          {t('pagerPrev')}
        </button>
        <span className="file-pager-status">{t('pagerPage', { page: listPage, pages: totalPages })}</span>
        <button
          type="button"
          className="file-pager-btn"
          disabled={listPage >= totalPages}
          onClick={() => goListPage(listPage + 1)}
        >
          {t('pagerNext')}
        </button>
        <label className="file-pager-jump">
          {t('pagerJump')}
          <input
            type="number"
            min={1}
            max={totalPages}
            value={jumpDraft}
            onChange={(e) => setJumpDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') submitJump()
            }}
          />
          <button type="button" className="file-pager-btn" onClick={submitJump}>
            {t('pagerGo')}
          </button>
        </label>
      </nav>
    )
  }

  // ── Plain view ────────────────────────────────────────

  function renderGlobalContent() {
    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="section-head">
            <span className="section-label">{t('secQuickStart')}</span>
          </div>
          {renderQuickCards()}
        </section>

        <section
          className="recents"
          aria-label={view === 'recent' ? t('secRecent') : t('secStarred')}
        >
          <div className="recents-toolbar">
            {selectedPaths.length > 0 ? (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: selectedPaths.length })}
                </span>
                <button
                  className="selection-action danger"
                  onClick={() => deleteFiles(selectedPaths)}
                >
                  {t('deleteFiles')}
                </button>
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            ) : (
              <div className="filter-pills" role="tablist" aria-label={t('filterAria')}>
                {FILTERS.map((f) => (
                  <button
                    key={f.key}
                    className={`filter-pill${filter === f.key ? ' active' : ''}`}
                    onClick={() => changeFilter(f.key)}
                  >
                    {t(f.label)}
                  </button>
                ))}
              </div>
            )}
            <div className="recents-heading">
              <span className="section-label">
                {view === 'recent' ? t('secRecent') : t('secStarred')}
              </span>
              <span className="file-count">{t(fileCountKey(listTotal), { n: listTotal })}</span>
            </div>
          </div>

          {entries.length === 0 ? (
            <p className="empty proj-empty">
              <svg
                className="proj-empty-icon"
                width="48"
                height="48"
                viewBox="0 0 24 24"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.29297 3.75H14.1729C14.4927 3.75 14.7979 3.88392 15.0146 4.11914L18.5566 7.96387C18.7512 8.17512 18.8593 8.45208 18.8594 8.73926V19.1055C18.8593 19.7376 18.346 20.25 17.7139 20.25H6.29297C5.66091 20.2499 5.14855 19.7375 5.14844 19.1055V4.89453C5.14855 4.26247 5.66091 3.75011 6.29297 3.75Z"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
                <path
                  d="M13.8984 4V7.11C13.8984 8.15382 14.7446 9 15.7884 9H18.8984"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
              </svg>
              <span className="empty-hint">
                {view === 'starred'
                  ? t('emptyStarred')
                  : view === 'shared'
                    ? t('emptyShared')
                    : navCounts.recent === 0
                      ? t('emptyRecent')
                      : t('emptyFiltered')}
              </span>
            </p>
          ) : (
            <div className={`recent-table${selectedPaths.length > 0 ? ' has-selection' : ''}`}>
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={allSelected}
                    onChange={toggleSelectAll}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span>{t('colLocation')}</span>
                {renderModifiedHeader()}
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {entries.map((entry) => renderFileRow(entry, 'global'))}
              </ul>
              {renderFilePager()}
            </div>
          )}
        </section>
      </main>
    )
  }

  return (
    <div className="home">
      <aside className="sidebar">
        <div className="sidebar-logo">
          <img className="logo-lockup" src={logoLockup} alt="GenOffice" />
        </div>

        {isMoreaiHomeEmbed() && (
          <>
            <nav className="sidebar-nav">
              <button
                className={`nav-item${view === 'recent' && !selectedProjectId && !cloudMode && !selectedAiSource ? ' active' : ''}`}
                onClick={() => {
                  changeView('recent')
                  setSelectedProjectId(null)
                  setSelectedAiSource(null)
                  setCloudMode(false)
                  void window.aiOffice.openAiSource?.(null)
                  setSelectedAiSessionKey(null)
                }}
              >
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <circle cx="8" cy="8" r="6.2" stroke="currentColor" strokeWidth="1.3" />
                  <path
                    d="M8 4.8V8l2.2 1.6"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinecap="round"
                  />
                </svg>
                <span className="nav-label">{t('navRecent')}</span>
                <span className="nav-count">{navCounts.recent}</span>
              </button>
              <button
                className={`nav-item${view === 'starred' && !selectedProjectId && !cloudMode && !selectedAiSource ? ' active' : ''}`}
                onClick={() => {
                  changeView('starred')
                  setSelectedProjectId(null)
                  setSelectedAiSource(null)
                  setCloudMode(false)
                  void window.aiOffice.openAiSource?.(null)
                  setSelectedAiSessionKey(null)
                }}
              >
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <path
                    d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinejoin="round"
                  />
                </svg>
                <span className="nav-label">{t('navStarred')}</span>
                <span className="nav-count">{navCounts.starred}</span>
              </button>
              <button
                className={`nav-item${view === 'shared' && !selectedProjectId && !cloudMode && !selectedAiSource ? ' active' : ''}`}
                onClick={() => {
                  changeView('shared')
                  setSelectedProjectId(null)
                  setSelectedAiSource(null)
                  setCloudMode(false)
                  void window.aiOffice.openAiSource?.(null)
                  setSelectedAiSessionKey(null)
                }}
              >
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <path
                    d="M6.15 10.55a3.15 3.15 0 0 1 0-4.45l1.85-1.85a3.15 3.15 0 0 1 4.45 4.45L11.1 10"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinecap="round"
                  />
                  <path
                    d="M9.85 5.45a3.15 3.15 0 0 1 0 4.45L8 11.75a3.15 3.15 0 1 1-4.45-4.45L4.9 6"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinecap="round"
                  />
                </svg>
                <span className="nav-label">{t('navShared')}</span>
                <span className="nav-count">{navCounts.shared}</span>
              </button>
            </nav>
            <div className="sidebar-divider" />
            <AiOutputsPanel
              selectedId={selectedAiSource}
              selectedSessionKey={selectedAiSessionKey}
              onSelect={(id) => {
                void (async () => {
                  setSelectedProjectId(null)
                  setCloudMode(false)
                  setSelected(new Set())
                  setRowMenu(null)
                  if (!id) {
                    await window.aiOffice.openAiSource?.(null)
                    setSelectedAiSource(null)
                    setSelectedAiSessionKey(null)
                    return
                  }
                  changeView('recent')
                  const sessionKey = (await window.aiOffice.openAiSource?.(id)) ?? null
                  // Batch source + session so file list reloads once (avoids expand flicker).
                  setSelectedAiSource(id)
                  setSelectedAiSessionKey(typeof sessionKey === 'string' ? sessionKey : null)
                })()
              }}
              onSelectSession={(sourceId, sessionKey) => {
                setSelectedAiSource(sourceId)
                setSelectedAiSessionKey(sessionKey)
                setSelectedProjectId(null)
                setCloudMode(false)
                void window.aiOffice.openAiSession?.(sourceId, sessionKey)
              }}
            />
          </>
        )}

        {!isMoreaiHomeEmbed() && (
          <nav className="sidebar-nav">
            <button
              className={`nav-item${view === 'recent' && !selectedProjectId && !cloudMode && !selectedAiSource ? ' active' : ''}`}
              onClick={() => {
                changeView('recent')
                setSelectedProjectId(null)
                setSelectedAiSource(null)
                setCloudMode(false)
                void window.aiOffice.openAiSource?.(null)
                setSelectedAiSessionKey(null)
              }}
            >
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <circle cx="8" cy="8" r="6.2" stroke="currentColor" strokeWidth="1.3" />
                <path
                  d="M8 4.8V8l2.2 1.6"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinecap="round"
                />
              </svg>
              <span className="nav-label">{t('navRecent')}</span>
              <span className="nav-count">{navCounts.recent}</span>
            </button>
            <button
              className={`nav-item${view === 'starred' && !selectedProjectId && !cloudMode && !selectedAiSource ? ' active' : ''}`}
              onClick={() => {
                changeView('starred')
                setSelectedProjectId(null)
                setSelectedAiSource(null)
                setCloudMode(false)
                void window.aiOffice.openAiSource?.(null)
                setSelectedAiSessionKey(null)
              }}
            >
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <path
                  d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinejoin="round"
                />
              </svg>
              <span className="nav-label">{t('navStarred')}</span>
              <span className="nav-count">{navCounts.starred}</span>
            </button>
            {loggedIn && (
              <button
                className={`nav-item${cloudMode && !selectedProjectId ? ' active' : ''}`}
                onClick={() => {
                  setCloudMode(true)
                  setSelectedProjectId(null)
                  setSelected(new Set())
                  setRowMenu(null)
                }}
              >
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <path
                    d="M8 1.8l1.55 4.65L14.2 8l-4.65 1.55L8 14.2 6.45 9.55 1.8 8l4.65-1.55z"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinejoin="round"
                  />
                </svg>
                <span className="nav-label">{t('navCloud')}</span>
                <svg
                  className="nav-external"
                  width="13"
                  height="13"
                  viewBox="0 0 16 16"
                  fill="none"
                  aria-hidden="true"
                >
                  <path
                    d="M6.5 3.5H4a1.5 1.5 0 0 0-1.5 1.5v7A1.5 1.5 0 0 0 4 13.5h7A1.5 1.5 0 0 0 12.5 12V9.5M9.5 2.5h4v4M13 3l-5.5 5.5"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
            )}
          </nav>
        )}

        {/* project sidebar */}
        {projectMode && (
          <>
            <div className="sidebar-divider" />
            <ProjectPanel
              projects={projects}
              selectedId={selectedProjectId}
              onSelect={(id) => {
                setSelectedProjectId(id)
                setSelectedAiSource(null)
                setSelectedAiSessionKey(null)
                void window.aiOffice.openAiSource?.(null)
                // reset list-selection state on any project switch (paths are
                // shared between the plain view and project views)
                setSelected(new Set())
                setRowMenu(null)
              }}
              onRefresh={refresh}
            />
          </>
        )}

        {isMoreaiHomeEmbed() ? (
          <MoreaiAccountModelEntry onStatusChange={handleAccountStatus} />
        ) : (
          <AccountEntry onStatusChange={handleAccountStatus} />
        )}
      </aside>

      {selectedProjectId ? (
        renderProjectContent()
      ) : selectedAiSource ? (
        renderAiBrowseContent()
      ) : cloudMode ? (
        <CloudProjectsView />
      ) : (
        renderGlobalContent()
      )}

      {confirmDelete && (
        <div className="modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('deleteModalTitle')}
            onClick={(event) => event.stopPropagation()}
          >
            <h3>{t('deleteModalTitle')}</h3>
            <p>
              {confirmDelete.length === 1
                ? t('deleteConfirmOne', { name: fileName(confirmDelete[0]) })
                : t('deleteConfirmMany', { n: confirmDelete.length })}
            </p>
            {confirmDelete.length > 1 && (
              <ul className="modal-file-list">
                {confirmDelete.slice(0, 6).map((p) => (
                  <li key={p}>{fileName(p)}</li>
                ))}
                {confirmDelete.length > 6 && (
                  <li>{t('deleteMoreCount', { n: confirmDelete.length })}</li>
                )}
              </ul>
            )}
            <div className="modal-buttons">
              <button
                className="btn btn-secondary"
                autoFocus
                onClick={() => setConfirmDelete(null)}
              >
                {t('cancel')}
              </button>
              <button className="btn btn-danger" onClick={confirmDeleteNow}>
                {t('delete')}
              </button>
            </div>
          </div>
        </div>
      )}

      {duplicating && (
        <div className="modal-overlay" role="presentation">
          <div
            className="modal modal-busy"
            role="alertdialog"
            aria-modal="true"
            aria-busy="true"
            aria-live="polite"
            aria-label={t('duplicatingTitle')}
          >
            <div className="modal-busy-row">
              <span className="load-more-spinner modal-busy-spinner" aria-hidden="true" />
              <div>
                <h3>{t('duplicatingTitle')}</h3>
                <p>{t('duplicatingHint')}</p>
              </div>
            </div>
          </div>
        </div>
      )}

      <DropToOpenOverlay />
    </div>
  )
}
