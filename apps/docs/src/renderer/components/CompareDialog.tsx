/**
 * Review → Compare: pick a second document without replacing the open file.
 * Style matches ProtectDialog (modal-backdrop + gs-form). Cloud/recent list
 * comes from listCompareDocs; local files use pickCompareDocx (not openDocx —
 * MoreAI's openDocx is the current cloud document).
 */
import { useEffect, useState } from 'react'
import type { CompareDocCandidate, OpenFileResult, OpenDocxResult } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { FieldError } from './PasswordInput'

export function CompareDialog({
  excludeId,
  onCancel,
  onPicked,
}: {
  /** Current document path / artifact id — hidden from the list */
  excludeId?: string | null
  onCancel: () => void
  onPicked: (file: OpenFileResult) => void
}) {
  const { t } = useI18n()
  const [source, setSource] = useState<'cloud' | 'recent'>('cloud')
  const [items, setItems] = useState<CompareDocCandidate[]>([])
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [localName, setLocalName] = useState<string | null>(null)
  const [localFile, setLocalFile] = useState<OpenFileResult | null>(null)
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  useEffect(() => {
    let alive = true
    void (async () => {
      try {
        const list = await window.desktop.listCompareDocs()
        if (!alive) return
        const next = (list?.items ?? []).filter((it) => {
          if (!it.id) return false
          if (!excludeId) return true
          if (it.id === excludeId) return false
          if (excludeId.includes(`/office/${it.id}/`)) return false
          return true
        })
        setSource(list?.source === 'recent' ? 'recent' : 'cloud')
        setItems(next)
      } catch {
        if (alive) setLoadError(true)
      } finally {
        if (alive) setLoading(false)
      }
    })()
    return () => {
      alive = false
    }
  }, [excludeId])

  const ready = !!(localFile || selectedId)
  const libraryTitle = source === 'recent' ? t('appCompareRecent') : t('appCompareCloud')

  const asOpenFile = (other: OpenDocxResult): OpenFileResult | null => {
    if (!other || 'needsPassword' in other) return null
    return other
  }

  const pickLocal = async () => {
    if (busy) return
    setBusy(true)
    setError('')
    try {
      const other = await window.desktop.pickCompareDocx()
      if (!other) return
      if ('needsPassword' in other) {
        setError(t('appCompareFailed', { error: t('appDocPwdTitle') }))
        return
      }
      setLocalFile(other)
      setLocalName(other.name)
      setSelectedId(null)
    } catch (err) {
      setError(t('appCompareFailed', { error: String(err) }))
    } finally {
      setBusy(false)
    }
  }

  const submit = async () => {
    if (busy || !ready) return
    if (localFile) {
      onPicked(localFile)
      return
    }
    if (!selectedId) return
    setBusy(true)
    setError('')
    try {
      const other = asOpenFile(await window.desktop.readCompareDocx(selectedId))
      if (!other) {
        setError(t('appCompareFailed', { error: t('appDocPwdTitle') }))
        return
      }
      onPicked(other)
    } catch (err) {
      setError(t('appCompareFailed', { error: String(err) }))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onCancel()}>
      <div
        className="modal protect-dialog compare-dialog gs-form"
        role="dialog"
        aria-labelledby="compare-dialog-title"
      >
        <h2 id="compare-dialog-title">{t('appCompareTitle')}</h2>
        <p className="modal-desc">{t('appCompareDesc')}</p>

        <h3 className="protect-section-title">{libraryTitle}</h3>
        {loading && <p className="fld-hint">{t('appCompareLoading')}</p>}
        {loadError && <p className="fld-hint">{t('appCompareLoadFailed')}</p>}
        {!loading && !loadError && items.length === 0 && (
          <p className="fld-hint">{t('appCompareNoDocs')}</p>
        )}
        {!loading && items.length > 0 && (
          <div className="compare-doc-list" role="listbox" aria-label={libraryTitle}>
            {items.map((item) => (
              <label key={item.id} className="protect-check">
                <input
                  type="radio"
                  name="compare-doc"
                  checked={selectedId === item.id && !localFile}
                  onChange={() => {
                    setSelectedId(item.id)
                    setLocalFile(null)
                    setLocalName(null)
                    setError('')
                  }}
                />
                <span className="ctl" aria-hidden="true" />
                <span className="compare-doc-name" title={item.name}>
                  {item.name}
                </span>
              </label>
            ))}
          </div>
        )}

        <h3 className="protect-section-title">{t('appCompareLocal')}</h3>
        <div className="compare-local-row">
          <button type="button" className="compare-pick-local" disabled={busy} onClick={() => void pickLocal()}>
            {t('appComparePickLocal')}
          </button>
          {localName && (
            <span className="compare-selected">{t('appCompareSelected', { name: localName })}</span>
          )}
        </div>

        {error && <FieldError>{error}</FieldError>}
        <div className="modal-actions">
          <button type="button" onClick={onCancel}>
            {t('appCancel')}
          </button>
          <button
            type="button"
            className="btn-primary"
            disabled={busy || !ready}
            onClick={() => void submit()}
          >
            {t('appCompareRun')}
          </button>
        </div>
      </div>
    </div>
  )
}
