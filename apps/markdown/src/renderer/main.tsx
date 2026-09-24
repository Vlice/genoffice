import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { waitForOfficeHostInjection } from '@genoffice/office-host'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import type { MarkdownApi, UiTheme } from '../shared/ipc'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/dropdown.css'
import 'katex/dist/katex.min.css'
import './styles.css'
import { installScreenTips } from '@genoffice/ui'

installScreenTips()

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

void (async () => {
  if (new URLSearchParams(window.location.search).get('readOnly') === '1') {
    document.body.setAttribute('data-moreai-embed-readonly', '1')
    document.documentElement.setAttribute('data-moreai-embed-readonly', '1')
  }
  try {
    await waitForOfficeHostInjection()
  } catch {
    /* Electron preload */
  }
  const injected = window.__OFFICE_HOST__ as MarkdownApi | undefined
  if (injected) window.markdownApi = injected

  const [lang, theme] = await Promise.all([
    window.markdownApi.getLanguage().catch(() => 'zh' as const),
    window.markdownApi.getTheme().catch(() => 'system' as const),
  ])
  document.documentElement.lang = htmlLang(lang as Lang)
  applyTheme(theme)
  window.markdownApi.onThemeChanged(applyTheme)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
})()
