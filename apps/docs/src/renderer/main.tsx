import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { waitForOfficeHostInjection } from '@genoffice/office-host'
import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import type { DesktopApi, UiTheme } from '../shared/ipc'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@genoffice/ui/dropdown.css'
import './styles.css'
import './fonts/fonts.css'
import { installScreenTips } from '@genoffice/ui'

installScreenTips()

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

async function bootstrap(): Promise<void> {
  // MoreAI SPA / iframe: inject window.__OFFICE_HOST__ then call __resolveOfficeHost().
  try {
    await waitForOfficeHostInjection()
  } catch {
    /* Electron preload already exposed window.desktop */
  }
  const injected = window.__OFFICE_HOST__ as DesktopApi | undefined
  if (injected) window.desktop = injected

  let lang: Lang = 'zh'
  let theme: UiTheme = 'system'
  try {
    ;[lang, theme] = await Promise.all([
      window.desktop.getLanguage().catch(() => 'zh' as const),
      window.desktop.getTheme().catch(() => 'system' as const),
    ])
  } catch {
    /* dev renderer without the preload bridge */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  applyTheme(theme)
  window.desktop?.onThemeChanged(applyTheme)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()
