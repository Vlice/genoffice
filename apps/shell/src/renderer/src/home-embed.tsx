/**
 * MoreAI / web embed entry: mount GenOffice Home only (no shell tab bar).
 * Expects the host to inject `window.aiOffice` (+ optional `aiOfficeProject`)
 * before or while this page waits.
 */
import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang } from '@genoffice/i18n'
import { Home } from './Home'
import { LocaleProvider } from './locale'
import { installScreenTips } from '@genoffice/ui'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/dropdown.css'
import './home.css'
import './home-embed.css'

declare global {
  interface Window {
    aiOffice?: import('../../shared/home-api').HomeApi
    aiOfficeProject?: import('../../shared/home-api').ProjectHomeApi
    __MOREAI_HOME_READY__?: boolean
  }
}

installScreenTips()
document.body.classList.add('moreai-home-embed')

async function waitForHomeApi(timeoutMs = 15000): Promise<NonNullable<typeof window.aiOffice>> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    if (window.aiOffice) return window.aiOffice
    await new Promise((r) => setTimeout(r, 40))
  }
  throw new Error('window.aiOffice was not injected by the MoreAI host')
}

void waitForHomeApi()
  .then(async (api) => {
    const lang = await api.getLanguage().catch(() => 'zh' as const)
    const theme = await api.getTheme().catch(() => 'system' as const)
    document.documentElement.lang = htmlLang(lang)
    if (theme !== 'system') {
      document.documentElement.setAttribute('data-theme', theme)
    }
    api.onThemeChanged((next) => {
      if (next === 'system') document.documentElement.removeAttribute('data-theme')
      else document.documentElement.setAttribute('data-theme', next)
    })
    window.__MOREAI_HOME_READY__ = true
    createRoot(document.getElementById('root')!).render(
      <React.StrictMode>
        <LocaleProvider initial={lang}>
          <Home />
        </LocaleProvider>
      </React.StrictMode>,
    )
  })
  .catch((err) => {
    const root = document.getElementById('root')
    if (root) {
      root.innerHTML = `<div style="padding:24px;font:14px system-ui;color:#b42318">${String(
        err instanceof Error ? err.message : err,
      )}</div>`
    }
  })
