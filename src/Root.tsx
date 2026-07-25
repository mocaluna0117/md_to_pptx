import { lazy, Suspense, useEffect, useState } from 'react'
import Home from './Home'

// Each app is its own chunk, so the launcher loads without Marp/MathJax/export libs.
const Deckdown = lazy(() => import('./App'))
const Docdown = lazy(() => import('./Docdown'))

export type Route = 'home' | 'slides' | 'docx'

function routeFromHash(): Route {
  const h = window.location.hash.replace(/^#\/?/, '')
  if (h === 'slides') return 'slides'
  if (h === 'docx') return 'docx'
  return 'home'
}

/** Navigate between the three pages (hash routing survives reloads on static hosting). */
export function navigate(route: Route): void {
  window.location.hash = route === 'home' ? '#/' : `#/${route}`
}

function Loading() {
  return <div className="route-loading">読み込み中…</div>
}

export default function Root() {
  const [route, setRoute] = useState<Route>(routeFromHash())
  useEffect(() => {
    const onHash = () => setRoute(routeFromHash())
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  if (route === 'slides') {
    return (
      <Suspense fallback={<Loading />}>
        <Deckdown />
      </Suspense>
    )
  }
  if (route === 'docx') {
    return (
      <Suspense fallback={<Loading />}>
        <Docdown />
      </Suspense>
    )
  }
  return <Home />
}
