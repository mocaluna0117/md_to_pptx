import { useEffect, useState } from 'react'
import Home from './Home'
import Deckdown from './App'
import Docdown from './Docdown'

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

export default function Root() {
  const [route, setRoute] = useState<Route>(routeFromHash())
  useEffect(() => {
    const onHash = () => setRoute(routeFromHash())
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  if (route === 'slides') return <Deckdown />
  if (route === 'docx') return <Docdown />
  return <Home />
}
