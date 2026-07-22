import { navigate } from './Root'
import './Home.css'

function SlideMark() {
  return (
    <svg viewBox="0 0 32 32" width="34" height="34" aria-hidden focusable="false">
      <rect x="4" y="7" width="24" height="18" rx="2.5" fill="currentColor" opacity="0.15" />
      <rect x="4" y="7" width="24" height="18" rx="2.5" fill="none" stroke="currentColor" strokeWidth="1.6" />
      <rect x="8" y="11" width="12" height="2.2" rx="1.1" fill="currentColor" />
      <rect x="8" y="16" width="16" height="1.8" rx="0.9" fill="currentColor" opacity="0.6" />
      <rect x="8" y="19.5" width="10" height="1.8" rx="0.9" fill="currentColor" opacity="0.6" />
    </svg>
  )
}

function DocMark() {
  return (
    <svg viewBox="0 0 32 32" width="34" height="34" aria-hidden focusable="false">
      <path
        d="M8 4h11l6 6v18a0 0 0 0 1 0 0H8a0 0 0 0 1 0 0V4z"
        fill="currentColor"
        opacity="0.15"
      />
      <path d="M8 4h11l6 6v18H8V4z" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" />
      <path d="M19 4v6h6" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" />
      <rect x="11" y="15" width="10" height="1.8" rx="0.9" fill="currentColor" opacity="0.7" />
      <rect x="11" y="19" width="10" height="1.8" rx="0.9" fill="currentColor" opacity="0.7" />
      <rect x="11" y="23" width="6" height="1.8" rx="0.9" fill="currentColor" opacity="0.7" />
    </svg>
  )
}

export default function Home() {
  return (
    <div className="home">
      <div className="home-inner">
        <h1 className="home-title">Markdown → Office</h1>
        <p className="home-sub">
          Markdown を、そのまま編集できる PowerPoint / Word に。ブラウザだけで完結し、内容は外部に送信しません。
        </p>
        <div className="home-cards">
          <button className="home-card" onClick={() => navigate('slides')}>
            <span className="home-card-icon slides">
              <SlideMark />
            </span>
            <span className="home-card-name">Deckdown</span>
            <span className="home-card-fmt">Markdown → PowerPoint（.pptx）</span>
            <span className="home-card-desc">スライドを PowerPoint のように直接編集して書き出し。</span>
          </button>
          <button className="home-card" onClick={() => navigate('docx')}>
            <span className="home-card-icon docs">
              <DocMark />
            </span>
            <span className="home-card-name">Docdown</span>
            <span className="home-card-fmt">Markdown → Word（.docx）</span>
            <span className="home-card-desc">文書としてプレビューして、編集できる Word に書き出し。</span>
          </button>
        </div>
      </div>
    </div>
  )
}
