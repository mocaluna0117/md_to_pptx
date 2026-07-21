import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
// `base` is '/' for local dev/build; the GitHub Pages workflow sets BASE_PATH
// to '/md_to_pptx/' so built asset URLs resolve under the project-site subpath.
export default defineConfig({
  base: process.env.BASE_PATH ?? '/',
  plugins: [react()],
})
