import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App, { LoginGate } from './App.jsx'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <LoginGate>
      <App />
    </LoginGate>
  </StrictMode>,
)
