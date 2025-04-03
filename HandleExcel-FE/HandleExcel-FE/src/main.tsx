import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { StyleProvider } from '@ant-design/cssinjs';
import './index.css'
import App from './App.tsx'

const root = createRoot(document.getElementById('root')!);

root.render(
  <StrictMode>
    <StyleProvider hashPriority="high">
      <App />
    </StyleProvider>
  </StrictMode>
);
