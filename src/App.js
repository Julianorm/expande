import { supabase } from './lib/supabase'

export default function App() {
  return (
    <div style={{ padding: 40, fontFamily: 'sans-serif' }}>
      <h1>CRM Rotas ✅</h1>
      <p>App carregado com sucesso!</p>
      <button onClick={() => supabase.auth.signOut()} style={{ padding: '10px 20px', marginTop: 20, cursor: 'pointer' }}>
        Sair
      </button>
    </div>
  )
}
