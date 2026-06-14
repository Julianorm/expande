import { useState } from 'react'
import { supabase } from '../lib/supabase'

const ACCENT = '#2563EB'
const BORDER = '#E2E8F0'
const DANGER = '#DC2626'
const MUTED = '#64748B'

export default function Login() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const [mode, setMode] = useState('login')
  const [success, setSuccess] = useState('')

  const handle = async () => {
    setError('')
    setSuccess('')
    setLoading(true)
    if (mode === 'login') {
      const { error } = await supabase.auth.signInWithPassword({ email, password })
      if (error) setError('E-mail ou senha incorretos.')
    } else {
      const { error } = await supabase.auth.signUp({ email, password })
      if (error) setError(error.message)
      else setSuccess('Conta criada! Verifique seu e-mail para confirmar.')
    }
    setLoading(false)
  }

  return (
    <div style={{
      minHeight: '100vh', display: 'flex', alignItems: 'center',
      justifyContent: 'center', background: '#F8FAFC', padding: 20,
    }}>
      <div style={{
        background: '#fff', border: `1px solid ${BORDER}`, borderRadius: 16,
        padding: '36px 32px', width: '100%', maxWidth: 380, boxShadow: '0 4px 24px #0001',
      }}>
        <div style={{ textAlign: 'center', marginBottom: 28 }}>
          <div style={{
            width: 48, height: 48, background: ACCENT, borderRadius: 12,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontSize: 24, margin: '0 auto 12px',
          }}>📦</div>
          <div style={{ fontWeight: 800, fontSize: 22 }}>CRM Rotas</div>
          <div style={{ color: MUTED, fontSize: 14, marginTop: 4 }}>
            {mode === 'login' ? 'Entre na sua conta' : 'Criar nova conta'}
          </div>
        </div>

        {error && (
          <div style={{ background: '#FEF2F2', border: `1px solid ${DANGER}33`, borderRadius: 8, padding: '10px 14px', color: DANGER, fontSize: 13, marginBottom: 16 }}>
            {error}
          </div>
        )}
        {success && (
          <div style={{ background: '#F0FDF4', border: '1px solid #16A34A33', borderRadius: 8, padding: '10px 14px', color: '#16A34A', fontSize: 13, marginBottom: 16 }}>
            {success}
          </div>
        )}

        <div style={{ marginBottom: 14 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: MUTED, display: 'block', marginBottom: 5 }}>E-MAIL</label>
          <input
            type="email" value={email} onChange={e => setEmail(e.target.value)}
            placeholder="seu@email.com"
            style={{ width: '100%', border: `1px solid ${BORDER}`, borderRadius: 8, padding: '10px 12px', fontSize: 14, boxSizing: 'border-box' }}
          />
        </div>

        <div style={{ marginBottom: 22 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: MUTED, display: 'block', marginBottom: 5 }}>SENHA</label>
          <input
            type="password" value={password} onChange={e => setPassword(e.target.value)}
            placeholder="••••••••"
            onKeyDown={e => e.key === 'Enter' && handle()}
            style={{ width: '100%', border: `1px solid ${BORDER}`, borderRadius: 8, padding: '10px 12px', fontSize: 14, boxSizing: 'border-box' }}
          />
        </div>

        <button
          onClick={handle} disabled={loading}
          style={{
            width: '100%', background: ACCENT, color: '#fff', border: 'none',
            borderRadius: 8, padding: '12px 0', fontWeight: 700, fontSize: 15,
            cursor: loading ? 'not-allowed' : 'pointer', opacity: loading ? 0.7 : 1,
          }}
        >
          {loading ? 'Aguarde…' : mode === 'login' ? 'Entrar' : 'Criar conta'}
        </button>

        <div style={{ textAlign: 'center', marginTop: 18, fontSize: 13, color: MUTED }}>
          {mode === 'login' ? 'Não tem conta? ' : 'Já tem conta? '}
          <span
            onClick={() => { setMode(mode === 'login' ? 'register' : 'login'); setError(''); setSuccess('') }}
            style={{ color: ACCENT, fontWeight: 600, cursor: 'pointer' }}
          >
            {mode === 'login' ? 'Criar conta' : 'Entrar'}
          </span>
        </div>
      </div>
    </div>
  )
}
