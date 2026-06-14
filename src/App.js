import { useState, useCallback, useMemo, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from './lib/supabase'
import { useAuth } from './lib/useAuth'

const ACCENT = '#2563EB'
const ACCENT_LIGHT = '#EFF6FF'
const SUCCESS = '#16A34A'
const WARNING = '#D97706'
const DANGER = '#DC2626'
const SURFACE = '#F8FAFC'
const CARD = '#FFFFFF'
const BORDER = '#E2E8F0'
const TEXT = '#0F172A'
const MUTED = '#64748B'

const fmt = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
const today = () => new Date().toISOString().split('T')[0]
const timeNow = () => new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })

const Badge = ({ color, children }) => (
  <span style={{ background: color + '18', color, border: `1px solid ${color}33`, borderRadius: 6, padding: '2px 10px', fontSize: 12, fontWeight: 600 }}>
    {children}
  </span>
)

const KpiCard = ({ label, value, sub, color = ACCENT }) => (
  <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 14, padding: '18px 22px', flex: 1, minWidth: 140 }}>
    <div style={{ fontSize: 12, color: MUTED, fontWeight: 600, marginBottom: 4, textTransform: 'uppercase', letterSpacing: 0.8 }}>{label}</div>
    <div style={{ fontSize: 26, fontWeight: 800, color }}>{value}</div>
    {sub && <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>{sub}</div>}
  </div>
)

export default function App() {
  const { user } = useAuth()
  const [clients, setClients] = useState([])
  const [sales, setSales] = useState([])
  const [routes, setRoutes] = useState([])
  const [selectedRoute, setSelectedRoute] = useState('')
  const [dailyGoal, setDailyGoal] = useState('')
  const [goalInput, setGoalInput] = useState('')
  const [selectedClient, setSelectedClient] = useState('')
  const [saleValue, setSaleValue] = useState('')
  const [saleNote, setSaleNote] = useState('')
  const [activeTab, setActiveTab] = useState('dashboard')
  const [toast, setToast] = useState(null)
  const [dragOver, setDragOver] = useState(false)
  const [clientSearch, setClientSearch] = useState('')
  const [tabSaleClient, setTabSaleClient] = useState(null)
  const [tabSaleClientInput, setTabSaleClientInput] = useState('')
  const [tabSaleValue, setTabSaleValue] = useState('')
  const [tabSaleNote, setTabSaleNote] = useState('')
  const [showSuggestions, setShowSuggestions] = useState(false)
  const [showPaste, setShowPaste] = useState(false)
  const [pasteText, setPasteText] = useState('')
  const [loading, setLoading] = useState(false)

  const showToast = (msg, type = 'success') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 3200)
  }

  const loadClients = useCallback(async () => {
    const { data, error } = await supabase.from('clients').select('*').eq('user_id', user.id).order('name')
    if (error) { showToast('Erro ao carregar clientes.', 'error'); return }
    setClients(data)
    setRoutes([...new Set(data.map(c => c.route))].sort())
  }, [user.id])

  const loadSales = useCallback(async () => {
    const { data, error } = await supabase.from('sales').select('*').eq('user_id', user.id).eq('date', today()).order('created_at')
    if (error) { showToast('Erro ao carregar vendas.', 'error'); return }
    setSales(data)
  }, [user.id])

  const loadGoal = useCallback(async (route) => {
    if (!route) return
    const { data } = await supabase.from('daily_goals').select('goal_value').eq('user_id', user.id).eq('route', route).eq('date', today()).single()
    setDailyGoal(data?.goal_value || '')
  }, [user.id])

  useEffect(() => { loadClients(); loadSales() }, [loadClients, loadSales])
  useEffect(() => { loadGoal(selectedRoute) }, [selectedRoute, loadGoal])

  const handleFile = useCallback((file) => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 })
        await importClients(data.slice(1).filter(r => r[0] && r[1]))
      } catch { showToast('Erro ao ler planilha.', 'error') }
    }
    reader.readAsBinaryString(file)
  }, [])

  const handlePaste = useCallback(async () => {
    try {
      const lines = pasteText.trim().split('\n').filter(Boolean)
      if (lines.length < 2) { showToast('Cole ao menos uma linha além do cabeçalho.', 'error'); return }
      const dataLines = lines[0].toLowerCase().includes('cliente') ? lines.slice(1) : lines
      const rows = dataLines.map(l => l.split('\t')).filter(c => c[0]?.trim() && c[1]?.trim())
      if (rows.length === 0) { showToast('Nenhum dado válido. Use Tab entre colunas.', 'error'); return }
      await importClients(rows)
      setShowPaste(false); setPasteText('')
    } catch { showToast('Erro ao processar dados.', 'error') }
  }, [pasteText])

  const importClients = async (rows) => {
    setLoading(true)
    await supabase.from('clients').delete().eq('user_id', user.id)
    const toInsert = rows.map(cols => ({
      user_id: user.id,
      name: String(cols[0]).trim(),
      route: String(cols[1]).trim(),
      inactive: cols[2] && String(cols[2]).trim().toLowerCase() === 'inativo',
    }))
    const { error } = await supabase.from('clients').insert(toInsert)
    if (error) { showToast('Erro ao salvar clientes.', 'error'); setLoading(false); return }
    await loadClients(); setSales([]); setSelectedRoute(''); setDailyGoal('')
    setLoading(false); showToast(`${toInsert.length} clientes importados!`)
  }

  const handleSetGoal = async () => {
    const v = parseFloat(goalInput)
    if (isNaN(v) || v <= 0) { showToast('Informe uma meta válida.', 'error'); return }
    const { error } = await supabase.from('daily_goals').upsert({ user_id: user.id, route: selectedRoute, goal_value: v, date: today() }, { onConflict: 'user_id,route,date' })
    if (error) { showToast('Erro ao salvar meta.', 'error'); return }
    setDailyGoal(v); setGoalInput(''); showToast(`Meta definida: ${fmt(v)}`)
  }

  const handleAddSale = async () => {
    if (!selectedClient || !saleValue || isNaN(parseFloat(saleValue))) { showToast('Selecione um cliente e informe o valor.', 'error'); return }
    const client = clients.find(c => c.id === selectedClient)
    const { data, error } = await supabase.from('sales').insert({ user_id: user.id, client_id: client.id, client_name: client.name, route: client.route, value: parseFloat(saleValue), note: saleNote, sale_time: timeNow(), date: today() }).select().single()
    if (error) { showToast('Erro ao registrar venda.', 'error'); return }
    setSales(prev => [...prev, data]); setSelectedClient(''); setSaleValue(''); setSaleNote('')
    showToast(`Venda de ${fmt(parseFloat(saleValue))} registrada!`)
  }

  const handleAddTabSale = async () => {
    if (!tabSaleClientInput.trim() || !tabSaleValue || isNaN(parseFloat(tabSaleValue))) { showToast('Informe o cliente e o valor.', 'error'); return }
    const value = parseFloat(tabSaleValue)
    const matched = tabSaleClient?.name === tabSaleClientInput ? tabSaleClient : null
    const { data, error } = await supabase.from('sales').insert({ user_id: user.id, client_id: matched?.id || null, client_name: tabSaleClientInput.trim(), route: matched?.route || selectedRoute || '—', value, note: tabSaleNote, sale_time: timeNow(), date: today() }).select().single()
    if (error) { showToast('Erro ao registrar venda.', 'error'); return }
    setSales(prev => [...prev, data]); setTabSaleClient(null); setTabSaleClientInput(''); setTabSaleValue(''); setTabSaleNote('')
    showToast(`Venda de ${fmt(value)} registrada!`)
  }

  const handleRemoveSale = async (id) => {
    const { error } = await supabase.from('sales').delete().eq('id', id)
    if (error) { showToast('Erro ao remover venda.', 'error'); return }
    setSales(prev => prev.filter(s => s.id !== id))
  }

  const routeClients = useMemo(() => selectedRoute ? clients.filter(c => c.route === selectedRoute) : [], [clients, selectedRoute])
  const routeSales = useMemo(() => sales.filter(s => s.route === selectedRoute), [sales, selectedRoute])
  const soldClientIds = useMemo(() => new Set(routeSales.map(s => s.client_id).filter(Boolean)), [routeSales])
  const inactiveSoldClients = useMemo(() => routeSales.filter(s => clients.find(c => c.id === s.client_id)?.inactive), [routeSales, clients])
  const activeRouteClients = useMemo(() => routeClients.filter(c => !c.inactive), [routeClients])
  const activeSoldIds = useMemo(() => new Set(routeSales.filter(s => { const c = clients.find(cl => cl.id === s.client_id); return !c || !c.inactive }).map(s => s.client_id).filter(Boolean)), [routeSales, clients])
  const totalSold = useMemo(() => routeSales.reduce((a, s) => a + s.value, 0), [routeSales])
  const remaining = activeRouteClients.length - activeSoldId
