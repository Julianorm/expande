import{useState,useCallback,useMemo,useEffect}from 'react'
import*as XLSX from 'xlsx'
import{supabase}from './lib/supabase'
import{useAuth}from './lib/useAuth'
const ACCENT='#2563EB',ACCENT_LIGHT='#EFF6FF',SUCCESS='#16A34A',WARNING='#D97706',DANGER='#DC2626',SURFACE='#F8FAFC',CARD='#FFFFFF',BORDER='#E2E8F0',TEXT='#0F172A',MUTED='#64748B'
const fmt=v=>v.toLocaleString('pt-BR',{style:'currency',currency:'BRL'})
const today=()=>new Date().toISOString().split('T')[0]
const timeNow=()=>new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})
const Badge=({color,children})=><span style={{background:color+'18',color,border:`1px solid ${color}33`,borderRadius:6,padding:'2px 10px',fontSize:12,fontWeight:600}}>{children}</span>
const KpiCard=({label,value,sub,color=ACCENT})=><div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,padding:'18px 22px',flex:1,minWidth:140}}><div style={{fontSize:12,color:MUTED,fontWeight:600,marginBottom:4,textTransform:'uppercase'}}>{label}</div><div style={{fontSize:26,fontWeight:800,color}}>{value}</div>{sub&&<div style={{fontSize:12,color:MUTED,marginTop:2}}>{sub}</div>}</div>
export default function App(){
const{user}=useAuth()
const[clients,setClients]=useState([])
const[sales,setSales]=useState([])
const[routes,setRoutes]=useState([])
const[selectedRoute,setSelectedRoute]=useState('')
const[dailyGoal,setDailyGoal]=useState('')
const[goalInput,setGoalInput]=useState('')
const[selectedClient,setSelectedClient]=useState('')
const[saleValue,setSaleValue]=useState('')
const[saleNote,setSaleNote]=useState('')
const[activeTab,setActiveTab]=useState('dashboard')
const[toast,setToast]=useState(null)
const[dragOver,setDragOver]=useState(false)
const[clientSearch,setClientSearch]=useState('')
const[tabSaleClient,setTabSaleClient]=useState(null)
const[tabSaleClientInput,setTabSaleClientInput]=useState('')
const[tabSaleValue,setTabSaleValue]=useState('')
const[tabSaleNote,setTabSaleNote]=useState('')
const[showSuggestions,setShowSuggestions]=useState(false)
const[showPaste,setShowPaste]=useState(false)
const[pasteText,setPasteText]=useState('')
const[loading,setLoading]=useState(false)
const[pedidoCliente,setPedidoCliente]=useState(null)
const[pedidoProdutos,setPedidoProdutos]=useState([])
const[pedidoSearch,setPedidoSearch]=useState('')
const[pedidoResultados,setPedidoResultados]=useState([])
const[pedidoFormaPgto,setPedidoFormaPgto]=useState('')
const[pedidoSituacao,setPedidoSituacao]=useState('Pedido S/ NFe')
const[pedidoLoading,setPedidoLoading]=useState(false)
const EGESTOR_API='https://qtogmmgkpnpkmvnkoxsz.supabase.co/functions/v1/egestor-api'
const showToast=(msg,type='success')=>{setToast({msg,type});setTimeout(()=>setToast(null),3200)}
const loadClients=useCallback(async()=>{if(!user?.id)return;const{data,error}=await supabase.from('clients').select('*').eq('user_id',user.id).order('name');if(error){showToast('Erro ao carregar clientes.','error');return}setClients(data);setRoutes([...new Set(data.map(c=>c.route))].sort())},[user?.id])
const loadSales=useCallback(async()=>{if(!user?.id)return;const{data,error}=await supabase.from('sales').select('*').eq('user_id',user.id).eq('date',today()).order('created_at');if(error){showToast('Erro ao carregar vendas.','error');return}setSales(data)},[user?.id])
const loadGoal=useCallback(async(route)=>{if(!route||!user?.id)return;const{data}=await supabase.from('daily_goals').select('goal_value').eq('user_id',user.id).eq('route',route).eq('date',today()).single();setDailyGoal(data?.goal_value||'')},[user?.id])
useEffect(()=>{if(user?.id){loadClients();loadSales()}},[loadClients,loadSales,user?.id])
useEffect(()=>{loadGoal(selectedRoute)},[selectedRoute,loadGoal])
const importClients=useCallback(async(rows)=>{if(!user?.id)return;setLoading(true);await supabase.from('clients').delete().eq('user_id',user.id);const toInsert=rows.map(cols=>({user_id:user.id,name:String(cols[0]).trim(),route:String(cols[1]).trim(),inactive:cols[2]&&String(cols[2]).trim().toLowerCase()==='inativo'}));const{error}=await supabase.from('clients').insert(toInsert);if(error){showToast('Erro ao salvar clientes.','error');setLoading(false);return}await loadClients();setSales([]);setSelectedRoute('');setDailyGoal('');setLoading(false);showToast(`${toInsert.length} clientes importados!`)},[user?.id,loadClients])
const handleFile=useCallback((file)=>{if(!file)return;const reader=new FileReader();reader.onload=async(e)=>{try{const wb=XLSX.read(e.target.result,{type:'binary'});const ws=wb.Sheets[wb.SheetNames[0]];const data=XLSX.utils.sheet_to_json(ws,{header:1});await importClients(data.slice(1).filter(r=>r[0]&&r[1]))}catch{showToast('Erro ao ler planilha.','error')}};reader.readAsBinaryString(file)},[importClients])
const handlePaste=useCallback(async()=>{try{const lines=pasteText.trim().split('\n').filter(Boolean);if(lines.length<2){showToast('Cole ao menos uma linha além do cabeçalho.','error');return}const dataLines=lines[0].toLowerCase().includes('cliente')?lines.slice(1):lines;const rows=dataLines.map(l=>l.split('\t')).filter(c=>c[0]?.trim()&&c[1]?.trim());if(rows.length===0){showToast('Nenhum dado válido.','error');return}await importClients(rows);setShowPaste(false);setPasteText('')}catch{showToast('Erro ao processar dados.','error')}},[pasteText,importClients])
const handleSetGoal=async()=>{if(!user?.id)return;const v=parseFloat(goalInput);if(isNaN(v)||v<=0){showToast('Informe uma meta válida.','error');return}const{error}=await supabase.from('daily_goals').upsert({user_id:user.id,route:selectedRoute,goal_value:v,date:today()},{onConflict:'user_id,route,date'});if(error){showToast('Erro ao salvar meta.','error');return}setDailyGoal(v);setGoalInput('');showToast(`Meta definida: ${fmt(v)}`)}
const handleAddSale=async()=>{if(!user?.id||!selectedClient||!saleValue||isNaN(parseFloat(saleValue))){showToast('Selecione um cliente e informe o valor.','error');return}const client=clients.find(c=>c.id===selectedClient);const{data,error}=await supabase.from('sales').insert({user_id:user.id,client_id:client.id,client_name:client.name,route:client.route,value:parseFloat(saleValue),note:saleNote,sale_time:timeNow(),date:today()}).select().single();if(error){showToast('Erro ao registrar venda.','error');return}setSales(prev=>[...prev,data]);setSelectedClient('');setSaleValue('');setSaleNote('');showToast(`Venda de ${fmt(parseFloat(saleValue))} registrada!`)}
const handleAddTabSale=async()=>{if(!user?.id||!tabSaleClientInput.trim()||!tabSaleValue||isNaN(parseFloat(tabSaleValue))){showToast('Informe o cliente e o valor.','error');return}const value=parseFloat(tabSaleValue);const matched=tabSaleClient?.name===tabSaleClientInput?tabSaleClient:null;const{data,error}=await supabase.from('sales').insert({user_id:user.id,client_id:matched?.id||null,client_name:tabSaleClientInput.trim(),route:matched?.route||selectedRoute||'—',value,note:tabSaleNote,sale_time:timeNow(),date:today()}).select().single();if(error){showToast('Erro ao registrar venda.','error');return}setSales(prev=>[...prev,data]);setTabSaleClient(null);setTabSaleClientInput('');setTabSaleValue('');setTabSaleNote('');showToast(`Venda de ${fmt(value)} registrada!`)}
const handleRemoveSale=async(id)=>{const{error}=await supabase.from('sales').delete().eq('id',id);if(error){showToast('Erro ao remover venda.','error');return}setSales(prev=>prev.filter(s=>s.id!==id))}
  const getSession=async()=>{const{data}=await supabase.auth.getSession();return data.session?.access_token||''}

const buscarProdutos=async(search)=>{
  if(search.length<2){setPedidoResultados([]);return}
  try{
    const token=await getSession()
    console.log('Token:',token?'ok':'vazio')
    const res=await fetch(`${EGESTOR_API}?action=produtos&search=${encodeURIComponent(search)}`,{headers:{'Authorization':`Bearer ${token}`}})
    console.log('Status:',res.status)
    const data=await res.json()
    console.log('Dados:',JSON.stringify(data))
    setPedidoResultados(Array.isArray(data)?data:[])
  }catch(err){
    console.error('Erro:',err.message)
    showToast('Erro ao buscar produtos','error')
  }
}

const addProdutoPedido=(produto)=>{
  setPedidoProdutos(prev=>{
    const existe=prev.find(p=>p.codigo===produto.codigo)
    if(existe)return prev.map(p=>p.codigo===produto.codigo?{...p,quant:p.quant+1}:p)
    return [...prev,{...produto,quant:1,vDesc:0}]
  })
  setPedidoSearch('')
  setPedidoResultados([])
}

const totalPedido=pedidoProdutos.reduce((acc,p)=>{
  const sub=p.precoVenda*p.quant
  const desc=sub*(p.vDesc||0)/100
  return acc+sub-desc
},0)

const confirmarPedido=async()=>{
  if(!pedidoCliente||pedidoProdutos.length===0||!pedidoFormaPgto){
    showToast('Preencha cliente, produtos e forma de pagamento.','error');return
  }
  setPedidoLoading(true)
  try{
    const res=await fetch(`${EGESTOR_API}?action=criar_venda`,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        codContato:pedidoCliente.erp_code,
        nomeContato:pedidoCliente.name,
        route:pedidoCliente.route,
        user_id:user.id,
        produtos:pedidoProdutos.map(p=>({codigo:p.codigo,quant:p.quant,preco:p.precoVenda,vDesc:p.vDesc||0})),
        codFormaPgto:parseInt(pedidoFormaPgto),
        situacaoOS:pedidoSituacao
      })
    })
    const result=await res.json()
    if(result.codigo){
      showToast(`Pedido #${result.codigo} criado no eGestor!`)
      setPedidoCliente(null);setPedidoProdutos([]);setPedidoFormaPgto('');setPedidoSituacao('Pedido S/ NFe')
      await loadSales()
    }else{
      showToast('Erro: '+JSON.stringify(result),'error')
    }
  }catch(err){
    showToast('Erro ao criar pedido','error')
  }
  setPedidoLoading(false)
}
  setPedidoLoading(true)
  const token=await getSession()
  const res=await fetch(`${EGESTOR_API}?action=criar_venda`,{
    method:'POST',
    headers:{'Authorization':`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify({
      codContato:pedidoCliente.erp_code,
      nomeContato:pedidoCliente.name,
      route:pedidoCliente.route,
      produtos:pedidoProdutos.map(p=>({codigo:p.codigo,quant:p.quant,preco:p.precoVenda,vDesc:p.vDesc||0})),
      codFormaPgto:parseInt(pedidoFormaPgto),
      situacaoOS:pedidoSituacao
    })
  })
  const result=await res.json()
  setPedidoLoading(false)
  if(result.codigo){
    showToast(`Pedido #${result.codigo} criado no eGestor!`)
    setPedidoCliente(null);setPedidoProdutos([]);setPedidoFormaPgto('');setPedidoSituacao('Pedido S/ NFe')
    await loadSales()
  }else{
    showToast('Erro ao criar pedido: '+JSON.stringify(result),'error')
  }
}
const routeClients=useMemo(()=>selectedRoute?clients.filter(c=>c.route===selectedRoute):[],[clients,selectedRoute])
const routeSales=useMemo(()=>sales.filter(s=>s.route===selectedRoute),[sales,selectedRoute])
const soldClientIds=useMemo(()=>new Set(routeSales.map(s=>s.client_id).filter(Boolean)),[routeSales])
const inactiveSoldClients=useMemo(()=>routeSales.filter(s=>clients.find(c=>c.id===s.client_id)?.inactive),[routeSales,clients])
const activeRouteClients=useMemo(()=>routeClients.filter(c=>!c.inactive),[routeClients])
const activeSoldIds=useMemo(()=>new Set(routeSales.filter(s=>{const c=clients.find(cl=>cl.id===s.client_id);return!c||!c.inactive}).map(s=>s.client_id).filter(Boolean)),[routeSales,clients])
const totalSold=useMemo(()=>routeSales.reduce((a,s)=>a+s.value,0),[routeSales])
const remaining=activeRouteClients.length-activeSoldIds.size
const avgTicket=activeSoldIds.size>0?totalSold/activeSoldIds.size:0
const goalProgress=dailyGoal?Math.min((totalSold/dailyGoal)*100,100):0
const clientSuggestions=useMemo(()=>{const pool=selectedRoute?routeClients:clients;if(!tabSaleClientInput.trim())return pool.slice(0,6);return pool.filter(c=>c.name.toLowerCase().includes(tabSaleClientInput.toLowerCase())).slice(0,6)},[clients,routeClients,selectedRoute,tabSaleClientInput])
const filteredClients=useMemo(()=>(selectedRoute?routeClients:clients).filter(c=>c.name.toLowerCase().includes(clientSearch.toLowerCase())),[routeClients,clients,selectedRoute,clientSearch])
const Tab=({id,label,icon})=><button onClick={()=>setActiveTab(id)} style={{background:activeTab===id?ACCENT:'transparent',color:activeTab===id?'#fff':MUTED,border:'none',borderRadius:8,padding:'8px 16px',fontWeight:600,fontSize:13,cursor:'pointer',display:'flex',alignItems:'center',gap:6}}><span>{icon}</span>{label}</button>
return(<div style={{minHeight:'100vh',background:SURFACE,fontFamily:"'Inter',system-ui,sans-serif",color:TEXT}}>
<div style={{background:CARD,borderBottom:`1px solid ${BORDER}`,padding:'0 20px'}}>
<div style={{maxWidth:1100,margin:'0 auto',display:'flex',alignItems:'center',gap:12,height:58}}>
<div style={{display:'flex',alignItems:'center',gap:8}}>
<div style={{width:30,height:30,background:ACCENT,borderRadius:8,display:'flex',alignItems:'center',justifyContent:'center',fontSize:16}}>📦</div>
<span style={{fontWeight:800,fontSize:17}}>CRM Rotas</span>
</div>
<button onClick={()=>{loadClients();loadSales()}} style={{background:'none',border:`1px solid ${BORDER}`,borderRadius:8,padding:'4px 10px',fontSize:13,color:MUTED,cursor:'pointer'}}>🔄</button>
<div style={{flex:1}}/>
<div style={{display:'flex',gap:4}}>
<Tab id="dashboard" label="Dashboard" icon="📊"/>
<Tab id="clientes" label="Clientes" icon="👥"/>
<Tab id="vendas" label="Vendas" icon="💰"/>
<Tab id="pedido" label="Pedido" icon="🛒"/>
</div>
<button onClick={()=>supabase.auth.signOut()} style={{background:'none',border:`1px solid ${BORDER}`,borderRadius:8,padding:'6px 12px',fontSize:12,color:MUTED,cursor:'pointer',fontWeight:600}}>Sair</button>
</div>
</div>
<div style={{maxWidth:1100,margin:'0 auto',padding:'24px 16px'}}>
{toast&&<div style={{position:'fixed',top:16,right:16,zIndex:1000,background:toast.type==='error'?DANGER:SUCCESS,color:'#fff',borderRadius:10,padding:'12px 18px',fontWeight:600,fontSize:14,boxShadow:'0 4px 20px #0003'}}>{toast.type==='error'?'❌':'✅'} {toast.msg}</div>}
<div style={{display:'flex',gap:12,marginBottom:20,flexWrap:'wrap'}}>
<div style={{flex:'0 0 auto',position:'relative'}}>
<label htmlFor="xl-input" style={{border:`2px dashed ${dragOver?ACCENT:BORDER}`,borderRadius:12,background:dragOver?ACCENT_LIGHT:CARD,padding:'12px 18px',display:'flex',alignItems:'center',gap:10,cursor:'pointer'}} onDrop={e=>{e.preventDefault();setDragOver(false);handleFile(e.dataTransfer.files[0])}} onDragOver={e=>{e.preventDefault();setDragOver(true)}} onDragLeave={()=>setDragOver(false)}>
<span style={{fontSize:20}}>{loading?'⏳':'📂'}</span>
<div><div style={{fontWeight:700,fontSize:13}}>{clients.length>0?`✓ ${clients.length} clientes`:'Importar Planilha'}</div><div style={{fontSize:11,color:MUTED}}>{clients.length>0?`${routes.length} rotas`:'Toque aqui ou use Colar Dados'}</div></div>
<input id="xl-input" type="file" accept=".xlsx,.xls" onChange={e=>handleFile(e.target.files[0])} style={{position:'absolute',width:1,height:1,opacity:0}}/>
</label>
<button onClick={()=>setShowPaste(v=>!v)} style={{marginTop:5,width:'100%',background:'none',border:`1px solid ${BORDER}`,borderRadius:8,padding:'6px 10px',fontSize:11,fontWeight:600,color:ACCENT,cursor:'pointer'}}>📋 {showPaste?'Fechar':'Colar Dados (celular)'}</button>
{showPaste&&<div style={{position:'absolute',top:'110%',left:0,zIndex:200,width:300,background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,boxShadow:'0 8px 32px #0003',padding:14}}>
<div style={{fontWeight:700,fontSize:13,marginBottom:6}}>📋 Colar dados</div>
<textarea rows={7} value={pasteText} onChange={e=>setPasteText(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'8px',fontSize:11,fontFamily:'monospace',resize:'vertical',boxSizing:'border-box'}}/>
<div style={{display:'flex',gap:8,marginTop:8}}>
<button onClick={handlePaste} style={{flex:1,background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'8px 0',fontWeight:700,fontSize:13,cursor:'pointer'}}>Importar</button>
<button onClick={()=>{setShowPaste(false);setPasteText('')}} style={{background:SURFACE,color:MUTED,border:`1px solid ${BORDER}`,borderRadius:8,padding:'8px 12px',fontWeight:600,fontSize:13,cursor:'pointer'}}>✕</button>
</div>
</div>}
</div>
{routes.length>0&&<div style={{flex:1,minWidth:180,background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'12px 16px',display:'flex',alignItems:'center',gap:10}}>
<span style={{fontSize:18}}>🗺️</span>
<div style={{flex:1}}><div style={{fontWeight:700,fontSize:12,marginBottom:4}}>Rota do Dia</div>
<select value={selectedRoute} onChange={e=>{setSelectedRoute(e.target.value);setDailyGoal('')}} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:7,padding:'6px 8px',fontSize:13,background:SURFACE,fontWeight:600,color:TEXT}}>
<option value="">Selecionar rota…</option>
{routes.map(r=><option key={r} value={r}>{r}</option>)}
</select></div>
</div>}
{selectedRoute&&<div style={{flex:'0 0 auto',background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'12px 16px',display:'flex',alignItems:'center',gap:10}}>
<span style={{fontSize:18}}>🎯</span>
<div><div style={{fontWeight:700,fontSize:12,marginBottom:4}}>Meta do Dia</div>
{dailyGoal?<div style={{display:'flex',alignItems:'center',gap:8}}><span style={{fontWeight:800,color:ACCENT,fontSize:15}}>{fmt(dailyGoal)}</span><button onClick={()=>setDailyGoal('')} style={{background:'none',border:'none',color:MUTED,cursor:'pointer'}}>✏️</button></div>
:<div style={{display:'flex',gap:6}}><input type="number" placeholder="R$ 0,00" value={goalInput} onChange={e=>setGoalInput(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleSetGoal()} style={{width:100,border:`1px solid ${BORDER}`,borderRadius:7,padding:'5px 8px',fontSize:13}}/><button onClick={handleSetGoal} style={{background:ACCENT,color:'#fff',border:'none',borderRadius:7,padding:'5px 12px',fontWeight:700,cursor:'pointer',fontSize:12}}>Definir</button></div>}
</div>
</div>}
</div>
{activeTab==='dashboard'&&<div>
{!selectedRoute?<div style={{textAlign:'center',padding:'70px 20px',color:MUTED}}><div style={{fontSize:44,marginBottom:10}}>🗺️</div><div style={{fontWeight:700,fontSize:17,color:TEXT}}>Selecione uma rota para começar</div></div>
:<>
<div style={{display:'flex',gap:12,marginBottom:16,flexWrap:'wrap'}}>
<KpiCard label="Clientes na Rota" value={activeRouteClients.length} sub={`Ativos • ${selectedRoute}`} color={ACCENT}/>
<KpiCard label="Atendidos" value={activeSoldIds.size} sub={`${activeRouteClients.length>0?Math.round((activeSoldIds.size/activeRouteClients.length)*100):0}% da rota`} color={SUCCESS}/>
<KpiCard label="Restantes" value={remaining} sub={remaining===0?'Rota concluída! 🎉':'a visitar'} color={remaining===0?SUCCESS:WARNING}/>
<KpiCard label="Ticket Médio" value={fmt(avgTicket)} sub={`${activeSoldIds.size} venda(s)`} color={ACCENT}/>
<KpiCard label="Total Vendido" value={fmt(totalSold)} sub={dailyGoal?`Meta: ${fmt(dailyGoal)}`:'sem meta'} color={totalSold>=(dailyGoal||Infinity)?SUCCESS:TEXT}/>
</div>
{dailyGoal>0&&<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,padding:'16px 20px',marginBottom:16}}>
<div style={{display:'flex',justifyContent:'space-between',marginBottom:8}}><span style={{fontWeight:700,fontSize:14}}>🎯 Progresso da Meta</span><span style={{fontWeight:800,color:goalProgress>=100?SUCCESS:ACCENT}}>{goalProgress.toFixed(1)}%</span></div>
<div style={{background:SURFACE,borderRadius:99,height:10,overflow:'hidden'}}><div style={{width:`${goalProgress}%`,height:'100%',background:goalProgress>=100?SUCCESS:ACCENT,borderRadius:99,transition:'width 0.4s'}}/></div>
<div style={{display:'flex',justifyContent:'space-between',marginTop:6,fontSize:12,color:MUTED}}><span>{fmt(totalSold)} vendidos</span><span>Faltam {fmt(Math.max(dailyGoal-totalSold,0))}</span></div>
</div>}
{inactiveSoldClients.length>0&&<div style={{background:'#FFF7ED',border:`1.5px solid ${WARNING}55`,borderRadius:14,padding:'16px 20px',marginBottom:16}}>
<div style={{display:'flex',alignItems:'center',gap:8,marginBottom:12}}>
<span>⚠️</span><span style={{fontWeight:700,fontSize:14,color:WARNING}}>Clientes Inativos Atendidos</span>
<Badge color={WARNING}>{inactiveSoldClients.length}</Badge><div style={{flex:1}}/><span style={{fontWeight:800,color:WARNING}}>{fmt(inactiveSoldClients.reduce((a,s)=>a+s.value,0))}</span>
</div>
<table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}><tbody>{inactiveSoldClients.map(s=><tr key={s.id}><td style={{padding:'6px 8px',fontWeight:600}}>{s.client_name}</td><td style={{padding:'6px 8px'}}><Badge color={WARNING}>{fmt(s.value)}</Badge></td><td style={{padding:'6px 8px',color:MUTED,fontSize:12}}>{s.note||'—'}</td></tr>)}</tbody></table>
<div style={{marginTop:8,fontSize:11,color:WARNING,fontStyle:'italic'}}>* Valores incluídos no total vendido.</div>
</div>}
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,padding:'18px 20px',marginBottom:16}}>
<div style={{fontWeight:700,fontSize:15,marginBottom:14}}>💰 Registrar Venda</div>
<div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'flex-end'}}>
<div style={{flex:2,minWidth:160}}><label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>CLIENTE</label>
<select value={selectedClient} onChange={e=>setSelectedClient(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,background:SURFACE}}>
<option value="">Selecionar…</option>
{routeClients.map(c=><option key={c.id} value={c.id}>{soldClientIds.has(c.id)?'✅ ':''}{c.name}{c.inactive?' ⛔':''}</option>)}
</select></div>
<div style={{flex:1,minWidth:110}}><label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>VALOR (R$)</label>
<input type="number" placeholder="0,00" value={saleValue} onChange={e=>setSaleValue(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleAddSale()} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,boxSizing:'border-box'}}/></div>
<div style={{flex:2,minWidth:140}}><label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>OBSERVAÇÃO</label>
<input type="text" placeholder="Opcional…" value={saleNote} onChange={e=>setSaleNote(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,boxSizing:'border-box'}}/></div>
<button onClick={handleAddSale} style={{background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'10px 20px',fontWeight:700,fontSize:13,cursor:'pointer',whiteSpace:'nowrap'}}>+ Registrar</button>
</div>
</div>
{routeSales.length>0&&<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,overflow:'hidden'}}>
<div style={{padding:'14px 18px',borderBottom:`1px solid ${BORDER}`,fontWeight:700,fontSize:14}}>Vendas de Hoje — {selectedRoute}</div>
<table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
<thead><tr style={{background:SURFACE}}>{['Hora','Cliente','Valor','Obs.',''].map(h=><th key={h} style={{padding:'9px 14px',textAlign:'left',fontWeight:600,fontSize:11,color:MUTED,borderBottom:`1px solid ${BORDER}`,textTransform:'uppercase'}}>{h}</th>)}</tr></thead>
<tbody>{[...routeSales].reverse().map((s,i)=><tr key={s.id} style={{borderBottom:`1px solid ${BORDER}`,background:i%2===0?CARD:SURFACE}}>
<td style={{padding:'10px 14px',color:MUTED,fontWeight:600}}>{s.sale_time}</td>
<td style={{padding:'10px 14px',fontWeight:600}}>{s.client_name}</td>
<td style={{padding:'10px 14px'}}><Badge color={SUCCESS}>{fmt(s.value)}</Badge></td>
<td style={{padding:'10px 14px',color:MUTED}}>{s.note||'—'}</td>
<td style={{padding:'10px 14px'}}><button onClick={()=>handleRemoveSale(s.id)} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:14}}>✕</button></td>
</tr>)}</tbody>
</table>
</div>}
</>}
</div>}
{activeTab==='clientes'&&<div>
{clients.length===0?<div style={{textAlign:'center',padding:'70px 20px',color:MUTED}}><div style={{fontSize:44,marginBottom:10}}>📂</div><div style={{fontWeight:700,fontSize:17,color:TEXT}}>Nenhuma planilha importada</div></div>
:<><div style={{display:'flex',gap:10,marginBottom:14,alignItems:'center'}}>
<input type="text" placeholder="🔍 Buscar cliente…" value={clientSearch} onChange={e=>setClientSearch(e.target.value)} style={{flex:1,border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 12px',fontSize:13}}/>
<Badge color={ACCENT}>{filteredClients.length} clientes</Badge>
</div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,overflow:'hidden'}}>
<table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
<thead><tr style={{background:SURFACE}}>{['Cliente','Rota','Cadastro','Atendimento'].map(h=><th key={h} style={{padding:'10px 14px',textAlign:'left',fontWeight:600,fontSize:11,color:MUTED,borderBottom:`1px solid ${BORDER}`,textTransform:'uppercase'}}>{h}</th>)}</tr></thead>
<tbody>{filteredClients.map((c,i)=><tr key={c.id} style={{borderBottom:`1px solid ${BORDER}`,background:c.inactive?'#FFF7ED':i%2===0?CARD:SURFACE}}>
<td style={{padding:'10px 14px',fontWeight:600}}>{c.name}</td>
<td style={{padding:'10px 14px'}}><Badge color={ACCENT}>{c.route}</Badge></td>
<td style={{padding:'10px 14px'}}>{c.inactive?<Badge color={WARNING}>⛔ Inativo</Badge>:<Badge color={SUCCESS}>✓ Ativo</Badge>}</td>
<td style={{padding:'10px 14px'}}>{soldClientIds.has(c.id)?<Badge color={SUCCESS}>✅ Vendido</Badge>:selectedRoute===c.route?<Badge color={MUTED}>⏳ Pendente</Badge>:<Badge color={MUTED}>—</Badge>}</td>
</tr>)}</tbody>
</table>
</div></>}
</div>}
{activeTab==='vendas'&&<div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,padding:'18px 20px',marginBottom:16}}>
<div style={{fontWeight:700,fontSize:15,marginBottom:14}}>💰 Registrar Venda</div>
<div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'flex-end'}}>
<div style={{flex:2,minWidth:180,position:'relative'}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>CLIENTE {selectedRoute&&<span style={{color:ACCENT}}>— {selectedRoute}</span>}</label>
<input type="text" placeholder={selectedRoute?`Buscar na ${selectedRoute}…`:'Selecione uma rota primeiro…'} value={tabSaleClientInput} disabled={!selectedRoute&&clients.length>0}
onChange={e=>{setTabSaleClientInput(e.target.value);setTabSaleClient(null);setShowSuggestions(true)}}
onFocus={()=>setShowSuggestions(true)} onBlur={()=>setTimeout(()=>setShowSuggestions(false),150)}
style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,boxSizing:'border-box'}}/>
{showSuggestions&&clientSuggestions.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,zIndex:100,background:CARD,border:`1px solid ${BORDER}`,borderRadius:8,boxShadow:'0 8px 24px #0002',marginTop:4,overflow:'hidden',maxHeight:240,overflowY:'auto'}}>
<div style={{padding:'5px 12px',fontSize:11,fontWeight:700,color:MUTED,borderBottom:`1px solid ${BORDER}`,background:SURFACE,textTransform:'uppercase'}}>{clientSuggestions.length} cliente(s) na rota</div>
{clientSuggestions.map(c=><div key={c.id} onMouseDown={()=>{setTabSaleClient(c);setTabSaleClientInput(c.name);setShowSuggestions(false)}}
style={{padding:'9px 12px',cursor:'pointer',fontSize:13,display:'flex',alignItems:'center',gap:8,borderBottom:`1px solid ${BORDER}`}}
onMouseEnter={e=>e.currentTarget.style.background=ACCENT_LIGHT}
onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
<span style={{fontWeight:600}}>{c.name}</span>
{c.inactive&&<Badge color={WARNING}>Inativo</Badge>}
{soldClientIds.has(c.id)&&<Badge color={SUCCESS}>✅</Badge>}
</div>)}
</div>}
</div>
<div style={{flex:1,minWidth:110}}><label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>VALOR (R$)</label>
<input type="number" placeholder="0,00" value={tabSaleValue} onChange={e=>setTabSaleValue(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleAddTabSale()} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,boxSizing:'border-box'}}/></div>
<div style={{flex:2,minWidth:130}}><label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>OBSERVAÇÃO</label>
<input type="text" placeholder="Opcional…" value={tabSaleNote} onChange={e=>setTabSaleNote(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,boxSizing:'border-box'}}/></div>
<button onClick={handleAddTabSale} style={{background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'10px 20px',fontWeight:700,fontSize:13,cursor:'pointer',whiteSpace:'nowrap'}}>+ Registrar</button>
</div>
</div>
{sales.length>0&&<div style={{display:'flex',gap:12,marginBottom:16,flexWrap:'wrap'}}>
<KpiCard label="Total de Vendas" value={sales.length} sub="hoje"/>
<KpiCard label="Receita Total" value={fmt(sales.reduce((a,s)=>a+s.value,0))} color={SUCCESS}/>
<KpiCard label="Ticket Médio" value={fmt(sales.reduce((a,s)=>a+s.value,0)/sales.length)} color={ACCENT}/>
</div>}
{sales.length===0?<div style={{textAlign:'center',padding:'48px 20px',color:MUTED,background:CARD,border:`1px solid ${BORDER}`,borderRadius:14}}><div style={{fontSize:36,marginBottom:8}}>📋</div><div style={{fontWeight:700,color:TEXT}}>Nenhuma venda hoje</div></div>
:<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,overflow:'hidden'}}>
<table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
<thead><tr style={{background:SURFACE}}>{['Hora','Cliente','Rota','Valor','Obs.',''].map(h=><th key={h} style={{padding:'9px 14px',textAlign:'left',fontWeight:600,fontSize:11,color:MUTED,borderBottom:`1px solid ${BORDER}`,textTransform:'uppercase'}}>{h}</th>)}</tr></thead>
<tbody>{[...sales].reverse().map((s,i)=><tr key={s.id} style={{borderBottom:`1px solid ${BORDER}`,background:i%2===0?CARD:SURFACE}}>
<td style={{padding:'10px 14px',color:MUTED,fontWeight:600}}>{s.sale_time}</td>
<td style={{padding:'10px 14px',fontWeight:600}}>{s.client_name}</td>
<td style={{padding:'10px 14px'}}><Badge color={ACCENT}>{s.route}</Badge></td>
<td style={{padding:'10px 14px'}}><Badge color={SUCCESS}>{fmt(s.value)}</Badge></td>
<td style={{padding:'10px 14px',color:MUTED}}>{s.note||'—'}</td>
<td style={{padding:'10px 14px'}}><button onClick={()=>handleRemoveSale(s.id)} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:14}}>✕</button></td>
</tr>)}</tbody>
</table>
</div>}
</div>}
  {activeTab==='pedido'&&<div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:14,padding:'18px 20px',marginBottom:16}}>
<div style={{fontWeight:700,fontSize:15,marginBottom:14}}>🛒 Novo Pedido</div>

<div style={{marginBottom:12}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>CLIENTE</label>
{pedidoCliente?<div style={{display:'flex',alignItems:'center',gap:8,background:ACCENT_LIGHT,borderRadius:8,padding:'8px 12px'}}>
<span style={{fontWeight:700,color:ACCENT,flex:1}}>{pedidoCliente.name}</span>
<button onClick={()=>setPedidoCliente(null)} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:14}}>✕</button>
</div>
:<select onChange={e=>{const c=clients.find(cl=>cl.id===e.target.value);setPedidoCliente(c||null)}} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,background:SURFACE}}>
<option value="">Selecionar cliente…</option>
{(selectedRoute?routeClients:clients).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
</select>}
</div>

<div style={{marginBottom:12}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>SITUAÇÃO</label>
<select value={pedidoSituacao} onChange={e=>setPedidoSituacao(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,background:SURFACE}}>
<option>Pedido S/ NFe</option>
<option>Pedido C/ NFe</option>
<option>Bonificação</option>
<option>Troca</option>
</select>
</div>

<div style={{marginBottom:12}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>FORMA DE PAGAMENTO</label>
<select value={pedidoFormaPgto} onChange={e=>setPedidoFormaPgto(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,background:SURFACE}}>
<option value="">Selecionar…</option>
<option value="1">Dinheiro</option>
<option value="2">Cheque</option>
<option value="8">Transf. eletrônica (Pix/Ted)</option>
<option value="16">Boleto Sicoob</option>
<option value="17">Débito em Conta</option>
</select>
</div>

<div style={{marginBottom:16,position:'relative'}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>BUSCAR PRODUTO</label>
<input type="text" placeholder="Digite o nome do produto…" value={pedidoSearch}
onChange={e=>{setPedidoSearch(e.target.value);buscarProdutos(e.target.value)}}
style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'9px 10px',fontSize:13,boxSizing:'border-box'}}/>
{pedidoResultados.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,zIndex:100,background:CARD,border:`1px solid ${BORDER}`,borderRadius:8,boxShadow:'0 8px 24px #0002',marginTop:4,maxHeight:240,overflowY:'auto'}}>
{pedidoResultados.map(p=><div key={p.codigo} onMouseDown={()=>addProdutoPedido(p)}
style={{padding:'10px 14px',cursor:'pointer',borderBottom:`1px solid ${BORDER}`,fontSize:13}}
onMouseEnter={e=>e.currentTarget.style.background=ACCENT_LIGHT}
onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
<div style={{fontWeight:600}}>{p.descricao}</div>
<div style={{fontSize:11,color:MUTED}}>Cód: {p.codigoProprio} • R$ {p.precoVenda?.toFixed(2)}</div>
</div>)}
</div>}
</div>

{pedidoProdutos.length>0&&<div style={{marginBottom:16}}>
<div style={{fontWeight:700,fontSize:13,marginBottom:8}}>Produtos do Pedido</div>
<table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
<thead><tr style={{background:SURFACE}}>{['Produto','Qtd','Preço','Desc%','Total',''].map(h=><th key={h} style={{padding:'8px 10px',textAlign:'left',fontWeight:600,fontSize:11,color:MUTED,borderBottom:`1px solid ${BORDER}`}}>{h}</th>)}</tr></thead>
<tbody>{pedidoProdutos.map((p,i)=>{
const sub=p.precoVenda*p.quant
const desc=sub*(p.vDesc||0)/100
const total=sub-desc
return<tr key={p.codigo} style={{borderBottom:`1px solid ${BORDER}`,background:i%2===0?CARD:SURFACE}}>
<td style={{padding:'8px 10px',fontWeight:600,fontSize:12}}>{p.descricao}</td>
<td style={{padding:'8px 10px'}}><input type="number" min="1" value={p.quant} onChange={e=>setPedidoProdutos(prev=>prev.map(x=>x.codigo===p.codigo?{...x,quant:parseFloat(e.target.value)||1}:x))} style={{width:60,border:`1px solid ${BORDER}`,borderRadius:6,padding:'4px 6px',fontSize:12}}/></td>
<td style={{padding:'8px 10px'}}><input type="number" min="0" step="0.01" value={p.precoVenda} onChange={e=>setPedidoProdutos(prev=>prev.map(x=>x.codigo===p.codigo?{...x,precoVenda:parseFloat(e.target.value)||0}:x))} style={{width:80,border:`1px solid ${BORDER}`,borderRadius:6,padding:'4px 6px',fontSize:12}}/></td>
<td style={{padding:'8px 10px'}}><input type="number" min="0" max="100" value={p.vDesc||0} onChange={e=>setPedidoProdutos(prev=>prev.map(x=>x.codigo===p.codigo?{...x,vDesc:parseFloat(e.target.value)||0}:x))} style={{width:60,border:`1px solid ${BORDER}`,borderRadius:6,padding:'4px 6px',fontSize:12}}/></td>
<td style={{padding:'8px 10px',fontWeight:700,color:SUCCESS}}>{fmt(total)}</td>
<td style={{padding:'8px 10px'}}><button onClick={()=>setPedidoProdutos(prev=>prev.filter(x=>x.codigo!==p.codigo))} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:14}}>✕</button></td>
</tr>})}
</tbody>
</table>
<div style={{textAlign:'right',marginTop:12,fontWeight:800,fontSize:18,color:ACCENT}}>
Total: {fmt(totalPedido)}
</div>
</div>}

<button onClick={confirmarPedido} disabled={pedidoLoading} style={{width:'100%',background:pedidoLoading?MUTED:SUCCESS,color:'#fff',border:'none',borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:15,cursor:pedidoLoading?'not-allowed':'pointer'}}>
{pedidoLoading?'Criando pedido…':'✅ Confirmar Pedido'}
</button>
</div>
</div>}
</div>
</div>)
}
