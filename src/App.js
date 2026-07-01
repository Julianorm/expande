import{useState,useCallback,useMemo,useEffect}from 'react'
import*as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import{supabase}from './lib/supabase'
import{useAuth}from './lib/useAuth'
const ACCENT='#2563EB',ACCENT_LIGHT='#EFF6FF',SUCCESS='#16A34A',WARNING='#D97706',DANGER='#DC2626',SURFACE='#F8FAFC',CARD='#FFFFFF',BORDER='#E2E8F0',TEXT='#0F172A',MUTED='#64748B'
const fmt=v=>v.toLocaleString('pt-BR',{style:'currency',currency:'BRL'})
const today=()=>new Date().toISOString().split('T')[0]
const timeNow=()=>new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})
const EGESTOR_API='https://qtogmmgkpnpkmvnkoxsz.supabase.co/functions/v1/egestor-api'
const ADMIN_API='https://qtogmmgkpnpkmvnkoxsz.supabase.co/functions/v1/admin-api'
const ADMIN_ID='7ad867ea-496c-412c-8acd-5cc7c21eca0e'
const Badge=({color,children})=><span style={{background:color+'18',color,border:`1px solid ${color}33`,borderRadius:6,padding:'2px 8px',fontSize:11,fontWeight:600}}>{children}</span>
const KpiCard=({label,value,sub,color=ACCENT})=><div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px',flex:1,minWidth:130}}><div style={{fontSize:11,color:MUTED,fontWeight:600,marginBottom:4,textTransform:'uppercase'}}>{label}</div><div style={{fontSize:22,fontWeight:800,color}}>{value}</div>{sub&&<div style={{fontSize:11,color:MUTED,marginTop:2}}>{sub}</div>}</div>
export default function App(){
const{user}=useAuth()
const[clients,setClients]=useState([])
const[sales,setSales]=useState([])
const[orders,setOrders]=useState([])
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
const[pedidoClienteSearch,setPedidoClienteSearch]=useState('')
const[pedidoPdf,setPedidoPdf]=useState(null)
const[modalProdutos,setModalProdutos]=useState(false)
const[modalProdutoSearch,setModalProdutoSearch]=useState('')
const[todosProdutos,setTodosProdutos]=useState([])
const[pedidoProdutos,setPedidoProdutos]=useState([])
const[pedidoSearch,setPedidoSearch]=useState('')
const[pedidoResultados,setPedidoResultados]=useState([])
const[pedidoFormaPgto,setPedidoFormaPgto]=useState('1')
const[pedidoVencimento,setPedidoVencimento]=useState(today())
const[pedidoSituacao,setPedidoSituacao]=useState('Pedido S/ NFe')
useEffect(()=>{
  if(pedidoSituacao==='Pedido S/ NFe')setPedidoFormaPgto('1')
  else if(pedidoSituacao==='Pedido C/ NFe')setPedidoFormaPgto('16')
  else setPedidoFormaPgto('')
},[pedidoSituacao])
const[pedidoLoading,setPedidoLoading]=useState(false)
const[clientePerfil,setClientePerfil]=useState(null)
const[perfilData,setPerfilData]=useState(null)
const[perfilLoading,setPerfilLoading]=useState(false)
const[configUsuarios,setConfigUsuarios]=useState([])
const[configVendedores,setConfigVendedores]=useState([])
const[novoEmail,setNovoEmail]=useState('')
const[novaSenha,setNovaSenha]=useState('')
const[novoNome,setNovoNome]=useState('')
const[novoEgestorCode,setNovoEgestorCode]=useState('')
const[configLoading,setConfigLoading]=useState(false)
const[senhaUser,setSenhaUser]=useState({})
const[produtoModal,setProdutoModal]=useState(null)
const[produtoModalQuant,setProdutoModalQuant]=useState(1)
const[produtoModalDesc,setProdutoModalDesc]=useState(0)
const[exportLoading,setExportLoading]=useState(false)
const[editandoOrder,setEditandoOrder]=useState(null)
const[editOrderProdutos,setEditOrderProdutos]=useState([])
const[editOrderSearch,setEditOrderSearch]=useState('')
const[editOrderResultados,setEditOrderResultados]=useState([])
const[editOrderFormaPgto,setEditOrderFormaPgto]=useState('')
const[editOrderSituacao,setEditOrderSituacao]=useState('Pedido S/ NFe')
const showToast=(msg,type='success')=>{setToast({msg,type});setTimeout(()=>setToast(null),3200)}
const loadClients=useCallback(async()=>{if(!user?.id)return;const{data:userCfg}=await supabase.from('user_config').select('rotas').eq('user_id',user.id).single();const rotasUser=userCfg?.rotas||[];const query=supabase.from('clients').select('*').order('route').order('ordem');const{data,error}=rotasUser.length>0?await query.in('route',rotasUser):await query;if(error){showToast('Erro ao carregar clientes.','error');return}setClients(data||[]);setRoutes([...new Set((data||[]).map(c=>c.route))].sort())},[user?.id])
const loadSales=useCallback(async()=>{if(!user?.id)return;const{data,error}=await supabase.from('sales').select('*').eq('user_id',user.id).eq('date',today()).order('created_at');if(error){showToast('Erro ao carregar vendas.','error');return}setSales(data)},[user?.id])
const loadOrders=useCallback(async()=>{if(!user?.id)return;const{data,error}=await supabase.from('orders').select('*').eq('user_id',user.id).eq('status','pendente').eq('date',today()).order('created_at');if(error){showToast('Erro ao carregar pedidos.','error');return}setOrders(data||[])},[user?.id])
const loadGoal=useCallback(async(route)=>{if(!route||!user?.id)return;const{data}=await supabase.from('daily_goals').select('goal_value').eq('user_id',user.id).eq('route',route).eq('date',today()).single();setDailyGoal(data?.goal_value||'')},[user?.id])
useEffect(()=>{if(user?.id){loadClients();loadSales();loadOrders()}},[loadClients,loadSales,loadOrders,user?.id])
useEffect(()=>{loadGoal(selectedRoute)},[selectedRoute,loadGoal])
const importClients=useCallback(async(rows)=>{if(!user?.id)return;setLoading(true);await supabase.from('clients').delete().eq('empresa_id','mageski');const toInsert=rows.map(cols=>({empresa_id:'mageski',name:String(cols[0]).trim(),route:String(cols[1]).trim(),inactive:cols[2]&&String(cols[2]).trim().toLowerCase()==='inativo'}));const{error}=await supabase.from('clients').insert(toInsert);if(error){showToast('Erro ao salvar clientes.','error');setLoading(false);return}await loadClients();setSales([]);setSelectedRoute('');setDailyGoal('');setLoading(false);showToast(`${toInsert.length} clientes importados!`)},[user?.id,loadClients])
const handleFile=useCallback((file)=>{if(!file)return;const reader=new FileReader();reader.onload=async(e)=>{try{const wb=XLSX.read(e.target.result,{type:'binary'});const ws=wb.Sheets[wb.SheetNames[0]];const data=XLSX.utils.sheet_to_json(ws,{header:1});await importClients(data.slice(1).filter(r=>r[0]&&r[1]))}catch{showToast('Erro ao ler planilha.','error')}};reader.readAsBinaryString(file)},[importClients])
const handlePaste=useCallback(async()=>{try{const lines=pasteText.trim().split('\n').filter(Boolean);if(lines.length<2){showToast('Cole ao menos uma linha além do cabeçalho.','error');return}const dataLines=lines[0].toLowerCase().includes('cliente')?lines.slice(1):lines;const rows=dataLines.map(l=>l.split('\t')).filter(c=>c[0]?.trim()&&c[1]?.trim());if(rows.length===0){showToast('Nenhum dado válido.','error');return}await importClients(rows);setShowPaste(false);setPasteText('')}catch{showToast('Erro ao processar dados.','error')}},[pasteText,importClients])
const handleSetGoal=async()=>{if(!user?.id)return;const v=parseFloat(goalInput);if(isNaN(v)||v<=0){showToast('Informe uma meta válida.','error');return}const{error}=await supabase.from('daily_goals').upsert({user_id:user.id,route:selectedRoute,goal_value:v,date:today()},{onConflict:'user_id,route,date'});if(error){showToast('Erro ao salvar meta.','error');return}setDailyGoal(v);setGoalInput('');showToast(`Meta definida: ${fmt(v)}`)}
const handleAddSale=async()=>{if(!user?.id||!selectedClient||!saleValue||isNaN(parseFloat(saleValue))){showToast('Selecione um cliente e informe o valor.','error');return}const client=clients.find(c=>c.id===selectedClient);const{data,error}=await supabase.from('sales').insert({user_id:user.id,client_id:client.id,client_name:client.name,route:client.route,value:parseFloat(saleValue),note:saleNote,sale_time:timeNow(),date:today()}).select().single();if(error){showToast('Erro ao registrar venda.','error');return}setSales(prev=>[...prev,data]);setSelectedClient('');setSaleValue('');setSaleNote('');showToast(`Venda de ${fmt(parseFloat(saleValue))} registrada!`)}
const handleAddTabSale=async()=>{if(!user?.id||!tabSaleClientInput.trim()||!tabSaleValue||isNaN(parseFloat(tabSaleValue))){showToast('Informe o cliente e o valor.','error');return}const value=parseFloat(tabSaleValue);const matched=tabSaleClient?.name===tabSaleClientInput?tabSaleClient:null;const{data,error}=await supabase.from('sales').insert({user_id:user.id,client_id:matched?.id||null,client_name:tabSaleClientInput.trim(),route:matched?.route||selectedRoute||'—',value,note:tabSaleNote,sale_time:timeNow(),date:today()}).select().single();if(error){showToast('Erro ao registrar venda.','error');return}setSales(prev=>[...prev,data]);setTabSaleClient(null);setTabSaleClientInput('');setTabSaleValue('');setTabSaleNote('');showToast(`Venda de ${fmt(value)} registrada!`)}
const handleRemoveSale=async(id)=>{const{error}=await supabase.from('sales').delete().eq('id',id);if(error){showToast('Erro ao remover venda.','error');return}setSales(prev=>prev.filter(s=>s.id!==id))}
const moverCliente=async(clienteId,direcao,rota)=>{const rotaClients=clients.filter(c=>c.route===rota).sort((a,b)=>a.ordem-b.ordem);const idx=rotaClients.findIndex(c=>c.id===clienteId);if(direcao==='up'&&idx===0)return;if(direcao==='down'&&idx===rotaClients.length-1)return;const outro=direcao==='up'?rotaClients[idx-1]:rotaClients[idx+1];const atual=rotaClients[idx];await supabase.from('clients').update({ordem:outro.ordem}).eq('id',atual.id);await supabase.from('clients').update({ordem:atual.ordem}).eq('id',outro.id);await loadClients()}
const gerarPdf=async(order)=>{
  const doc=new jsPDF()
  const nomeArquivo=`Pedido_${order.client_name.replace(/[^a-zA-Z0-9]/g,'_')}_${order.date}.pdf`
  try{
    const toBase64=url=>fetch(url).then(r=>r.blob()).then(b=>new Promise((res,rej)=>{const reader=new FileReader();reader.onload=()=>res(reader.result);reader.onerror=rej;reader.readAsDataURL(b)}))
    const logoBase64=await toBase64('/logo_alimentos_4.png')
    doc.addImage(logoBase64,'PNG',10,10,60,20)
  }catch(e){}
  doc.setFontSize(9)
  doc.setTextColor(100)
  doc.text('CNPJ: 18.520.142/0001-45',200,12,{align:'right'})
  doc.text('Rua Anália Vieira de Souza, nº 38',200,17,{align:'right'})
  doc.text('Bairro São Vicente - Afonso Cláudio ES',200,22,{align:'right'})
  doc.text('Tel: 27 99852-2632',200,27,{align:'right'})
  doc.setDrawColor(217,119,6)
  doc.setLineWidth(0.5)
  doc.line(10,35,200,35)
  doc.setFontSize(16)
  doc.setTextColor(30,41,59)
  doc.setFont('helvetica','bold')
  doc.text('Pedido de Venda',10,44)
  doc.setFontSize(8)
  doc.setTextColor(100)
  doc.setFont('helvetica','bold')
  doc.text('CLIENTE',10,54)
  doc.text('DATA',105,54)
  doc.setFont('helvetica','normal')
  doc.setTextColor(30,41,59)
  doc.text(order.client_name,10,59)
  doc.text(new Date(order.date+'T12:00:00').toLocaleDateString('pt-BR'),105,59)
  doc.setFont('helvetica','bold')
  doc.setTextColor(100)
  doc.text('SITUAÇÃO',10,66)
  doc.text('FORMA DE PAGAMENTO',105,66)
  doc.setFont('helvetica','normal')
  doc.setTextColor(30,41,59)
  doc.text(order.situacao,10,71)
  doc.text({1:'Dinheiro',2:'Cheque',8:'Pix/Ted',16:'Boleto Sicoob',17:'Débito em Conta'}[order.forma_pgto]||'-',105,71)
  doc.setFont('helvetica','bold')
  doc.setTextColor(100)
  doc.text('VENCIMENTO',10,78)
  doc.text('ROTA',105,78)
  doc.setFont('helvetica','normal')
  doc.setTextColor(30,41,59)
  doc.text(order.vencimento?new Date(order.vencimento+'T12:00:00').toLocaleDateString('pt-BR'):'-',10,83)
  doc.text(order.route||'-',105,83)
  doc.setFillColor(30,41,59)
  doc.rect(10,90,190,8,'F')
  doc.setTextColor(255,255,255)
  doc.setFontSize(9)
  doc.setFont('helvetica','bold')
  doc.text('Produto',12,95.5)
  doc.text('Qtd',130,95.5)
  doc.text('Preço',148,95.5)
  doc.text('Desc%',168,95.5)
  doc.text('Total',185,95.5)
  let y=104
  doc.setFont('helvetica','normal')
  doc.setFontSize(9)
  order.produtos.forEach((p,i)=>{
    if(i%2===0)doc.setFillColor(248,250,252)
    else doc.setFillColor(255,255,255)
    doc.rect(10,y-5,190,8,'F')
    doc.setTextColor(30,41,59)
    doc.text(p.descricao.substring(0,45),12,y)
    doc.text(String(p.quant),130,y)
    doc.text(p.precoVenda.toLocaleString('pt-BR',{style:'currency',currency:'BRL'}),148,y)
    doc.text(`${p.vDesc||0}%`,168,y)
    doc.text((p.precoVenda*p.quant*(1-(p.vDesc||0)/100)).toLocaleString('pt-BR',{style:'currency',currency:'BRL'}),185,y)
    y+=8
  })
  doc.setDrawColor(226,232,240)
  doc.line(10,y,200,y)
  y+=8
  doc.setFont('helvetica','bold')
  doc.setFontSize(12)
  doc.setTextColor(37,99,235)
  doc.text(`Total: ${order.total.toLocaleString('pt-BR',{style:'currency',currency:'BRL'})}`,200,y,{align:'right'})
  doc.setFontSize(8)
  doc.setTextColor(148,163,184)
  doc.setFont('helvetica','normal')
  doc.text('Mageski Alimentos — 18.520.142/0001-45 — Rua Anália Vieira de Souza, nº 38, Bairro São Vicente, Afonso Cláudio ES — Tel: 27 99852-2632',105,285,{align:'center'})
  doc.save(nomeArquivo)
}const abrirModalProdutos=async()=>{
  setModalProdutos(true)
  setModalProdutoSearch('')
  if(todosProdutos.length===0){
    try{
      const{data}=await supabase.from('products').select('*').gt('preco_venda',0).or('tags.cs.{"PROPRIO"},tags.cs.{"TERCEIROS"}').order('descricao')
      setTodosProdutos(data||[])
    }catch(err){showToast('Erro ao carregar produtos','error')}
  }
}
const abrirPerfil=async(cliente)=>{setClientePerfil(cliente);setPerfilData(null);setPerfilLoading(true);try{const res=await fetch(`${EGESTOR_API}?action=perfil_cliente&codContato=${cliente.erp_code}`);const data=await res.json();setPerfilData(data)}catch(err){showToast('Erro ao buscar perfil','error')}setPerfilLoading(false)}
const salvarPedido=async()=>{if(!pedidoCliente||pedidoProdutos.length===0){showToast('Preencha cliente e produtos.','error');return}if(['Pedido S/ NFe','Pedido C/ NFe'].includes(pedidoSituacao)&&!pedidoFormaPgto){showToast('Selecione a forma de pagamento.','error');return}setPedidoLoading(true);try{const total=pedidoProdutos.reduce((acc,p)=>{const sub=p.precoVenda*p.quant;const desc=sub*(p.vDesc||0)/100;return acc+sub-desc},0);const{error}=await supabase.from('orders').insert({user_id:user.id,client_id:pedidoCliente.id||null,client_name:pedidoCliente.name,client_erp_code:pedidoCliente.erp_code,route:pedidoCliente.route||'',situacao:pedidoSituacao,forma_pgto:pedidoFormaPgto?parseInt(pedidoFormaPgto):null,vencimento:pedidoVencimento,produtos:pedidoProdutos,total,date:today()});if(error){showToast('Erro ao salvar pedido','error')}else{showToast('Pedido salvo!');setPedidoCliente(null);setPedidoProdutos([]);setPedidoFormaPgto('1');setPedidoSituacao('Pedido S/ NFe');setPedidoVencimento(today());await loadOrders()}}catch(err){showToast('Erro ao salvar pedido','error')}setPedidoLoading(false)}
const excluirPedido=async(id)=>{const{error}=await supabase.from('orders').delete().eq('id',id);if(error){showToast('Erro ao excluir pedido','error');return}setOrders(prev=>prev.filter(o=>o.id!==id));showToast('Pedido excluído')}
const exportarPedidos=async()=>{const pendentes=orders.filter(o=>o.status==='pendente');if(pendentes.length===0){showToast('Nenhum pedido pendente','error');return}setExportLoading(true);let ok=0;let erros=0;for(const order of pendentes){try{const res=await fetch(`${EGESTOR_API}?action=criar_venda`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({codContato:order.client_erp_code,nomeContato:order.client_name,route:order.route,user_id:user.id,produtos:order.produtos.map(p=>({codigo:p.codigo,quant:p.quant,preco:p.precoVenda,vDesc:p.vDesc||0})),codFormaPgto:order.forma_pgto,vencimento:order.vencimento||today(),situacaoOS:order.situacao})});const result=await res.json();if(result.codigo){const{data:saleData}=await supabase.from('sales').select('id').eq('erp_code',result.codigo).single();if(saleData?.id){const items=order.produtos.map(p=>({sale_id:saleData.id,user_id:user.id,client_erp_code:order.client_erp_code,erp_code:p.codigo,descricao:p.descricao,codigo_proprio:p.codigoProprio||'',quant:p.quant,preco:p.precoVenda,vdesc:p.vDesc||0,total:p.precoVenda*p.quant*(1-(p.vDesc||0)/100),date:today()}));await supabase.from('sales_items').insert(items)}await supabase.from('orders').delete().eq('id',order.id);ok++}else{erros++}}catch(err){erros++}}await loadOrders();await loadSales();setExportLoading(false);if(erros===0){showToast(`${ok} pedido(s) exportado(s)!`)}else{showToast(`${ok} exportado(s), ${erros} com erro`,'error')}}
const abrirEdicaoOrder=async(order)=>{setEditandoOrder(order);setEditOrderProdutos(order.produtos||[]);setEditOrderFormaPgto(String(order.forma_pgto||'1'));setEditOrderSituacao(order.situacao||'Pedido S/ NFe')}
const salvarEdicaoOrder=async()=>{if(!editandoOrder||editOrderProdutos.length===0){showToast('Adicione ao menos um produto','error');return}const total=editOrderProdutos.reduce((acc,p)=>acc+p.precoVenda*p.quant*(1-(p.vDesc||0)/100),0);const{error}=await supabase.from('orders').update({produtos:editOrderProdutos,forma_pgto:parseInt(editOrderFormaPgto),situacao:editOrderSituacao,total}).eq('id',editandoOrder.id);if(error){showToast('Erro ao salvar','error');return}showToast('Pedido atualizado!');setEditandoOrder(null);await loadOrders()}
const buscarProdutos=async(search,setResultados)=>{if(search.length<2){setResultados([]);return}try{const res=await fetch(`${EGESTOR_API}?action=produtos&search=${encodeURIComponent(search)}`);const data=await res.json();const filtrado=(Array.isArray(data)?data:[]).filter(p=>p.descricao?.toLowerCase().includes(search.toLowerCase())||p.codigoProprio?.toLowerCase().includes(search.toLowerCase()));setResultados(filtrado)}catch(err){showToast('Erro ao buscar produtos','error')}}
const addProduto=(produto,setProdutos,setSearch,setResultados)=>{setProdutoModal({produto,setProdutos,setSearch,setResultados});setProdutoModalQuant(1);setProdutoModalDesc(0);setSearch('');setResultados([])}
const confirmarProdutoModal=()=>{if(!produtoModal)return;const{produto,setProdutos}=produtoModal;setProdutos(prev=>{const existe=prev.find(p=>p.codigo===produto.codigo);if(existe)return prev.map(p=>p.codigo===produto.codigo?{...p,quant:p.quant+produtoModalQuant,vDesc:produtoModalDesc}:p);return[...prev,{...produto,quant:produtoModalQuant,vDesc:produtoModalDesc}]});setProdutoModal(null)}
const carregarConfig=async()=>{setConfigLoading(true);try{const usersRes=await fetch(`${ADMIN_API}?action=listar_usuarios`);const users=await usersRes.json();setConfigUsuarios(Array.isArray(users)?users:[]);setConfigVendedores([{codigo:1,nome:'Vicon Soluções Empresariais'},{codigo:3,nome:'JULIANO RODRIGO MAGESKI'},{codigo:4,nome:'LUCAS DE PAULO OLIVEIRA'},{codigo:5,nome:'ELIESIMO ADRIANO PEREIRA'},{codigo:7,nome:'OTAVIO DOS SANTOS RIBEIRO'}])}catch(err){showToast('Erro ao carregar configurações','error')}setConfigLoading(false)}
const criarUsuario=async()=>{if(!novoEmail||!novaSenha||!novoNome||!novoEgestorCode){showToast('Preencha todos os campos','error');return}setConfigLoading(true);try{const res=await fetch(`${ADMIN_API}?action=criar_usuario`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:novoEmail,password:novaSenha,name:novoNome,egestor_code:parseInt(novoEgestorCode)})});const data=await res.json();if(data.error){showToast('Erro: '+data.error,'error')}else{showToast('Usuário criado!');setNovoEmail('');setNovaSenha('');setNovoNome('');setNovoEgestorCode('');await carregarConfig()}}catch(err){showToast('Erro ao criar usuário','error')}setConfigLoading(false)}
const deletarUsuario=async(user_id)=>{if(!window.confirm('Confirma exclusão do usuário?'))return;setConfigLoading(true);try{const res=await fetch(`${ADMIN_API}?action=deletar_usuario`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({user_id})});const data=await res.json();if(data.error){showToast('Erro: '+data.error,'error')}else{showToast('Usuário excluído!');await carregarConfig()}}catch(err){showToast('Erro ao excluir','error')}setConfigLoading(false)}
const atualizarSenha=async(user_id)=>{const nova=senhaUser[user_id];if(!nova||nova.length<6){showToast('Senha deve ter ao menos 6 caracteres','error');return}setConfigLoading(true);try{const res=await fetch(`${ADMIN_API}?action=atualizar_senha`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({user_id,password:nova})});const data=await res.json();if(data.error){showToast('Erro: '+data.error,'error')}else{showToast('Senha atualizada!');setSenhaUser(prev=>({...prev,[user_id]:''}))} }catch(err){showToast('Erro ao atualizar senha','error')}setConfigLoading(false)}
const totalPedido=pedidoProdutos.reduce((acc,p)=>acc+p.precoVenda*p.quant*(1-(p.vDesc||0)/100),0)
const routeClients=useMemo(()=>selectedRoute?clients.filter(c=>c.route===selectedRoute):[],[clients,selectedRoute])
const routeSales=useMemo(()=>sales.filter(s=>s.route===selectedRoute),[sales,selectedRoute])
const soldClientIds=useMemo(()=>new Set(routeSales.map(s=>s.client_id).filter(Boolean)),[routeSales])
const inactiveSoldClients=useMemo(()=>routeSales.filter(s=>clients.find(c=>c.id===s.client_id)?.inactive),[routeSales,clients])
const activeRouteClients=useMemo(()=>routeClients.filter(c=>!c.inactive),[routeClients])
const activeSoldIds=useMemo(()=>new Set(routeSales.filter(s=>{const c=clients.find(cl=>cl.id===s.client_id);return!c||!c.inactive}).map(s=>s.client_id).filter(Boolean)),[routeSales,clients])
const totalSold=useMemo(()=>routeSales.filter(s=>!['Bonificação','Troca'].includes(s.note)).reduce((a,s)=>a+s.value,0),[routeSales])
const remaining=activeRouteClients.length-activeSoldIds.size
const avgTicket=activeSoldIds.size>0?totalSold/activeSoldIds.size:0
const goalProgress=dailyGoal?Math.min((totalSold/dailyGoal)*100,100):0
const clientSuggestions=useMemo(()=>{const pool=selectedRoute?routeClients:clients;if(!tabSaleClientInput.trim())return pool.slice(0,6);return pool.filter(c=>c.name.toLowerCase().includes(tabSaleClientInput.toLowerCase())).slice(0,6)},[clients,routeClients,selectedRoute,tabSaleClientInput])
const filteredClients=useMemo(()=>(selectedRoute?routeClients:clients).filter(c=>c.name.toLowerCase().includes(clientSearch.toLowerCase())),[routeClients,clients,selectedRoute,clientSearch])
const Tab=({id,label,icon,badge})=><button onClick={()=>setActiveTab(id)} style={{background:activeTab===id?ACCENT:'transparent',color:activeTab===id?'#fff':MUTED,border:'none',borderRadius:8,padding:'6px 10px',fontWeight:600,fontSize:11,cursor:'pointer',display:'flex',flexDirection:'column',alignItems:'center',gap:2,minWidth:56,position:'relative'}}>
{badge>0&&<span style={{position:'absolute',top:2,right:8,background:DANGER,color:'#fff',borderRadius:99,fontSize:9,fontWeight:700,padding:'1px 5px'}}>{badge}</span>}
<span style={{fontSize:18}}>{icon}</span><span>{label}</span></button>
const FormaPgtoLabel={1:'Dinheiro',2:'Cheque',8:'Pix/Ted',16:'Boleto Sicoob',17:'Déb. Conta'}
const ProdutoCard=({p,onChange,onRemove})=><div style={{background:SURFACE,border:`1px solid ${BORDER}`,borderRadius:8,padding:'8px 12px',marginBottom:6}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
<span style={{fontWeight:600,fontSize:12,flex:1}}>{p.descricao}</span>
<button onClick={onRemove} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:14}}>✕</button>
</div>
<div style={{display:'flex',gap:6}}>
<div style={{flex:1}}><div style={{fontSize:10,color:MUTED,marginBottom:2}}>QTD</div><input type="number" min="1" value={p.quant} onChange={e=>onChange({...p,quant:parseFloat(e.target.value)||1})} onFocus={e=>e.target.select()} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:6,padding:'6px 8px',fontSize:13,boxSizing:'border-box'}}/></div>
<div style={{flex:1}}><div style={{fontSize:10,color:MUTED,marginBottom:2}}>PREÇO</div><input type="number" min="0" step="0.01" value={p.precoVenda} onChange={e=>onChange({...p,precoVenda:parseFloat(e.target.value)||0})} onFocus={e=>e.target.select()} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:6,padding:'6px 8px',fontSize:13,boxSizing:'border-box'}}/></div>
<div style={{flex:1}}><div style={{fontSize:10,color:MUTED,marginBottom:2}}>DESC%</div><input type="number" min="0" max="100" value={p.vDesc||0} onChange={e=>onChange({...p,vDesc:parseFloat(e.target.value)||0})} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:6,padding:'6px 8px',fontSize:13,boxSizing:'border-box'}}/></div>
<div style={{flex:1,textAlign:'right'}}><div style={{fontSize:10,color:MUTED,marginBottom:2}}>TOTAL</div><div style={{fontWeight:700,color:SUCCESS,fontSize:12}}>{fmt(p.precoVenda*p.quant*(1-(p.vDesc||0)/100))}</div></div>
</div>
</div>
return(<div style={{minHeight:'100vh',background:SURFACE,fontFamily:"'Inter',system-ui,sans-serif",color:TEXT,paddingBottom:72}}>
<div style={{background:'#1e293b',borderBottom:`1px solid ${BORDER}`,padding:'0 16px',position:'sticky',top:0,zIndex:100}}>
<div style={{display:'flex',alignItems:'center',gap:8,height:52}}>
<svg viewBox="0 0 200 40" xmlns="http://www.w3.org/2000/svg" width="160" height="32">
<circle cx="20" cy="20" r="18" fill="white"/>
<circle cx="20" cy="20" r="15" fill="#2563EB"/>
<polyline points="11,27 16,19 21,23 29,13" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
<polygon points="26,10 32,12 30,18" fill="white"/>
<text x="44" y="27" fontFamily="Arial, sans-serif" fontSize="21" fontWeight="300" fill="white">e</text>
<text x="56" y="27" fontFamily="Arial, sans-serif" fontSize="21" fontWeight="900" fill="white" stroke="#2563EB" strokeWidth="0.5">X</text>
<text x="70" y="27" fontFamily="Arial, sans-serif" fontSize="21" fontWeight="300" fill="white">pande</text>
</svg>
<button onClick={()=>supabase.auth.signOut()} style={{background:'none',border:`1px solid #ffffff55`,borderRadius:8,padding:'8px 16px',fontSize:13,color:'#e2e8f0',cursor:'pointer',fontWeight:600,marginLeft:'auto'}}>Sair</button></div>
</div>
<div style={{position:'fixed',bottom:0,left:0,right:0,background:CARD,borderTop:`1px solid ${BORDER}`,display:'flex',justifyContent:'space-around',padding:'6px 0',zIndex:100}}>
<Tab id="dashboard" label="Dashboard" icon="📊"/>
<Tab id="clientes" label="Clientes" icon="👥"/>
<Tab id="vendas" label="Vendas" icon="💰"/>
<Tab id="pedido" label="Pedido" icon="🛒" badge={orders.length}/>
{user?.id===ADMIN_ID&&<Tab id="config" label="Config" icon="⚙️"/>}
</div>
<div style={{padding:'12px 16px'}}>
{clientePerfil&&<div style={{position:'fixed',inset:0,background:'#0008',zIndex:500,display:'flex',alignItems:'flex-end'}}>
<div style={{background:CARD,borderRadius:'16px 16px 0 0',padding:20,width:'100%',maxHeight:'90vh',overflowY:'auto'}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
<div style={{fontWeight:800,fontSize:16}}>{clientePerfil.name}</div>
<button onClick={()=>setClientePerfil(null)} style={{background:'none',border:'none',fontSize:20,cursor:'pointer',color:MUTED}}>✕</button>
</div>
<div style={{marginBottom:14}}>
<Badge color={clientePerfil.inactive?WARNING:SUCCESS}>{clientePerfil.inactive?'⛔ Inativo':'✓ Ativo'}</Badge>
<span style={{marginLeft:6}}><Badge color={ACCENT}>{clientePerfil.route}</Badge></span>
</div>
{perfilLoading?<div style={{textAlign:'center',padding:'30px 0',color:MUTED}}>⏳ Carregando histórico...</div>
:perfilData?<>
<div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:14}}>
<div style={{background:ACCENT_LIGHT,borderRadius:10,padding:'10px 12px',textAlign:'center'}}>
<div style={{fontSize:22,fontWeight:800,color:ACCENT}}>{perfilData.mixAtual}/{perfilData.mixTotal}</div>
<div style={{fontSize:11,color:MUTED,fontWeight:600}}>MIX DE PRODUTOS</div>
</div>
<div style={{background:'#F0FDF4',borderRadius:10,padding:'10px 12px',textAlign:'center'}}>
<div style={{fontSize:22,fontWeight:800,color:SUCCESS}}>{perfilData.numPedidos}</div>
<div style={{fontSize:11,color:MUTED,fontWeight:600}}>PEDIDOS NO CRM</div>
</div>
<div style={{background:SURFACE,borderRadius:10,padding:'10px 12px',textAlign:'center'}}>
<div style={{fontSize:16,fontWeight:800,color:TEXT}}>{fmt(perfilData.totalGeral)}</div>
<div style={{fontSize:11,color:MUTED,fontWeight:600}}>TOTAL COMPRADO</div>
</div>
<div style={{background:SURFACE,borderRadius:10,padding:'10px 12px',textAlign:'center'}}>
<div style={{fontSize:16,fontWeight:800,color:TEXT}}>{perfilData.numPedidos>0?fmt(perfilData.totalGeral/perfilData.numPedidos):fmt(0)}</div>
<div style={{fontSize:11,color:MUTED,fontWeight:600}}>TICKET MÉDIO</div>
</div>
</div>
{perfilData.comprados.length>0&&<>
<div style={{fontWeight:700,fontSize:13,marginBottom:8,color:SUCCESS}}>✅ Produtos que compra ({perfilData.comprados.length})</div>
{perfilData.comprados.map(p=><div key={p.erp_code} style={{background:'#F0FDF4',border:`1px solid ${SUCCESS}22`,borderRadius:8,padding:'8px 12px',marginBottom:6,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
<div><div style={{fontWeight:600,fontSize:12}}>{p.descricao}</div><div style={{fontSize:10,color:MUTED}}>Qtd total: {p.quant_total} • Última: {p.ultima_compra}</div></div>
<div style={{textAlign:'right'}}><div style={{fontWeight:700,color:SUCCESS,fontSize:12}}>{fmt(p.total_comprado)}</div></div>
</div>)}
</>}
{perfilData.naoComprados.length>0&&<>
<div style={{fontWeight:700,fontSize:13,marginBottom:8,marginTop:12,color:WARNING}}>💡 Ainda não compra ({perfilData.naoComprados.length})</div>
{perfilData.naoComprados.map(p=><div key={p.erp_code} style={{background:'#FFFBEB',border:`1px solid ${WARNING}22`,borderRadius:8,padding:'8px 12px',marginBottom:6}}>
<div style={{fontWeight:600,fontSize:12}}>{p.descricao}</div>
{p.codigo_proprio&&<div style={{fontSize:10,color:MUTED}}>Cód: {p.codigo_proprio}</div>}
</div>)}
</>}
{perfilData.comprados.length===0&&perfilData.naoComprados.length===0&&<div style={{textAlign:'center',padding:'20px 0',color:MUTED}}>Nenhum histórico ainda.<br/>Os dados aparecem após exportar pedidos.</div>}
</>:null}
<div style={{display:'flex',gap:8,marginTop:16}}>
<button onClick={()=>setClientePerfil(null)} style={{flex:1,background:SURFACE,color:MUTED,border:`1px solid ${BORDER}`,borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>Fechar</button>
<button onClick={()=>{setPedidoCliente(clientePerfil);setClientePerfil(null);setActiveTab('pedido')}} style={{flex:2,background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>🛒 Fazer Pedido</button>
</div>
</div>
</div>}
{modalProdutos&&<div style={{position:'fixed',inset:0,background:'#0008',zIndex:600,display:'flex',alignItems:'flex-end'}}>
<div style={{background:CARD,borderRadius:'16px 16px 0 0',padding:20,width:'100%',height:'90vh',display:'flex',flexDirection:'column'}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
<div style={{fontWeight:800,fontSize:16}}>📦 Produtos</div>
<button onClick={()=>setModalProdutos(false)} style={{background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'8px 20px',fontWeight:700,fontSize:14,cursor:'pointer'}}>Salvar</button>
</div>
<input type="text" placeholder="🔍 Buscar produto…" value={modalProdutoSearch} onChange={e=>setModalProdutoSearch(e.target.value)}
style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,boxSizing:'border-box',marginBottom:12}}/>
<div style={{overflowY:'auto',flex:1}}>
{todosProdutos.filter(p=>p.descricao?.toLowerCase().includes(modalProdutoSearch.toLowerCase())||p.codigo_proprio?.toLowerCase().includes(modalProdutoSearch.toLowerCase())).map(p=>{
const prod={codigo:p.erp_code,descricao:p.descricao,codigoProprio:p.codigo_proprio,precoVenda:p.preco_venda}
const noLista=pedidoProdutos.find(x=>x.codigo===prod.codigo)
return<div key={prod.codigo} style={{background:noLista?ACCENT_LIGHT:SURFACE,border:`1px solid ${noLista?ACCENT:BORDER}`,borderRadius:10,padding:'10px 12px',marginBottom:8,display:'flex',alignItems:'center',gap:8}}>
<div style={{flex:1}}>
<div style={{fontWeight:600,fontSize:13,color:noLista?ACCENT:TEXT}}>{prod.descricao}</div>
<div style={{fontSize:11,color:MUTED}}>{fmt(prod.precoVenda||0)}</div>
</div>
<div style={{display:'flex',alignItems:'center',gap:6}}>
<button onMouseDown={()=>{if(noLista&&noLista.quant<=1){setPedidoProdutos(prev=>prev.filter(x=>x.codigo!==prod.codigo))}else if(noLista){setPedidoProdutos(prev=>prev.map(x=>x.codigo===prod.codigo?{...x,quant:x.quant-1}:x))}}} style={{width:32,height:32,background:noLista?DANGER:'#eee',color:noLista?'#fff':MUTED,border:'none',borderRadius:6,fontSize:18,fontWeight:700,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center'}}>−</button>
<input type="number" min="0" value={noLista?.quant||0} onChange={e=>{const q=parseFloat(e.target.value)||0;if(q===0){setPedidoProdutos(prev=>prev.filter(x=>x.codigo!==prod.codigo))}else if(noLista){setPedidoProdutos(prev=>prev.map(x=>x.codigo===prod.codigo?{...x,quant:q}:x))}else{setPedidoProdutos(prev=>[...prev,{...prod,quant:q,vDesc:0}])}}} onFocus={e=>e.target.select()} style={{width:44,textAlign:'center',border:`1px solid ${BORDER}`,borderRadius:6,padding:'6px 4px',fontSize:14,fontWeight:700}}/>
<button onMouseDown={()=>{if(noLista){setPedidoProdutos(prev=>prev.map(x=>x.codigo===prod.codigo?{...x,quant:x.quant+1}:x))}else{setPedidoProdutos(prev=>[...prev,{...prod,quant:1,vDesc:0}])}}} style={{width:32,height:32,background:ACCENT,color:'#fff',border:'none',borderRadius:6,fontSize:18,fontWeight:700,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center'}}>+</button>
</div>
</div>})}
</div>
</div>
</div>}
{produtoModal&&<div style={{position:'fixed',inset:0,background:'#0008',zIndex:600,display:'flex',alignItems:'center',justifyContent:'center',padding:16}}>
<div style={{background:CARD,borderRadius:16,padding:20,width:'100%',maxWidth:340}}>
<div style={{fontWeight:700,fontSize:15,marginBottom:4}}>{produtoModal.produto.descricao}</div>
<div style={{fontSize:12,color:MUTED,marginBottom:16}}>{fmt(produtoModal.produto.precoVenda)} por unidade</div>
<div style={{marginBottom:12}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>QUANTIDADE</label>
<input type="number" min="1" value={produtoModalQuant} autoFocus onChange={e=>setProdutoModalQuant(parseFloat(e.target.value)||1)} onFocus={e=>e.target.select()} onKeyDown={e=>e.key==='Enter'&&confirmarProdutoModal()} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:18,boxSizing:'border-box',textAlign:'center',fontWeight:700}}/>
</div>
<div style={{marginBottom:16}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>DESCONTO %</label>
<input type="number" min="0" max="100" value={produtoModalDesc} onChange={e=>setProdutoModalDesc(parseFloat(e.target.value)||0)} onFocus={e=>e.target.select()} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:16,boxSizing:'border-box',textAlign:'center'}}/>
</div>
<div style={{background:ACCENT_LIGHT,borderRadius:8,padding:'10px 12px',textAlign:'center',marginBottom:16}}>
<div style={{fontSize:11,color:MUTED,marginBottom:2}}>TOTAL</div>
<div style={{fontWeight:800,fontSize:20,color:ACCENT}}>{fmt(produtoModal.produto.precoVenda*produtoModalQuant*(1-produtoModalDesc/100))}</div>
</div>
<div style={{display:'flex',gap:8}}>
<button onClick={()=>setProdutoModal(null)} style={{flex:1,background:SURFACE,color:MUTED,border:`1px solid ${BORDER}`,borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>Cancelar</button>
<button onClick={confirmarProdutoModal} style={{flex:2,background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>+ Adicionar</button>
</div>
</div>
</div>}
{toast&&<div style={{position:'fixed',top:60,left:16,right:16,zIndex:1000,background:toast.type==='error'?DANGER:SUCCESS,color:'#fff',borderRadius:10,padding:'12px 16px',fontWeight:600,fontSize:13,boxShadow:'0 4px 20px #0003',textAlign:'center'}}>{toast.type==='error'?'❌':'✅'} {toast.msg}</div>}
<div style={{display:'flex',gap:8,marginBottom:12,flexWrap:'wrap'}}>
{routes.length>0&&<div style={{flex:1,minWidth:150,background:CARD,border:`1px solid ${BORDER}`,borderRadius:10,padding:'8px 12px'}}>
<div style={{fontWeight:700,fontSize:11,marginBottom:3,color:MUTED}}>ROTA DO DIA</div>
<select value={selectedRoute} onChange={e=>{setSelectedRoute(e.target.value);setDailyGoal('')}} style={{width:'100%',border:'none',background:'transparent',fontWeight:700,fontSize:13,color:TEXT,outline:'none'}}>
<option value="">Selecionar…</option>
{routes.map(r=><option key={r} value={r}>{r}</option>)}
</select>
</div>}
{selectedRoute&&<div style={{flex:'0 0 auto',background:CARD,border:`1px solid ${BORDER}`,borderRadius:10,padding:'8px 12px'}}>
<div style={{fontWeight:700,fontSize:11,marginBottom:3,color:MUTED}}>META DO DIA</div>
{dailyGoal?<div style={{display:'flex',alignItems:'center',gap:6}}><span style={{fontWeight:800,color:ACCENT,fontSize:14}}>{fmt(dailyGoal)}</span><button onClick={()=>setDailyGoal('')} style={{background:'none',border:'none',color:MUTED,cursor:'pointer',fontSize:12}}>✏️</button></div>
:<div style={{display:'flex',gap:4}}><input type="number" placeholder="0,00" value={goalInput} onChange={e=>setGoalInput(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleSetGoal()} style={{width:80,border:`1px solid ${BORDER}`,borderRadius:6,padding:'4px 6px',fontSize:12}}/><button onClick={handleSetGoal} style={{background:ACCENT,color:'#fff',border:'none',borderRadius:6,padding:'4px 8px',fontWeight:700,cursor:'pointer',fontSize:11}}>OK</button></div>}
</div>}
</div>
{activeTab==='dashboard'&&<div>
{!selectedRoute?<div style={{textAlign:'center',padding:'60px 20px',color:MUTED}}><div style={{fontSize:48,marginBottom:12}}>🗺️</div><div style={{fontWeight:700,fontSize:16,color:TEXT}}>Selecione uma rota para começar</div></div>
:<>
<div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:12}}>
<KpiCard label="Na Rota" value={activeRouteClients.length} sub="clientes ativos" color={ACCENT}/>
<KpiCard label="Atendidos" value={activeSoldIds.size} sub={`${activeRouteClients.length>0?Math.round((activeSoldIds.size/activeRouteClients.length)*100):0}%`} color={SUCCESS}/>
<KpiCard label="Restantes" value={remaining} sub={remaining===0?'Concluído! 🎉':'a visitar'} color={remaining===0?SUCCESS:WARNING}/>
<KpiCard label="Ticket Médio" value={fmt(avgTicket)} sub={`${activeSoldIds.size} venda(s)`} color={ACCENT}/>
</div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px',marginBottom:12}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
<span style={{fontWeight:700,fontSize:13}}>💰 Total Vendido</span>
<span style={{fontWeight:800,fontSize:18,color:totalSold>=(dailyGoal||Infinity)?SUCCESS:ACCENT}}>{fmt(totalSold)}</span>
</div>
{dailyGoal>0&&<><div style={{display:'flex',justifyContent:'space-between',fontSize:11,color:MUTED,marginBottom:6}}><span>Meta: {fmt(dailyGoal)}</span><span>{goalProgress.toFixed(1)}%</span></div>
<div style={{background:SURFACE,borderRadius:99,height:8,overflow:'hidden'}}><div style={{width:`${goalProgress}%`,height:'100%',background:goalProgress>=100?SUCCESS:ACCENT,borderRadius:99,transition:'width 0.4s'}}/></div>
<div style={{textAlign:'right',marginTop:4,fontSize:11,color:MUTED}}>Faltam {fmt(Math.max(dailyGoal-totalSold,0))}</div></>}
</div>
{inactiveSoldClients.length>0&&<div style={{background:'#FFF7ED',border:`1.5px solid ${WARNING}55`,borderRadius:12,padding:'12px 14px',marginBottom:12}}>
<div style={{display:'flex',alignItems:'center',gap:6,marginBottom:8}}>
<span>⚠️</span><span style={{fontWeight:700,fontSize:13,color:WARNING}}>Inativos Atendidos</span>
<Badge color={WARNING}>{inactiveSoldClients.length}</Badge>
<span style={{marginLeft:'auto',fontWeight:800,color:WARNING,fontSize:13}}>{fmt(inactiveSoldClients.reduce((a,s)=>a+s.value,0))}</span>
</div>
{inactiveSoldClients.map(s=><div key={s.id} style={{display:'flex',justifyContent:'space-between',fontSize:12,padding:'4px 0',borderTop:`1px solid ${WARNING}22`}}>
<span style={{fontWeight:600}}>{s.client_name}</span><Badge color={WARNING}>{fmt(s.value)}</Badge>
</div>)}
</div>}
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,overflow:'hidden',marginBottom:12}}>
<div style={{padding:'12px 14px',borderBottom:`1px solid ${BORDER}`,fontWeight:700,fontSize:13,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
<span>🗺️ Mapa da Rota</span>
<Badge color={ACCENT}>{routeClients.length} clientes</Badge>
</div>
{routeClients.length===0?<div style={{textAlign:'center',padding:'20px',color:MUTED}}>Nenhum cliente nesta rota</div>
:routeClients.map((c,i)=>{
const vendido=soldClientIds.has(c.id)||orders.some(o=>o.client_id===c.id||o.client_erp_code===c.erp_code)
const temPedido=orders.some(o=>o.client_id===c.id||o.client_erp_code===c.erp_code)
return<div key={c.id} onClick={()=>abrirPerfil(c)} style={{padding:'10px 14px',borderBottom:`1px solid ${BORDER}`,display:'flex',alignItems:'center',gap:10,cursor:'pointer',background:vendido?'#F0FDF4':c.inactive?'#FFF7ED':'transparent'}}>
<div style={{width:28,height:28,borderRadius:99,background:temPedido?WARNING:vendido?SUCCESS:c.inactive?'#FEE2E2':BORDER,display:'flex',alignItems:'center',justifyContent:'center',fontSize:13,flexShrink:0,color:'#fff',fontWeight:700}}>
{temPedido?'📋':vendido?'✓':c.inactive?'⛔':i+1}
</div>
<div style={{flex:1}}>
<div style={{fontWeight:600,fontSize:13}}>{c.name}</div>
<div style={{fontSize:11,color:MUTED}}>{temPedido?'Pedido pendente':vendido?'Atendido':'Aguardando visita'}</div>
</div>
<span style={{color:MUTED,fontSize:16}}>›</span>
</div>})}
</div>
{routeSales.length>0&&<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,overflow:'hidden'}}>
<div style={{padding:'12px 14px',borderBottom:`1px solid ${BORDER}`,fontWeight:700,fontSize:13}}>Vendas de Hoje</div>
{[...routeSales].reverse().map((s,i)=><div key={s.id} style={{padding:'10px 14px',borderBottom:`1px solid ${BORDER}`,background:i%2===0?CARD:SURFACE,display:'flex',alignItems:'center',gap:8}}>
<div style={{flex:1}}><div style={{fontWeight:600,fontSize:13}}>{s.client_name}</div><div style={{fontSize:11,color:MUTED}}>{s.sale_time}{s.note?' • '+s.note:''}</div></div>
<Badge color={SUCCESS}>{fmt(s.value)}</Badge>
<button onClick={()=>handleRemoveSale(s.id)} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:16}}>✕</button>
</div>)}
</div>}
</>}
</div>}
{activeTab==='clientes'&&<div>
{clients.length===0?<div style={{textAlign:'center',padding:'60px 20px',color:MUTED}}><div style={{fontSize:48,marginBottom:12}}>📂</div><div style={{fontWeight:700,fontSize:16,color:TEXT}}>Nenhuma planilha importada</div></div>
:<><div style={{display:'flex',gap:8,marginBottom:12,alignItems:'center'}}>
<input type="text" placeholder="🔍 Buscar cliente…" value={clientSearch} onChange={e=>setClientSearch(e.target.value)} style={{flex:1,border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14}}/>
<Badge color={ACCENT}>{filteredClients.length}</Badge>
</div>
{filteredClients.map((c,i)=><div key={c.id} style={{background:c.inactive?'#FFF7ED':CARD,border:`1px solid ${BORDER}`,borderRadius:10,padding:'10px 14px',marginBottom:6,display:'flex',alignItems:'center',gap:8}}>
<div style={{display:'flex',flexDirection:'column',gap:2,marginRight:4}}>
<button onClick={()=>moverCliente(c.id,'up',c.route)} style={{background:'none',border:`1px solid ${BORDER}`,borderRadius:4,padding:'2px 6px',fontSize:10,cursor:'pointer',color:MUTED,lineHeight:1}}>▲</button>
<button onClick={()=>moverCliente(c.id,'down',c.route)} style={{background:'none',border:`1px solid ${BORDER}`,borderRadius:4,padding:'2px 6px',fontSize:10,cursor:'pointer',color:MUTED,lineHeight:1}}>▼</button>
</div>
<div style={{flex:1,cursor:'pointer'}} onClick={()=>abrirPerfil(c)}>
<div style={{fontWeight:600,fontSize:13}}>{c.name}</div>
<div style={{fontSize:11,color:MUTED,marginTop:2}}><Badge color={ACCENT}>{c.route}</Badge></div>
</div>
<div style={{display:'flex',flexDirection:'column',alignItems:'flex-end',gap:4}}>
{c.inactive?<Badge color={WARNING}>⛔ Inativo</Badge>:<Badge color={SUCCESS}>✓ Ativo</Badge>}
{soldClientIds.has(c.id)?<Badge color={SUCCESS}>✅ Vendido</Badge>:selectedRoute===c.route?<Badge color={MUTED}>⏳ Pendente</Badge>:null}
</div>
</div>)}
</>}
</div>}
{activeTab==='vendas'&&<div>
{(()=>{
const totalExportado=sales.filter(s=>!['Bonificação','Troca'].includes(s.note)).reduce((a,s)=>a+s.value,0)
const totalPendente=orders.filter(o=>!['Bonificação','Troca'].includes(o.situacao)).reduce((a,o)=>a+o.total,0)
const totalGeral=totalExportado+totalPendente
const progresso=dailyGoal>0?Math.min((totalGeral/dailyGoal)*100,100):0
return<>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px',marginBottom:12}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
<span style={{fontWeight:700,fontSize:14}}>💰 Total do Dia</span>
<span style={{fontWeight:800,fontSize:18,color:ACCENT}}>{fmt(totalGeral)}</span>
</div>
{dailyGoal>0&&<>
<div style={{display:'flex',justifyContent:'space-between',fontSize:11,color:MUTED,marginBottom:6}}>
<span>Meta: {fmt(dailyGoal)}</span><span>{progresso.toFixed(1)}%</span>
</div>
<div style={{background:SURFACE,borderRadius:99,height:8,overflow:'hidden'}}>
<div style={{width:`${progresso}%`,height:'100%',background:progresso>=100?SUCCESS:ACCENT,borderRadius:99,transition:'width 0.4s'}}/>
</div>
<div style={{display:'flex',justifyContent:'space-between',marginTop:6,fontSize:11,color:MUTED}}>
<span>✅ Exportado: {fmt(totalExportado)}</span>
<span>📋 Pendente: {fmt(totalPendente)}</span>
</div>
</>}
<div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,marginTop:12}}>
<div style={{background:SURFACE,borderRadius:8,padding:'8px',textAlign:'center'}}>
<div style={{fontWeight:800,fontSize:16,color:SUCCESS}}>{sales.length}</div>
<div style={{fontSize:10,color:MUTED,fontWeight:600}}>EXPORTADOS</div>
</div>
<div style={{background:SURFACE,borderRadius:8,padding:'8px',textAlign:'center'}}>
<div style={{fontWeight:800,fontSize:16,color:WARNING}}>{orders.length}</div>
<div style={{fontSize:10,color:MUTED,fontWeight:600}}>PENDENTES</div>
</div>
<div style={{background:SURFACE,borderRadius:8,padding:'8px',textAlign:'center'}}>
<div style={{fontWeight:800,fontSize:16,color:ACCENT}}>{sales.length+orders.length}</div>
<div style={{fontSize:10,color:MUTED,fontWeight:600}}>TOTAL</div>
</div>
</div>
</div>
{sales.length>0&&<>
<div style={{fontWeight:700,fontSize:13,marginBottom:8,color:SUCCESS}}>✅ Exportados para eGestor ({sales.length})</div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,overflow:'hidden',marginBottom:12}}>
{[...sales].reverse().map((s,i)=><div key={s.id} style={{padding:'10px 14px',borderBottom:`1px solid ${BORDER}`,background:i%2===0?CARD:SURFACE,display:'flex',alignItems:'center',gap:8}}>
<div style={{flex:1}}>
<div style={{fontWeight:600,fontSize:13}}>{s.client_name}</div>
<div style={{fontSize:11,color:MUTED}}>{s.sale_time}{s.route?' • '+s.route:''}{s.note?' • '+s.note:''}</div>
</div>
<Badge color={SUCCESS}>{fmt(s.value)}</Badge>
</div>)}
</div>
</>}
{orders.length>0&&<>
<div style={{fontWeight:700,fontSize:13,marginBottom:8,color:WARNING}}>📋 A Exportar ({orders.length})</div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,overflow:'hidden',marginBottom:12}}>
{orders.map((o,i)=><div key={o.id} style={{padding:'10px 14px',borderBottom:`1px solid ${BORDER}`,background:i%2===0?CARD:SURFACE,display:'flex',alignItems:'center',gap:8}}>
<div style={{flex:1}}>
<div style={{fontWeight:600,fontSize:13}}>{o.client_name}</div>
<div style={{fontSize:11,color:MUTED}}>{o.situacao} • {o.produtos?.length} produto(s)</div>
</div>
<Badge color={WARNING}>{fmt(o.total)}</Badge>
</div>)}
</div>
{orders.length>0&&<button onClick={()=>setActiveTab('pedido')} style={{width:'100%',background:SUCCESS,color:'#fff',border:'none',borderRadius:10,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>🚀 Ir para Exportar</button>}
</>}
{sales.length===0&&orders.length===0&&<div style={{textAlign:'center',padding:'40px 20px',color:MUTED,background:CARD,border:`1px solid ${BORDER}`,borderRadius:12}}><div style={{fontSize:36,marginBottom:8}}>📋</div><div style={{fontWeight:700}}>Nenhuma venda hoje</div></div>}
</>})()}
</div>}
{activeTab==='pedido'&&<div>
{editandoOrder&&<div style={{position:'fixed',inset:0,background:'#0008',zIndex:500,display:'flex',alignItems:'flex-end'}}>
<div style={{background:CARD,borderRadius:'16px 16px 0 0',padding:20,width:'100%',maxHeight:'90vh',overflowY:'auto'}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:14}}>
<div style={{fontWeight:800,fontSize:15}}>✏️ Editar Pedido</div>
<button onClick={()=>setEditandoOrder(null)} style={{background:'none',border:'none',fontSize:20,cursor:'pointer',color:MUTED}}>✕</button>
</div>
<div style={{fontWeight:600,fontSize:13,color:ACCENT,marginBottom:12}}>{editandoOrder.client_name}</div>
<select value={editOrderSituacao} onChange={e=>setEditOrderSituacao(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,background:SURFACE,marginBottom:8}}>
<option>Pedido S/ NFe</option><option>Pedido C/ NFe</option><option>Bonificação</option><option>Troca</option>
</select>
<select value={editOrderFormaPgto} onChange={e=>setEditOrderFormaPgto(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,background:SURFACE,marginBottom:8}}>
<option value="1">Dinheiro</option><option value="2">Cheque</option><option value="8">Pix/Ted</option><option value="16">Boleto Sicoob</option><option value="17">Débito em Conta</option>
</select>
<button onClick={abrirModalProdutos} style={{width:'100%',background:SURFACE,border:`1px solid ${BORDER}`,borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,color:ACCENT,cursor:'pointer',marginBottom:8}}>
+ Adicionar Produto
</button>
{editOrderProdutos.map(p=><ProdutoCard key={p.codigo} p={p} onChange={np=>setEditOrderProdutos(prev=>prev.map(x=>x.codigo===p.codigo?np:x))} onRemove={()=>setEditOrderProdutos(prev=>prev.filter(x=>x.codigo!==p.codigo))}/>)}
<div style={{textAlign:'right',fontWeight:800,fontSize:16,color:ACCENT,margin:'8px 0 12px'}}>Total: {fmt(editOrderProdutos.reduce((acc,p)=>acc+p.precoVenda*p.quant*(1-(p.vDesc||0)/100),0))}</div>
<div style={{display:'flex',gap:8}}>
<button onClick={()=>setEditandoOrder(null)} style={{flex:1,background:SURFACE,color:MUTED,border:`1px solid ${BORDER}`,borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>Cancelar</button>
<button onClick={salvarEdicaoOrder} style={{flex:2,background:SUCCESS,color:'#fff',border:'none',borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>💾 Salvar</button>
</div>
</div>
</div>}
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px',marginBottom:12}}>
<div style={{fontWeight:700,fontSize:14,marginBottom:12}}>🛒 Novo Pedido</div>
{pedidoCliente?<div style={{display:'flex',alignItems:'center',gap:8,background:ACCENT_LIGHT,borderRadius:8,padding:'10px 12px',marginBottom:8}}>
<span style={{fontWeight:700,color:ACCENT,flex:1,fontSize:13}}>{pedidoCliente.name}</span>
<button onClick={()=>setPedidoCliente(null)} style={{background:'none',border:'none',color:DANGER,cursor:'pointer',fontSize:16}}>✕</button>
</div>
:<div style={{position:'relative',marginBottom:8}}>
<input type="text" placeholder="🔍 Buscar cliente…" value={pedidoClienteSearch||''} onChange={e=>setPedidoClienteSearch(e.target.value)}
style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,boxSizing:'border-box'}}/>
{pedidoClienteSearch&&(selectedRoute?routeClients:clients).filter(c=>c.name.toLowerCase().includes(pedidoClienteSearch.toLowerCase())).length>0&&
<div style={{position:'absolute',top:'100%',left:0,right:0,zIndex:100,background:CARD,border:`1px solid ${BORDER}`,borderRadius:8,boxShadow:'0 8px 24px #0002',marginTop:4,maxHeight:220,overflowY:'auto'}}>
{(selectedRoute?routeClients:clients).filter(c=>c.name.toLowerCase().includes(pedidoClienteSearch.toLowerCase())).slice(0,8).map(c=>
<div key={c.id} onMouseDown={()=>{setPedidoCliente(c);setPedidoClienteSearch('')}}
style={{padding:'10px 14px',cursor:'pointer',borderBottom:`1px solid ${BORDER}`,fontSize:13}}
onMouseEnter={e=>e.currentTarget.style.background=ACCENT_LIGHT}
onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
<div style={{fontWeight:600}}>{c.name}</div>
<div style={{fontSize:11,color:MUTED}}>{c.route}</div>
</div>)}
</div>}
</div>}
<select value={pedidoSituacao} onChange={e=>setPedidoSituacao(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,background:SURFACE,marginBottom:8}}>
<option>Pedido S/ NFe</option><option>Pedido C/ NFe</option><option>Bonificação</option><option>Troca</option>
</select>
{['Pedido S/ NFe','Pedido C/ NFe'].includes(pedidoSituacao)&&<>
<select value={pedidoFormaPgto} onChange={e=>setPedidoFormaPgto(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,background:SURFACE,marginBottom:8}}>
<option value="1">Dinheiro</option><option value="2">Cheque</option><option value="8">Pix/Ted</option><option value="16">Boleto Sicoob</option><option value="17">Débito em Conta</option>
</select>
<div style={{marginBottom:8}}>
<label style={{fontSize:11,fontWeight:600,color:MUTED,display:'block',marginBottom:4}}>PRAZO DE VENCIMENTO</label>
<div style={{display:'flex',gap:8}}>
{[7,14,21,28].map(dias=>{
const dataVenc=new Date(Date.now()+dias*86400000).toISOString().split('T')[0]
const selecionado=pedidoVencimento===dataVenc
return<button key={dias} onClick={()=>setPedidoVencimento(dataVenc)} style={{flex:1,background:selecionado?ACCENT:SURFACE,color:selecionado?'#fff':MUTED,border:`1px solid ${selecionado?ACCENT:BORDER}`,borderRadius:8,padding:'10px 0',fontSize:13,fontWeight:700,cursor:'pointer'}}>{dias}d</button>
})}
</div>
{pedidoVencimento&&<div style={{fontSize:11,color:MUTED,marginTop:6,textAlign:'center'}}>Vencimento: {new Date(pedidoVencimento+'T12:00:00').toLocaleDateString('pt-BR')}</div>}
</div>
</>}
<button onClick={abrirModalProdutos} style={{width:'100%',background:SURFACE,border:`1px solid ${BORDER}`,borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,color:ACCENT,cursor:'pointer',marginBottom:8}}>+ Adicionar Produto</button>
{pedidoProdutos.map(p=><ProdutoCard key={p.codigo} p={p} onChange={np=>setPedidoProdutos(prev=>prev.map(x=>x.codigo===p.codigo?np:x))} onRemove={()=>setPedidoProdutos(prev=>prev.filter(x=>x.codigo!==p.codigo))}/>)}
{pedidoProdutos.length>0&&<div style={{textAlign:'right',fontWeight:800,fontSize:18,color:ACCENT,margin:'8px 0 12px'}}>Total: {fmt(totalPedido)}</div>}
<button onClick={salvarPedido} disabled={pedidoLoading} style={{width:'100%',background:pedidoLoading?MUTED:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'14px 0',fontWeight:700,fontSize:15,cursor:pedidoLoading?'not-allowed':'pointer'}}>
{pedidoLoading?'Salvando…':'💾 Salvar Pedido'}
</button>
</div>
{orders.length>0&&<>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>
<div style={{fontWeight:700,fontSize:14}}>Pedidos Pendentes ({orders.length})</div>
<div style={{fontWeight:800,fontSize:13,color:ACCENT}}>{fmt(orders.reduce((a,o)=>a+o.total,0))}</div>
</div>
{orders.map(o=><div key={o.id} style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'12px 14px',marginBottom:8}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:6}}>
<div><div style={{fontWeight:700,fontSize:13}}>{o.client_name}</div><div style={{fontSize:11,color:MUTED}}>{o.situacao} • {FormaPgtoLabel[o.forma_pgto]||o.forma_pgto}</div></div>
<div style={{textAlign:'right'}}><div style={{fontWeight:800,color:SUCCESS,fontSize:15}}>{fmt(o.total)}</div><div style={{fontSize:10,color:MUTED}}>{o.produtos?.length} produto(s)</div></div>
</div>
<div style={{fontSize:11,color:MUTED,marginBottom:8}}>{o.produtos?.map(p=>`${p.descricao} x${p.quant}`).join(', ')}</div>
<div style={{display:'flex',gap:6}}>
<button onClick={()=>abrirEdicaoOrder(o)} style={{flex:1,background:SURFACE,border:`1px solid ${BORDER}`,borderRadius:7,padding:'7px 0',fontSize:12,fontWeight:600,color:TEXT,cursor:'pointer'}}>✏️ Editar</button>
<button onClick={()=>gerarPdf(o)} style={{flex:1,background:'#EFF6FF',border:`1px solid ${ACCENT}33`,borderRadius:7,padding:'7px 0',fontSize:12,fontWeight:600,color:ACCENT,cursor:'pointer'}}>📄 PDF</button>
<button onClick={()=>excluirPedido(o.id)} style={{flex:1,background:'#FEF2F2',border:`1px solid ${DANGER}33`,borderRadius:7,padding:'7px 0',fontSize:12,fontWeight:600,color:DANGER,cursor:'pointer'}}>🗑️ Excluir</button>
</div>
</div>)}
<button onClick={exportarPedidos} disabled={exportLoading} style={{width:'100%',background:exportLoading?MUTED:SUCCESS,color:'#fff',border:'none',borderRadius:10,padding:'14px 0',fontWeight:800,fontSize:15,cursor:exportLoading?'not-allowed':'pointer',marginTop:4}}>
{exportLoading?'Exportando…':'🚀 Exportar para eGestor'}
</button>
</>}
</div>}
{activeTab==='config'&&user?.id===ADMIN_ID&&<div>
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px',marginBottom:12}}>
<div style={{fontWeight:700,fontSize:14,marginBottom:12}}>⚙️ Configurações</div>
<button onClick={carregarConfig} disabled={configLoading} style={{width:'100%',background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'10px 0',fontWeight:700,fontSize:14,cursor:'pointer',marginBottom:12}}>
{configLoading?'Carregando…':'🔄 Carregar Usuários'}
</button>
<div style={{fontWeight:700,fontSize:13,marginBottom:8}}>➕ Novo Usuário</div>
<input type="text" placeholder="Nome completo" value={novoNome} onChange={e=>setNovoNome(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,boxSizing:'border-box',marginBottom:8}}/>
<input type="email" placeholder="E-mail" value={novoEmail} onChange={e=>setNovoEmail(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,boxSizing:'border-box',marginBottom:8}}/>
<input type="password" placeholder="Senha" value={novaSenha} onChange={e=>setNovaSenha(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,boxSizing:'border-box',marginBottom:8}}/>
<select value={novoEgestorCode} onChange={e=>setNovoEgestorCode(e.target.value)} style={{width:'100%',border:`1px solid ${BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:14,background:SURFACE,marginBottom:8}}>
<option value="">Selecionar vendedor eGestor…</option>
{configVendedores.map(v=><option key={v.codigo} value={v.codigo}>{v.nome}</option>)}
</select>
<button onClick={criarUsuario} disabled={configLoading} style={{width:'100%',background:SUCCESS,color:'#fff',border:'none',borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>✅ Criar Usuário</button>
</div>
{configUsuarios.length>0&&<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px',marginBottom:12}}>
<div style={{fontWeight:700,fontSize:13,marginBottom:12}}>👥 Usuários Cadastrados</div>
{configUsuarios.map(u=><div key={u.id} style={{border:`1px solid ${BORDER}`,borderRadius:10,padding:'12px',marginBottom:8}}>
<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
<div>
<div style={{fontWeight:600,fontSize:13}}>{u.config?.name||'Sem nome'}</div>
<div style={{fontSize:11,color:MUTED}}>{u.email}</div>
<div style={{fontSize:11,color:MUTED}}>Vendedor: {configVendedores.find(v=>v.codigo===u.config?.egestor_vendedor_code)?.nome||'Não vinculado'}</div>
</div>
{u.id!==ADMIN_ID&&<button onClick={()=>deletarUsuario(u.id)} style={{background:'#FEF2F2',border:`1px solid ${DANGER}33`,borderRadius:7,padding:'6px 10px',fontSize:11,fontWeight:600,color:DANGER,cursor:'pointer'}}>🗑️</button>}
</div>
{u.id!==ADMIN_ID&&<div style={{marginBottom:8}}>
<div style={{fontSize:11,fontWeight:600,color:MUTED,marginBottom:6}}>ROTAS DO VENDEDOR</div>
<div style={{display:'flex',flexWrap:'wrap',gap:6,marginBottom:8}}>
{['AFONSO CLAUDIO','BREJETUBA','DOMINGOS MARTINS','LARANJA DA TERRA','SANTA MARIA','SANTA TERESA','VENDA NOVA'].map(rota=>{const temRota=(u.config?.rotas||[]).includes(rota);return<button key={rota} onClick={async()=>{const novasRotas=temRota?(u.config?.rotas||[]).filter(r=>r!==rota):[...(u.config?.rotas||[]),rota];const{error}=await supabase.from('user_config').update({rotas:novasRotas}).eq('user_id',u.id);if(error){showToast('Erro: '+error.message,'error');return}showToast('Rota atualizada!');await carregarConfig()}} style={{background:temRota?ACCENT:SURFACE,color:temRota?'#fff':MUTED,border:`1px solid ${temRota?ACCENT:BORDER}`,borderRadius:6,padding:'4px 10px',fontSize:11,fontWeight:600,cursor:'pointer'}}>{rota}</button>})}</div>
</div>}
{u.id!==ADMIN_ID&&<div style={{display:'flex',gap:8}}>
<input type="password" placeholder="Nova senha…" value={senhaUser[u.id]||''} onChange={e=>setSenhaUser(prev=>({...prev,[u.id]:e.target.value}))} style={{flex:1,border:`1px solid ${BORDER}`,borderRadius:8,padding:'8px 10px',fontSize:13,boxSizing:'border-box'}}/>
<button onClick={()=>atualizarSenha(u.id)} style={{background:ACCENT,color:'#fff',border:'none',borderRadius:8,padding:'8px 12px',fontWeight:600,fontSize:12,cursor:'pointer'}}>Alterar</button>
</div>}
</div>)}
</div>}
<div style={{background:CARD,border:`1px solid ${BORDER}`,borderRadius:12,padding:'14px 16px'}}>
<div style={{fontWeight:700,fontSize:13,marginBottom:8}}>🔄 Sincronizar Produtos</div>
<div style={{fontSize:12,color:MUTED,marginBottom:12}}>Atualiza todos os produtos do eGestor no CRM.</div>
<button onClick={async()=>{setConfigLoading(true);try{await fetch('https://qtogmmgkpnpkmvnkoxsz.supabase.co/functions/v1/sync-produtos');showToast('Produtos sincronizados!')}catch{showToast('Erro ao sincronizar','error')}setConfigLoading(false)}} disabled={configLoading} style={{width:'100%',background:WARNING,color:'#fff',border:'none',borderRadius:8,padding:'12px 0',fontWeight:700,fontSize:14,cursor:'pointer'}}>
📦 Sincronizar Produtos
</button>
</div>
</div>}
</div>
</div>)
}

