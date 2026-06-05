/* JM3D marketplace profit analysis module. Loaded before the main inline app. */

function getRowsFromSheet(ws){
  return XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});
}
function workbookSheets(wb){
  return wb.SheetNames.map(name=>({name,rows:getRowsFromSheet(wb.Sheets[name])}));
}
function detectShopeeStore(sheets,fileName){
  const rows=sheets.flatMap(s=>s.rows).slice(0,80);
  for(const r of rows){
    const i=r.findIndex(c=>norm(c).includes('nome de usuario vendedor'));
    if(i>-1&&r[i+1])return cleanCell(r[i+1]);
  }
  const n=norm(fileName);
  if(n.includes('kaline'))return '_kaline98';
  if(n.includes('mateus'))return 'mateusoliver98';
  return 'Detectar';
}
function detectImportKind(fileName,sheets){
  const n=norm(fileName);
  const joined=sheets.flatMap(s=>s.rows.slice(0,8)).flat().map(norm).join(' ');
  if(n.includes('income lancado')||joined.includes('quantia total lancada'))return 'shopee_income';
  if(n.includes('order completed')||joined.includes('status do pedido'))return 'shopee_completed';
  if(n.includes('enviado pedido')||joined.includes('order substatus'))return 'tiktok_orders';
  if(n.startsWith('income')||joined.includes('valor total a ser liquidado'))return 'tiktok_income';
  return 'unknown';
}
function makeImportId(o){
  const item=o.sku||o.productName||'pedido';
  return [o.marketplace,o.store,o.orderId,item].map(slug).join('__');
}
function mergeOrderLine(map,o){
  if(!o.orderId)return;
  o.id=makeImportId(o);
  const old=map[o.id]||{};
  map[o.id]=Object.assign({},old,o,{
    gross:o.gross||old.gross||0,
    net:o.net||old.net||0,
    fees:Math.abs(o.fees||old.fees||0),
    shipping:o.shipping||old.shipping||0,
    discounts:Math.abs(o.discounts||old.discounts||0),
    qty:o.qty||old.qty||1,
    sourceFiles:[...new Set([...(old.sourceFiles||[]),...(o.sourceFiles||[])])]
  });
}
function parseShopeeIncome(sheets,fileName){
  const store=detectShopeeStore(sheets,fileName), out=[];
  sheets.forEach(sheet=>{
    const hi=findHeaderRow(sheet.rows,['id do pedido','quantia total lancada']);
    if(hi<0)return;
    const headers=sheet.rows[hi];
    sheet.rows.slice(hi+1).forEach(row=>{
      const obj=rowToObj(headers,row);
      const view=cleanCell(pick(obj,['Ver']));
      if(view&&norm(view)!=='sku')return;
      const orderId=cleanCell(pick(obj,['ID do pedido']));
      if(!orderId)return;
      const gross=parseMoneyBR(pick(obj,['Preço do produto','Preco do produto','Quantia paga pelo comprador']));
      const net=parseMoneyBR(pick(obj,['Quantia total lançada (R$)','Quantia total lancada']));
      const fees=Math.abs(sumMoney(obj,['Taxa de comissão líquida','Taxa de servico liquida','Taxa de transação','Taxa de comissao Afiliados do Vendedor','Taxa de Devolução Fácil Shopee','Taxa da Recarga Automática (Pedido)']));
      const discounts=Math.abs(sumMoney(obj,['Cupom','Voucher subsidiado pelo Seller','Voucher compartilhado subsidiado pelo Seller','Coin Cashback subsidiado pelo Seller','Coin Cashback compartilhado subsidiado pelo Seller','Valor do Reembolso']));
      out.push({
        marketplace:'shopee',store,kind:'Shopee Income',orderId,
        date:parseDateFlexible(pick(obj,['Data de conclusão do pagamento','Data de criação do pedido'])),
        paidDate:parseDateFlexible(pick(obj,['Data de conclusão do pagamento'])),
        sku:cleanCell(pick(obj,['SKU'])).replace(/^[-/]$/,''),
        productName:cleanCell(pick(obj,['Nome do produto'])),
        qty:1,gross,net,fees,discounts,
        shipping:sumMoney(obj,['Frete cobrado pelo parceiro logístico','Desconto de frete pela Shopee','Taxa de envio reverso']),
        status:net?'Pago':'Pendente',
        sourceFiles:[fileName]
      });
    });
  });
  return out.filter(o=>o.productName||o.net||o.gross);
}
function parseShopeeCompleted(sheets,fileName){
  const rows=sheets.flatMap(s=>s.rows);
  const hi=findHeaderRow(rows,['id do pedido','nome do produto']);
  if(hi<0)return [];
  const headers=rows[hi], out=[];
  rows.slice(hi+1).forEach(row=>{
    const obj=rowToObj(headers,row);
    const orderId=cleanCell(pick(obj,['ID do pedido']));
    if(!orderId)return;
    const method=cleanCell(pick(obj,['Opção de envio','Opcao de envio']));
    const store=method.includes('CPF')?'mateusoliver98':'_kaline98';
    out.push({
      marketplace:'shopee',store,kind:'Shopee Order Completed',orderId,
      date:parseDateFlexible(pick(obj,['Data de criação do pedido','Hora do pagamento do pedido'])),
      paidDate:parseDateFlexible(pick(obj,['Hora do pagamento do pedido'])),
      sku:cleanCell(pick(obj,['Número de referência SKU','Nº de referência do SKU principal'])),
      productName:cleanCell(pick(obj,['Nome do Produto'])),
      qty:parseInt(pick(obj,['Quantidade','Número de produtos pedidos']))||1,
      gross:parseMoneyBR(pick(obj,['Valor Total','Subtotal do produto','Preço acordado'])),
      net:0,fees:0,
      shipping:parseMoneyBR(pick(obj,['Taxa de envio pagas pelo comprador'])),
      discounts:Math.abs(sumMoney(obj,['Desconto do vendedor','Cupom do vendedor','Cupom','Compensar Moedas Shopee'])),
      status:cleanCell(pick(obj,['Status do pedido']))||'Concluído',
      sourceFiles:[fileName]
    });
  });
  return out;
}
function parseTikTokOrders(sheets,fileName){
  const rows=sheets.flatMap(s=>s.rows);
  const hi=findHeaderRow(rows,['order id','product name']);
  if(hi<0)return [];
  const headers=rows[hi], out=[];
  rows.slice(hi+1).forEach(row=>{
    const obj=rowToObj(headers,row);
    const orderId=cleanCell(pick(obj,['Order ID']));
    if(!orderId)return;
    const gross=parseMoneyBR(pick(obj,['SKU Subtotal Before Discount','Order Amount']));
    const net=parseMoneyBR(pick(obj,['Order Amount','SKU Subtotal After Discount']));
    out.push({
      marketplace:'tiktok',store:'TikTok Shop',kind:'TikTok Orders',orderId,
      date:parseDateFlexible(pick(obj,['Created Time','Paid Time'])),
      paidDate:parseDateFlexible(pick(obj,['Paid Time'])),
      sku:cleanCell(pick(obj,['Seller SKU','SKU ID'])),
      productName:cleanCell(pick(obj,['Product Name'])),
      qty:parseInt(pick(obj,['Quantity']))||1,
      gross,net,fees:Math.max(0,gross-net),
      shipping:parseMoneyBR(pick(obj,['Shipping Fee After Discount','Original Shipping Fee'])),
      discounts:Math.abs(sumMoney(obj,['SKU Platform Discount','SKU Seller Discount','Shipping Fee Seller Discount','Shipping Fee Platform Discount','Payment platform discount'])),
      status:cleanCell(pick(obj,['Order Status','Order Substatus'])),
      sourceFiles:[fileName]
    });
  });
  return out;
}
function parseTikTokIncome(sheets,fileName){
  const out=[];
  sheets.forEach(sheet=>{
    const hi=findHeaderRow(sheet.rows,['id do pedido','nome do produto','valor total a ser liquidado']);
    if(hi<0)return;
    const headers=sheet.rows[hi];
    sheet.rows.slice(hi+1).forEach(row=>{
      const obj=rowToObj(headers,row);
      const orderId=cleanCell(pick(obj,['ID do pedido','ID do demonstrativo']));
      if(!orderId)return;
      const productName=cleanCell(pick(obj,['Nome do produto']));
      const net=parseMoneyBR(pick(obj,['Valor total a ser liquidado','Valor do pagamento']));
      const gross=parseMoneyBR(pick(obj,['Vendas líquidas','Subtotal do item antes dos descontos','Preço do SKU']))||net;
      const fees=Math.abs(sumMoney(obj,['Taxas','Taxa de comissão','Taxa da plataforma','Taxa de transação']));
      out.push({
        marketplace:'tiktok',store:'TikTok Shop',kind:'TikTok Income',orderId,
        date:parseDateFlexible(pick(obj,['Data do demonstrativo','Data de início do pagamento','Data de criacao do pedido'])),
        paidDate:parseDateFlexible(pick(obj,['Data de conclusão do pagamento'])),
        sku:cleanCell(pick(obj,['SKU ID','ID do SKU'])),
        productName,qty:parseInt(pick(obj,['Quantidade']))||1,
        gross,net,fees,
        shipping:parseMoneyBR(pick(obj,['Frete'])),
        discounts:Math.abs(sumMoney(obj,['Descontos','Cupom','Ajustes'])),
        status:cleanCell(pick(obj,['Status']))||'Pago',
        sourceFiles:[fileName]
      });
    });
  });
  return out.filter(o=>o.productName||o.net||o.gross);
}
function parseCsvRows(text){
  const rows=[], cur=[];let val='',q=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i],nx=text[i+1];
    if(ch==='"'&&q&&nx==='"'){val+='"';i++;}
    else if(ch==='"')q=!q;
    else if(ch===','&&!q){cur.push(val);val='';}
    else if((ch==='\n'||ch==='\r')&&!q){
      if(ch==='\r'&&nx==='\n')i++;
      cur.push(val);
      if(cur.some(x=>x!==''))rows.push(cur.splice(0));
      val='';
    }else val+=ch;
  }
  cur.push(val);if(cur.some(x=>x!==''))rows.push(cur);
  return rows;
}
async function fileToSheets(file){
  if(file.name.toLowerCase().endsWith('.csv')){
    const text=await file.text();
    return [{name:file.name,rows:parseCsvRows(text)}];
  }
  if(!window.XLSX)throw new Error('Biblioteca XLSX não carregou. Verifique a internet.');
  const data=await file.arrayBuffer();
  return workbookSheets(XLSX.read(data,{type:'array',cellDates:true}));
}
async function processMarketplaceFiles(files){
  const map={}, summaries=[], errors=[];
  for(const file of files){
    try{
      const sheets=await fileToSheets(file);
      const kind=detectImportKind(file.name,sheets);
      let rows=[];
      if(kind==='shopee_income')rows=parseShopeeIncome(sheets,file.name);
      else if(kind==='shopee_completed')rows=parseShopeeCompleted(sheets,file.name);
      else if(kind==='tiktok_orders')rows=parseTikTokOrders(sheets,file.name);
      else if(kind==='tiktok_income')rows=parseTikTokIncome(sheets,file.name);
      else errors.push(`${file.name}: arquivo não reconhecido`);
      rows.forEach(o=>mergeOrderLine(map,o));
      summaries.push({file:file.name,kind,orders:rows.length,recognized:kind!=='unknown'});
    }catch(e){errors.push(`${file.name}: ${e.message}`);}
  }
  const existing=new Set(marketplaceOrders.map(o=>o.id));
  const orders=Object.values(map).map(o=>enrichOrder(o));
  importReview={
    orders,
    newOrders:orders.filter(o=>!existing.has(o.id)),
    duplicates:orders.filter(o=>existing.has(o.id)),
    summaries,errors,createdAt:Date.now()
  };
  renderImportReview();
}
function findProductForOrder(o){
  const rule=productMatchRules[slug(o.sku||o.productName)]||productMatchRules[slug(o.productName)];
  if(rule){
    const p=localProducts.find(x=>x.id===rule.productId);
    if(p)return p;
  }
  const on=norm(o.productName);
  if(!on)return null;
  let best=null,bestScore=0;
  localProducts.forEach(p=>{
    const pn=norm(p.name);
    let score=pn===on?100:(on.includes(pn)||pn.includes(on)?70:0);
    if(!score){
      const a=new Set(on.split(' ').filter(x=>x.length>2));
      const b=new Set(pn.split(' ').filter(x=>x.length>2));
      const inter=[...a].filter(x=>b.has(x)).length;
      score=b.size?inter/b.size*60:0;
    }
    if(score>bestScore){best=p;bestScore=score;}
  });
  return bestScore>=35?best:null;
}
function productUnitCost(product,meiUnit=0){
  const s=getMarketplaceSettings(), cfg=getLocalConfig();
  if(!product)return {total:meiUnit,filament:0,energy:0,packaging:0,bubble:0,extra:0,mei:meiUnit};
  const weight=Number(product.weight)||0, printH=Number(product.printH)||0;
  const filament=weight/1000*(s.filamentKgPrice||95);
  const energy=printH*(s.printerKwhHour||0.1)*(s.energyKwhPrice||1.05);
  const packaging=Number(product.breakdown?.pkg)||0;
  const bubble=Number(product.breakdown?.bubble)||cfg.bubbleUnit||0;
  const maint=Number(product.breakdown?.maint)||cfg.maintenance||0;
  const base=filament+energy+packaging+bubble+maint+(s.labelCost||0)+(s.postProcessCost||0)+(s.nozzleWearCost||0)+(s.otherProductionCost||0);
  const loss=base*((s.lossPct||0)/100);
  return {total:base+loss+meiUnit,filament,energy,packaging,bubble,extra:maint+loss+(s.labelCost||0)+(s.postProcessCost||0)+(s.nozzleWearCost||0)+(s.otherProductionCost||0),mei:meiUnit};
}
function enrichOrder(o,meiUnit=0){
  const product=findProductForOrder(o);
  const c=productUnitCost(product,meiUnit);
  const qty=Number(o.qty)||1;
  const net=o.net||(o.gross-Math.abs(o.fees||0)-Math.abs(o.discounts||0)-Math.abs(o.shipping||0));
  const cost=c.total*qty;
  const profit=net-cost;
  const margin=net>0?profit/net*100:0;
  return Object.assign({},o,{id:o.id||makeImportId(o),linkedProductId:product?.id||'',linkedProductName:product?.name||'',unitCost:c.total,totalCost:cost,net,profit,margin,health:profit<0?'bad':margin<15?'warn':'ok',costParts:c});
}
function analyzedOrders(){
  const base=marketplaceOrders.map(o=>Object.assign({},o));
  const units=base.reduce((t,o)=>t+(Number(o.qty)||1),0)||1, count=base.length||1;
  const s=getMarketplaceSettings();
  const meiUnit=s.meiAllocation==='ignore'?0:(s.meiAllocation==='order'?(s.dasMeiMonthly||0)/count:(s.dasMeiMonthly||0)/units);
  return base.map(o=>enrichOrder(o,meiUnit));
}
function filteredOrders(){
  return analyzedOrders().filter(o=>{
    if(analysisFilters.marketplace!=='all'&&o.marketplace!==analysisFilters.marketplace)return false;
    if(analysisFilters.store!=='all'&&o.store!==analysisFilters.store)return false;
    if(analysisFilters.health==='bad'&&o.health!=='bad')return false;
    if(analysisFilters.health==='unlinked'&&o.linkedProductId)return false;
    const q=norm(analysisFilters.q);
    if(q&&!norm(`${o.orderId} ${o.productName} ${o.linkedProductName}`).includes(q))return false;
    return true;
  });
}
function summarizeOrders(orders){
  return orders.reduce((a,o)=>{
    const qty=Number(o.qty)||1;
    a.gross+=Number(o.gross)||0;a.net+=Number(o.net)||0;a.fees+=Math.abs(Number(o.fees)||0);
    a.cost+=Number(o.totalCost)||0;a.profit+=Number(o.profit)||0;a.qty+=qty;
    a.energy+=Number(o.costParts?.energy||0)*qty;a.filament+=Number(o.costParts?.filament||0)*qty;
    a.packaging+=(Number(o.costParts?.packaging||0)+Number(o.costParts?.bubble||0))*qty;
    a.mei+=Number(o.costParts?.mei||0)*qty;
    if(!o.linkedProductId)a.unlinked++; if(o.health==='bad')a.bad++;
    if(!/pago|conclu/i.test(o.status||''))a.pending++;
    return a;
  },{gross:0,net:0,fees:0,cost:0,profit:0,qty:0,energy:0,filament:0,packaging:0,mei:0,unlinked:0,bad:0,pending:0});
}
function renderMetric(label,value,cls=''){
  return `<div class="metric ${cls}"><div class="ml">${label}</div><div class="mv">${value}</div></div>`;
}
function renderAnalysis(){
  const content=document.getElementById('content'); if(!content)return;
  const tabs=`<div class="seg">
    <button class="${analysisMode==='dashboard'?'on':''}" onclick="analysisMode='dashboard';renderAnalysis()">Dashboard</button>
    <button class="${analysisMode==='import'?'on':''}" onclick="analysisMode='import';renderAnalysis()">Importar</button>
    <button class="${analysisMode==='orders'?'on':''}" onclick="analysisMode='orders';renderAnalysis()">Pedidos</button>
    <button class="${analysisMode==='pricing'?'on':''}" onclick="analysisMode='pricing';renderAnalysis()">Precificação</button>
  </div>`;
  content.innerHTML=tabs+(analysisMode==='import'?renderImportView():analysisMode==='orders'?renderOrdersView():analysisMode==='pricing'?renderPricingView():renderDashboardView());
}
function renderDashboardView(){
  const orders=filteredOrders(), s=summarizeOrders(orders), margin=s.net?s.profit/s.net*100:0;
  const stores=['all','_kaline98','mateusoliver98','TikTok Shop'];
  return `<div class="filter-row">
    <select class="fi" onchange="analysisFilters.store=this.value;renderAnalysis()">${stores.map(x=>`<option value="${x}" ${analysisFilters.store===x?'selected':''}>${x==='all'?'Todas as lojas':x}</option>`).join('')}</select>
    <select class="fi" onchange="analysisFilters.marketplace=this.value;renderAnalysis()"><option value="all" ${analysisFilters.marketplace==='all'?'selected':''}>Todos marketplaces</option><option value="shopee" ${analysisFilters.marketplace==='shopee'?'selected':''}>Shopee</option><option value="tiktok" ${analysisFilters.marketplace==='tiktok'?'selected':''}>TikTok Shop</option></select>
  </div>
  <div class="dash-grid">
    ${renderMetric('Vendas brutas',brl(s.gross))}
    ${renderMetric('Líquido recebido',brl(s.net),'good')}
    ${renderMetric('Taxas totais',brl(s.fees),'warn')}
    ${renderMetric('Custos produção',brl(s.cost))}
    ${renderMetric('Filamento',brl(s.filament))}
    ${renderMetric('Energia',brl(s.energy))}
    ${renderMetric('Embalagens',brl(s.packaging))}
    ${renderMetric('DAS MEI rateado',brl(s.mei))}
    ${renderMetric('Lucro líquido',brl(s.profit),s.profit>=0?'good':'bad')}
    ${renderMetric('Margem média',pct(margin),margin>=15?'good':margin>=0?'warn':'bad')}
    ${renderMetric('Pendentes',s.pending,'warn')}
    ${renderMetric('Com prejuízo',s.bad,'bad')}
  </div>
  ${renderProductSummary(orders)}
  <button class="btn btn-export" onclick="exportAnalysisCsv()">Exportar CSV da análise</button>`;
}
function renderProductSummary(orders){
  const map={};
  orders.forEach(o=>{
    const k=o.linkedProductName||o.productName||'Produto não vinculado';
    map[k]=map[k]||{name:k,qty:0,gross:0,net:0,cost:0,profit:0};
    map[k].qty+=Number(o.qty)||1;map[k].gross+=o.gross||0;map[k].net+=o.net||0;map[k].cost+=o.totalCost||0;map[k].profit+=o.profit||0;
  });
  const rows=Object.values(map).sort((a,b)=>b.profit-a.profit).slice(0,8);
  if(!rows.length)return '<div class="empty"><div class="ei">📊</div><div>Importe relatórios para ver a análise real.</div></div>';
  return `<div class="sec"><div class="sec-hdr"><span>🏆</span><span class="sec-title">Análise por Produto</span></div><div class="sec-body"><table class="mini-table"><thead><tr><th>Produto</th><th>Qtd</th><th>Lucro</th><th>Margem</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.name)}</td><td class="num">${r.qty}</td><td class="num">${brl(r.profit)}</td><td class="num">${pct(r.net?r.profit/r.net*100:0)}</td></tr>`).join('')}</tbody></table></div></div>`;
}
function renderImportView(){
  return `<div class="import-box"><div class="sec-title">Importação em lote Shopee/TikTok</div>
    <div class="muted-note">Selecione vários arquivos XLSX e CSV juntos. O app tenta detectar plataforma, loja, pedidos, taxas e valores automaticamente antes de salvar.</div>
    <div class="file-drop"><input class="fi" type="file" id="market-files" accept=".xlsx,.csv" multiple onchange="processMarketplaceFiles([...this.files])"><div class="muted-note">Formatos suportados: Shopee Income, Shopee Order.completed, TikTok income e TikTok Enviado pedido CSV.</div></div>
  </div><div id="import-review">${importReview?renderImportReviewHtml():''}</div>`;
}
function renderImportReview(){
  const el=document.getElementById('import-review');
  if(el)el.innerHTML=renderImportReviewHtml(); else renderAnalysis();
}
function renderImportReviewHtml(){
  const r=importReview, orders=r.orders.map(o=>enrichOrder(o)), s=summarizeOrders(orders);
  return `<div class="sec"><div class="sec-hdr"><span>✅</span><span class="sec-title">Revisão da importação</span></div><div class="sec-body">
    <div class="dash-grid">${renderMetric('Arquivos lidos',r.summaries.length)}${renderMetric('Reconhecidos',r.summaries.filter(x=>x.recognized).length,'good')}${renderMetric('Novos pedidos',r.newOrders.length,'good')}${renderMetric('Duplicados ignorados',r.duplicates.length,'warn')}${renderMetric('Não vinculados',orders.filter(o=>!o.linkedProductId).length,'warn')}${renderMetric('Erros',r.errors.length,r.errors.length?'bad':'good')}${renderMetric('Valor bruto',brl(s.gross))}${renderMetric('Valor líquido',brl(s.net),'good')}</div>
    <div class="review-list"><table class="mini-table"><thead><tr><th>Arquivo</th><th>Tipo</th><th>Pedidos</th></tr></thead><tbody>${r.summaries.map(x=>`<tr><td>${esc(x.file)}</td><td>${esc(x.kind)}</td><td class="num">${x.orders}</td></tr>`).join('')}</tbody></table></div>
    ${r.errors.length?`<div class="muted-note" style="color:var(--red)">${r.errors.map(esc).join('<br>')}</div>`:''}
    <button class="btn btn-save" onclick="confirmMarketplaceImport()">Confirmar importação</button><button class="btn btn-secondary" onclick="analysisMode='orders';renderAnalysis()">Corrigir vínculos</button><button class="btn btn-danger" onclick="importReview=null;renderAnalysis()">Cancelar</button>
  </div></div>`;
}
async function confirmMarketplaceImport(){
  if(!importReview)return;
  const save=importReview.newOrders.map(o=>enrichOrder(o));
  try{
    if(marketplaceOrdersRef){
      const updates={}; save.forEach(o=>updates[o.id]=sanitizeForFirebase(o));
      await marketplaceOrdersRef.update(updates);
    }else{
      marketplaceOrders=[...marketplaceOrders,...save];
      localStorage.setItem('jm3d_market_orders',JSON.stringify(marketplaceOrders));
    }
    setSyncStatus('ok',`✅ ${save.length} pedidos importados`);
    importReview=null;analysisMode='dashboard';renderAnalysis();
  }catch(e){setSyncStatus('err','Erro ao salvar importação: '+e.message);}
}
function renderOrdersView(){
  const orders=filteredOrders();
  return `<div class="filter-row"><input class="fi" placeholder="Buscar pedido/produto" value="${ea(analysisFilters.q)}" oninput="analysisFilters.q=this.value;renderAnalysis()"><select class="fi" onchange="analysisFilters.health=this.value;renderAnalysis()"><option value="all">Todos</option><option value="bad" ${analysisFilters.health==='bad'?'selected':''}>Lucro negativo</option><option value="unlinked" ${analysisFilters.health==='unlinked'?'selected':''}>Não vinculados</option></select></div>
  ${orders.length?orders.map(renderOrderCard).join(''):'<div class="empty"><div class="ei">📦</div><div>Nenhum pedido importado ainda.</div></div>'}`;
}
function renderOrderCard(o){
  const h=o.health==='bad'?'bad':o.health==='warn'?'warn':'ok';
  const hl=o.health==='bad'?'prejuízo':o.health==='warn'?'margem baixa':'saudável';
  return `<div class="order-card"><div class="order-top"><div><div class="order-title">${esc(o.productName||'Produto sem nome')}</div><div class="order-meta">${esc(o.marketplace)} • ${esc(o.store)} • ${esc(o.orderId)} • ${esc(o.date||'sem data')}</div></div><span class="health ${h}">${hl}</span></div>
    <div class="chips"><span class="chip c-price">Líquido ${brl(o.net)}</span><span class="chip c-cost">Custo ${brl(o.totalCost)}</span><span class="chip c-pct">Lucro ${brl(o.profit)} (${pct(o.margin)})</span></div>
    <div class="muted-note">Vínculo: ${o.linkedProductName?esc(o.linkedProductName):'<b>produto não vinculado</b>'}</div><button class="btn btn-secondary" style="font-size:12px;padding:9px" onclick="openLinkProduct('${o.id}')">Vincular produto importado</button></div>`;
}
function openLinkProduct(orderId){
  const o=analyzedOrders().find(x=>x.id===orderId); if(!o)return;
  openSht('Vincular produto',`<div class="muted-note">Produto no relatório:<br><b>${esc(o.productName)}</b></div><select class="fi" id="link-product">${localProducts.map(p=>`<option value="${p.id}" ${p.id===o.linkedProductId?'selected':''}>${esc(p.name)}</option>`).join('')}</select><button class="btn btn-save" onclick="saveProductLink('${o.id}')">Salvar vínculo</button><button class="btn btn-secondary" onclick="closeSht()">Cancelar</button>`);
}
function saveProductLink(orderId){
  const o=marketplaceOrders.find(x=>x.id===orderId)||importReview?.orders.find(x=>x.id===orderId); if(!o)return;
  const productId=document.getElementById('link-product')?.value; if(!productId)return;
  const key=slug(o.sku||o.productName);
  productMatchRules[key]={productId,source:o.productName||o.sku,savedTs:Date.now()};
  if(matchRulesRef)matchRulesRef.child(key).set(sanitizeForFirebase(productMatchRules[key])).catch(()=>{});
  closeSht();renderAnalysis();
}
function renderPricingView(){
  const cfg=getMarketplaceSettings();
  return `<div class="sec"><div class="sec-hdr"><span>⚙️</span><span class="sec-title">Custos Globais Profissionais</span></div><div class="sec-body">
    <div class="filter-row"><input class="fi" id="mc-fil" type="number" value="${cfg.filamentKgPrice}" placeholder="Filamento R$/kg"><input class="fi" id="mc-kwh" type="number" value="${cfg.energyKwhPrice}" placeholder="Energia R$/kWh"></div>
    <div class="filter-row"><input class="fi" id="mc-cons" type="number" value="${cfg.printerKwhHour}" placeholder="kWh/h impressora"><input class="fi" id="mc-das" type="number" value="${cfg.dasMeiMonthly}" placeholder="DAS MEI"></div>
    <div class="filter-row"><input class="fi" id="mc-loss" type="number" value="${cfg.lossPct}" placeholder="Perdas %"><input class="fi" id="mc-post" type="number" value="${cfg.postProcessCost}" placeholder="Pós-processamento"></div>
    <div class="filter-row"><input class="fi" id="mc-shp" type="number" value="${cfg.shopeeCommissionPct}" placeholder="Shopee %"><input class="fi" id="mc-tkt" type="number" value="${cfg.tiktokCommissionPct}" placeholder="TikTok %"></div>
    <button class="btn btn-save" onclick="saveMarketCfgUI()">Salvar custos globais</button></div></div>${renderPricingProducts()}`;
}
function saveMarketCfgUI(){
  const old=getMarketplaceSettings();
  saveMarketplaceSettings(Object.assign(old,{
    filamentKgPrice:parseMoneyBR(document.getElementById('mc-fil').value),
    energyKwhPrice:parseMoneyBR(document.getElementById('mc-kwh').value),
    printerKwhHour:parseMoneyBR(document.getElementById('mc-cons').value),
    dasMeiMonthly:parseMoneyBR(document.getElementById('mc-das').value),
    lossPct:parseMoneyBR(document.getElementById('mc-loss').value),
    postProcessCost:parseMoneyBR(document.getElementById('mc-post').value),
    shopeeCommissionPct:parseMoneyBR(document.getElementById('mc-shp').value),
    tiktokCommissionPct:parseMoneyBR(document.getElementById('mc-tkt').value)
  }));
  setSyncStatus('ok','✅ Custos globais salvos');renderAnalysis();
}
function renderPricingProducts(){
  if(!localProducts.length)return '<div class="empty"><div class="ei">🧮</div><div>Cadastre produtos para simular preços profissionais.</div></div>';
  return localProducts.slice(0,20).map(p=>{
    const c=productUnitCost(p,0), min=priceForMargin(c.total,'shopee',0), sh10=priceForMargin(c.total,'shopee',10), sh30=priceForMargin(c.total,'shopee',30), tk30=priceForMargin(c.total,'tiktok',30);
    const current=Number(p.price)||0, health=current&&current<min?'bad':current&&current<sh10?'warn':'ok';
    return `<div class="order-card"><div class="order-top"><div><div class="order-title">${esc(p.name)}</div><div class="order-meta">Custo técnico ${brl(c.total)} • peso ${p.weight||0}g • ${p.printH||0}h</div></div><span class="health ${health}">${health==='bad'?'abaixo do mínimo':health==='warn'?'margem baixa':'preço saudável'}</span></div>
    <div class="cost-bars"><span style="width:${Math.min(100,c.filament/c.total*100||0)}%;background:var(--cyan)"></span><span style="width:${Math.min(100,c.energy/c.total*100||0)}%;background:var(--orange)"></span><span style="width:${Math.min(100,(c.packaging+c.bubble)/c.total*100||0)}%;background:var(--purple)"></span></div>
    <table class="mini-table"><tr><th>Canal</th><th>Mínimo</th><th>10%</th><th>30%</th><th>50%</th></tr><tr><td>Shopee</td><td>${brl(min)}</td><td>${brl(sh10)}</td><td>${brl(sh30)}</td><td>${brl(priceForMargin(c.total,'shopee',50))}</td></tr><tr><td>TikTok</td><td>${brl(priceForMargin(c.total,'tiktok',0))}</td><td>${brl(priceForMargin(c.total,'tiktok',10))}</td><td>${brl(tk30)}</td><td>${brl(priceForMargin(c.total,'tiktok',50))}</td></tr><tr><td>Direta</td><td>${brl(c.total)}</td><td>${brl(priceForMargin(c.total,'direct',10))}</td><td>${brl(priceForMargin(c.total,'direct',30))}</td><td>${brl(priceForMargin(c.total,'direct',50))}</td></tr></table></div>`;
  }).join('');
}
function exportAnalysisCsv(){
  const rows=[['marketplace','loja','data','pedido','produto','qtd','bruto','liquido','taxas','custo','lucro','margem','produto_vinculado','status']];
  analyzedOrders().forEach(o=>rows.push([o.marketplace,o.store,o.date,o.orderId,o.productName,o.qty,o.gross,o.net,o.fees,o.totalCost,o.profit,o.margin,o.linkedProductName,o.status]));
  const csv=rows.map(r=>r.map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download='jm3d_analise_lucro.csv';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
}
