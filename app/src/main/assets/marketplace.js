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
function canonicalMarketplace(v){
  const n=norm(v);
  if(n.includes('tiktok'))return 'tiktokShop';
  if(n.includes('shopee'))return 'shopee';
  if(n.includes('direct'))return 'direct';
  return v||'unknown';
}
function marketplaceFilterValue(v){
  const c=canonicalMarketplace(v);
  return c==='tiktokShop'?'tiktokShop':c;
}
function isIncomeKind(kind){
  return norm(kind).includes('income');
}
function makeImportId(o){
  return [canonicalMarketplace(o.marketplace),o.store,o.orderId].map(slug).join('_');
}
function mergeOrderLine(map,o){
  if(!o.orderId)return;
  o.id=makeImportId(o);
  const old=map[o.id]||{};
  const hasOld=!!old.id;
  const sameKind=!hasOld||old.kind===o.kind;
  const preferNewFinancial=isIncomeKind(o.kind)&&!isIncomeKind(old.kind);
  const preferOldFinancial=isIncomeKind(old.kind)&&!isIncomeKind(o.kind);
  const money=(field,abs=false)=>{
    const oldVal=Number(old[field])||0, newVal=Number(o[field])||0;
    if(!hasOld)return abs?Math.abs(newVal):newVal;
    if(sameKind)return abs?Math.abs(oldVal)+Math.abs(newVal):oldVal+newVal;
    if(preferNewFinancial)return abs?Math.abs(newVal):newVal;
    if(preferOldFinancial)return abs?Math.abs(oldVal):oldVal;
    const chosen=Math.abs(newVal)>Math.abs(oldVal)?newVal:oldVal;
    return abs?Math.abs(chosen):chosen;
  };
  const oldNames=old.importedProductNames||[];
  const oldSkus=old.importedSkus||[];
  const names=[...oldNames,old.productName,o.productName].filter(Boolean);
  const skus=[...oldSkus,old.sku,o.sku].filter(Boolean);
  map[o.id]=Object.assign({},old,o,{
    id:o.id,
    productName:[...new Set(names)].join(' + '),
    sku:[...new Set(skus)].filter(Boolean)[0]||'',
    importedProductNames:[...new Set(names)],
    importedSkus:[...new Set(skus)],
    gross:money('gross'),
    net:money('net'),
    fees:money('fees',true),
    shipping:money('shipping'),
    discounts:money('discounts',true),
    qty:sameKind?(old.qty||0)+(o.qty||1):Math.max(old.qty||1,o.qty||1),
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
        itemId:cleanCell(pick(obj,['SKU'])).replace(/^[-/]$/,''),
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
      itemId:cleanCell(pick(obj,['Número de referência SKU','Nº de referência do SKU principal'])),
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
      marketplace:'tiktokShop',store:'TikTok Shop',kind:'TikTok Orders',orderId,
      date:parseDateFlexible(pick(obj,['Created Time','Paid Time'])),
      paidDate:parseDateFlexible(pick(obj,['Paid Time'])),
      sku:cleanCell(pick(obj,['Seller SKU','SKU ID'])),
      itemId:cleanCell(pick(obj,['SKU ID'])),
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
  const statements={};
  sheets.forEach(sheet=>{
    const hi=findHeaderRow(sheet.rows,['id do demonstrativo','valor total a ser liquidado']);
    if(hi<0)return;
    const headers=sheet.rows[hi];
    sheet.rows.slice(hi+1).forEach(row=>{
      const obj=rowToObj(headers,row);
      const statementId=cleanCell(pick(obj,['ID do demonstrativo']));
      if(!statementId)return;
      statements[statementId]={
        statementId,
        paymentId:cleanCell(pick(obj,['ID do pagamento'])),
        status:cleanCell(pick(obj,['Status']))||'Pagos',
        date:parseDateFlexible(pick(obj,['Data do demonstrativo'])),
        net:parseMoneyBR(pick(obj,['Valor total a ser liquidado','Valor do pagamento'])),
        gross:parseMoneyBR(pick(obj,['Vendas líquidas','Vendas liquidas','Subtotal do item antes dos descontos'])),
        fees:Math.abs(parseMoneyBR(pick(obj,['Taxas']))),
        shipping:parseMoneyBR(pick(obj,['Frete','Custo líquido de frete'])),
        discounts:Math.abs(parseMoneyBR(pick(obj,['Ajustes','Descontos'])))
      };
    });
  });

  const detailRows=[];
  sheets.forEach(sheet=>{
    const hi=findHeaderRow(sheet.rows,['id do demonstrativo','nome do produto']);
    if(hi<0)return;
    const headers=sheet.rows[hi];
    sheet.rows.slice(hi+1).forEach(row=>{
      const obj=rowToObj(headers,row);
      const statementId=cleanCell(pick(obj,['ID do demonstrativo']));
      const productName=cleanCell(pick(obj,['Nome do produto']));
      const orderId=cleanCell(pick(obj,['ID do pedido/ajuste','ID do pedido','ID do pedido ajuste']));
      if(!statementId||!productName||!orderId)return;
      detailRows.push({
        statementId,orderId,productName,
        date:parseDateFlexible(pick(obj,['Data de criação do pedido','Data de criacao do pedido','Data do demonstrativo'])),
        qty:parseInt(pick(obj,['Quantidade']))||1,
        sku:cleanCell(pick(obj,['ID do SKU','SKU ID'])),
        itemId:cleanCell(pick(obj,['ID do SKU','SKU ID'])),
        status:cleanCell(pick(obj,['Status']))||statements[statementId]?.status||'Pagos'
      });
    });
  });
  if(detailRows.length){
    const totals={};
    detailRows.forEach(r=>{totals[r.statementId]=(totals[r.statementId]||0)+r.qty;});
    detailRows.forEach(r=>{
      const st=statements[r.statementId]||{};
      const share=(r.qty||1)/(totals[r.statementId]||1);
      out.push({
        marketplace:'tiktokShop',store:'TikTok Shop',kind:'TikTok Income',orderId:r.orderId,
        date:r.date||st.date,paidDate:st.date,
        sku:r.sku,itemId:r.itemId,productName:r.productName,qty:r.qty,
        gross:(st.gross||0)*share,net:(st.net||0)*share,fees:(st.fees||0)*share,
        shipping:(st.shipping||0)*share,discounts:(st.discounts||0)*share,
        status:r.status,sourceFiles:[fileName]
      });
    });
    return out.filter(o=>o.productName||o.net||o.gross);
  }

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
        marketplace:'tiktokShop',store:'TikTok Shop',kind:'TikTok Income',orderId,
        date:parseDateFlexible(pick(obj,['Data do demonstrativo','Data de início do pagamento','Data de criacao do pedido'])),
        paidDate:parseDateFlexible(pick(obj,['Data de conclusão do pagamento'])),
        sku:cleanCell(pick(obj,['SKU ID','ID do SKU'])),
        itemId:cleanCell(pick(obj,['ID do pedido','ID do SKU','SKU ID'])),
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
function findProductMatch(o){
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  const rule=productMatchRules[slug(o.sku||o.productName)]||productMatchRules[slug(o.itemId||'')]||productMatchRules[slug(o.productName)];
  if(rule){
    const p=catalog.find(x=>x.id===rule.productId||x.legacyProductId===rule.productId);
    if(p)return {product:p,confidence:1,method:'manual'};
  }
  const sku=norm(o.sku);
  if(sku){
    const p=catalog.find(x=>norm(x.sku)===sku);
    if(p)return {product:p,confidence:.98,method:'sku'};
  }
  const item=norm(o.itemId);
  if(item){
    const p=catalog.find(x=>Object.values(x.links||{}).some(link=>norm(link).includes(item))||norm(x.sku).includes(item));
    if(p)return {product:p,confidence:.9,method:'itemId/link'};
  }
  const on=norm(o.productName);
  if(!on)return {product:null,confidence:0,method:'none'};
  for(const p of catalog){
    const aliases=[...(p.aliases||[]),p.name].map(norm);
    if(aliases.some(a=>a&&on.includes(a)))return {product:p,confidence:.86,method:'alias'};
  }
  let best=null,bestScore=0;
  catalog.forEach(p=>{
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
  return bestScore>=35?{product:best,confidence:Math.min(.84,bestScore/100),method:'nome parecido'}:{product:null,confidence:0,method:'none'};
}
function findProductForOrder(o){
  return findProductMatch(o).product;
}
function productUnitCost(product,meiUnit=0){
  const s=getMarketplaceSettings(), cfg=getLocalConfig();
  if(!product)return {total:meiUnit,filament:0,energy:0,packaging:0,bubble:0,label:0,maintenance:0,waste:0,postProcess:0,nozzleWear:0,extra:0,mei:meiUnit,other:0};
  const pc=product.costs||{};
  const weight=Number(pc.weightGrams ?? product.weight)||0;
  const printH=Number(pc.printTimeHours ?? product.printH)||0;
  const filament=weight/1000*(Number(pc.filamentCostPerKg)||s.filamentKgPrice||95);
  const energy=printH*(Number(pc.energyKwhPerHour)||s.printerKwhHour||0.1)*(Number(pc.energyCostPerKwh)||s.energyKwhPrice||1.05);
  const packaging=Number(pc.packaging ?? product.breakdown?.pkg)||0;
  const bubble=Number(pc.bubbleWrap ?? product.breakdown?.bubble)||cfg.bubbleUnit||0;
  const label=Number(pc.label)||s.labelCost||0;
  const maintenance=Number(pc.maintenance ?? product.breakdown?.maint)||cfg.maintenance||0;
  const extra=Number(pc.extraCost)||s.otherProductionCost||0;
  const postProcess=Number(pc.postProcessCost)||s.postProcessCost||0;
  const nozzleWear=Number(pc.nozzleWearCost)||s.nozzleWearCost||0;
  const base=filament+energy+packaging+bubble+label+maintenance+extra+postProcess+nozzleWear;
  const loss=base*((Number(pc.wastePercent)||s.lossPct||0)/100);
  const other=label+maintenance+loss+extra+postProcess+nozzleWear;
  return {total:base+loss+meiUnit,filament,energy,packaging,bubble,label,maintenance,waste:loss,postProcess,nozzleWear,extra,mei:meiUnit,other};
}
function enrichOrder(o,meiUnit=0){
  o=Object.assign({},o,{marketplace:canonicalMarketplace(o.marketplace),orderId:o.orderId||o.id,productName:o.productName||o.importedProductName,sku:o.sku||o.importedSku,qty:o.qty||o.quantity||1,date:o.date||o.orderDate,paymentDate:o.paymentDate||o.paidDate});
  const match=findProductMatch(o);
  const product=match.product;
  const c=productUnitCost(product,meiUnit);
  const qty=Number(o.qty)||1;
  const fin=o.financial||{};
  const gross=Number(fin.grossAmount ?? o.gross)||0;
  const fees=Math.abs(Number(fin.commissionFee||0)+Number(fin.fixedFee||0)+Number(fin.transactionFee||0)+Number(fin.affiliateFee||0)+Number(fin.otherFees||0))||Math.abs(Number(o.fees)||0);
  const discounts=Math.abs(Number(fin.marketplaceDiscount||0)+Number(fin.sellerDiscount||0))||Math.abs(Number(o.discounts)||0);
  const shipping=Number(fin.shippingCostToSeller ?? o.shipping)||0;
  const net=Number(fin.netReceived ?? o.net)||(gross-fees-discounts-Math.abs(shipping));
  const cost=c.total*qty;
  const profit=net-cost;
  const margin=net>0?profit/net*100:0;
  const health=profit<0?'prejuizo':margin<15?'margem_baixa':margin<25?'atenção':'healthy';
  return Object.assign({},o,{
    id:o.id||makeImportId(o),
    orderDate:o.date||o.orderDate||'',
    paymentDate:o.paymentDate||o.paidDate||o.date||'',
    importedProductName:o.productName||'',
    importedSku:o.sku||'',
    importedItemId:o.itemId||o.importedItemId||'',
    linkedProductId:product?.id||'',
    linkedProductName:product?.name||'',
    linkConfidence:match.confidence,
    linkMethod:match.method,
    quantity:qty,
    financial:{
      grossAmount:gross,marketplaceDiscount:Math.abs(Number(fin.marketplaceDiscount)||0),
      sellerDiscount:Math.abs(Number(fin.sellerDiscount)||0),shippingCharged:Number(fin.shippingCharged)||0,
      shippingCostToSeller:shipping,commissionFee:Number(fin.commissionFee)||0,fixedFee:Number(fin.fixedFee)||0,
      transactionFee:Number(fin.transactionFee)||0,affiliateFee:Number(fin.affiliateFee)||0,
      otherFees:fees,netReceived:net
    },
    productionCost:{
      filament:c.filament,energy:c.energy,packaging:c.packaging,bubbleWrap:c.bubble,
      label:c.label,maintenance:c.maintenance,waste:c.waste,postProcess:c.postProcess,
      nozzleWear:c.nozzleWear,extraCost:c.extra,meiAllocated:c.mei,totalUnitCost:c.total,totalCost:cost
    },
    profit:{netProfit:profit,marginPercent:margin,status:health},
    unitCost:c.total,totalCost:cost,net,profit,margin,
    health:profit<0?'bad':margin<15?'warn':'ok',
    costParts:c,
    updatedAt:new Date().toISOString()
  });
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
    if(analysisFilters.marketplace!=='all'&&marketplaceFilterValue(o.marketplace)!==analysisFilters.marketplace)return false;
    if(analysisFilters.store!=='all'&&o.store!==analysisFilters.store)return false;
    if(analysisFilters.health==='bad'&&o.health!=='bad')return false;
    if(analysisFilters.health==='unlinked'&&o.linkedProductId)return false;
    const q=norm(analysisFilters.q);
    if(q&&!norm(`${o.orderId} ${o.productName} ${o.linkedProductName}`).includes(q))return false;
    if(!isInsidePeriod(o))return false;
    return true;
  });
}
function isInsidePeriod(o){
  const d=parseDateFlexible(o.paymentDate||o.paidDate||o.date||o.orderDate);
  if(!d)return true;
  const day=new Date(d+'T00:00:00');
  const today=new Date(); today.setHours(0,0,0,0);
  if(analysisFilters.period==='today')return day.getTime()===today.getTime();
  if(analysisFilters.period==='7d')return day>=new Date(today.getTime()-6*86400000);
  if(analysisFilters.period==='30d')return day>=new Date(today.getTime()-29*86400000);
  if(analysisFilters.period==='custom'){
    const from=analysisFilters.from?new Date(analysisFilters.from+'T00:00:00'):null;
    const to=analysisFilters.to?new Date(analysisFilters.to+'T23:59:59'):null;
    return (!from||day>=from)&&(!to||day<=to);
  }
  return true;
}
function summarizeOrders(orders){
  return orders.reduce((a,o)=>{
    const qty=Number(o.qty)||1;
    a.gross+=Number(o.gross)||Number(o.financial?.grossAmount)||0;a.net+=Number(o.net)||Number(o.financial?.netReceived)||0;
    a.fees+=Math.abs(Number(o.financial?.otherFees ?? o.fees)||0);
    a.cost+=Number(o.totalCost)||0;a.profit+=Number(o.profit)||0;a.qty+=qty;
    a.energy+=Number((o.productionCost?.energy ?? o.costParts?.energy)||0)*qty;
    a.filament+=Number((o.productionCost?.filament ?? o.costParts?.filament)||0)*qty;
    a.packaging+=(Number(o.costParts?.packaging||0)+Number(o.costParts?.bubble||0))*qty;
    a.mei+=Number((o.productionCost?.meiAllocated ?? o.costParts?.mei)||0)*qty;
    a.other+=Number(o.costParts?.other ?? (
      Number(o.productionCost?.label||0)+Number(o.productionCost?.maintenance||0)+
      Number(o.productionCost?.waste||0)+Number(o.productionCost?.postProcess||0)+
      Number(o.productionCost?.nozzleWear||0)+Number(o.productionCost?.extraCost||0)
    ))*qty;
    if(marketplaceFilterValue(o.marketplace)==='shopee')a.shopeeNet+=Number(o.net)||0,a.shopeeProfit+=Number(o.profit)||0;
    if(marketplaceFilterValue(o.marketplace)==='tiktokShop')a.tiktokNet+=Number(o.net)||0,a.tiktokProfit+=Number(o.profit)||0;
    if(o.store==='_kaline98')a.kalineNet+=Number(o.net)||0;
    if(o.store==='mateusoliver98')a.mateusNet+=Number(o.net)||0;
    if(!o.linkedProductId)a.unlinked++; if(o.health==='bad')a.bad++;
    if(!/pago|conclu/i.test(o.status||''))a.pending++;
    return a;
  },{gross:0,net:0,fees:0,cost:0,profit:0,qty:0,energy:0,filament:0,packaging:0,mei:0,other:0,shopeeNet:0,tiktokNet:0,kalineNet:0,mateusNet:0,shopeeProfit:0,tiktokProfit:0,unlinked:0,bad:0,pending:0});
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
  return `<div class="analysis-shell">
    <div class="period-row">
      ${['today','7d','30d','custom'].map(p=>`<button class="${analysisFilters.period===p?'on':''}" onclick="analysisFilters.period='${p}';renderAnalysis()">${p==='today'?'Hoje':p==='7d'?'7 dias':p==='30d'?'30 dias':'Período'}</button>`).join('')}
    </div>
    ${analysisFilters.period==='custom'?`<div class="filter-row"><input class="fi" type="date" value="${analysisFilters.from}" onchange="analysisFilters.from=this.value;renderAnalysis()"><input class="fi" type="date" value="${analysisFilters.to}" onchange="analysisFilters.to=this.value;renderAnalysis()"></div>`:''}
    <div class="filter-row">
      <select class="fi" onchange="analysisFilters.store=this.value;renderAnalysis()">${stores.map(x=>`<option value="${x}" ${analysisFilters.store===x?'selected':''}>${x==='all'?'Todas as lojas':x}</option>`).join('')}</select>
      <select class="fi" onchange="analysisFilters.marketplace=this.value;renderAnalysis()"><option value="all" ${analysisFilters.marketplace==='all'?'selected':''}>Todos marketplaces</option><option value="shopee" ${analysisFilters.marketplace==='shopee'?'selected':''}>Shopee</option><option value="tiktokShop" ${analysisFilters.marketplace==='tiktokShop'?'selected':''}>TikTok Shop</option></select>
    </div>
    <section class="dash-section"><h3>Dinheiro que entrou <span>${orders.length} pedidos</span></h3><div class="finance-grid">
      ${moneyCard('Recebido Shopee _kaline98',brl(s.kalineNet),'in')}
      ${moneyCard('Recebido Shopee mateusoliver98',brl(s.mateusNet),'in')}
      ${moneyCard('Recebido TikTok Shop',brl(s.tiktokNet),'in')}
      ${moneyCard('Recebido Total',brl(s.net),'in','liquido já descontado')}
    </div></section>
    <section class="dash-section"><h3>Dinheiro que saiu <span>custos reais estimados</span></h3><div class="finance-grid">
      ${moneyCard('Taxas marketplace',brl(s.fees),'out')}
      ${moneyCard('Filamento',brl(s.filament),'out')}
      ${moneyCard('Energia',brl(s.energy),'out')}
      ${moneyCard('Embalagens',brl(s.packaging),'out')}
      ${moneyCard('DAS MEI',brl(s.mei),'out')}
      ${moneyCard('Outros custos',brl(s.other),'out')}
    </div></section>
    <section class="dash-section"><h3>Dinheiro que sobrou <span>resultado operacional</span></h3><div class="finance-grid">
      ${moneyCard('Lucro líquido',brl(s.profit),s.profit>=0?'net':'risk')}
      ${moneyCard('Margem líquida',pct(margin),margin>=15?'net':'risk')}
      ${moneyCard('Lucro Shopee',brl(s.shopeeProfit),s.shopeeProfit>=0?'net':'risk')}
      ${moneyCard('Lucro TikTok Shop',brl(s.tiktokProfit),s.tiktokProfit>=0?'net':'risk')}
    </div></section>
    <section class="dash-section"><h3>Gráficos <span>visão diária e composição</span></h3>
      <div class="chart-card"><div class="chart-title">Recebimento por dia</div>${renderLineChart(dailySeries(orders,'net'),'#00CFFF')}</div>
      <div class="chart-card"><div class="chart-title">Lucro líquido por dia</div>${renderLineChart(dailySeries(orders,'profit'),'#00E676')}</div>
      <div class="chart-card"><div class="chart-title">Shopee vs TikTok Shop</div>${renderBarCompare([{label:'Shopee',value:s.shopeeProfit,color:'#00CFFF'},{label:'TikTok',value:s.tiktokProfit,color:'#FF8C00'}])}</div>
      <div class="chart-card"><div class="chart-title">Composição dos custos</div>${renderCostDonut(s)}</div>
    </section>
    <section class="dash-section"><h3>Alertas <span>ações prioritárias</span></h3>${renderAlerts(orders)}</section>
    ${renderProductSummary(orders)}
    <button class="btn btn-export" onclick="exportAnalysisCsv()">Exportar CSV da análise</button>
  </div>`;
}
function moneyCard(label,value,cls,sub=''){
  return `<div class="money-card ${cls}"><div class="tag">${label}</div><div class="amount">${value}</div>${sub?`<div class="sub">${sub}</div>`:''}</div>`;
}
function dailySeries(orders,field){
  const map={};
  orders.forEach(o=>{
    const d=parseDateFlexible(o.paymentDate||o.paidDate||o.date||o.orderDate)||'sem data';
    map[d]=(map[d]||0)+(Number(o[field])||0);
  });
  return Object.keys(map).sort().slice(-14).map(k=>({label:k.slice(5),value:map[k]}));
}
function renderLineChart(data,color){
  if(!data.length)return '<div class="muted-note">Sem dados no período.</div>';
  const w=320,h=130,p=18,max=Math.max(...data.map(d=>Math.abs(d.value)),1);
  const pts=data.map((d,i)=>{
    const x=p+(data.length===1?0:i*(w-p*2)/(data.length-1));
    const y=h-p-((d.value+max)/(max*2))*(h-p*2);
    return `${x},${y}`;
  }).join(' ');
  return `<svg class="chart-svg" viewBox="0 0 ${w} ${h}" aria-label="gráfico de linha">
    <line x1="${p}" y1="${h/2}" x2="${w-p}" y2="${h/2}" stroke="#263044"/>
    <polyline fill="none" stroke="${color}" stroke-width="3" points="${pts}"/>
    ${data.map((d,i)=>`<text x="${p+i*(w-p*2)/Math.max(1,data.length-1)}" y="${h-3}" fill="#8899BB" font-size="9" text-anchor="middle">${d.label}</text>`).join('')}
  </svg>`;
}
function renderBarCompare(items){
  const max=Math.max(...items.map(i=>Math.abs(i.value)),1);
  return `<svg class="chart-svg" viewBox="0 0 320 140">${items.map((it,i)=>{
    const x=45+i*130, bw=70, bh=Math.abs(it.value)/max*85, y=105-bh;
    return `<rect x="${x}" y="${y}" width="${bw}" height="${bh}" rx="5" fill="${it.color}"/><text x="${x+bw/2}" y="124" fill="#8899BB" font-size="11" text-anchor="middle">${it.label}</text><text x="${x+bw/2}" y="${Math.max(14,y-6)}" fill="#fff" font-size="11" text-anchor="middle">${brl(it.value)}</text>`;
  }).join('')}</svg>`;
}
function renderCostDonut(s){
  const parts=[
    {label:'Filamento',value:s.filament,color:'#00CFFF'},
    {label:'Energia',value:s.energy,color:'#FF8C00'},
    {label:'Embalagem',value:s.packaging,color:'#7B2FFF'},
    {label:'Marketplace',value:s.fees,color:'#FF3B8B'},
    {label:'MEI',value:s.mei,color:'#00E676'}
  ];
  const total=parts.reduce((t,p)=>t+p.value,0)||1;
  let acc=0;
  const circles=parts.map(p=>{
    const dash=p.value/total*100, off=25-acc; acc+=dash;
    return `<circle cx="60" cy="60" r="42" fill="none" stroke="${p.color}" stroke-width="16" stroke-dasharray="${dash} ${100-dash}" stroke-dashoffset="${off}"/>`;
  }).join('');
  return `<div class="donut-wrap"><svg width="120" height="120" viewBox="0 0 120 120">${circles}<circle cx="60" cy="60" r="25" fill="#0F1522"/><text x="60" y="64" fill="#fff" font-size="12" text-anchor="middle">${brl(total)}</text></svg><div class="legend">${parts.map(p=>`<div><span style="background:${p.color}"></span>${p.label}: <b>${brl(p.value)}</b></div>`).join('')}</div></div>`;
}
function renderAlerts(orders){
  const low=orders.filter(o=>o.health==='warn').length;
  const bad=orders.filter(o=>o.health==='bad').length;
  const unlinked=orders.filter(o=>!o.linkedProductId).length;
  const best=[...orders].sort((a,b)=>b.profit-a.profit)[0];
  return `<div class="alert-list">
    <div class="alert-item"><b>Produtos com margem baixa</b><span>${low}</span></div>
    <div class="alert-item"><b>Produtos com prejuízo</b><span>${bad}</span></div>
    <div class="alert-item"><b>Produtos sem vínculo</b><span>${unlinked}</span></div>
    <div class="alert-item"><b>Produto mais lucrativo</b><span>${best?esc(best.linkedProductName||best.productName):'sem dados'}</span></div>
  </div>`;
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
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  openSht('Vincular produto',`<div class="muted-note">Produto no relatório:<br><b>${esc(o.productName)}</b><br>SKU/item: ${esc(o.sku||o.itemId||'-')}</div><select class="fi" id="link-product">${catalog.map(p=>`<option value="${p.id}" ${p.id===o.linkedProductId?'selected':''}>${esc(p.name)}${p.sku?' • '+esc(p.sku):''}</option>`).join('')}</select><button class="btn btn-save" onclick="saveProductLink('${o.id}')">Salvar vínculo permanente</button><button class="btn btn-secondary" onclick="closeSht()">Cancelar</button>`);
}
function saveProductLink(orderId){
  const o=marketplaceOrders.find(x=>x.id===orderId)||importReview?.orders.find(x=>x.id===orderId); if(!o)return;
  const productId=document.getElementById('link-product')?.value; if(!productId)return;
  const keys=[o.sku,o.itemId,o.productName].filter(Boolean).map(slug);
  keys.forEach(key=>{productMatchRules[key]={productId,source:o.productName||o.sku||o.itemId,savedTs:Date.now()};});
  if(matchRulesRef){
    const updates={}; keys.forEach(key=>updates[key]=sanitizeForFirebase(productMatchRules[key]));
    matchRulesRef.update(updates).catch(()=>{});
  }
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
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  if(!catalog.length)return '<div class="empty"><div class="ei">🧮</div><div>Cadastre produtos para simular preços profissionais.</div></div>';
  return catalog.slice(0,20).map(p=>{
    const c=productUnitCost(p,0), min=priceForMargin(c.total,'shopee',0), sh10=priceForMargin(c.total,'shopee',10), sh30=priceForMargin(c.total,'shopee',30), tk30=priceForMargin(c.total,'tiktok',30);
    const current=Number(p.marketplaceSettings?.shopee?.salePrice ?? p.price)||0, health=current&&current<min?'bad':current&&current<sh10?'warn':'ok';
    const weight=p.costs?.weightGrams??p.weight??0, hours=p.costs?.printTimeHours??p.printH??0;
    return `<div class="order-card"><div class="order-top"><div><div class="order-title">${esc(p.name)}</div><div class="order-meta">SKU ${esc(p.sku||'-')} • Custo técnico ${brl(c.total)} • peso ${weight}g • ${hours}h</div></div><span class="health ${health}">${health==='bad'?'abaixo do mínimo':health==='warn'?'margem baixa':'preço saudável'}</span></div>
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
