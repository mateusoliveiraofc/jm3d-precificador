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

/* Final overrides: product recognition must prioritize real marketplace identity before names. */
function findProductMatch(o){
  const catalog=catalogProducts();
  for(const key of manualRuleKeys(o)){
    const rule=productMatchRules[key];
    if(!rule)continue;
    const p=catalog.find(x=>x.id===rule.productId||x.legacyProductId===rule.productId);
    if(p)return {product:p,confidence:1,method:'regra permanente'};
  }
  const urlInfo=extractOrderUrlInfo(o);
  const itemId=cleanCell(o.itemId||o.importedItemId||urlInfo.itemId);
  if(itemId){
    const p=catalog.find(x=>productHasItemId(x,itemId));
    if(p)return {product:p,confidence:.99,method:'itemId'};
  }
  const sku=norm(o.sku||o.importedSku);
  if(sku){
    const p=catalog.find(x=>norm(x.sku)===sku);
    if(p)return {product:p,confidence:.98,method:'sku'};
  }
  const importedUrl=cleanCell(o.importedProductUrl||o.productUrl||urlInfo.url);
  if(importedUrl){
    const p=catalog.find(x=>productHasExactUrl(x,importedUrl));
    if(p)return {product:p,confidence:.96,method:'link do anúncio'};
  }
  const shopId=cleanCell(o.shopId||o.importedShopId||urlInfo.shopId);
  if(shopId&&itemId){
    const p=catalog.find(x=>productHasShopItem(x,shopId,itemId));
    if(p)return {product:p,confidence:.95,method:'shopId + itemId'};
  }
  const on=norm(o.productName||o.importedProductName);
  if(on){
    const exact=catalog.find(p=>norm(p.name)===on);
    if(exact)return {product:exact,confidence:.88,method:'nome exato'};
    for(const p of catalog){
      const aliases=(p.aliases||[]).map(norm).filter(Boolean);
      if(aliases.some(a=>a===on||on.includes(a)||a.includes(on)))return {product:p,confidence:.84,method:'alias'};
    }
  }
  const importedPhoto=o.importedPhotoUrl||o.photoUrl||'';
  if(importedPhoto){
    const sig=photoSignal(importedPhoto);
    const p=catalog.find(x=>{
      const photo=x.photoUrl||x.photo||'';
      return photo&&((photo===importedPhoto)||photoSignal(photo)===sig);
    });
    if(p)return {product:p,confidence:.76,method:'foto semelhante'};
  }
  if(!on)return {product:null,confidence:0,method:'none'};
  let best=null,bestScore=0;
  catalog.forEach(p=>{
    const pn=norm(p.name);
    let score=on.includes(pn)||pn.includes(on)?58:0;
    if(!score){
      const a=new Set(on.split(' ').filter(x=>x.length>2));
      const b=new Set(pn.split(' ').filter(x=>x.length>2));
      const inter=[...a].filter(x=>b.has(x)).length;
      score=b.size?inter/b.size*52:0;
    }
    if(score>bestScore){best=p;bestScore=score;}
  });
  return bestScore>=42?{product:best,confidence:Math.min(.70,bestScore/100),method:'nome parecido'}:{product:null,confidence:0,method:'none'};
}
function openLinkProductLegacy(orderId){
  const o=analyzedOrders().find(x=>x.id===orderId)||marketplaceOrders.find(x=>x.id===orderId)||importReview?.orders.find(x=>x.id===orderId);
  if(!o)return;
  const catalog=catalogProducts();
  const suggestions=[...catalog].map(p=>Object.assign({score:linkSuggestionScore(o,p)},p)).sort((a,b)=>b.score-a.score).slice(0,8);
  openSht('Vincular produto',`<div class="muted-note">Produto importado:<br><b>${esc(o.productName||o.importedProductName||'Produto sem nome')}</b><br>${marketplaceLabel(o.marketplace)} • ${esc(o.store)} • Pedido ${esc(o.orderId)}<br>SKU ${esc(o.sku||o.importedSku||'-')} • Item ${esc(o.itemId||o.importedItemId||'-')} • Recebido ${brl(o.net||0)}</div>
    <div class="link-suggestion-grid">${suggestions.map(p=>`<button class="link-suggestion" onclick="saveProductLink('${ea(orderId)}','${ea(p.id)}')">${p.photoUrl||p.photo?`<img src="${ea(p.photoUrl||p.photo)}">`:'<div class="ph">📦</div>'}<b>${esc(p.name)}</b><span>SKU ${esc(p.sku||'-')} • Item ${esc(firstMarketplaceId(p)||'-')}<br>${p.score?Math.round(p.score)+'% de confiança':'catálogo'}</span></button>`).join('')}</div>
    <select class="fi" id="link-product">${catalog.map(p=>`<option value="${ea(p.id)}" ${p.id===o.linkedProductId?'selected':''}>${esc(p.name)}${p.sku?' • '+esc(p.sku):''}</option>`).join('')}</select>
    <button class="btn btn-save" onclick="saveProductLink('${ea(orderId)}')">Salvar vínculo permanente</button><button class="btn btn-secondary" onclick="openStoreProductImport()">Cadastrar produto da loja</button><button class="btn btn-secondary" onclick="closeSht()">Cancelar</button>`);
}
function saveProductLink(orderId,productId){
  const o=marketplaceOrders.find(x=>x.id===orderId)||importReview?.orders.find(x=>x.id===orderId)||analyzedOrders().find(x=>x.id===orderId);
  if(!o)return;
  const selected=productId||document.getElementById('link-product')?.value; if(!selected)return;
  const p=catalogProducts().find(x=>x.id===selected||x.legacyProductId===selected);
  const keys=[...new Set(manualRuleKeys(o))];
  keys.forEach(key=>{productMatchRules[key]={productId:selected,productName:p?.name||'',source:o.productName||o.sku||o.itemId,savedTs:Date.now()};});
  o.linkedProductId=selected;
  o.linkedProductName=p?.name||o.linkedProductName||'';
  o.linkConfidence=1;
  o.linkMethod='manual permanente';
  if(matchRulesRef){
    const updates={}; keys.forEach(key=>updates[key]=sanitizeForFirebase(productMatchRules[key]));
    matchRulesRef.update(updates).catch(()=>{});
  }else{
    localStorage.setItem('jm3d_match_rules',JSON.stringify(productMatchRules));
  }
  if(marketplaceOrdersRef&&marketplaceOrders.some(x=>x.id===orderId)){
    marketplaceOrdersRef.child(orderId).update(sanitizeForFirebase({
      linkedProductId:selected,
      linkedProductName:p?.name||'',
      linkConfidence:1,
      linkMethod:'manual permanente',
      updatedAt:new Date().toISOString()
    })).catch(()=>{});
  }else{
    localStorage.setItem('jm3d_market_orders',JSON.stringify(marketplaceOrders));
  }
  closeSht();renderContent();
}

function catalogProductIds(p){
  return Object.values(p?.marketplaceIds||{}).map(cleanCell).filter(Boolean);
}
function catalogProductShopIds(p){
  return Object.values(p?.shopIds||{}).map(cleanCell).filter(Boolean);
}
function catalogProductLinks(p){
  return Object.values(p?.links||{}).map(cleanCell).filter(Boolean);
}
function extractOrderUrlInfo(o){
  return marketplaceProductUrlInfo(o.importedProductUrl||o.productUrl||o.link||'');
}
function productHasItemId(p,itemId){
  const item=norm(itemId);
  if(!item)return false;
  if(catalogProductIds(p).some(id=>norm(id)===item))return true;
  return catalogProductLinks(p).some(link=>norm(link).includes(item));
}
function productHasShopItem(p,shopId,itemId){
  const shop=norm(shopId), item=norm(itemId);
  if(!shop||!item)return false;
  return catalogProductShopIds(p).some(id=>norm(id)===shop)&&productHasItemId(p,item);
}
function productHasExactUrl(p,url){
  const u=norm(url);
  return !!u&&catalogProductLinks(p).some(link=>norm(link)===u);
}
function manualRuleKeys(o){
  const urlInfo=extractOrderUrlInfo(o);
  const keys=[
    o.itemId,o.importedItemId,urlInfo.itemId,
    o.sku,o.importedSku,
    o.importedProductUrl,o.productUrl,
    urlInfo.shopId&&urlInfo.itemId?`${urlInfo.shopId}_${urlInfo.itemId}`:'',
    o.productName,o.importedProductName
  ];
  return keys.filter(Boolean).map(slug);
}
function findProductMatch(o){
  const catalog=catalogProducts();
  for(const key of manualRuleKeys(o)){
    const rule=productMatchRules[key];
    if(!rule)continue;
    const p=catalog.find(x=>x.id===rule.productId||x.legacyProductId===rule.productId);
    if(p)return {product:p,confidence:1,method:'regra permanente'};
  }
  const urlInfo=extractOrderUrlInfo(o);
  const itemId=cleanCell(o.itemId||o.importedItemId||urlInfo.itemId);
  if(itemId){
    const p=catalog.find(x=>productHasItemId(x,itemId));
    if(p)return {product:p,confidence:.99,method:'itemId'};
  }
  const sku=norm(o.sku||o.importedSku);
  if(sku){
    const p=catalog.find(x=>norm(x.sku)===sku);
    if(p)return {product:p,confidence:.98,method:'sku'};
  }
  const importedUrl=cleanCell(o.importedProductUrl||o.productUrl||urlInfo.url);
  if(importedUrl){
    const p=catalog.find(x=>productHasExactUrl(x,importedUrl));
    if(p)return {product:p,confidence:.96,method:'link do anúncio'};
  }
  const shopId=cleanCell(o.shopId||o.importedShopId||urlInfo.shopId);
  if(shopId&&itemId){
    const p=catalog.find(x=>productHasShopItem(x,shopId,itemId));
    if(p)return {product:p,confidence:.95,method:'shopId + itemId'};
  }
  const on=norm(o.productName||o.importedProductName);
  if(on){
    const exact=catalog.find(p=>norm(p.name)===on);
    if(exact)return {product:exact,confidence:.88,method:'nome exato'};
    for(const p of catalog){
      const aliases=(p.aliases||[]).map(norm).filter(Boolean);
      if(aliases.some(a=>a===on||on.includes(a)||a.includes(on)))return {product:p,confidence:.84,method:'alias'};
    }
  }
  const importedPhoto=o.importedPhotoUrl||o.photoUrl||'';
  if(importedPhoto){
    const sig=photoSignal(importedPhoto);
    const p=catalog.find(x=>{
      const photo=x.photoUrl||x.photo||'';
      return photo&&((photo===importedPhoto)||photoSignal(photo)===sig);
    });
    if(p)return {product:p,confidence:.76,method:'foto semelhante'};
  }
  if(!on)return {product:null,confidence:0,method:'none'};
  let best=null,bestScore=0;
  catalog.forEach(p=>{
    const pn=norm(p.name);
    let score=on.includes(pn)||pn.includes(on)?58:0;
    if(!score){
      const a=new Set(on.split(' ').filter(x=>x.length>2));
      const b=new Set(pn.split(' ').filter(x=>x.length>2));
      const inter=[...a].filter(x=>b.has(x)).length;
      score=b.size?inter/b.size*52:0;
    }
    if(score>bestScore){best=p;bestScore=score;}
  });
  return bestScore>=42?{product:best,confidence:Math.min(.70,bestScore/100),method:'nome parecido'}:{product:null,confidence:0,method:'none'};
}
function linkSuggestionScore(order,p){
  const direct=findProductMatch(Object.assign({},order,{itemId:order.itemId||order.importedItemId,sku:order.sku||order.importedSku}));
  if(direct.product&&(direct.product.id===p.id||direct.product.legacyProductId===p.id))return direct.confidence*100;
  const on=norm(order.productName||order.importedProductName), pn=norm(p.name);
  const alias=(p.aliases||[]).some(a=>on&&norm(a)&&on.includes(norm(a)));
  return alias?70:(on&&pn&&on.includes(pn)?55:0);
}
function openLinkProduct(orderId){
  const o=analyzedOrders().find(x=>x.id===orderId)||marketplaceOrders.find(x=>x.id===orderId)||importReview?.orders.find(x=>x.id===orderId);
  if(!o)return;
  const catalog=catalogProducts();
  const suggestions=[...catalog].map(p=>Object.assign({score:linkSuggestionScore(o,p)},p)).sort((a,b)=>b.score-a.score).slice(0,8);
  openSht('Vincular produto',`<div class="muted-note">Produto importado:<br><b>${esc(o.productName||o.importedProductName||'Produto sem nome')}</b><br>${marketplaceLabel(o.marketplace)} • ${esc(o.store)} • Pedido ${esc(o.orderId)}<br>SKU ${esc(o.sku||o.importedSku||'-')} • Item ${esc(o.itemId||o.importedItemId||'-')} • Recebido ${brl(o.net||0)}</div>
    <div class="link-suggestion-grid">${suggestions.map(p=>`<button class="link-suggestion" onclick="saveProductLink('${ea(orderId)}','${ea(p.id)}')">${p.photoUrl||p.photo?`<img src="${ea(p.photoUrl||p.photo)}">`:'<div class="ph">📦</div>'}<b>${esc(p.name)}</b><span>SKU ${esc(p.sku||'-')} • Item ${esc(firstMarketplaceId(p)||'-')}<br>${p.score?Math.round(p.score)+'% de confiança':'catálogo'}</span></button>`).join('')}</div>
    <select class="fi" id="link-product">${catalog.map(p=>`<option value="${ea(p.id)}" ${p.id===o.linkedProductId?'selected':''}>${esc(p.name)}${p.sku?' • '+esc(p.sku):''}</option>`).join('')}</select>
    <button class="btn btn-save" onclick="saveProductLink('${ea(orderId)}')">Salvar vínculo permanente</button><button class="btn btn-secondary" onclick="openStoreProductImport()">Cadastrar produto da loja</button><button class="btn btn-secondary" onclick="closeSht()">Cancelar</button>`);
}
function saveProductLink(orderId,productId){
  const o=marketplaceOrders.find(x=>x.id===orderId)||importReview?.orders.find(x=>x.id===orderId)||analyzedOrders().find(x=>x.id===orderId);
  if(!o)return;
  const selected=productId||document.getElementById('link-product')?.value; if(!selected)return;
  const p=catalogProducts().find(x=>x.id===selected||x.legacyProductId===selected);
  const keys=[...new Set(manualRuleKeys(o))];
  keys.forEach(key=>{productMatchRules[key]={productId:selected,productName:p?.name||'',source:o.productName||o.sku||o.itemId,savedTs:Date.now()};});
  o.linkedProductId=selected;
  o.linkedProductName=p?.name||o.linkedProductName||'';
  o.linkConfidence=1;
  o.linkMethod='manual permanente';
  if(matchRulesRef){
    const updates={}; keys.forEach(key=>updates[key]=sanitizeForFirebase(productMatchRules[key]));
    matchRulesRef.update(updates).catch(()=>{});
  }else{
    localStorage.setItem('jm3d_match_rules',JSON.stringify(productMatchRules));
  }
  if(marketplaceOrdersRef&&marketplaceOrders.some(x=>x.id===orderId)){
    marketplaceOrdersRef.child(orderId).update(sanitizeForFirebase({
      linkedProductId:selected,
      linkedProductName:p?.name||'',
      linkConfidence:1,
      linkMethod:'manual permanente',
      updatedAt:new Date().toISOString()
    })).catch(()=>{});
  }else{
    localStorage.setItem('jm3d_market_orders',JSON.stringify(marketplaceOrders));
  }
  closeSht();renderContent();
}
function findProductForOrder(o){
  return findProductMatch(o).product;
}
function firstPositiveNumber(...values){
  for(const v of values){
    const n=Number(v);
    if(Number.isFinite(n)&&n>0)return n;
  }
  return 0;
}
function explicitProductUnitCost(product){
  const pc=product?.costs||{};
  return firstPositiveNumber(
    pc.totalUnitCost,
    pc.unitCost,
    pc.totalCost,
    product?.totalCost,
    product?.unitCost,
    product?.cost
  );
}
function productUnitCost(product,meiUnit=0){
  const s=getMarketplaceSettings(), cfg=getLocalConfig();
  if(!product)return {total:meiUnit,filament:0,energy:0,packaging:0,bubble:0,label:0,maintenance:0,waste:0,postProcess:0,nozzleWear:0,extra:0,mei:meiUnit,other:0};
  const pc=product.costs||{};
  const explicitUnitCost=explicitProductUnitCost(product);
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
  const productionTotal=explicitUnitCost||base+loss;
  return {total:productionTotal+meiUnit,filament,energy,packaging,bubble,label,maintenance,waste:loss,postProcess,nozzleWear,extra,mei:meiUnit,other};
}
function enrichOrder(o,meiUnit=0){
  o=Object.assign({},o,{marketplace:canonicalMarketplace(o.marketplace),orderId:o.orderId||o.id,productName:o.productName||o.importedProductName,sku:o.sku||o.importedSku,qty:o.qty||o.quantity||1,date:o.date||o.orderDate,paymentDate:o.paymentDate||o.paidDate});
  const match=findProductMatch(o);
  const product=match.product;
  const qty=Number(o.qty)||1;
  const fin=o.financial||{};
  const gross=Number(fin.grossAmount ?? o.gross)||0;
  const fees=Math.abs(Number(fin.commissionFee||0)+Number(fin.fixedFee||0)+Number(fin.transactionFee||0)+Number(fin.affiliateFee||0)+Number(fin.otherFees||0))||Math.abs(Number(o.fees)||0);
  const discounts=Math.abs(Number(fin.marketplaceDiscount||0)+Number(fin.sellerDiscount||0))||Math.abs(Number(o.discounts)||0);
  const shipping=Number(fin.shippingCostToSeller ?? o.shipping)||0;
  const net=Number(fin.netReceived ?? o.net)||(gross-fees-discounts-Math.abs(shipping));
  let c=productUnitCost(product,meiUnit);
  if(!product)c={total:qty?net/qty:0,filament:0,energy:0,packaging:0,bubble:0,label:0,maintenance:0,waste:0,postProcess:0,nozzleWear:0,extra:0,mei:0,other:0};
  const cost=c.total*qty;
  const profit=product?net-cost:0;
  const margin=product&&net>0?profit/net*100:0;
  const health=!product?'sem_vinculo':profit<0?'prejuizo':margin<15?'margem_baixa':margin<25?'atenção':'healthy';
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
    health:!product?'unlinked':profit<0?'bad':margin<15?'warn':'ok',
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
      <div class="chart-card"><div class="chart-title">Recebimento por dia</div>${renderLineChart(dailySeries(orders,'net'),'#00C2FF')}</div>
      <div class="chart-card"><div class="chart-title">Lucro líquido por dia</div>${renderLineChart(dailySeries(orders,'profit'),'#00C2FF')}</div>
      <div class="chart-card"><div class="chart-title">Shopee vs TikTok Shop</div>${renderBarCompare([{label:'Shopee',value:s.shopeeProfit,color:'#00C2FF'},{label:'TikTok',value:s.tiktokProfit,color:'#F97316'}])}</div>
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
    {label:'Filamento',value:s.filament,color:'#00C2FF'},
    {label:'Energia',value:s.energy,color:'#F97316'},
    {label:'Embalagem',value:s.packaging,color:'#7B2FFF'},
    {label:'Marketplace',value:s.fees,color:'#FF3B8B'},
    {label:'MEI',value:s.mei,color:'#00C2FF'}
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
function openLinkProductLegacyOrdersFinal(orderId){
  const o=analyzedOrders().find(x=>x.id===orderId); if(!o)return;
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  openSht('Vincular produto',`<div class="muted-note">Produto no relatório:<br><b>${esc(o.productName)}</b><br>SKU/item: ${esc(o.sku||o.itemId||'-')}</div><select class="fi" id="link-product">${catalog.map(p=>`<option value="${p.id}" ${p.id===o.linkedProductId?'selected':''}>${esc(p.name)}${p.sku?' • '+esc(p.sku):''}</option>`).join('')}</select><button class="btn btn-save" onclick="saveProductLink('${o.id}')">Salvar vínculo permanente</button><button class="btn btn-secondary" onclick="closeSht()">Cancelar</button>`);
}
function saveProductLinkLegacy(orderId){
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

/* ERP experience layer */
var erpProductRank='profit';
var erpPricingPrice=39.90;
var erpPricingProductId='';
var storeProductDraft=null;

function erpPeriodOptions(){
  return [
    ['today','Hoje'],['7d','7 dias'],['30d','30 dias'],
    ['thisMonth','Este mês'],['previousMonth','Mês anterior'],['custom','Personalizado']
  ];
}
function erpStores(){
  return [['all','Todas'],['_kaline98','Shopee _kaline98'],['mateusoliver98','Shopee mateusoliver98'],['TikTok Shop','TikTok Shop']];
}
function erpPeriodBounds(period=analysisFilters.period){
  const today=new Date(); today.setHours(0,0,0,0);
  const startOfMonth=new Date(today.getFullYear(),today.getMonth(),1);
  const endOfMonth=new Date(today.getFullYear(),today.getMonth()+1,0,23,59,59,999);
  if(period==='today')return {from:today,to:new Date(today.getTime()+86399999)};
  if(period==='7d')return {from:new Date(today.getTime()-6*86400000),to:new Date(today.getTime()+86399999)};
  if(period==='30d')return {from:new Date(today.getTime()-29*86400000),to:new Date(today.getTime()+86399999)};
  if(period==='thisMonth')return {from:startOfMonth,to:endOfMonth};
  if(period==='previousMonth')return {from:new Date(today.getFullYear(),today.getMonth()-1,1),to:new Date(today.getFullYear(),today.getMonth(),0,23,59,59,999)};
  if(period==='custom')return {
    from:analysisFilters.from?new Date(analysisFilters.from+'T00:00:00'):null,
    to:analysisFilters.to?new Date(analysisFilters.to+'T23:59:59'):null
  };
  return {from:null,to:null};
}
function isInsidePeriod(o){
  const d=parseDateFlexible(o.paymentDate||o.paidDate||o.date||o.orderDate);
  if(!d)return true;
  const day=new Date(d+'T12:00:00');
  const b=erpPeriodBounds();
  return (!b.from||day>=b.from)&&(!b.to||day<=b.to);
}
function erpPreviousOrders(){
  const b=erpPeriodBounds();
  if(!b.from||!b.to)return [];
  const span=b.to.getTime()-b.from.getTime()+1;
  const prevTo=new Date(b.from.getTime()-1);
  const prevFrom=new Date(prevTo.getTime()-span+1);
  return analyzedOrders().filter(o=>{
    const d=parseDateFlexible(o.paymentDate||o.paidDate||o.date||o.orderDate);
    if(!d)return false;
    const day=new Date(d+'T12:00:00');
    if(day<prevFrom||day>prevTo)return false;
    if(analysisFilters.marketplace!=='all'&&marketplaceFilterValue(o.marketplace)!==analysisFilters.marketplace)return false;
    if(analysisFilters.store!=='all'&&o.store!==analysisFilters.store)return false;
    return true;
  });
}
function erpTrend(current,previous){
  if(!previous&&current)return 'novo período';
  if(!previous)return 'sem comparação';
  const delta=(current-previous)/Math.abs(previous)*100;
  return `${delta>=0?'subiu':'caiu'} ${Math.abs(delta).toFixed(1)}%`;
}
function erpReserveTotal(s){return s.filament+s.energy+s.packaging+s.mei+s.other;}
function erpFreeProfit(s){return s.net-erpReserveTotal(s)-s.fees;}
function erpTicket(s,orders){return orders.length?s.net/orders.length:0;}
function erpRoi(s){return s.cost?s.profit/s.cost*100:0;}
function erpHealth(margin,profit=1){
  if(profit<0||margin<0)return {cls:'bad',label:'Prejuízo'};
  if(margin<15)return {cls:'bad',label:'Margem baixa'};
  if(margin<25)return {cls:'warn',label:'Atenção'};
  if(margin<45)return {cls:'good',label:'Saudável'};
  return {cls:'good',label:'Excelente'};
}
function erpTopProduct(orders,mode='profit'){
  return productPerformance(orders).sort((a,b)=>(b[mode]||0)-(a[mode]||0))[0]||null;
}
function erpTopStore(orders){
  const map={};
  orders.forEach(o=>{
    const k=o.store||'Sem loja';
    map[k]=map[k]||{name:k,net:0,profit:0,orders:0,qty:0};
    map[k].net+=Number(o.net)||0; map[k].profit+=Number(o.profit)||0; map[k].orders++; map[k].qty+=Number(o.qty)||1;
  });
  return Object.values(map).sort((a,b)=>b.profit-a.profit)[0]||null;
}
function erpTopMarketplace(orders){
  const map={};
  orders.forEach(o=>{
    const k=marketplaceFilterValue(o.marketplace);
    map[k]=map[k]||{name:k,net:0,profit:0,orders:0};
    map[k].net+=Number(o.net)||0; map[k].profit+=Number(o.profit)||0; map[k].orders++;
  });
  return Object.values(map).sort((a,b)=>b.profit-a.profit)[0]||null;
}
function erpCard(icon,label,value,cls='info',hint=''){
  return `<div class="erp-card ${cls}"><div class="ico">${icon}</div><div class="lbl">${label}</div><div class="val">${value}</div>${hint?`<div class="hint">${hint}</div>`:''}</div>`;
}
function erpStoreStats(orders){
  return erpStores().filter(x=>x[0]!=='all').map(([id,label])=>{
    const subset=orders.filter(o=>o.store===id), s=summarizeOrders(subset), margin=s.net?s.profit/s.net*100:0;
    return {id,label,orders:subset.length,net:s.net,profit:s.profit,qty:s.qty,margin,ticket:erpTicket(s,subset)};
  }).sort((a,b)=>b.profit-a.profit);
}
function production3dMetrics(orders){
  const catalog=catalogProducts();
  return orders.reduce((a,o)=>{
    const product=catalog.find(p=>p.id===o.linkedProductId||p.legacyProductId===o.linkedProductId)||findProductMatch(o).product||{};
    const qty=Number(o.qty)||1;
    const pc=product.costs||{};
    const weight=Number(pc.weightGrams ?? product.weight)||0;
    const hours=Number(pc.printTimeHours ?? product.printH)||0;
    a.grams+=weight*qty;
    a.hours+=hours*qty;
    a.products+=qty;
    return a;
  },{grams:0,hours:0,products:0});
}
function catalogProducts(){return typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;}
function erpFilters(){
  return `<div class="erp-section"><div class="erp-pills">${erpPeriodOptions().map(([id,label])=>`<button class="${analysisFilters.period===id?'on':''}" onclick="analysisFilters.period='${id}';renderContent()">${label}</button>`).join('')}</div>
  ${analysisFilters.period==='custom'?`<div class="filter-row" style="margin-top:10px"><input class="fi" type="date" value="${analysisFilters.from}" onchange="analysisFilters.from=this.value;renderContent()"><input class="fi" type="date" value="${analysisFilters.to}" onchange="analysisFilters.to=this.value;renderContent()"></div>`:''}
  <div class="erp-pills" style="margin-top:8px">${erpStores().map(([id,label])=>`<button class="${analysisFilters.store===id?'on':''}" onclick="analysisFilters.store='${id}';renderContent()">${label}</button>`).join('')}</div></div>`;
}
function renderAnalysis(){
  const content=document.getElementById('content'); if(!content)return;
  content.innerHTML=renderFinanceiroView();
}
function renderFinanceiroView(){
  const orders=filteredOrders(), s=summarizeOrders(orders), prev=summarizeOrders(erpPreviousOrders());
  const margin=s.net?s.profit/s.net*100:0, free=erpFreeProfit(s), reserve=erpReserveTotal(s);
  const bestProduct=erpTopProduct(orders,'profit'), bestStore=erpTopStore(orders), bestMarketplace=erpTopMarketplace(orders);
  return `<div class="erp-shell">
    ${erpFilters()}
    <div class="erp-hero">
      <div class="erp-eyebrow">Caiu na sua conta</div>
      <div class="erp-big">${brl(s.net)}</div>
      <div class="erp-sub">Valor líquido recebido da Shopee e TikTok Shop. ${erpTrend(s.net,prev.net)} em relação ao período anterior.</div>
    </div>
    <div class="erp-grid">
      ${erpCard('💼','Lucro livre',brl(free),free>=0?'good':'bad','Depois de separar produção, taxas e MEI.')}
      ${erpCard('📈','Margem real',pct(margin),margin>=25?'good':margin>=15?'warn':'bad',`${orders.length} pedidos no período`)}
      ${erpCard('🧾','Ticket médio',brl(erpTicket(s,orders)),'info','Recebido líquido por pedido.')}
      ${erpCard('🚀','ROI',pct(erpRoi(s)),erpRoi(s)>=50?'good':'warn','Lucro sobre custo de produção.')}
    </div>
    <section class="erp-section"><h3>Você precisa separar <small>${brl(reserve)} no total</small></h3>
      <div class="reserve-list">
        ${reserveRow('Filamento',s.filament)}${reserveRow('Energia',s.energy)}${reserveRow('Embalagens',s.packaging)}
        ${reserveRow('MEI',s.mei)}${reserveRow('Manutenção e outros',s.other)}
      </div>
    </section>
    <section class="erp-section"><h3>Campeões do período <small>decisão rápida</small></h3>
      <div class="erp-grid">
        ${renderChampionProduct(bestProduct)}
        ${erpCard('🏬','Loja campeã',bestStore?esc(bestStore.name):'Sem dados','info',bestStore?`${bestStore.orders} pedidos • ${brl(bestStore.profit)} de lucro`:'')}
        ${erpCard('🛒','Marketplace campeão',bestMarketplace?marketplaceLabel(bestMarketplace.name):'Sem dados','info',bestMarketplace?`${brl(bestMarketplace.net)} recebidos`:'' )}
        ${erpCard('⚠','Produtos sem vínculo',String(s.unlinked),s.unlinked?'warn':'good',s.unlinked?'Corrija vínculos para melhorar custos.':'Tudo vinculado no período.')}
      </div>
    </section>
    <section class="erp-section"><h3>Fluxo de caixa <small>recebido, lucro e separação</small></h3>${renderCashflowChart(orders)}</section>
    <section class="erp-section"><h3>Comparação entre lojas <small>performance real</small></h3>${renderStoreComparison(orders)}</section>
    <section class="erp-section"><h3>Alertas de negócio <small>o que agir primeiro</small></h3>${renderBusinessAlerts(orders)}</section>
    <button class="btn btn-export" onclick="exportAnalysisCsv()">Exportar CSV da análise</button>
  </div>`;
}
function renderFinanceiroView(){
  const orders=filteredOrders(), s=summarizeOrders(orders), prev=summarizeOrders(erpPreviousOrders());
  const margin=s.net?s.profit/s.net*100:0, reserve=erpReserveTotal(s), available=s.net-reserve, prod=production3dMetrics(orders);
  const bestProduct=erpTopProduct(orders,'profit'), bestStore=erpTopStore(orders), bestMarketplace=erpTopMarketplace(orders);
  return `<div class="erp-shell">
    ${erpFilters()}
    <section class="finance-hero-grid">
      <div class="erp-hero hero-available">
        <div class="erp-eyebrow">Dinheiro disponível</div>
        <div class="erp-big">${brl(available)}</div>
        <div class="erp-sub">Recebido menos o que precisa separar. ${erpTrend(s.net,prev.net)} em recebimento no período.</div>
      </div>
      <div class="erp-card received-card info"><div class="ico">💳</div><div class="lbl">Recebido</div><div class="val">${brl(s.net)}</div><div class="hint">Caiu na conta pela Shopee/TikTok.</div></div>
      <div class="erp-card separate-now warn"><div class="ico">🧾</div><div class="lbl">Separar agora</div><div class="val">${brl(reserve)}</div><div class="hint">Filamento, energia, embalagem, MEI e manutenção.</div></div>
    </section>
    <div class="erp-grid kpi-strip">
      ${erpCard('💼','Lucro livre',brl(s.profit),s.profit>=0?'good':'bad','Resultado real dos pedidos importados.')}
      ${erpCard('📈','Margem real',pct(margin),margin>=25?'good':margin>=15?'warn':'bad',`${orders.length} pedidos no período`)}
      ${erpCard('🧺','Produtos vendidos',String(s.qty),'info','Unidades vendidas no filtro atual.')}
      ${erpCard('🎯','Ticket médio',brl(erpTicket(s,orders)),'info','Recebido líquido por pedido.')}
    </div>
    <section class="erp-section"><h3>Você precisa separar <small>${brl(reserve)} no total</small></h3>
      <div class="reserve-list">
        ${reserveRow('Filamento',s.filament)}${reserveRow('Energia',s.energy)}${reserveRow('Embalagens',s.packaging)}
        ${reserveRow('MEI',s.mei)}${reserveRow('Manutenção e outros',s.other)}
      </div>
    </section>
    <section class="erp-section"><h3>Produção 3D <small>operação do período</small></h3>
      <div class="prod-kpi-grid">
        ${productionKpi('Filamento consumido',(prod.grams/1000).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+' kg')}
        ${productionKpi('Horas impressas',prod.hours.toLocaleString('pt-BR',{maximumFractionDigits:1})+' h')}
        ${productionKpi('Produtos vendidos',String(prod.products))}
        ${productionKpi('Ticket médio',brl(erpTicket(s,orders)))}
      </div>
    </section>
    <section class="erp-section"><h3>Alertas importantes <small>o que agir primeiro</small></h3>${renderBusinessAlerts(orders)}</section>
    <section class="erp-section"><h3>Destaques do período <small>decisão rápida</small></h3>
      <div class="highlight-grid">
        ${renderChampionProduct(bestProduct)}
        ${erpCard('🏬','Loja campeã',bestStore?esc(bestStore.name):'Sem dados','info',bestStore?`${bestStore.orders} pedidos - ${brl(bestStore.profit)} de lucro`:'')}
        ${erpCard('🛒','Marketplace campeão',bestMarketplace?marketplaceLabel(bestMarketplace.name):'Sem dados','info',bestMarketplace?`${brl(bestMarketplace.net)} recebidos`:'' )}
      </div>
    </section>
    <section class="erp-section"><h3>Fluxo de caixa <small>recebido, lucro e separação</small></h3>${renderCashflowChart(orders)}</section>
    <section class="erp-section"><h3>Ranking das lojas <small>performance real</small></h3>${renderStoreComparison(orders)}</section>
    <button class="btn btn-export" onclick="exportAnalysisCsv()">Exportar CSV da análise</button>
  </div>`;
}
function renderChampionProduct(p){
  if(!p)return erpCard('🏆','Produto campeão','Sem dados','good','Importe relatórios para descobrir.');
  return `<div class="erp-product-card" style="margin:0;grid-template-columns:74px 1fr">
    ${p.photo?`<img class="erp-product-img" style="width:74px;height:74px" src="${p.photo}">`:'<div class="erp-product-fallback" style="width:74px;height:74px">🏆</div>'}
    <div><div class="erp-eyebrow">Produto campeão</div><div class="erp-product-name">${esc(p.name)}</div><div class="erp-product-meta">${p.qty} vendas • ${brl(p.profit)} lucro • margem ${pct(p.margin)}</div><div class="badge-row"><span class="biz-badge good">Mais lucrativo</span></div></div>
  </div>`;
}
function reserveRow(label,value){return `<div class="reserve-row"><span>${label}</span><b>${brl(value)}</b></div>`;}
function productionKpi(label,value){return `<div class="prod-kpi"><span>${label}</span><b>${value}</b></div>`;}
function marketplaceLabel(v){return v==='tiktokShop'?'TikTok Shop':v==='shopee'?'Shopee':v;}
function dailySeries3(orders){
  const map={};
  orders.forEach(o=>{
    const d=parseDateFlexible(o.paymentDate||o.paidDate||o.date||o.orderDate)||'sem data';
    map[d]=map[d]||{label:d.slice(5),net:0,profit:0,reserve:0};
    map[d].net+=Number(o.net)||0; map[d].profit+=Number(o.profit)||0;
    const c=o.costParts||{};
    map[d].reserve+=(Number(c.filament)||0)+(Number(c.energy)||0)+(Number(c.packaging)||0)+(Number(c.bubble)||0)+(Number(c.mei)||0)+(Number(c.other)||0);
  });
  return Object.keys(map).sort().slice(-14).map(k=>map[k]);
}
function renderCashflowChart(orders){
  const data=dailySeries3(orders);
  if(!data.length)return '<div class="muted-note">Importe relatórios para ver o fluxo de caixa.</div>';
  const w=340,h=170,p=24,max=Math.max(...data.flatMap(d=>[d.net,d.profit,d.reserve].map(Math.abs)),1);
  const pathFor=(field)=>data.map((d,i)=>{
    const x=p+(data.length===1?0:i*(w-p*2)/(data.length-1));
    const y=h-p-(Math.max(0,d[field])/max)*(h-p*2);
    return `${i?'L':'M'}${x.toFixed(1)} ${y.toFixed(1)}`;
  }).join(' ');
  return `<div class="cashflow-card"><svg class="cashflow-svg" viewBox="0 0 ${w} ${h}" aria-label="Fluxo de caixa">
    <defs><linearGradient id="cashFill" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#00C2FF" stop-opacity=".32"/><stop offset="1" stop-color="#3B82F6" stop-opacity="0"/></linearGradient></defs>
    <path d="${pathFor('net')} L ${w-p} ${h-p} L ${p} ${h-p} Z" fill="url(#cashFill)"/>
    <path d="${pathFor('net')}" fill="none" stroke="#00C2FF" stroke-width="4" stroke-linecap="round"/>
    <path d="${pathFor('profit')}" fill="none" stroke="#3B82F6" stroke-width="3" stroke-linecap="round"/>
    <path d="${pathFor('reserve')}" fill="none" stroke="#F97316" stroke-width="3" stroke-linecap="round" stroke-dasharray="5 6"/>
    ${data.map((d,i)=>`<text x="${p+i*(w-p*2)/Math.max(1,data.length-1)}" y="${h-5}" fill="#8899BB" font-size="9" text-anchor="middle">${d.label}</text>`).join('')}
  </svg><div class="legend"><div><span style="background:#00C2FF"></span>Recebido por dia</div><div><span style="background:#00C2FF"></span>Lucro por dia</div><div><span style="background:#F97316"></span>Separação por dia</div></div></div>`;
}
function renderStoreComparison(orders){
  const cards=erpStores().filter(x=>x[0]!=='all').map(([id,label])=>{
    const subset=orders.filter(o=>o.store===id), s=summarizeOrders(subset), margin=s.net?s.profit/s.net*100:0;
    return `<div class="store-card"><div><b>${label}</b><div class="store-meta">${subset.length} pedidos • margem ${pct(margin)}<br>Ticket ${brl(erpTicket(s,subset))}</div></div><div class="store-val">${brl(s.profit)}<br><small>${brl(s.net)}</small></div></div>`;
  }).join('');
  return `<div class="store-grid">${cards}</div>`;
}
function renderBusinessAlerts(orders){
  const products=productPerformance(orders);
  const low=products.filter(p=>p.margin<20&&p.qty>0).slice(0,3);
  const bad=products.filter(p=>p.profit<0).slice(0,3);
  const unlinked=orders.filter(o=>!o.linkedProductId).slice(0,3);
  const items=[
    ...low.map(p=>['warn','Preço abaixo do ideal',`${p.name}: margem ${pct(p.margin)}`]),
    ...bad.map(p=>['bad','Produto em prejuízo',`${p.name}: ${brl(p.profit)}`]),
    ...unlinked.map(o=>['warn','Produto sem vínculo',o.productName||o.orderId])
  ].slice(0,6);
  if(!items.length)return '<div class="alert-item"><b>Tudo certo no período</b><span>sem alertas críticos</span></div>';
  return `<div class="alert-list">${items.map(([cls,title,body])=>`<div class="alert-item"><b>${title}</b><span class="${cls}">${esc(body)}</span></div>`).join('')}</div>`;
}
function renderStoreComparison(orders){
  const stats=erpStoreStats(orders), max=Math.max(...stats.map(s=>Math.max(0,s.profit)),1);
  if(!stats.length)return '<div class="muted-note">Sem lojas para comparar no período.</div>';
  return `<div class="store-rank-grid">${stats.map((s,i)=>`<div class="store-rank-card ${i===0?'top':''}">
    <div class="store-rank-head"><span>${i+1}</span><b>${esc(s.label)}</b></div>
    <div class="store-rank-value">${brl(s.profit)}</div>
    <div class="store-rank-meta">Recebido ${brl(s.net)} - ${s.orders} pedidos - margem ${pct(s.margin)}</div>
    <div class="store-rank-bar"><i style="width:${Math.min(100,Math.max(4,s.profit/max*100))}%"></i></div>
    <div class="store-rank-foot"><span>Ticket ${brl(s.ticket)}</span><span>${s.qty} un.</span></div>
  </div>`).join('')}</div>`;
}
function renderBusinessAlerts(orders){
  const products=productPerformance(orders);
  const low=products.filter(p=>p.margin<20&&p.qty>0).slice(0,2);
  const bad=products.filter(p=>p.profit<0).slice(0,2);
  const unlinked=orders.filter(o=>!o.linkedProductId).slice(0,4);
  const cards=[
    ...unlinked.map(o=>({cls:'warn',title:'Produto sem vínculo',body:`${o.productName||o.importedProductName||o.orderId} - ${marketplaceLabel(o.marketplace)} - ${o.store||''}`,action:`<button onclick="openLinkProduct('${ea(o.id)}')">Vincular</button>`})),
    ...bad.map(p=>({cls:'bad',title:'Produto em prejuízo',body:`${p.name}: ${brl(p.profit)}`,action:''})),
    ...low.map(p=>({cls:'warn',title:'Margem baixa',body:`${p.name}: margem ${pct(p.margin)}`,action:''}))
  ].slice(0,6);
  if(!cards.length)return '<div class="alert-item"><b>Tudo certo no período</b><span>sem alertas críticos</span></div>';
  return `<div class="alert-list alert-action-list">${cards.map(c=>`<div class="alert-item action ${c.cls}"><div><b>${esc(c.title)}</b><span>${esc(c.body)}</span></div>${c.action}</div>`).join('')}</div>`;
}
function refFilters(){
  return `<div class="ref-filterbar"><div class="ref-searchbar"><input class="fi" placeholder="Buscar produto, pedido, SKU ou item..." value="${ea(analysisFilters.q)}" oninput="analysisFilters.q=this.value;renderContent()" autocomplete="off"><button class="btn btn-secondary" onclick="analysisFilters.q='';renderContent()">Limpar</button></div>
  <div class="ref-pills">${erpPeriodOptions().map(([id,label])=>`<button class="${analysisFilters.period===id?'on':''}" onclick="analysisFilters.period='${id}';renderContent()">${label}</button>`).join('')}</div>
  <div class="ref-pills">${erpStores().map(([id,label])=>`<button class="${analysisFilters.store===id?'on':''}" onclick="analysisFilters.store='${id}';renderContent()">${label}</button>`).join('')}</div>
  ${analysisFilters.period==='custom'?`<div class="ref-dates"><input class="fi" type="date" value="${analysisFilters.from}" onchange="analysisFilters.from=this.value;renderContent()"><input class="fi" type="date" value="${analysisFilters.to}" onchange="analysisFilters.to=this.value;renderContent()"></div>`:''}</div>`;
}
function refSparkline(values,color='#00C2FF',fill=false){
  const data=(values||[]).map(Number).filter(Number.isFinite);
  const w=150,h=54,left=4,right=4,top=7,bottom=8, id='spark'+Math.random().toString(36).slice(2,8);
  const fallback=[18,16,20,14,28,31,26,36,33,43];
  const series=data.length?data:fallback;
  const min=Math.min(...series,0),max=Math.max(...series,1),span=max-min||1;
  const points=series.map((v,i)=>({
    x:left+(series.length===1?0:i*(w-left-right)/(series.length-1)),
    y:h-bottom-((v-min)/span)*(h-top-bottom)
  }));
  const line=jmSmoothPath(points);
  const area=`${line} L ${points[points.length-1].x.toFixed(1)} ${h-bottom} L ${points[0].x.toFixed(1)} ${h-bottom} Z`;
  const markerEvery=Math.max(1,Math.floor(points.length/4));
  return `<svg class="ref-spark ref-spark-premium" viewBox="0 0 ${w} ${h}" aria-hidden="true">
    <defs>
      <linearGradient id="${id}Area" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="${color}" stop-opacity=".34"/><stop offset=".72" stop-color="${color}" stop-opacity=".10"/><stop offset="1" stop-color="${color}" stop-opacity="0"/></linearGradient>
      <linearGradient id="${id}Line" x1="0" x2="1" y1="0" y2="0"><stop offset="0" stop-color="#3B82F6"/><stop offset=".48" stop-color="${color}"/><stop offset="1" stop-color="#22D3EE"/></linearGradient>
    </defs>
    <line class="spark-grid" x1="${left}" y1="${top+8}" x2="${w-right}" y2="${top+8}"/>
    <line class="spark-grid" x1="${left}" y1="${top+22}" x2="${w-right}" y2="${top+22}"/>
    <line class="spark-grid" x1="${left}" y1="${top+36}" x2="${w-right}" y2="${top+36}"/>
    ${fill?`<path d="${area}" fill="url(#${id}Area)"/>`:''}
    <path class="spark-line" d="${line}" fill="none" stroke="url(#${id}Line)" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/>
    ${points.filter((_,i)=>i%markerEvery===0||i===points.length-1).map(p=>`<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="2.6" fill="#22D3EE" stroke="#0B0F14" stroke-width="1.6"/>`).join('')}
  </svg>`;
}
function refBars(values){
  const data=(values||[]).map(v=>Math.max(0,Number(v)||0)).slice(-10), max=Math.max(...data,1);
  return `<div class="ref-bars">${data.map(v=>`<i style="height:${Math.max(12,v/max*100)}%"></i>`).join('')}</div>`;
}
function refDailyValues(orders,field='net'){
  return dailySeries3(orders).map(d=>Number(d[field])||0);
}
function refProductImage(src,alt='Produto'){
  return src?`<img class="ref-product-img" src="${ea(src)}" alt="${ea(alt)}">`:`<div class="ref-product-img ref-product-fallback">${esc(String(alt||'Produto').slice(0,2).toUpperCase())}</div>`;
}
function refTopSummary(orders,s,available,margin){
  const netSeries=refDailyValues(orders,'net'), profitSeries=refDailyValues(orders,'profit');
  return `<section class="ref-card ref-summary">
    <div class="ref-summary-block">
      <div class="ref-label">Lucro livre</div>
      <div class="ref-money green">${brl(s.profit)}</div>
      ${refSparkline(profitSeries,'#00C2FF',true)}
      <div class="spark-caption"><i style="background:#00C2FF"></i>Evolucao do lucro</div>
      <div class="ref-summary-foot"><span>Caiu na sua conta:<b>${brl(s.net)}</b></span><span>Pedidos:<b>${orders.length}</b></span></div>
    </div>
    <div class="ref-summary-block available">
      <div class="ref-label">Dinheiro disponível</div>
      <div class="ref-money">${brl(available)}</div>
      ${refSparkline(netSeries,'#3B82F6',true)}
      <div class="spark-caption"><i style="background:#3B82F6"></i>Evolucao do recebido</div>
      <div class="ref-summary-foot"><span>Margem real:<b>${pct(margin)}</b></span><span>Produtos:<b>${s.qty}</b></span></div>
    </div>
  </section>`;
}
function refCostCard(s,reserve,orders){
  const prod=production3dMetrics(orders);
  return `<section class="ref-card ref-costs"><h3>Análise de Custos e Separação</h3>
    <div class="ref-cost-list">
      ${reserveRow('Filamento',s.filament)}${reserveRow('Energia',s.energy)}${reserveRow('Embalagens',s.packaging)}
      ${reserveRow('MEI',s.mei)}${reserveRow('Manutenção e outros',s.other)}
    </div>
    <div class="ref-separate">Preciso separar:<b>${brl(reserve)}</b></div>
    <div class="ref-production mini"><span>kg <b>${(prod.grams/1000).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</b></span><span>horas <b>${prod.hours.toLocaleString('pt-BR',{maximumFractionDigits:1})}</b></span><span>vendidos <b>${prod.products}</b></span><span>ticket <b>${brl(erpTicket(s,orders))}</b></span></div>
  </section>`;
}
function refUnlinkedCard(orders){
  const unlinked=orders.filter(o=>!o.linkedProductId), rows=unlinked.slice(0,5);
  return `<section class="ref-card ref-unlinked"><h3><span>Produtos sem vínculo (${unlinked.length})</span></h3>
    ${rows.length?`<div class="ref-table-wrap"><table class="ref-table"><thead><tr><th></th><th>Nome</th><th>SKU ID</th><th>Loja</th><th>Ação</th></tr></thead><tbody>${rows.map((o,i)=>{
      const match=findProductMatch(o).product||{}, photo=o.importedPhotoUrl||match.photoUrl||match.photo||'';
      return `<tr><td>${i+1}</td><td><div class="ref-row-product">${refProductImage(photo,o.productName||o.importedProductName)}<b>${esc(o.productName||o.importedProductName||'Produto sem nome')}</b></div></td><td>${esc(o.sku||o.importedSku||o.itemId||o.importedItemId||'-')}</td><td>${esc(o.store||marketplaceLabel(o.marketplace))}</td><td><button onclick="openLinkProduct('${ea(o.id)}')">Vincular</button></td></tr>`;
    }).join('')}</tbody></table></div>`:'<div class="ref-empty">Tudo vinculado no período.</div>'}
  </section>`;
}
function refHighlights(bestProduct,bestStore,bestMarketplace){
  return `<aside class="ref-card ref-highlights"><h3>Destaques do Período</h3>
    <div class="ref-feature">
      <div class="ref-feature-label">Produto campeão do período</div>
      ${refProductImage(bestProduct?.photo,bestProduct?.name||'Produto campeão')}
      <h4>${esc(bestProduct?.name||'Sem dados')}</h4>
      <div class="ref-feature-grid"><span>${bestProduct?.qty||0} vendas</span><b>${brl(bestProduct?.profit||0)} lucro</b><span>margem ${pct(bestProduct?.margin||0)}</span><b>${brl(bestProduct?.qty?bestProduct.profit/bestProduct.qty:0)}</b></div>
    </div>
    <div class="ref-mini-grid">
      <div class="ref-mini-card"><span>Loja campeã</span><b>${bestStore?esc(bestStore.name):'Sem dados'}</b></div>
      <div class="ref-mini-card"><span>Marketplace campeão</span><b>${bestMarketplace?marketplaceLabel(bestMarketplace.name):'Sem dados'}</b></div>
    </div>
  </aside>`;
}
function storeLogoClass(id){return id==='TikTok Shop'?'tiktok':'shopee';}
function refStoreCards(orders){
  const stats=erpStoreStats(orders);
  return `<section class="ref-stores">${stats.map(s=>{
    const subset=orders.filter(o=>o.store===s.id), series=refDailyValues(subset,'profit');
    return `<div class="ref-store-card"><div class="ref-store-left"><div class="ref-store-logo ${storeLogoClass(s.id)}">${s.id==='TikTok Shop'?'♪':'S'}</div><b>${esc(s.label)}</b></div>
    <div class="ref-store-main"><span>Profit:</span><strong>${brl(s.profit)}</strong>${refSparkline(series,'#00C2FF')}</div>
    <div class="ref-store-metrics"><span>recebido <b>${brl(s.net)}</b></span><span>custos <b>${brl(Math.max(0,s.net-s.profit))}</b></span><span>margem <b>${pct(s.margin)}</b></span><span>ticket <b>${brl(s.ticket)}</b></span></div></div>`;
  }).join('')}</section>`;
}
function refChampionProduct(orders){
  const rows=productPerformance(orders).filter(p=>p.qty>0).sort((a,b)=>b.profit-a.profit);
  auditProductCalculations('Financeiro/Dashboard',rows);
  return rows.find(p=>p.photo&&p.linked)||rows.find(p=>p.photo)||rows[0]||null;
}
function renderFinanceiroView(){
  const orders=filteredOrders(), s=summarizeOrders(orders), reserve=erpReserveTotal(s), available=s.net-reserve, margin=s.net?s.profit/s.net*100:0;
  const bestProduct=refChampionProduct(orders), bestStore=erpTopStore(orders), bestMarketplace=erpTopMarketplace(orders);
  return `<div class="ref-dashboard">
    ${refFilters()}
    <div class="ref-board">
      ${refTopSummary(orders,s,available,margin)}
      ${refHighlights(bestProduct,bestStore,bestMarketplace)}
      ${refCostCard(s,reserve,orders)}
      ${refUnlinkedCard(orders)}
      ${refStoreCards(orders)}
    </div>
    <button class="btn btn-export" onclick="exportAnalysisCsv()">Exportar CSV da análise</button>
  </div>`;
}
function orderQuantity(o){return Number(o.qty ?? o.quantity)||1;}
function orderGross(o){return Number(o.gross ?? o.financial?.grossAmount)||0;}
function orderNet(o){return Number(o.net ?? o.financial?.netReceived)||0;}
function resolveProductPerformanceUnitCost(product,orders){
  const explicit=explicitProductUnitCost(product);
  if(explicit>0)return explicit;
  const computed=productUnitCost(product,0).total;
  if(computed>0)return computed;
  const qty=(orders||[]).reduce((t,o)=>t+orderQuantity(o),0);
  const cost=(orders||[]).reduce((t,o)=>t+(Number(o.totalCost)||Number(o.productionCost?.totalCost)||0),0);
  return qty?cost/qty:0;
}
function auditProductCalculations(context,products){
  if(!Array.isArray(products))return;
  const rows=products.filter(p=>Number(p.qty)>0).map(p=>{
    const qty=Number(p.qty)||0, revenue=Number(p.net)||0, unit=Number(p.unitCost)||0;
    const expectedCost=unit*qty, expectedProfit=revenue-expectedCost, expectedMargin=revenue?expectedProfit/revenue*100:0;
    const cardCost=Number(p.cost)||0, cardProfit=Number(p.profit)||0, cardMargin=Number(p.margin)||0;
    return {
      contexto:context,
      nome:p.name,
      quantidade:qty,
      receita:Number(revenue.toFixed(2)),
      custo_unitario:Number(unit.toFixed(2)),
      custo_total:Number(expectedCost.toFixed(2)),
      lucro:Number(expectedProfit.toFixed(2)),
      margem:Number(expectedMargin.toFixed(2)),
      card_custo:Number(cardCost.toFixed(2)),
      card_lucro:Number(cardProfit.toFixed(2)),
      card_margem:Number(cardMargin.toFixed(2)),
      divergencia:(Math.abs(expectedCost-cardCost)>0.01||Math.abs(expectedProfit-cardProfit)>0.01||Math.abs(expectedMargin-cardMargin)>0.01)?'SIM':''
    };
  });
  if(!rows.length)return;
  window.__jm3dCalculationAudit={context,rows,updatedAt:new Date().toISOString()};
  window.__jm3dCalculationAudits=[...(window.__jm3dCalculationAudits||[]),window.__jm3dCalculationAudit].slice(-10);
  console.groupCollapsed(`[JM3D auditoria de calculo] ${context}`);
  console.log('[JM3D auditoria de calculo dados]', JSON.stringify(rows));
  console.table(rows);
  console.groupEnd();
}
function productPerformance(orders=filteredOrders()){
  const catalog=catalogProducts();
  const byId={};
  orders.forEach(o=>{
    const id=o.linkedProductId||slug(o.linkedProductName||o.productName||'sem_vinculo');
    const product=catalog.find(p=>p.id===id||p.legacyProductId===id)||{};
    byId[id]=byId[id]||{id,name:o.linkedProductName||product.name||o.productName||'Produto não vinculado',sku:product.sku||o.sku||'',itemId:o.itemId||o.importedItemId||firstMarketplaceId(product)||'',category:product.category||'',photo:product.photoUrl||product.photo||o.importedPhotoUrl||'',linked:!!o.linkedProductId,qty:0,orders:0,gross:0,net:0,cost:0,profit:0,margin:0,mainMarketplace:'',marketplaces:{}};
    byId[id].unitCost=byId[id].unitCost||0;
    byId[id].marketplaceStats=byId[id].marketplaceStats||{};
    byId[id]._orders=byId[id]._orders||[];
    const p=byId[id], qty=orderQuantity(o), m=marketplaceFilterValue(o.marketplace);
    p.qty+=qty; p.orders++; p.gross+=orderGross(o); p.net+=orderNet(o); p._orders.push(o);
    p.marketplaceStats[m]=p.marketplaceStats[m]||{qty:0,net:0,profit:0};
    p.marketplaceStats[m].qty+=qty;
    p.marketplaceStats[m].net+=orderNet(o);
  });
  Object.values(byId).forEach(p=>{
    const product=catalog.find(x=>x.id===p.id||x.legacyProductId===p.id)||{};
    p.unitCost=p.linked?resolveProductPerformanceUnitCost(product,p._orders):(p.qty?p.net/p.qty:0);
    p.cost=p.unitCost*p.qty;
    p.profit=p.net-p.cost;
    Object.keys(p.marketplaceStats||{}).forEach(m=>{
      const stat=p.marketplaceStats[m];
      stat.profit=stat.net-(p.unitCost*stat.qty);
      p.marketplaces[m]=stat.profit;
    });
    p.margin=p.net?p.profit/p.net*100:0;
    p.mainMarketplace=Object.keys(p.marketplaces).sort((a,b)=>p.marketplaces[b]-p.marketplaces[a])[0]||'-';
    delete p._orders;
  });
  catalog.forEach(product=>{
    const id=product.id||product.legacyProductId;
    if(!id||byId[id])return;
    byId[id]={id,name:product.name||'Produto sem nome',sku:product.sku||'',itemId:firstMarketplaceId(product),category:product.category||'',photo:product.photoUrl||product.photo||'',linked:true,qty:0,orders:0,gross:0,net:0,cost:0,profit:0,margin:0,unitCost:resolveProductPerformanceUnitCost(product,[]),mainMarketplace:mainMarketplaceFromProduct(product),marketplaces:{},marketplaceStats:{}};
  });
  return Object.values(byId);
}
function firstMarketplaceId(p){
  const ids=p?.marketplaceIds||{};
  return ids.shopee_kaline98||ids.shopee_mateusoliver98||ids.tiktokShop||ids.shopee||'';
}
function mainMarketplaceFromProduct(p){
  const links=p?.links||{};
  if(links.tiktokShop)return 'tiktokShop';
  if(links.shopee_kaline98||links.shopee_mateusoliver98)return 'shopee';
  return '-';
}
function renderErpProducts(){
  const content=document.getElementById('content'); if(!content)return;
  const orders=filteredOrders(), products=productPerformance(orders);
  const rankers={
    sold:(a,b)=>b.qty-a.qty, profit:(a,b)=>b.profit-a.profit, margin:(a,b)=>b.margin-a.margin,
    lowMargin:(a,b)=>a.margin-b.margin, loss:(a,b)=>a.profit-b.profit, slow:(a,b)=>a.qty-b.qty
  };
  const labels={sold:'Mais vendidos',profit:'Mais lucrativos',margin:'Maior margem',lowMargin:'Menor margem',loss:'Prejuízo',slow:'Pouca saída'};
  const query=norm(analysisFilters.q);
  const visibleProducts=query?products.filter(p=>norm(`${p.name} ${p.sku} ${p.itemId} ${p.category} ${p.mainMarketplace}`).includes(query)):products;
  const rows=[...visibleProducts].sort(rankers[erpProductRank]||rankers.profit);
  auditProductCalculations('Produtos',rows);
  content.innerHTML=`<div class="erp-shell">${erpFilters()}<section class="erp-section"><h3>Produtos <small>${rows.length} no catálogo</small></h3>
    <div class="compact-actions" style="margin-bottom:10px"><button class="btn btn-primary" onclick="openStoreProductImport()">Importar produtos da loja</button><button class="btn btn-secondary" onclick="swTab('calc')">Cadastrar manual</button></div>
    <div class="product-rank-tabs">${Object.keys(labels).map(k=>`<button class="${erpProductRank===k?'on':''}" onclick="erpProductRank='${k}';renderErpProducts()">${labels[k]}</button>`).join('')}</div>
    ${rows.length?rows.map(renderErpProductCard).join(''):'<div class="empty"><div class="ei">📦</div><div>Importe relatórios para ver performance por produto.</div></div>'}
  </section></div>`;
}
function renderErpProductCard(p){
  const h=erpHealth(p.margin,p.profit);
  const badges=[];
  if(p.qty>=20)badges.push(['good','🏆 Mais vendido']);
  if(p.profit>0&&p.margin>=45)badges.push(['good','⭐ Excelente']);
  if(p.margin<20)badges.push(['warn','⚠ Margem baixa']);
  if(p.profit<0)badges.push(['bad','📉 Prejuízo']);
  if(!p.linked)badges.push(['bad','Sem vínculo']);
  if(!p.sku)badges.push(['warn','Sem SKU']);
  return `<div class="erp-product-card" onclick="openProductDashboard('${ea(p.id)}')">
    ${p.photo?`<img class="erp-product-img" src="${p.photo}">`:'<div class="erp-product-fallback">📦</div>'}
    <div><div class="erp-product-name">${esc(p.name)}</div><div class="erp-product-meta">${marketplaceLabel(p.mainMarketplace)} • SKU ${esc(p.sku||'-')} • Item ${esc(p.itemId||'-')}<br>${p.qty} vendas</div>
    <div class="erp-product-numbers"><div><span>Recebido</span><b>${brl(p.net)}</b></div><div><span>Custou</span><b>${brl(p.cost)}</b></div><div><span>Lucro</span><b>${brl(p.profit)}</b></div><div><span>Margem</span><b>${pct(p.margin)}</b></div></div>
    <div class="badge-row"><span class="biz-badge ${h.cls}">${h.label}</span>${badges.map(([cls,b])=>`<span class="biz-badge ${cls}">${b}</span>`).join('')}</div></div>
  </div>`;
}
function renderUnlinkedOrdersPanel(orders){
  const unlinked=orders.filter(o=>!o.linkedProductId).slice(0,8);
  if(!unlinked.length)return '';
  return `<section class="erp-section"><h3>Produtos sem vínculo <small>${unlinked.length} pendentes</small></h3>${unlinked.map(o=>`<div class="order-card"><div class="order-top"><div><div class="order-title">${esc(o.productName||o.importedProductName||'Produto sem nome')}</div><div class="order-meta">${marketplaceLabel(o.marketplace)} • ${esc(o.store)} • Pedido ${esc(o.orderId)} • Item ${esc(o.itemId||o.importedItemId||'-')}</div></div><span class="health warn">${brl(o.net)}</span></div><button class="btn btn-secondary" style="font-size:12px;padding:9px" onclick="openLinkProduct('${o.id}')">Vincular</button></div>`).join('')}</section>`;
}
function openProductDashboard(id){
  const p=productPerformance(filteredOrders()).find(x=>x.id===id); if(!p)return;
  const c=productUnitCost((typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts).find(x=>x.id===id)||{},0);
  const ideal=priceForMargin(c.total,'shopee',30);
  openSht('Dashboard do produto',`${p.photo?`<img src="${p.photo}" style="width:100%;max-height:220px;object-fit:cover;border-radius:16px;margin-bottom:12px">`:''}
    <div class="erp-product-name">${esc(p.name)}</div><div class="muted-note">SKU ${esc(p.sku||'-')} • Categoria ${esc(p.category||'-')}</div>
    <div class="erp-grid">${erpCard('💰','Lucro acumulado',brl(p.profit),p.profit>=0?'good':'bad')}${erpCard('📦','Quantidade vendida',String(p.qty),'info')}${erpCard('📈','Margem',pct(p.margin),p.margin>=25?'good':'warn')}${erpCard('🛒','Mais rentável',marketplaceLabel(p.mainMarketplace),'info')}</div>
    <section class="erp-section" style="margin-top:10px"><h3>Preço e oportunidades</h3><div class="ideal-grid">${idealCard('Preço ideal',ideal)}${idealCard('+ R$ 2 no preço',p.qty*2)}${idealCard('+ R$ 5 no preço',p.qty*5)}${idealCard('Lucro unitário',p.qty?p.profit/p.qty:0)}</div></section>
    <button class="btn btn-secondary" onclick="closeSht()">Fechar</button>`);
}
function idealCard(label,value){return `<div class="ideal-card"><span>${label}</span><b>${brl(value)}</b></div>`;}
function renderImportsPage(){
  const content=document.getElementById('content'); if(!content)return;
  content.innerHTML=`<div class="erp-shell">${renderImportView()}${importReview?'<div id="import-review">'+renderImportReviewHtml()+'</div>':'<div id="import-review"></div>'}
  <section class="erp-section"><h3>Histórico de importações <small>${marketplaceImports.length} lotes</small></h3>${renderImportHistory()}</section></div>`;
}
function renderImportView(){
  return `<section class="erp-section"><h3>Importação em massa <small>Shopee + TikTok</small></h3>
    <div class="import-stepper"><div class="import-step"><b>1.</b> Selecione CSV e XLSX misturados</div><div class="import-step"><b>2.</b> O sistema categoriza e cruza pedidos</div><div class="import-step"><b>3.</b> Duplicados são ignorados e o lucro é calculado</div></div>
    <div class="file-drop"><input class="fi" type="file" id="market-files" accept=".xlsx,.csv" multiple onchange="processMarketplaceFiles([...this.files])"><div class="muted-note">Formatos suportados: Shopee Income, Shopee Order.completed, TikTok income e TikTok Enviado pedido CSV.</div></div>
  </section>`;
}
function renderImportHistory(){
  if(!marketplaceImports.length)return '<div class="empty"><div class="ei">📥</div><div>Nenhuma importação confirmada ainda.</div></div>';
  return marketplaceImports.slice(0,20).map(i=>`<div class="store-card"><div><b>${esc(i.name||i.id)}</b><div class="store-meta">${new Date(i.createdAt||Date.now()).toLocaleString('pt-BR')} • ${i.files||0} arquivos<br>${i.newOrders||0} novos • ${i.duplicates||0} duplicados • ${i.unlinked||0} pendentes</div></div><div class="store-val">${brl(i.net||0)}<br><small>${esc(i.status||'importado')}</small></div></div>`).join('');
}
function renderImportReviewHtml(){
  const r=importReview, orders=r.orders.map(o=>enrichOrder(o)), s=summarizeOrders(orders);
  return `<section class="erp-section"><h3>Resumo final da importação <small>revise antes de salvar</small></h3>
    <div class="erp-grid">${erpCard('📄','Arquivos analisados',String(r.summaries.length),'info')}${erpCard('🧾','Pedidos encontrados',String(orders.length),'info')}${erpCard('✅','Pedidos novos',String(r.newOrders.length),'good')}${erpCard('♻','Duplicados ignorados',String(r.duplicates.length),r.duplicates.length?'warn':'good')}${erpCard('🔗','Produtos vinculados',String(orders.filter(o=>o.linkedProductId).length),'good')}${erpCard('⚠','Produtos pendentes',String(orders.filter(o=>!o.linkedProductId).length),orders.some(o=>!o.linkedProductId)?'warn':'good')}${erpCard('💳','Valor importado',brl(s.net),'info')}${erpCard('💼','Lucro calculado',brl(s.profit),s.profit>=0?'good':'bad')}</div>
    <div class="review-list" style="margin-top:10px"><table class="mini-table"><thead><tr><th>Arquivo</th><th>Tipo</th><th>Pedidos</th></tr></thead><tbody>${r.summaries.map(x=>`<tr><td>${esc(x.file)}</td><td>${esc(x.kind)}</td><td class="num">${x.orders}</td></tr>`).join('')}</tbody></table></div>
    ${r.errors.length?`<div class="muted-note" style="color:var(--red)">${r.errors.map(esc).join('<br>')}</div>`:''}
    <button class="btn btn-save" onclick="confirmMarketplaceImport()">Confirmar importação</button><button class="btn btn-secondary" onclick="analysisFilters.health='unlinked';swTab('produtos')">Corrigir vínculos</button><button class="btn btn-danger" onclick="importReview=null;renderContent()">Cancelar</button>
  </section>`;
}
async function confirmMarketplaceImport(){
  if(!importReview)return;
  const batchId='batch_'+Date.now();
  const save=importReview.newOrders.map(o=>Object.assign(enrichOrder(o),{source:Object.assign({},o.source||{},{importBatchId:batchId})}));
  const all=importReview.orders.map(o=>enrichOrder(o)), s=summarizeOrders(all);
  const record=sanitizeForFirebase({id:batchId,name:'Importação '+new Date().toLocaleString('pt-BR'),createdAt:Date.now(),files:importReview.summaries.length,summaries:importReview.summaries,newOrders:importReview.newOrders.length,duplicates:importReview.duplicates.length,errors:importReview.errors,unlinked:all.filter(o=>!o.linkedProductId).length,gross:s.gross,net:s.net,profit:s.profit,status:'importado'});
  try{
    if(marketplaceOrdersRef){
      const updates={}; save.forEach(o=>updates[o.id]=sanitizeForFirebase(o));
      await marketplaceOrdersRef.update(updates);
      if(marketplaceImportsRef)await marketplaceImportsRef.child(batchId).set(record);
    }else{
      marketplaceOrders=[...marketplaceOrders,...save];
      marketplaceImports=[record,...marketplaceImports];
      localStorage.setItem('jm3d_market_orders',JSON.stringify(marketplaceOrders));
      localStorage.setItem('jm3d_market_imports',JSON.stringify(marketplaceImports));
    }
    setSyncStatus('ok',`✅ ${save.length} pedidos importados`);
    importReview=null; swTab('financeiro');
  }catch(e){setSyncStatus('err','Erro ao salvar importação: '+e.message);}
}
function renderPricingView(){
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  if(!erpPricingProductId&&catalog[0])erpPricingProductId=catalog[0].id;
  const product=catalog.find(p=>p.id===erpPricingProductId)||catalog[0]||null;
  const cost=product?productUnitCost(product,0).total:0;
  const price=Number(erpPricingPrice)||0;
  return `<div class="erp-shell"><section class="erp-section"><h3>Simulador profissional <small>quanto sobra se vender por...</small></h3>
    <select class="fi" onchange="erpPricingProductId=this.value;renderPricingView()">${catalog.map(p=>`<option value="${p.id}" ${product&&product.id===p.id?'selected':''}>${esc(p.name)}</option>`).join('')}</select>
    <label class="fl">Preço de venda</label><input class="fi" type="number" value="${price}" oninput="erpPricingPrice=parseMoneyBR(this.value);renderPricingView()">
    <div class="pricing-sim-grid">${renderChannelSim('Shopee','shopee',price,cost)}${renderChannelSim('TikTok Shop','tiktok',price,cost)}${renderChannelSim('Venda direta','direct',price,cost)}</div>
  </section><section class="erp-section"><h3>Preço ideal <small>linguagem de negócio</small></h3><div class="ideal-grid">${idealCard('Preço mínimo',priceForMargin(cost,'shopee',0))}${idealCard('Preço saudável',priceForMargin(cost,'shopee',20))}${idealCard('Preço recomendado',priceForMargin(cost,'shopee',30))}${idealCard('Preço premium',priceForMargin(cost,'shopee',50))}</div></section>
  <button class="btn btn-primary" onclick="swTab('calc')">Abrir calculadora/cadastro completo</button>${renderPricingProducts()}</div>`;
}
function renderChannelSim(label,marketplace,price,cost){
  const f=feeRuleFor(marketplace), fees=price*(f.pct/100)+f.fixed, net=price-fees, profit=net-cost, margin=net?profit/net*100:0, h=erpHealth(margin,profit);
  return `<div class="pricing-sim-card"><h4>${label}</h4><div class="sim-line"><span>Preço</span><b>${brl(price)}</b></div><div class="sim-line"><span>Taxas</span><b>${brl(fees)}</b></div><div class="sim-line"><span>Recebe</span><b>${brl(net)}</b></div><div class="sim-line"><span>Custos</span><b>${brl(cost)}</b></div><div class="sim-line"><span>Lucro</span><b>${brl(profit)}</b></div><div class="sim-line"><span>Margem</span><b>${pct(margin)}</b></div><div class="badge-row"><span class="biz-badge ${h.cls}">${h.label}</span></div></div>`;
}
function cfgField(id,label,value,hint,step='0.01'){
  return `<label class="cfg-field"><span>${label}</span><input class="fi" id="${id}" type="number" step="${step}" value="${ea(value)}"><small>${hint}</small></label>`;
}
function renderSettingsPage(){
  const cfg=getMarketplaceSettings();
  return `<div class="erp-shell"><section class="erp-section"><h3>Configurações financeiras <small>custos globais</small></h3>
    <div class="filter-row"><input class="fi" id="mc-fil" type="number" value="${cfg.filamentKgPrice}" placeholder="Filamento R$/kg"><input class="fi" id="mc-kwh" type="number" value="${cfg.energyKwhPrice}" placeholder="Energia R$/kWh"></div>
    <div class="filter-row"><input class="fi" id="mc-cons" type="number" value="${cfg.printerKwhHour}" placeholder="kWh/h impressora"><input class="fi" id="mc-das" type="number" value="${cfg.dasMeiMonthly}" placeholder="DAS MEI"></div>
    <div class="filter-row"><input class="fi" id="mc-loss" type="number" value="${cfg.lossPct}" placeholder="Perdas %"><input class="fi" id="mc-post" type="number" value="${cfg.postProcessCost}" placeholder="Pós-processamento"></div>
    <div class="filter-row"><input class="fi" id="mc-shp" type="number" value="${cfg.shopeeCommissionPct}" placeholder="Shopee %"><input class="fi" id="mc-tkt" type="number" value="${cfg.tiktokCommissionPct}" placeholder="TikTok %"></div>
    <button class="btn btn-save" onclick="saveMarketCfgUI()">Salvar custos globais</button><button class="btn btn-secondary" onclick="openConfig()">Abrir configurações antigas</button>
  </section><section class="erp-section"><h3>Lojas ativas</h3><div class="store-grid">${erpStores().filter(x=>x[0]!=='all').map(([id,label])=>`<div class="store-card"><div><b>${label}</b><div class="store-meta">Usada em filtros, dashboards e importações.</div></div><div class="store-val">Ativa</div></div>`).join('')}</div></section></div>`;
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
  setSyncStatus('ok','✅ Custos globais salvos');renderContent();
}
function marketplaceProductUrlInfo(rawUrl){
  const info={url:cleanCell(rawUrl),marketplace:'',store:'',storeKey:'',itemId:'',shopId:'',price:0};
  if(!info.url)return info;
  try{
    const u=new URL(info.url);
    const host=norm(u.hostname);
    const pathParts=u.pathname.split('/').filter(Boolean).map(x=>decodeURIComponent(x));
    const search=u.searchParams;
    if(host.includes('shopee')){
      info.marketplace='shopee';
      info.itemId=cleanCell(search.get('itemId')||search.get('itemid')||'');
      info.shopId=cleanCell(search.get('shopId')||search.get('shopid')||'');
      const path=decodeURIComponent(u.pathname);
      const iMatch=path.match(/(?:^|[.\-/])i[.\-/](\d+)[.\-/](\d+)/i)||path.match(/(\d{5,})[.\-/](\d{5,})$/);
      if(iMatch){info.shopId=info.shopId||iMatch[1];info.itemId=info.itemId||iMatch[2];}
      const storeName=(pathParts[0]||'').replace(/^@/,'');
      info.store=storeName||'Shopee';
      if(norm(storeName).includes('kaline'))info.storeKey='shopee_kaline98';
      else if(norm(storeName).includes('mateus'))info.storeKey='shopee_mateusoliver98';
      else info.storeKey='shopee';
    }else if(host.includes('tiktok')){
      info.marketplace='tiktokShop';
      info.store='TikTok Shop';
      info.storeKey='tiktokShop';
      info.itemId=cleanCell(search.get('itemId')||search.get('item_id')||pathParts.find(p=>/^\d{6,}$/.test(p))||'');
      info.shopId=cleanCell(search.get('shopId')||search.get('shop_id')||'');
    }
  }catch(e){}
  return info;
}
function automaticAliases(name,sku,itemId){
  const base=[name,sku,itemId].filter(Boolean).map(cleanCell);
  const words=norm(name).split(' ').filter(w=>w.length>3);
  for(let i=2;i<=Math.min(5,words.length);i++)base.push(words.slice(0,i).join(' '));
  return [...new Set(base.filter(Boolean))];
}
async function inspectStoreProductUrl(){
  const url=document.getElementById('store-product-url')?.value||'';
  const info=marketplaceProductUrlInfo(url);
  storeProductDraft=Object.assign({},storeProductDraft||{},info);
  const status=document.getElementById('store-product-status');
  if(status)status.textContent='Analisando link...';
  try{
    const res=await fetch(url,{mode:'cors',credentials:'omit'});
    const html=await res.text();
    const meta=(prop)=>{
      const re=new RegExp(`<meta[^>]+(?:property|name)=["']${prop}["'][^>]+content=["']([^"']+)["']`,'i');
      return cleanCell((html.match(re)||[])[1]||'');
    };
    storeProductDraft.name=meta('og:title')||storeProductDraft.name||'';
    storeProductDraft.photoUrl=meta('og:image')||storeProductDraft.photoUrl||'';
    storeProductDraft.price=parseMoneyBR(meta('product:price:amount')||meta('og:price:amount'));
    fillStoreProductForm(storeProductDraft);
    if(status)status.textContent='Dados encontrados. Revise antes de salvar.';
  }catch(e){
    fillStoreProductForm(storeProductDraft);
    if(status)status.textContent='A Shopee pode bloquear captura automática. Preencha/revise manualmente e salve com link, foto e itemId.';
  }
}
function fillStoreProductForm(d){
  const set=(id,v)=>{const el=document.getElementById(id);if(el&&v!==undefined)el.value=v||'';};
  set('store-product-name',d.name);
  set('store-product-sku',d.sku);
  set('store-product-item',d.itemId);
  set('store-product-shop',d.shopId);
  set('store-product-photo',d.photoUrl);
  set('store-product-price',d.price||'');
  const store=document.getElementById('store-product-store'); if(store&&d.storeKey)store.value=d.storeKey;
}
function onStoreProductPhotoSelected(ev){
  const f=ev.target.files&&ev.target.files[0]; if(!f)return;
  const rd=new FileReader();
  rd.onload=()=>{storeProductDraft=Object.assign({},storeProductDraft||{},{photoUrl:rd.result});fillStoreProductForm(storeProductDraft);};
  rd.readAsDataURL(f);
}
function openStoreProductImport(){
  storeProductDraft={};
  openSht('Importar produto da loja',`<div class="muted-note">Cole um link da Shopee/TikTok ou cadastre manualmente. O catálogo salvo usa itemId, SKU, link e foto para vincular pedidos sem depender só do nome.</div>
    <label class="fl">Link da loja ou anúncio</label><input class="fi" id="store-product-url" placeholder="https://shopee.com.br/...itemId=..." oninput="storeProductDraft=Object.assign({},storeProductDraft||{},marketplaceProductUrlInfo(this.value))">
    <div class="compact-actions"><button class="btn btn-primary" onclick="inspectStoreProductUrl()">Tentar capturar dados</button><button class="btn btn-secondary" onclick="document.getElementById('store-photo-file').click()">Anexar foto</button></div>
    <input type="file" id="store-photo-file" accept="image/*" style="display:none" onchange="onStoreProductPhotoSelected(event)">
    <div class="muted-note" id="store-product-status">Se a captura for bloqueada, preencha os campos abaixo.</div>
    <label class="fl">Nome real do anúncio</label><input class="fi" id="store-product-name" placeholder="Nome real do produto">
    <div class="filter-row"><input class="fi" id="store-product-sku" placeholder="SKU"><input class="fi" id="store-product-item" placeholder="ItemId do anúncio"></div>
    <div class="filter-row"><input class="fi" id="store-product-shop" placeholder="ShopId"><input class="fi" id="store-product-price" type="number" placeholder="Preço anunciado"></div>
    <label class="fl">Foto principal</label><input class="fi" id="store-product-photo" placeholder="URL da foto ou anexe uma imagem">
    <label class="fl">Loja</label><select class="fi" id="store-product-store"><option value="shopee_kaline98">Shopee _kaline98</option><option value="shopee_mateusoliver98">Shopee mateusoliver98</option><option value="tiktokShop">TikTok Shop</option><option value="shopee">Shopee genérica</option></select>
    <label class="fl">Aliases</label><textarea class="fi" id="store-product-aliases" rows="3" placeholder="Um alias por linha"></textarea>
    <button class="btn btn-save" onclick="saveStoreProduct()">Salvar no catálogo</button><button class="btn btn-secondary" onclick="closeSht()">Cancelar</button>`);
}
async function saveStoreProduct(){
  const url=cleanCell(document.getElementById('store-product-url')?.value);
  const info=Object.assign({},marketplaceProductUrlInfo(url),storeProductDraft||{});
  const name=cleanCell(document.getElementById('store-product-name')?.value);
  const sku=cleanCell(document.getElementById('store-product-sku')?.value);
  const itemId=cleanCell(document.getElementById('store-product-item')?.value||info.itemId);
  const shopId=cleanCell(document.getElementById('store-product-shop')?.value||info.shopId);
  const photoUrl=cleanCell(document.getElementById('store-product-photo')?.value||info.photoUrl);
  const storeKey=document.getElementById('store-product-store')?.value||info.storeKey||'shopee';
  const aliasesManual=String(document.getElementById('store-product-aliases')?.value||'').split('\n').map(cleanCell).filter(Boolean);
  if(!name){alert('Informe o nome real do anúncio.');return;}
  if(!photoUrl){alert('Informe ou anexe a foto principal do produto.');return;}
  if(!sku&&!itemId&&!url){alert('Informe SKU, itemId ou link do anúncio para reconhecer pedidos automaticamente.');return;}
  const links={}, marketplaceIds={}, shopIds={};
  if(url)links[storeKey]=url;
  if(itemId)marketplaceIds[storeKey]=itemId;
  if(shopId)shopIds[storeKey]=shopId;
  const now=new Date().toISOString();
  const id='prod_'+slug(sku||itemId||name);
  const old=(catalogProducts().find(p=>p.id===id)||{});
  const oldUnitCost=explicitProductUnitCost(old);
  const channel=storeKey.startsWith('tiktok')?'tiktokShop':'shopee';
  const oldSettings=old.marketplaceSettings||{};
  const product=Object.assign({},old,{
    id,name,sku:sku||old.sku||'',photoUrl:photoUrl||old.photoUrl||old.photo||'',category:old.category||'',
    totalCost:oldUnitCost||0,
    unitCost:oldUnitCost||0,
    notes:old.notes||'Produto importado/cadastrado a partir da loja.',
    aliases:[...new Set([...(old.aliases||[]),...automaticAliases(name,sku,itemId),...aliasesManual])],
    links:Object.assign({},old.links||{},links),
    marketplaceIds:Object.assign({},old.marketplaceIds||{},marketplaceIds),
    shopIds:Object.assign({},old.shopIds||{},shopIds),
    variations:old.variations||[],
    source:Object.assign({},old.source||{},{type:'store_product_import',url,lastStoreKey:storeKey}),
    marketplaceSettings:Object.assign({},oldSettings,{[channel]:Object.assign({},oldSettings[channel]||{},{active:true,salePrice:parseMoneyBR(document.getElementById('store-product-price')?.value)})}),
    costs:Object.assign({},old.costs||{},oldUnitCost?{totalUnitCost:oldUnitCost,unitCost:oldUnitCost,totalCost:oldUnitCost}:{}),
    createdAt:old.createdAt||now,updatedAt:now
  });
  try{
    if(productCatalogRef)await productCatalogRef.child(product.id).set(sanitizeForFirebase(product));
    productCatalog=[product,...productCatalog.filter(p=>p.id!==product.id)];
    localStorage.setItem('jm3d_product_catalog',JSON.stringify(productCatalog));
    setSyncStatus('ok','Produto da loja salvo no catálogo');
    closeSht();renderContent();
  }catch(e){
    productCatalog=[product,...productCatalog.filter(p=>p.id!==product.id)];
    localStorage.setItem('jm3d_product_catalog',JSON.stringify(productCatalog));
    setSyncStatus('err','Produto salvo localmente. Firebase será sincronizado quando possível.');
    closeSht();renderContent();
  }
}
function photoSignal(v){
  const s=String(v||'');
  if(!s)return '';
  return slug(s.replace(/^data:image\/[^;]+;base64,/,'').slice(0,160)+'_'+s.length);
}
function findProductMatchLegacyFinal(o){
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
  const item=norm(o.itemId||o.importedItemId);
  if(item){
    const p=catalog.find(x=>Object.values(x.links||{}).some(link=>norm(link).includes(item))||norm(x.sku).includes(item));
    if(p)return {product:p,confidence:.9,method:'itemId/link'};
  }
  const importedUrl=norm(o.importedProductUrl||o.productUrl||'');
  if(importedUrl){
    const p=catalog.find(x=>Object.values(x.links||{}).some(link=>link&&norm(link)===importedUrl));
    if(p)return {product:p,confidence:.88,method:'url do anúncio'};
  }
  const on=norm(o.productName);
  if(on){
    for(const p of catalog){
      const aliases=[...(p.aliases||[]),p.name].map(norm);
      if(aliases.some(a=>a&&on.includes(a)))return {product:p,confidence:.86,method:'alias'};
    }
  }
  const importedPhoto=o.importedPhotoUrl||o.photoUrl||'';
  if(importedPhoto){
    const sig=photoSignal(importedPhoto);
    const p=catalog.find(x=>{
      const photo=x.photoUrl||x.photo||'';
      return photo&&((photo===importedPhoto)||photoSignal(photo)===sig);
    });
    if(p)return {product:p,confidence:.74,method:'foto semelhante'};
  }
  if(!on)return {product:null,confidence:0,method:'none'};
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

/* Full-width graphite/cyan SaaS UI overrides */
function jmPageHead(eyebrow,title,subtitle,actions=''){
  return `<section class="ref-card jm-page-head"><div class="jm-page-title"><span>${esc(eyebrow)}</span><h2>${esc(title)}</h2><p>${esc(subtitle)}</p></div>${actions?`<div class="jm-page-actions">${actions}</div>`:''}</section>`;
}
function jmInitials(text){
  return esc(String(text||'JM').trim().slice(0,2).toUpperCase()||'JM');
}
function renderErpProducts(){
  const content=document.getElementById('content'); if(!content)return;
  const orders=filteredOrders(), products=productPerformance(orders);
  const rankers={
    sold:(a,b)=>b.qty-a.qty, profit:(a,b)=>b.profit-a.profit, margin:(a,b)=>b.margin-a.margin,
    lowMargin:(a,b)=>a.margin-b.margin, loss:(a,b)=>a.profit-b.profit, slow:(a,b)=>a.qty-b.qty
  };
  const labels={sold:'Mais vendidos',profit:'Mais lucrativos',margin:'Maior margem',lowMargin:'Menor margem',loss:'Prejuízo',slow:'Pouca saída'};
  const query=norm(analysisFilters.q);
  const visibleProducts=query?products.filter(p=>norm(`${p.name} ${p.sku} ${p.itemId} ${p.category} ${p.mainMarketplace}`).includes(query)):products;
  const rows=[...visibleProducts].sort(rankers[erpProductRank]||rankers.profit);
  auditProductCalculations('Produtos',rows);
  const actions=`<button class="btn btn-primary" onclick="openStoreProductImport()">Importar produtos da loja</button><button class="btn btn-secondary" onclick="swTab('calc')">Cadastrar manual</button>`;
  content.innerHTML=`<div class="ref-dashboard jm-page jm-products-page">
    ${jmPageHead('Produtos','Central de produtos','Fotos reais, SKU, itemId, vendas, custos e margem por produto.',actions)}
    ${refFilters()}
    <section class="ref-card jm-panel"><h3>Catálogo de gestão <small>${rows.length} produtos</small></h3>
      <div class="product-rank-tabs">${Object.keys(labels).map(k=>`<button class="${erpProductRank===k?'on':''}" onclick="erpProductRank='${k}';renderErpProducts()">${labels[k]}</button>`).join('')}</div>
      ${rows.length?`<div class="jm-products-grid">${rows.map(renderErpProductCard).join('')}</div>`:'<div class="empty"><div class="ei">JM</div><div>Importe relatórios para ver performance por produto.</div></div>'}
    </section>
  </div>`;
}
function renderErpProductCard(p){
  const h=erpHealth(p.margin,p.profit);
  const badges=[];
  if(p.qty>=20)badges.push(['good','Mais vendido']);
  if(p.profit>0&&p.margin>=45)badges.push(['good','Excelente']);
  if(p.margin<20)badges.push(['warn','Margem baixa']);
  if(p.profit<0)badges.push(['bad','Prejuízo']);
  if(!p.linked)badges.push(['bad','Sem vínculo']);
  if(!p.sku)badges.push(['warn','Sem SKU']);
  return `<div class="erp-product-card" onclick="openProductDashboard('${ea(p.id)}')">
    ${p.photo?`<img class="erp-product-img" src="${ea(p.photo)}" alt="${ea(p.name)}">`:`<div class="erp-product-fallback">${jmInitials(p.name)}</div>`}
    <div><div class="erp-product-name">${esc(p.name)}</div><div class="erp-product-meta">${marketplaceLabel(p.mainMarketplace)} | SKU ${esc(p.sku||'-')} | Item ${esc(p.itemId||'-')}<br>${p.qty} vendas</div>
    <div class="erp-product-numbers"><div><span>Recebido</span><b>${brl(p.net)}</b></div><div><span>Custou</span><b>${brl(p.cost)}</b></div><div><span>Lucro</span><b>${brl(p.profit)}</b></div><div><span>Margem</span><b>${pct(p.margin)}</b></div></div>
    <div class="badge-row"><span class="biz-badge ${h.cls}">${h.label}</span>${badges.map(([cls,b])=>`<span class="biz-badge ${cls}">${b}</span>`).join('')}</div></div>
  </div>`;
}
function renderImportsPage(){
  const content=document.getElementById('content'); if(!content)return;
  content.innerHTML=`<div class="ref-dashboard jm-page jm-imports-page">
    ${jmPageHead('Importações','Importação em massa','CSV e XLSX misturados da Shopee e TikTok, com revisão antes de salvar.')}
    <div class="jm-import-grid">
      <div class="jm-import-main">${renderImportView()}${importReview?`<div id="import-review">${renderImportReviewHtml()}</div>`:'<div id="import-review"></div>'}</div>
      <section class="ref-card jm-panel jm-import-side"><h3>Histórico de importações <small>${marketplaceImports.length} lotes</small></h3>${renderImportHistory()}</section>
    </div>
  </div>`;
}
function renderImportView(){
  return `<section class="ref-card jm-panel"><h3>Selecionar arquivos <small>Shopee + TikTok</small></h3>
    <div class="import-stepper"><div class="import-step"><b>1.</b> Selecione vários CSV/XLSX juntos.</div><div class="import-step"><b>2.</b> O app categoriza marketplace, loja e período.</div><div class="import-step"><b>3.</b> Duplicados são ignorados e o lucro é calculado.</div></div>
    <div class="file-drop"><input class="fi" type="file" id="market-files" accept=".xlsx,.csv" multiple onchange="processMarketplaceFiles([...this.files])"><div class="muted-note">Suporta Shopee Income, Order.completed, Order.shipping, Enviado pedido CSV e TikTok income.</div></div>
  </section>`;
}
function renderImportHistory(){
  if(!marketplaceImports.length)return '<div class="empty"><div class="ei">JM</div><div>Nenhuma importação confirmada ainda.</div></div>';
  return `<div class="jm-history-list">${marketplaceImports.slice(0,20).map(i=>`<div class="store-card"><div><b>${esc(i.name||i.id)}</b><div class="store-meta">${new Date(i.createdAt||Date.now()).toLocaleString('pt-BR')} | ${i.files||0} arquivos<br>${i.newOrders||0} novos | ${i.duplicates||0} duplicados | ${i.unlinked||0} pendentes</div></div><div class="store-val">${brl(i.net||0)}<br><small>${esc(i.status||'importado')}</small></div></div>`).join('')}</div>`;
}
function renderImportReviewHtml(){
  const r=importReview, orders=r.orders.map(o=>enrichOrder(o)), s=summarizeOrders(orders);
  return `<section class="ref-card jm-panel"><h3>Resumo final da importação <small>revise antes de salvar</small></h3>
    <div class="erp-grid">${erpCard('ARQ','Arquivos analisados',String(r.summaries.length),'info')}${erpCard('PED','Pedidos encontrados',String(orders.length),'info')}${erpCard('NOV','Pedidos novos',String(r.newOrders.length),'good')}${erpCard('DUP','Duplicados ignorados',String(r.duplicates.length),r.duplicates.length?'warn':'good')}${erpCard('VIN','Produtos vinculados',String(orders.filter(o=>o.linkedProductId).length),'good')}${erpCard('PEN','Produtos pendentes',String(orders.filter(o=>!o.linkedProductId).length),orders.some(o=>!o.linkedProductId)?'warn':'good')}${erpCard('R$','Valor importado',brl(s.net),'info')}${erpCard('LUC','Lucro calculado',brl(s.profit),s.profit>=0?'good':'bad')}</div>
    <div class="review-list" style="margin-top:10px"><table class="mini-table"><thead><tr><th>Arquivo</th><th>Tipo</th><th>Pedidos</th></tr></thead><tbody>${r.summaries.map(x=>`<tr><td>${esc(x.file)}</td><td>${esc(x.kind)}</td><td class="num">${x.orders}</td></tr>`).join('')}</tbody></table></div>
    ${r.errors.length?`<div class="muted-note" style="color:var(--red)">${r.errors.map(esc).join('<br>')}</div>`:''}
    <div class="jm-action-row"><button class="btn btn-save" onclick="confirmMarketplaceImport()">Confirmar importação</button><button class="btn btn-secondary" onclick="analysisFilters.health='unlinked';swTab('produtos')">Corrigir vínculos</button><button class="btn btn-danger" onclick="importReview=null;renderContent()">Cancelar</button></div>
  </section>`;
}
function renderPricingView(){
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  if(!erpPricingProductId&&catalog[0])erpPricingProductId=catalog[0].id;
  const product=catalog.find(p=>p.id===erpPricingProductId)||catalog[0]||null;
  const cost=product?productUnitCost(product,0).total:0;
  const price=Number(erpPricingPrice)||0;
  const actions=`<button class="btn btn-primary" onclick="swTab('calc')">Abrir cadastro completo</button>`;
  return `<div class="ref-dashboard jm-page jm-pricing-page">
    ${jmPageHead('Precificação','Simulador profissional','Compare Shopee, TikTok e venda direta pelo dinheiro que realmente sobra.',actions)}
    <div class="jm-pricing-grid">
      <section class="ref-card jm-panel jm-pricing-main"><h3>Quanto sobra se vender por...</h3>
        <select class="fi" onchange="erpPricingProductId=this.value;renderContent()">${catalog.map(p=>`<option value="${p.id}" ${product&&product.id===p.id?'selected':''}>${esc(p.name)}</option>`).join('')}</select>
        <label class="fl">Preço de venda</label><input class="fi" type="number" value="${price}" oninput="erpPricingPrice=parseMoneyBR(this.value);renderContent()">
        <div class="pricing-sim-grid">${renderChannelSim('Shopee','shopee',price,cost)}${renderChannelSim('TikTok Shop','tiktok',price,cost)}${renderChannelSim('Venda direta','direct',price,cost)}</div>
      </section>
      <section class="ref-card jm-panel jm-pricing-side"><h3>Preço ideal <small>Shopee como base</small></h3><div class="ideal-grid jm-ideal-grid">${idealCard('Preço mínimo',priceForMargin(cost,'shopee',0))}${idealCard('Preço saudável',priceForMargin(cost,'shopee',20))}${idealCard('Recomendado',priceForMargin(cost,'shopee',30))}${idealCard('Premium',priceForMargin(cost,'shopee',50))}</div></section>
    </div>
    ${renderPricingProducts()}
  </div>`;
}
function renderPricingProducts(){
  const catalog=typeof mergedProductCatalog==='function'?mergedProductCatalog():localProducts;
  if(!catalog.length)return '<section class="ref-card jm-panel"><div class="empty"><div class="ei">JM</div><div>Cadastre produtos para simular preços profissionais.</div></div></section>';
  const query=norm(analysisFilters.q);
  const rows=(query?catalog.filter(p=>norm(`${p.name} ${p.sku} ${firstMarketplaceId(p)} ${p.category}`).includes(query)):catalog).slice(0,20);
  return `<section class="ref-card jm-panel"><h3>Comparacao real por produto <small>${rows.length} de ${catalog.length} no catalogo</small></h3>
    <div class="pricing-tools"><div class="ref-searchbar"><input class="fi" placeholder="Buscar produto, SKU ou item..." value="${ea(analysisFilters.q)}" oninput="analysisFilters.q=this.value;renderContent()" autocomplete="off"><button class="btn btn-secondary" onclick="analysisFilters.q='';renderContent()">Limpar</button></div>${pricingLegendHtml()}</div>
    <div class="pricing-help-grid">
      <div><b>Minimo</b><span>preco para nao ter prejuizo.</span></div>
      <div><b>Shopee 30%</b><span>preco sugerido com taxas da Shopee e 30% de margem.</span></div>
      <div><b>TikTok 30%</b><span>preco sugerido com taxas do TikTok e 30% de margem.</span></div>
      <div><b>Status</b><span>saude do preco atual cadastrado.</span></div>
    </div>
    <div class="jm-products-grid">${rows.length?rows.map(renderPricingProductCard).join(''):'<div class="empty"><div class="ei">JM</div><div>Nenhum produto encontrado para a busca.</div></div>'}</div></section>`;
}
function pricingCostParts(c){
  const total=Number(c.total)||0;
  const packaging=(Number(c.packaging)||0)+(Number(c.bubble)||0);
  const known=(Number(c.filament)||0)+(Number(c.energy)||0)+packaging;
  const other=Math.max(0,total-known);
  return [
    {label:'Filamento',hint:'material consumido pela peca',value:Number(c.filament)||0,color:'#00C2FF'},
    {label:'Energia',hint:'kWh da impressora no tempo de impressao',value:Number(c.energy)||0,color:'#F97316'},
    {label:'Embalagem',hint:'embalagem + plastico bolha',value:packaging,color:'#3B82F6'},
    {label:'Outros',hint:'perdas, manutencao, acabamento e adicionais',value:other,color:'#A855F7'}
  ];
}
function pricingLegendHtml(){
  return `<div class="cost-legend-panel"><div class="legend-title">Legenda dos custos</div><div class="cost-legend-row">
    ${pricingCostParts({}).map(p=>`<span><i style="background:${p.color}"></i><b>${p.label}</b><small>${p.hint}</small></span>`).join('')}
  </div></div>`;
}
function renderPricingProductCard(p){
  const c=productUnitCost(p,0), min=priceForMargin(c.total,'shopee',0), sh30=priceForMargin(c.total,'shopee',30), tk30=priceForMargin(c.total,'tiktok',30);
  const current=Number(p.marketplaceSettings?.shopee?.salePrice ?? p.price)||0, health=current&&current<min?'bad':current&&current<priceForMargin(c.total,'shopee',10)?'warn':'ok';
  const weight=p.costs?.weightGrams??p.weight??0, hours=p.costs?.printTimeHours??p.printH??0;
  const costParts=pricingCostParts(c), total=Number(c.total)||0;
  return `<div class="erp-product-card"><div>${p.photoUrl||p.photo?`<img class="erp-product-img" src="${ea(p.photoUrl||p.photo)}" alt="${ea(p.name)}">`:`<div class="erp-product-fallback">${jmInitials(p.name)}</div>`}</div><div><div class="erp-product-name">${esc(p.name)}</div><div class="erp-product-meta">SKU ${esc(p.sku||'-')} | custo técnico ${brl(c.total)} | ${weight}g | ${hours}h</div>
    <div class="erp-product-numbers"><div><span>Mínimo</span><b>${brl(min)}</b></div><div><span>Shopee 30%</span><b>${brl(sh30)}</b></div><div><span>TikTok 30%</span><b>${brl(tk30)}</b></div><div><span>Status</span><b>${health==='bad'?'Baixo':health==='warn'?'Atenção':'Saudável'}</b></div></div>
    <div class="cost-bars" title="Composicao do custo tecnico">${costParts.map(part=>`<span style="width:${Math.max(0.5,Math.min(100,total?part.value/total*100:0))}%;background:${part.color}"></span>`).join('')}</div>
    <div class="cost-card-legend">${costParts.filter(p=>p.value>0).map(part=>`<span><i style="background:${part.color}"></i>${part.label} ${brl(part.value)}</span>`).join('')||'<span>Custo tecnico nao detalhado</span>'}</div></div></div>`;
}
function renderSettingsPage(){
  const cfg=getMarketplaceSettings();
  return `<div class="ref-dashboard jm-page jm-settings-page">
    ${jmPageHead('Configurações','Custos e lojas','Ajustes globais usados em precificação, importação e dashboard financeiro.')}
    <div class="jm-settings-grid">
      <section class="ref-card jm-panel jm-settings-main"><h3>Custos globais <small>produção e fiscal</small></h3>
        <div class="cfg-field-grid">
          ${cfgField('mc-fil','Filamento R$/kg',cfg.filamentKgPrice,'Preco medio do kg de filamento usado no custo tecnico.')}
          ${cfgField('mc-kwh','Energia R$/kWh',cfg.energyKwhPrice,'Valor do kWh da sua conta de luz.')}
          ${cfgField('mc-cons','Consumo kWh/h',cfg.printerKwhHour,'Consumo medio da impressora por hora de impressao.')}
          ${cfgField('mc-das','DAS MEI mensal',cfg.dasMeiMonthly,'Valor mensal rateado por unidade ou pedido conforme configuracao.')}
          ${cfgField('mc-loss','Perdas/refugo %',cfg.lossPct,'Percentual para falhas, testes e margem de erro da impressao.')}
          ${cfgField('mc-post','Pos-processamento R$',cfg.postProcessCost,'Cola, pintura, acabamento ou trabalho manual por unidade.')}
          ${cfgField('mc-shp','Taxa Shopee %',cfg.shopeeCommissionPct,'Percentual usado nas simulacoes da Shopee.')}
          ${cfgField('mc-tkt','Taxa TikTok %',cfg.tiktokCommissionPct,'Percentual usado nas simulacoes do TikTok Shop.')}
        </div>
        ${pricingLegendHtml()}
        <div class="jm-action-row"><button class="btn btn-save" onclick="saveMarketCfgUI()">Salvar custos globais</button><button class="btn btn-secondary" onclick="openConfig()">Abrir configurações antigas</button></div>
      </section>
      <section class="ref-card jm-panel jm-settings-side"><h3>Lojas ativas</h3><div class="store-grid">${erpStores().filter(x=>x[0]!=='all').map(([id,label])=>`<div class="store-card"><div><b>${label}</b><div class="store-meta">Usada em filtros, dashboards, importações e vínculos.</div></div><div class="store-val">Ativa</div></div>`).join('')}</div></section>
    </div>
  </div>`;
}
function jmIsoDay(d){
  return new Date(d.getFullYear(),d.getMonth(),d.getDate()).toISOString().slice(0,10);
}
function jmChartRangeDays(){
  if(analysisFilters.period==='today')return 7;
  if(analysisFilters.period==='7d')return 7;
  if(analysisFilters.period==='30d')return 30;
  if(analysisFilters.period==='custom'&&analysisFilters.from&&analysisFilters.to){
    const a=new Date(analysisFilters.from+'T00:00:00'), b=new Date(analysisFilters.to+'T00:00:00');
    return Math.max(1,Math.min(365,Math.round((b-a)/86400000)+1));
  }
  return 30;
}
function jmChartActiveRange(){
  if(analysisFilters.period==='7d')return '7d';
  if(analysisFilters.period==='30d')return '30d';
  if(analysisFilters.period==='custom'&&analysisFilters.from&&analysisFilters.to){
    const d=jmChartRangeDays();
    if(d>=355)return '12m';
    if(d>=170)return '6m';
    if(d>=80)return '3m';
  }
  return analysisFilters.period==='today'?'7d':'30d';
}
function jmSetChartPeriod(range){
  const today=new Date(); today.setHours(0,0,0,0);
  const setCustom=(days)=>{
    const from=new Date(today.getTime()-(days-1)*86400000);
    analysisFilters.period='custom';
    analysisFilters.from=jmIsoDay(from);
    analysisFilters.to=jmIsoDay(today);
  };
  if(range==='7d')analysisFilters.period='7d';
  else if(range==='30d')analysisFilters.period='30d';
  else if(range==='3m')setCustom(90);
  else if(range==='6m')setCustom(180);
  else if(range==='12m')setCustom(365);
  renderContent();
}
function jmMoneyShort(v){
  const n=Math.abs(Number(v)||0), sign=Number(v)<0?'-':'';
  if(n>=1000000)return sign+'R$ '+(n/1000000).toLocaleString('pt-BR',{maximumFractionDigits:1})+' mi';
  if(n>=1000)return sign+'R$ '+(n/1000).toLocaleString('pt-BR',{maximumFractionDigits:1})+' mil';
  return sign+'R$ '+n.toLocaleString('pt-BR',{maximumFractionDigits:0});
}
function jmChartDateLabel(iso){
  if(!iso)return '';
  const d=new Date(iso+'T00:00:00');
  return d.toLocaleDateString('pt-BR',{day:'2-digit',month:'short'}).replace('.','');
}
function jmFinanceChartSeries(orders){
  const days=jmChartRangeDays();
  const today=new Date(); today.setHours(0,0,0,0);
  const start=new Date(today.getTime()-(days-1)*86400000);
  const byDay={};
  for(let t=start.getTime();t<=today.getTime();t+=86400000)byDay[jmIsoDay(new Date(t))]={net:0,profit:0,reserve:0};
  orders.forEach(o=>{
    const iso=parseDateFlexible(o.paymentDate||o.paidDate||o.date||o.orderDate);
    if(!iso||!byDay[iso])return;
    byDay[iso].net+=Number(o.net)||0;
    byDay[iso].profit+=Number(o.profit)||0;
    const c=o.costParts||{};
    byDay[iso].reserve+=(Number(c.filament)||0)+(Number(c.energy)||0)+(Number(c.packaging)||0)+(Number(c.bubble)||0)+(Number(c.mei)||0)+(Number(c.other)||0);
  });
  let cumulative=0;
  return Object.keys(byDay).sort().map(iso=>{
    cumulative+=byDay[iso].net;
    return {iso,label:jmChartDateLabel(iso),net:byDay[iso].net,profit:byDay[iso].profit,reserve:byDay[iso].reserve,value:cumulative};
  });
}
function jmSmoothPath(points){
  if(!points.length)return '';
  if(points.length===1)return `M${points[0].x} ${points[0].y}`;
  let d=`M${points[0].x.toFixed(1)} ${points[0].y.toFixed(1)}`;
  for(let i=0;i<points.length-1;i++){
    const p0=points[i-1]||points[i], p1=points[i], p2=points[i+1], p3=points[i+2]||p2;
    const cp1={x:p1.x+(p2.x-p0.x)/6,y:p1.y+(p2.y-p0.y)/6};
    const cp2={x:p2.x-(p3.x-p1.x)/6,y:p2.y-(p3.y-p1.y)/6};
    d+=` C${cp1.x.toFixed(1)} ${cp1.y.toFixed(1)},${cp2.x.toFixed(1)} ${cp2.y.toFixed(1)},${p2.x.toFixed(1)} ${p2.y.toFixed(1)}`;
  }
  return d;
}
function jmFinanceChart(orders){
  const series=jmFinanceChartSeries(orders), s=summarizeOrders(orders), active=jmChartActiveRange();
  if(!series.length)return '';
  const w=920,h=320,left=54,right=24,top=28,bottom=44;
  const vals=series.map(d=>d.value), max=Math.max(...vals,1), min=Math.min(...vals,0), span=max-min||1;
  const points=series.map((d,i)=>({
    x:left+(series.length===1?0:i*(w-left-right)/(series.length-1)),
    y:h-bottom-((d.value-min)/span)*(h-top-bottom),
    data:d
  }));
  const line=jmSmoothPath(points), area=`${line} L ${points[points.length-1].x.toFixed(1)} ${h-bottom} L ${points[0].x.toFixed(1)} ${h-bottom} Z`;
  const yTicks=[0,.25,.5,.75,1].map(t=>({y:h-bottom-t*(h-top-bottom),v:min+t*span}));
  const step=Math.max(1,Math.floor((series.length-1)/5));
  const xTicks=series.filter((_,i)=>i%step===0||i===series.length-1).slice(-6);
  const markerEvery=Math.max(1,Math.floor(series.length/5));
  const rangeBtns=[['7d','7D'],['30d','30D'],['3m','3M'],['6m','6M'],['12m','12M']];
  return `<section class="ref-card jm-chart-card">
    <div class="jm-chart-head"><div><div class="jm-chart-title">Evolução do caixa</div><div class="jm-chart-sub">Recebido acumulado com base nas importações do período.</div></div>
      <div class="jm-chart-tabs">${rangeBtns.map(([id,label])=>`<button class="${active===id?'on':''}" onclick="jmSetChartPeriod('${id}')">${label}</button>`).join('')}</div></div>
    <svg class="jm-chart-svg" viewBox="0 0 ${w} ${h}" role="img" aria-label="Evolução do caixa">
      <defs>
        <linearGradient id="jmCashArea" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#00C2FF" stop-opacity=".34"/><stop offset=".72" stop-color="#00C2FF" stop-opacity=".10"/><stop offset="1" stop-color="#00C2FF" stop-opacity="0"/></linearGradient>
        <linearGradient id="jmCashLine" x1="0" x2="1" y1="0" y2="0"><stop offset="0" stop-color="#3B82F6"/><stop offset=".45" stop-color="#00C2FF"/><stop offset="1" stop-color="#22D3EE"/></linearGradient>
      </defs>
      ${yTicks.map(t=>`<line class="grid-line" x1="${left}" y1="${t.y.toFixed(1)}" x2="${w-right}" y2="${t.y.toFixed(1)}"/><text class="axis-label" x="${left-8}" y="${(t.y+4).toFixed(1)}" text-anchor="end">${jmMoneyShort(t.v)}</text>`).join('')}
      ${xTicks.map(d=>{const p=points.find(x=>x.data.iso===d.iso)||points[0];return `<line class="grid-line" x1="${p.x.toFixed(1)}" y1="${top}" x2="${p.x.toFixed(1)}" y2="${h-bottom}"/><text class="axis-label" x="${p.x.toFixed(1)}" y="${h-12}" text-anchor="middle">${d.label}</text>`;}).join('')}
      <path d="${area}" fill="url(#jmCashArea)"/>
      <path class="jm-line" d="${line}" fill="none" stroke="url(#jmCashLine)" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
      ${points.filter((_,i)=>i%markerEvery===0||i===points.length-1).map(p=>`<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4.2" fill="#00F5FF" stroke="#0B0F14" stroke-width="2"/>`).join('')}
    </svg>
    <div class="jm-chart-footer"><div><b>Movimentações recentes</b><span>Recebido ${brl(s.net)} | lucro ${brl(s.profit)} | separar ${brl(erpReserveTotal(s))}</span></div><div class="jm-chart-time">${new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})}</div></div>
  </section>`;
}
function renderFinanceiroView(){
  const orders=filteredOrders(), s=summarizeOrders(orders), reserve=erpReserveTotal(s), available=s.net-reserve, margin=s.net?s.profit/s.net*100:0;
  const bestProduct=refChampionProduct(orders), bestStore=erpTopStore(orders), bestMarketplace=erpTopMarketplace(orders);
  return `<div class="ref-dashboard">
    ${refFilters()}
    <div class="ref-board">
      ${refTopSummary(orders,s,available,margin)}
      ${refHighlights(bestProduct,bestStore,bestMarketplace)}
      ${refCostCard(s,reserve,orders)}
      ${refUnlinkedCard(orders)}
      ${refStoreCards(orders)}
    </div>
    ${jmFinanceChart(orders)}
    <button class="btn btn-export" onclick="exportAnalysisCsv()">Exportar CSV da análise</button>
  </div>`;
}
