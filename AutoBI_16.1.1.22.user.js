// ==UserScript==
// @name         AutoBI 16.1.1.22
// @namespace    https://github.com/PhamngocNDH/AutoBI
// @version      16.1.1.22
// @description  AutoBI 16.1.1.22
// @author       38967 - Mr Phạm
// @match        https://crm.thegioididong.com/*
// @match        https://baocao.dienmayxanh.com/*
// @include      https://baocao.dienmayxanh.com/*
// @connect      *
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_16.1.1.15.user.js
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window) return;
  const DATA_KEY='tgdd_data_cache_v30',RESET_KEY='autobi_tc22_reset_done',CAPTURE_KEY='autobi_tc22_capture_signature';
  const get=(k,d=null)=>{try{return GM_getValue(k,d)}catch(_){return d}},set=(k,v)=>{try{GM_setValue(k,v)}catch(_){}};
  const obj=v=>{if(typeof v==='string')try{v=JSON.parse(v)}catch(_){}return v&&typeof v==='object'?v:{}};
  const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const text=e=>(e?.innerText||e?.textContent||'').trim();
  const num=v=>{const n=parseFloat(String(v??'').replace(/,/g,'').replace(/%/g,'').trim());return Number.isFinite(n)?n:0};
  function toast(msg){try{const w=typeof unsafeWindow!=='undefined'?unsafeWindow:window;if(w.UI?.showToast)return w.UI.showToast(msg,4500)}catch(_){}console.log('[AutoBI 22]',msg)}
  function cleanBadTrialDataOnce(){if(get(RESET_KEY,false))return;const z=obj(get(DATA_KEY,{}));z.link7=z.link7||{};for(const k of ['shop1','shop2','shop3','shop4','shop5']){const x=z.link7[k];if(!x||typeof x!=='object')continue;delete x._SHOP_DT;delete x._SHOP_HC;delete x._SHOP_FE;delete x._SHOP_TOTAL_}set(DATA_KEY,z);set(RESET_KEY,true)}
  function findCorrectTable(){for(const t of document.querySelectorAll('table')){const s=norm(text(t));if(!s.includes('dt tra gop')||!s.includes('dt sieu thi')||!s.includes('homecredit')||!s.includes('fecredit'))continue;const rows=[...t.querySelectorAll('tbody tr')].map(r=>norm(text(r)));if(rows.some(r=>r.includes('hai anh'))&&rows.some(r=>r.includes('truc cuong')))return t}return null}
  function keyForRow(s){s=norm(s);if(s.includes('hai anh'))return'shop1';if(s.includes('truc cuong'))return'shop2';return null}
  function capture(){if(!location.pathname.includes('/dashboard/tra-cham'))return false;const table=findCorrectTable();if(!table)return false;const z=obj(get(DATA_KEY,{}));z.link7=z.link7||{};z.link2=z.link2||{};const got=[];let sumST=0,sumTra=0;for(const tr of table.querySelectorAll('tbody tr')){const td=[...tr.querySelectorAll('td')];if(td.length<7)continue;const key=keyForRow(text(td[0]));if(!key)continue;const tra=num(text(td[1])),st=num(text(td[2])),pct=num(text(td[3])),hc=num(text(td[4])),fe=num(text(td[6]));if(!(st>0))continue;z.link7[key]=z.link7[key]||{};z.link7[key]._SHOP_DT=st;z.link7[key]._SHOP_HC=hc;z.link7[key]._SHOP_FE=fe;z.link7[key]._SHOP_TOTAL_={tg:pct,dt_sieu_thi:st,dt_tragop:tra};z.link2[key]=z.link2[key]||{};z.link2[key].tg=pct;sumST+=st;sumTra+=tra;got.push(`${key}:${st}/${tra}/${hc}/${fe}/${pct}`)}if(got.length!==2)return false;z.link2.total=z.link2.total||{};z.link2.total.tg=sumST?sumTra/sumST*100:0;set(DATA_KEY,z);const sig=got.join('|');if(get(CAPTURE_KEY,'')!==sig){set(CAPTURE_KEY,sig);toast('✅ Đã lấy đúng Trả chậm Hải Anh + Trực Cường. Bấm AutoBI để kiểm tra.')}return true}
  cleanBadTrialDataOnce();
  if(location.pathname.includes('/dashboard/tra-cham')){setTimeout(capture,1200);new MutationObserver(()=>capture()).observe(document.documentElement,{childList:true,subtree:true});setInterval(capture,2000)}
})();
