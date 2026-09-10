// ==UserScript==
// @name         AutoBI 16.1.1.20
// @namespace    https://github.com/PhamngocNDH/AutoBI
// @version      16.1.1.20
// @description  AutoBI 16.1.1.20
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
// @require      https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_16.1.1.16.user.js
// ==/UserScript==

/* AutoBI 16.1.1.20 */
(function(){
'use strict';
if(window.top!==window)return;
const DK='tgdd_data_cache_v30',CK='TGDD_BI_STORE_PERMANENT_CONFIG_GM_V1',ST='autobi_tc20_state',LAST='autobi_tc20_last';
const sl=m=>new Promise(r=>setTimeout(r,m));
const nm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const tx=e=>(e?.innerText||e?.textContent||'').trim();
const nu=v=>{const n=parseFloat(String(v??'').replace(/,/g,'').replace(/%/g,'').trim());return Number.isFinite(n)?n:0};
const gv=(k,d=null)=>{try{return GM_getValue(k,d)}catch(_){return d}},sv=(k,v)=>{try{GM_setValue(k,v)}catch(_){}},dv=k=>{try{GM_deleteValue(k)}catch(_){}};
function obj(v){if(typeof v==='string')try{v=JSON.parse(v)}catch(_){}return v&&typeof v==='object'?v:{}}
function cfg(){let c=obj(gv('tgdd_active_run_config'));if(Object.keys(c).length)return c;c=obj(gv(CK));if(!Object.keys(c).length)try{c=obj(localStorage.getItem(CK))}catch(_){}return c}
function shops(c){let a=[];for(let i=1;i<=5;i++)if(c['shop'+i])a.push({key:'shop'+i,name:String(c['shop'+i])});return a}
function toast(m){try{const w=typeof unsafeWindow!=='undefined'?unsafeWindow:window;if(w.UI?.showToast)return w.UI.showToast(m,3200)}catch(_){}console.log('[AutoBI 20]',m)}
async function wait(fn,to=22000){let e=Date.now()+to;while(Date.now()<e){try{let v=fn();if(v)return v}catch(_){}await sl(300)}return null}
function matchShop(row,c){const r=nm(row);for(let i=1;i<=5;i++){if(!c['shop'+i])continue;const f=nm(c['shop'+i]),parts=f.split(' '),tail=parts.slice(-2).join(' '),code=(f.match(/\b(?:hha|tni)\b/)||[])[0];if((f&&(r.includes(f)||f.includes(r)))||(tail&&r.includes(tail))||(code&&r.includes(code)))return'shop'+i}if(r.includes('hai anh'))return'shop1';if(r.includes('truc cuong'))return'shop2';return null}
function totalTable(){return[...document.querySelectorAll('table')].find(t=>{const s=nm(tx(t));return s.includes('dt tra gop')&&s.includes('dt sieu thi')&&s.includes('homecredit')&&s.includes('fe')})||null}
function save(t,c){let z=obj(gv(DK,{}));z.link7=z.link7||{};z.link2=z.link2||{};let n=0,st=0,tg=0,log=[];for(const r of t.querySelectorAll('tbody tr')){const a=[...r.querySelectorAll('td')];if(a.length<7)continue;const name=tx(a[0]),k=matchShop(name,c);if(!k)continue;const tra=nu(tx(a[1])),sieuthi=nu(tx(a[2])),pct=nu(tx(a[3])),hc=nu(tx(a[4])),fe=nu(tx(a[6]));if(!(sieuthi>0))continue;z.link7[k]=z.link7[k]||{};Object.assign(z.link7[k],{_SHOP_DT:sieuthi,_SHOP_HC:hc,_SHOP_FE:fe,_SHOP_TOTAL_:{tg:pct,dt_sieu_thi:sieuthi,dt_tragop:tra}});z.link2[k]=z.link2[k]||{};z.link2[k].tg=pct;st+=sieuthi;tg+=tra;n++;log.push(k+': '+sieuthi+'/'+tra+'/'+hc+'/'+fe)}if(n){z.link2.total=z.link2.total||{};z.link2.total.tg=st?tg/st*100:0;sv(DK,z);sv(LAST,Date.now());sv('autobi_tc20_debug',log.join(' | '));}return{n,log}}
function currentKey(c){const all=nm(document.body.innerText);for(let i=1;i<=5;i++){const s=c['shop'+i];if(!s)continue;const f=nm(s),tail=f.split(' ').slice(-2).join(' ');if(all.includes('ket qua kinh doanh cum')&&(all.includes(f)||all.includes(tail)))return'shop'+i}return'shop1'}
function setNext(label,val){const q=nm(label);for(const e of document.querySelectorAll('td')){if(nm(tx(e))===q){const v=e.nextElementSibling;if(v){v.textContent=String(Math.round(Number(val)||0));return true}}}return false}
function syncHealth(){const c=cfg(),z=obj(gv(DK,{}));if(!Object.keys(c).length||!z.link7)return false;const body=nm(document.body.innerText);if(!body.includes('suc khoe st')||!body.includes('dt sieu thi')||!body.includes('dt tra cham'))return false;const k=currentKey(c),x=z.link7[k]||{},p=Number(z.link2?.[k]?.tg)||Number(x._SHOP_TOTAL_?.tg)||0,dt=Number(x._SHOP_DT)||0,hc=Number(x._SHOP_HC)||0,fe=Number(x._SHOP_FE)||0,tra=Number(x._SHOP_TOTAL_?.dt_tragop)||(dt*p/100);if(!(dt>0))return false;setNext('DT siêu thị',dt);setNext('DT trả chậm',tra);setNext('Homecredit',hc);setNext('FECredit',fe);const cells=[...document.querySelectorAll('td')];const head=cells.find(e=>nm(tx(e))==='ty trong');if(head){let tr=head.closest('tr'),table=tr?.closest('table');if(table){const candidates=[...table.querySelectorAll('td')].filter(e=>/^\s*\d+(?:\.\d+)?%\s*$/.test(tx(e)));const near=candidates.find(e=>Math.abs(e.getBoundingClientRect().left-head.getBoundingClientRect().left)<180);if(near)near.textContent=Math.round(p)+'%'}}return true}
async function tra(){const c=await wait(()=>{const x=cfg();return shops(x).length?x:null},12000);if(!c)return;const t=await wait(totalTable,25000);if(!t){toast('⚠️ Không thấy bảng Trả chậm');return}await sl(1000);const r=save(t,c);if(!r.n){toast('⚠️ Thấy bảng nhưng chưa đọc được dòng siêu thị');return}toast('✅ Đã lưu Trả chậm: '+r.log.join(' | '));const st=obj(gv(ST)),back=st.back||'/dashboard/home';sv(ST,{done:true,back});await sl(1400);location.href=back}
async function main(){const c=await wait(()=>{const x=cfg();return shops(x).length?x:null},12000);if(!c)return;let st=obj(gv(ST));if(st.done){dv(ST);let tries=0;const tm=setInterval(()=>{tries++;if(syncHealth()||tries>20)clearInterval(tm)},500);return}let z=obj(gv(DK,{}));if(shops(c).every(s=>Number(z.link7?.[s.key]?._SHOP_DT)>0)){let tries=0;const tm=setInterval(()=>{tries++;if(syncHealth()||tries>20)clearInterval(tm)},500);return}if(Date.now()-Number(gv(LAST,0)||0)<20000)return;sv(ST,{back:location.pathname+location.search+location.hash});toast('🔄 Đang lấy dữ liệu Báo cáo trả chậm...');await sl(700);location.href='/dashboard/tra-cham'}
if(location.pathname.includes('/dashboard/tra-cham'))setTimeout(tra,1500);else setTimeout(main,3000);
new MutationObserver(()=>{if(!location.pathname.includes('/dashboard/tra-cham'))syncHealth()}).observe(document.documentElement,{childList:true,subtree:true});
})();
