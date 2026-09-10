// ==UserScript==
// @name         AutoBI 16.1.1.19
// @namespace    https://github.com/PhamngocNDH/AutoBI
// @version      16.1.1.19
// @description  AutoBI 16.1.1.19
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

/* AutoBI 16.1.1.19 - đọc trực tiếp bảng Tổng Báo cáo trả chậm, không cần lọc từng ST */
(function () {
'use strict';
if (window.top !== window) return;
const DATA_KEY='tgdd_data_cache_v30',CFG_KEY='TGDD_BI_STORE_PERMANENT_CONFIG_GM_V1',STATE='autobi_tc19_state',LAST='autobi_tc19_last';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const text=e=>(e?.innerText||e?.textContent||'').trim();
const num=v=>{const n=parseFloat(String(v??'').replace(/,/g,'').replace(/%/g,'').trim());return Number.isFinite(n)?n:0};
const get=(k,d=null)=>{try{return GM_getValue(k,d)}catch(_){return d}},set=(k,v)=>{try{GM_setValue(k,v)}catch(_){}},del=k=>{try{GM_deleteValue(k)}catch(_){}};
function obj(v){if(typeof v==='string')try{v=JSON.parse(v)}catch(_){}return v&&typeof v==='object'?v:{}}
function cfg(){let c=obj(get('tgdd_active_run_config'));if(Object.keys(c).length)return c;c=obj(get(CFG_KEY));if(!Object.keys(c).length)try{c=obj(localStorage.getItem(CFG_KEY))}catch(_){}if(Object.keys(c).length)set('tgdd_active_run_config',c);return c}
function shops(c){const a=[];for(let i=1;i<=5;i++)if(c['shop'+i])a.push({key:'shop'+i,name:String(c['shop'+i])});return a}
function toast(m){try{const w=typeof unsafeWindow!=='undefined'?unsafeWindow:window;if(w.UI?.showToast)return w.UI.showToast(m,3200)}catch(_){}console.log('[AutoBI 19]',m)}
async function wait(fn,to=20000){const end=Date.now()+to;while(Date.now()<end){try{const v=fn();if(v)return v}catch(_){}await sleep(300)}return null}
function shopKey(rowName,c){const r=norm(rowName);for(let i=1;i<=5;i++){const s=c['shop'+i];if(!s)continue;const f=norm(s),tail=f.split(' ').slice(-3).join(' ');if(r===f||r.includes(f)||f.includes(r)||(tail&&r.includes(tail)))return'shop'+i}return null}
function findTotalTable(){return[...document.querySelectorAll('table')].find(t=>{const h=norm(text(t.querySelector('thead'))||text(t));return h.includes('sieu thi')&&h.includes('dt tra gop')&&h.includes('dt sieu thi')&&h.includes('homecredit')})||null}
function parseAndSave(table,c){const z=obj(get(DATA_KEY,{}));z.link7=z.link7||{};z.link2=z.link2||{};let count=0,sumTraGop=0,sumSieuThi=0;for(const r of table.querySelectorAll('tbody tr')){const a=[...r.querySelectorAll('td')];if(a.length<7)continue;const k=shopKey(text(a[0]),c);if(!k)continue;const dtTraGop=num(text(a[1])),dtSieuThi=num(text(a[2])),tyTrong=num(text(a[3])),hc=num(text(a[4])),fe=num(text(a[6]));z.link7[k]=z.link7[k]||{};z.link7[k]._SHOP_DT=dtSieuThi;z.link7[k]._SHOP_HC=hc;z.link7[k]._SHOP_FE=fe;z.link7[k]._SHOP_TOTAL_={tg:tyTrong,dt_sieu_thi:dtSieuThi,dt_tragop:dtTraGop};z.link2[k]=z.link2[k]||{};z.link2[k].tg=tyTrong;sumTraGop+=dtTraGop;sumSieuThi+=dtSieuThi;count++}if(count){z.link2.total=z.link2.total||{};if(sumSieuThi>0)z.link2.total.tg=sumTraGop/sumSieuThi*100;set(DATA_KEY,z);set(LAST,Date.now())}return count}
function ready(c){const z=obj(get(DATA_KEY,{})),ss=shops(c);return ss.length&&ss.every(s=>Number(z.link7?.[s.key]?._SHOP_DT)>0)}
async function onTraCham(){const c=await wait(()=>{const x=cfg();return shops(x).length?x:null},12000);if(!c)return;const clickExact=s=>{const q=norm(s),e=[...document.querySelectorAll('button,a,span,div')].find(x=>norm(text(x))===q);if(e){try{e.click()}catch(_){}return true}return false};clickExact('Tỷ trọng trả góp');await sleep(350);clickExact('Lũy kế');await sleep(700);const t=await wait(findTotalTable,22000);if(!t){toast('⚠️ Không tìm thấy bảng Tổng Trả chậm');return}await sleep(1000);const n=parseAndSave(t,c);if(!n){toast('⚠️ Đã thấy bảng nhưng chưa khớp tên siêu thị');return}toast('✅ Đã lấy Trả chậm '+n+' siêu thị');const st=obj(get(STATE)),back=st.back||'/dashboard/thi-dua';set(STATE,{done:true,back});await sleep(800);location.href=back}
async function onThiDua(){const c=await wait(()=>{const x=cfg();return shops(x).length?x:null},12000);if(!c)return;const st=obj(get(STATE));if(st.done){del(STATE);toast('✅ Trả chậm đã cập nhật vào Sức khỏe ST');setTimeout(()=>location.reload(),700);return}if(ready(c))return;if(Date.now()-Number(get(LAST,0)||0)<30000)return;set(STATE,{back:location.pathname+location.search+location.hash});toast('🔄 Sang Báo cáo trả chậm lấy bảng Tổng...');await sleep(700);location.href='/dashboard/tra-cham'}
if(location.pathname.includes('/dashboard/tra-cham'))setTimeout(onTraCham,1200);else if(location.pathname.includes('/dashboard/thi-dua'))setTimeout(onThiDua,3500);
})();
