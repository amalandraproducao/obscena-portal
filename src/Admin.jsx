import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  "https://yztciaxsneabtayfxijq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl6dGNpYXhzbmVhYnRheWZ4aWpxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk2NDc2NDIsImV4cCI6MjA5NTIyMzY0Mn0.01tfOk1mIH3_zyVxO2DGCa0gOMxrhWUblpuFah5QjAs"
);

const PASS = "malandra2025";
const COMM = 0.10;
const p2 = n => String(n || 0).padStart(2, "0");
const p3 = n => String(n || 0).padStart(3, "0");
const mkCode = (a, b) => `${p2(a)}-${p3(b)}`;
const fmt = n => (+n || 0).toFixed(2);
const METODOS = ["Dinheiro", "MB Way", "Transferência"];
const FONT = `@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');`;

const ChevR  = () => <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"/></svg>;
const ChevD  = () => <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 12 15 18 9"/></svg>;
const ChevU  = () => <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="18 15 12 9 6 15"/></svg>;
const Check  = () => <svg width="9"  height="9"  viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3"   strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>;
const X      = () => <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>;
const Plus   = () => <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>;
const Dl     = () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>;
const Refresh= () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>;

const PAGES = [
  { id: "vendas",     label: "Registo de venda",  icon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/></svg> },
  { id: "realizadas", label: "Vendas realizadas", icon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> },
  { id: "faturas",    label: "Faturas",            icon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg> },
  { id: "inventario", label: "Inventário",         icon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/></svg> },
  { id: "artistas",   label: "Artistas",           icon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg> },
  { id: "relatorio",  label: "Relatório final",    icon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> },
];

const CSS = `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Inter', sans-serif; font-size: 13px; background: #f4f4f2; color: #111; height: 100vh; overflow: hidden; }
.app { display: flex; height: 100vh; }

.nav { width: 192px; flex-shrink: 0; background: #fff; border-right: 1px solid #e0e0e0; display: flex; flex-direction: column; }
.nav-brand { padding: 16px 16px 12px; border-bottom: 1px solid #e0e0e0; }
.nav-brand-name { font-size: 12px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; color: #111; }
.nav-brand-sub { font-size: 11px; color: #777; margin-top: 2px; }
.nav-items { padding: 8px 6px; flex: 1; }
.nav-item { display: flex; align-items: center; gap: 9px; padding: 8px 10px; border-radius: 8px; cursor: pointer; font-size: 12px; color: #555; transition: background .12s, color .12s; margin-bottom: 1px; user-select: none; }
.nav-item:hover { background: #f4f4f2; color: #111; }
.nav-item.on { background: #f0f0ee; color: #111; font-weight: 500; }
.nav-item svg { flex-shrink: 0; width: 14px; height: 14px; }
.nav-badge { margin-left: auto; background: #fff3e0; color: #8a3a00; font-size: 10px; font-weight: 600; padding: 1px 6px; border-radius: 10px; }
.nav-footer { padding: 10px 12px; border-top: 1px solid #e0e0e0; display: flex; align-items: center; justify-content: space-between; gap: 4px; }
.btn-logout  { background: none; border: none; font-family: 'Inter', sans-serif; font-size: 11px; color: #bbb; cursor: pointer; padding: 4px 8px; border-radius: 5px; }
.btn-logout:hover { color: #666; }
.btn-refresh { background: none; border: none; font-family: 'Inter', sans-serif; font-size: 11px; color: #aaa; cursor: pointer; padding: 4px 8px; border-radius: 5px; display: flex; align-items: center; gap: 4px; }
.btn-refresh:hover { color: #555; }

.main { flex: 1; overflow: hidden; display: flex; flex-direction: column; min-width: 0; }
.page-hdr { padding: 20px 20px 0; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; }
.page-title { font-size: 15px; font-weight: 500; color: #111; }
.content { padding: 14px 20px 24px; flex: 1; overflow-y: auto; }

.card { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; padding: 14px 16px; margin-bottom: 10px; }
.card:last-child { margin-bottom: 0; }
.sec-lbl { font-size: 10px; font-weight: 600; letter-spacing: .6px; text-transform: uppercase; color: #666; margin-bottom: 8px; }

.vendas-layout { display: grid; grid-template-columns: 1fr 300px; gap: 12px; height: calc(100vh - 110px); overflow: hidden; }
.vendas-left { overflow-y: auto; height: 100%; padding-right: 2px; }
.vendas-right { height: 100%; display: flex; flex-direction: column; overflow: hidden; }

.code-row { display: flex; gap: 8px; align-items: center; margin-bottom: 5px; }
.code-inp { background: #f0f0ee; border: 1px solid transparent; border-radius: 8px; padding: 8px 12px; font-family: 'Inter', sans-serif; font-size: 17px; font-weight: 500; letter-spacing: 2px; width: 130px; outline: none; color: #111; transition: border-color .15s; }
.code-inp:focus { border-color: #aaa; background: #e8e8e5; }
.btn-add { background: #f0f0ee; border: 1px solid transparent; border-radius: 8px; padding: 8px 13px; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; color: #444; cursor: pointer; display: flex; align-items: center; gap: 5px; transition: background .12s; white-space: nowrap; }
.btn-add:hover { background: #e4e4e1; }
.code-err { font-size: 11px; color: #c00; margin-top: 4px; font-weight: 500; }
.hint { font-size: 11px; color: #888; margin-top: 4px; }

.browser-sep { display: flex; align-items: center; gap: 10px; margin: 14px 0 12px; }
.browser-sep-line { flex: 1; height: 1px; background: #e8e8e8; }
.browser-sep-txt { font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #888; }
.browser { border: 1px solid #e4e4e4; border-radius: 10px; overflow: hidden; background: #fff; margin-bottom: 10px; }
.bar-artista { display: flex; align-items: center; gap: 10px; padding: 10px 13px; cursor: pointer; transition: background .12s; user-select: none; border-top: 1px solid #f0f0ee; }
.bar-artista:first-child { border-top: none; }
.bar-artista:hover { background: #f8f8f6; }
.bar-artista.on { background: #f4f4f2; }
.ba-code { font-size: 11px; font-weight: 600; color: #888; min-width: 20px; }
.ba-nome { font-weight: 500; font-size: 12px; flex: 1; color: #111; }
.ba-arte { font-size: 11px; color: #777; }
.ba-chev { color: #bbb; transition: transform .15s; flex-shrink: 0; display: flex; }
.bar-artista.on .ba-chev { transform: rotate(90deg); }
.browser-artigos { background: #fafaf8; border-top: 1px solid #ececea; }
.bar-artigo { display: flex; align-items: center; gap: 8px; padding: 8px 13px 8px 32px; border-top: 1px solid #f2f2f0; transition: background .12s; }
.bar-artigo:first-child { border-top: none; }
.bar-artigo:not(.esgotado) { cursor: pointer; }
.bar-artigo:not(.esgotado):hover { background: #eeeeeb; }
.bar-artigo.esgotado { opacity: .4; }
.baa-code { font-size: 10px; font-weight: 600; color: #888; min-width: 48px; font-variant-numeric: tabular-nums; }
.baa-titulo { flex: 1; font-size: 12px; font-weight: 500; color: #111; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.baa-preco { font-size: 12px; font-weight: 500; min-width: 42px; text-align: right; flex-shrink: 0; color: #111; }
.baa-stock { font-size: 11px; min-width: 26px; text-align: right; flex-shrink: 0; }
.in-cart-pill { font-size: 10px; font-weight: 600; background: #e8eeff; color: #334499; padding: 2px 7px; border-radius: 10px; flex-shrink: 0; }

.carr-panel { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; display: flex; flex-direction: column; overflow: hidden; flex: 1; min-height: 0; }
.carr-hdr { padding: 12px 14px 10px; border-bottom: 1px solid #ececea; flex-shrink: 0; display: flex; align-items: center; justify-content: space-between; }
.carr-title { font-size: 11px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #555; }
.carr-count { font-size: 11px; color: #888; }
.carr-items { flex: 1; overflow-y: auto; padding: 0 14px; }
.carr-empty { display: flex; align-items: center; justify-content: center; height: 80px; font-size: 12px; color: #aaa; }
.carr-item { display: flex; align-items: center; gap: 8px; padding: 8px 0; border-top: 1px solid #f4f4f2; }
.carr-item:first-child { border-top: none; }
.carr-info { flex: 1; min-width: 0; }
.carr-titulo { font-size: 12px; font-weight: 500; color: #111; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.carr-sub { font-size: 11px; color: #777; margin-top: 1px; }
.carr-qty { display: flex; align-items: center; gap: 5px; flex-shrink: 0; }
.qty-btn { width: 20px; height: 20px; border-radius: 4px; border: 1px solid #ddd; background: #f4f4f2; font-size: 13px; line-height: 1; display: flex; align-items: center; justify-content: center; cursor: pointer; color: #444; transition: background .12s; flex-shrink: 0; user-select: none; }
.qty-btn:hover { background: #e8e8e5; }
.qty-val { font-size: 12px; font-weight: 500; min-width: 16px; text-align: center; color: #111; }
.carr-preco { font-size: 12px; font-weight: 500; min-width: 42px; text-align: right; flex-shrink: 0; color: #111; }
.carr-rm { background: none; border: none; cursor: pointer; color: #bbb; padding: 2px 3px; line-height: 1; transition: color .12s; flex-shrink: 0; display: flex; }
.carr-rm:hover { color: #c00; }
.oferta-pill { font-size: 10px; font-weight: 600; background: #e4f0e4; color: #2a5a1a; padding: 2px 7px; border-radius: 10px; flex-shrink: 0; }

.carr-footer { padding: 12px 14px; border-top: 1px solid #ececea; flex-shrink: 0; }
.total-row { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 12px; }
.total-lbl { font-size: 12px; color: #555; font-weight: 500; }
.total-val { font-size: 18px; font-weight: 500; color: #111; }
.metodos { display: flex; gap: 5px; margin-bottom: 10px; }
.met-btn { flex: 1; background: #f0f0ee; border: 1px solid transparent; border-radius: 7px; padding: 7px 4px; font-family: 'Inter', sans-serif; font-size: 11px; font-weight: 500; color: #555; cursor: pointer; transition: all .12s; text-align: center; }
.met-btn:hover { background: #e8e8e5; }
.met-btn.on { background: #111; color: #fff; }
.toggle-row { display: flex; align-items: center; gap: 9px; padding: 8px 11px; background: #f4f4f2; border-radius: 8px; cursor: pointer; margin-bottom: 8px; transition: background .12s; user-select: none; }
.toggle-row:hover { background: #ececea; }
.tbox { width: 14px; height: 14px; border-radius: 4px; background: #d8d8d8; border: 1px solid #bbb; flex-shrink: 0; display: flex; align-items: center; justify-content: center; transition: all .15s; }
.tbox.on { background: #111; border-color: #111; }
.toggle-lbl { font-size: 12px; color: #444; font-weight: 500; }
.fat-fields { background: #f4f4f2; border-radius: 8px; padding: 10px; margin-bottom: 8px; display: flex; flex-direction: column; gap: 7px; }
.fld-lbl { font-size: 10px; color: #555; font-weight: 600; margin-bottom: 3px; }
.fld-inp { width: 100%; background: #fff; border: 1px solid #ddd; border-radius: 6px; padding: 6px 9px; font-family: 'Inter', sans-serif; font-size: 12px; color: #111; outline: none; transition: border-color .15s; }
.fld-inp:focus { border-color: #999; }
.fld-inp.e { border-color: #e00; background: #fce8e8; }
.g2 { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
.btn-confirmar { width: 100%; background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 500; padding: 11px; border-radius: 8px; cursor: pointer; transition: opacity .12s; }
.btn-confirmar:hover { opacity: .85; }
.btn-confirmar:disabled { opacity: .4; cursor: not-allowed; }
.btn-confirmar-empty { width: 100%; background: #f0f0ee; border: none; color: #999; font-family: 'Inter', sans-serif; font-size: 12px; padding: 11px; border-radius: 8px; cursor: not-allowed; }
.venda-ok { background: #e4f0e4; border-radius: 10px; padding: 13px; display: flex; align-items: center; gap: 10px; font-size: 13px; color: #1e4a1e; font-weight: 500; margin-bottom: 8px; }
.ok-ic { width: 24px; height: 24px; background: #4a8a2a; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #fff; flex-shrink: 0; }

.real-layout { display: grid; grid-template-columns: 1fr 220px; gap: 16px; align-items: start; }
.stats-list { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; padding: 14px 16px; position: sticky; top: 0; }
.stats-list-title { font-size: 11px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #555; margin-bottom: 12px; }
.stat-row { display: flex; justify-content: space-between; align-items: baseline; padding: 7px 0; border-top: 1px solid #f0f0ee; }
.stat-row:first-of-type { border-top: none; }
.stat-lbl { font-size: 12px; color: #555; }
.stat-val { font-size: 14px; font-weight: 500; color: #111; }
.venda-list { display: flex; flex-direction: column; gap: 8px; }
.venda-card { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; overflow: hidden; }
.venda-card-hdr { display: flex; align-items: center; gap: 9px; padding: 10px 14px; background: #f8f8f6; flex-wrap: wrap; border-bottom: 1px solid #f0f0ee; }
.vh-num { font-size: 11px; font-weight: 600; color: #555; }
.vh-hora { font-size: 11px; color: #777; }
.vh-met { font-size: 11px; background: #ececea; color: #444; padding: 2px 8px; border-radius: 8px; font-weight: 500; }
.vh-fat { font-size: 10px; font-weight: 600; background: #fff3e0; color: #7a3300; padding: 2px 8px; border-radius: 8px; }
.vh-total { font-size: 13px; font-weight: 500; margin-left: auto; color: #111; }
.venda-art-row { display: flex; align-items: center; gap: 9px; padding: 8px 14px; border-top: 1px solid #f4f4f2; font-size: 12px; }
.var-code { font-size: 11px; font-weight: 600; color: #777; min-width: 52px; font-variant-numeric: tabular-nums; }
.var-info { flex: 1; min-width: 0; }
.var-titulo { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 500; color: #111; }
.var-artista { color: #777; font-size: 11px; }
.var-qty { font-size: 11px; color: #666; flex-shrink: 0; font-weight: 500; }
.var-preco { font-weight: 500; flex-shrink: 0; color: #111; }

.fat-layout { display: grid; grid-template-columns: 1fr 200px; gap: 16px; align-items: start; }
.fat-table { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; overflow: hidden; }
.fat-thead { background: #f8f8f6; border-bottom: 1px solid #e4e4e4; }
.fat-th { display: grid; grid-template-columns: 28px 1fr 1fr 130px 80px 90px; gap: 10px; padding: 9px 16px; font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #555; }
.fat-row-item { display: grid; grid-template-columns: 28px 1fr 1fr 130px 80px 90px; gap: 10px; padding: 10px 16px; border-top: 1px solid #f2f2f0; align-items: center; font-size: 12px; transition: background .12s; }
.fat-row-item:hover { background: #fafaf8; }
.fat-row-item.done { opacity: .5; }
.fat-cell { color: #444; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.fat-cell.bold { font-weight: 500; color: #111; }
.fat-cell.mono { font-variant-numeric: tabular-nums; font-size: 11px; color: #666; font-weight: 600; }
.btn-proc { background: #f0f0ee; border: none; border-radius: 6px; padding: 4px 10px; font-family: 'Inter', sans-serif; font-size: 11px; font-weight: 600; color: #444; cursor: pointer; transition: background .12s; white-space: nowrap; }
.btn-proc:hover { background: #e4e4e1; }
.done-badge { font-size: 10px; font-weight: 600; background: #e4f0e4; color: #2a5a1a; padding: 3px 8px; border-radius: 6px; display: inline-flex; align-items: center; gap: 4px; }
.fat-stats { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; padding: 14px 16px; position: sticky; top: 0; }
.fat-stats-title { font-size: 11px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #555; margin-bottom: 12px; }

.inv-container { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; overflow: hidden; }
.inv-hdr-row { display: grid; grid-template-columns: 90px 1fr 90px 80px 80px; position: sticky; top: 0; z-index: 2; border-bottom: 1px solid #e0e0e0; }
.inv-hdr-row > span { font-size: 10px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: #666; padding: 9px 14px; background: #f8f8f6; }
.inv-section { border-top: 1px solid #e4e4e4; }
.inv-section:first-of-type { border-top: none; }
.inv-artist-bar { display: flex; align-items: center; gap: 10px; padding: 10px 14px; cursor: pointer; user-select: none; transition: background .12s; }
.inv-artist-bar:hover { background: #f8f8f6; }
.inv-acode { font-size: 11px; font-weight: 600; color: #777; min-width: 20px; flex-shrink: 0; }
.inv-anome { font-weight: 500; color: #111; }
.inv-aarte { font-size: 11px; color: #666; }
.inv-astats { font-size: 11px; color: #666; flex-shrink: 0; }
.inv-artigo-row { display: grid; grid-template-columns: 90px 1fr 90px 80px 80px; align-items: center; }
.inv-artigo-row > span { padding: 8px 14px; border-top: 1px solid #f2f2f0; font-size: 12px; color: #333; background: #fafaf8; transition: background .1s; }
.inv-artigo-row:hover > span { background: #f4f4f2; }
.stk { display: inline-flex; align-items: center; justify-content: center; min-width: 22px; height: 18px; border-radius: 4px; font-size: 11px; font-weight: 600; background: #ececea; padding: 0 5px; color: #444; }
.stk.low { background: #fff3e0; color: #7a3300; }
.stk.zero { background: #fce8e8; color: #900; }

.art-list { display: flex; flex-direction: column; gap: 7px; }
.art-card { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; padding: 14px 16px; }
.art-top { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
.art-code { font-size: 11px; font-weight: 600; color: #777; min-width: 20px; }
.art-nome { font-weight: 500; flex: 1; color: #111; }
.art-arte { font-size: 12px; color: #666; }
.art-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0; }
.art-col { padding: 0 20px; border-left: 1px solid #ececea; }
.art-col:first-child { padding-left: 0; border-left: none; }
.art-col:last-child { padding-right: 0; }
.art-f { display: flex; flex-direction: column; padding: 6px 0; border-top: 1px solid #f4f4f2; }
.art-f:first-child { border-top: none; }
.art-fl { font-size: 10px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: #888; margin-bottom: 3px; }
.art-fv { font-size: 12px; color: #333; }

.login { min-height: 100vh; display: flex; align-items: center; justify-content: center; background: #f4f4f2; }
.login-box { background: #fff; border-radius: 16px; padding: 40px; width: 100%; max-width: 360px; border: 1px solid #e4e4e4; }
.login-title { font-size: 18px; font-weight: 500; margin-bottom: 4px; }
.login-sub { font-size: 12px; color: #999; margin-bottom: 28px; }
.lbl { display: block; font-size: 10px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: #666; margin-bottom: 6px; }
.inp { width: 100%; background: #f0f0ee; border: 1px solid transparent; outline: none; color: #000; font-family: 'Inter', sans-serif; font-size: 13px; padding: 10px 14px; border-radius: 8px; transition: background .15s; }
.inp:focus { background: #e8e8e5; border-color: #aaa; }
.inp.e { background: #fce8e8; border-color: #e00; }
.em { font-size: 11px; color: #d00; margin-top: 4px; }
.btn-primary { width: 100%; background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 500; padding: 11px; border-radius: 8px; cursor: pointer; margin-top: 16px; transition: background .12s; }
.btn-primary:hover { background: #333; }

.empty { text-align: center; padding: 52px 20px; font-size: 13px; color: #999; }
.btn-export { background: #f0f0ee; border: none; color: #333; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; padding: 9px 14px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 6px; transition: background .12s; }
.btn-export:hover { background: #e4e4e1; }
.rel-card { background: #fff; border-radius: 12px; border: 1px solid #e4e4e4; padding: 18px 20px; margin-bottom: 10px; }
.rel-title { font-size: 14px; font-weight: 500; margin-bottom: 4px; }
.rel-desc { font-size: 12px; color: #888; margin-bottom: 14px; line-height: 1.7; }
.rel-row { display: flex; justify-content: space-between; align-items: center; padding: 11px 0; border-top: 1px solid #f4f4f2; font-size: 13px; }
.rel-row:first-of-type { border-top: none; }
`;

export default function Admin() {
  const [auth, setAuth]   = useState(() => sessionStorage.getItem("obs_adm") === "1");
  const [pw, setPw]       = useState("");
  const [pwErr, setPwErr] = useState(false);
  const [page, setPage]   = useState("vendas");

  const [artistas, setArtistas]               = useState([]);
  const [artigos,  setArtigos]                = useState([]);
  const [vendas,   setVendas]                 = useState([]);
  const [loading,  setLoading]                = useState(true);
  const [expanded, setExpanded]               = useState(new Set());
  const [faturasProcessadas, setFatProc]      = useState(new Set());

  const [carrinho,       setCarrinho]         = useState([]);
  const [codeInput,      setCodeInput]        = useState("");
  const [codeErr,        setCodeErr]          = useState("");
  const [metodo,         setMetodo]           = useState("Dinheiro");
  const [querFatura,     setQuerFatura]       = useState(false);
  const [fatura,         setFatura]           = useState({ nome: "", email: "", nif: "" });
  const [fatErr,         setFatErr]           = useState({});
  const [vendaOk,        setVendaOk]          = useState(false);
  const [saving,         setSaving]           = useState(false);
  const [browserArtista, setBrowserArtista]   = useState(null);
  const codeRef = useRef(null);

  const login  = () => { if (pw === PASS) { sessionStorage.setItem("obs_adm","1"); setAuth(true); } else setPwErr(true); };
  const logout = () => { sessionStorage.removeItem("obs_adm"); setAuth(false); };

  useEffect(() => { if (auth) load(); }, [auth]);

  const load = async () => {
    setLoading(true);
    const [{ data: a }, { data: ar }, { data: v }] = await Promise.all([
      supabase.from("artistas").select("*").order("codigo"),
      supabase.from("artigos").select("*").order("artista_id,codigo"),
      supabase.from("vendas").select("*").order("created_at", { ascending: true }),
    ]);
    setArtistas(a || []);
    setArtigos(ar || []);
    setVendas(v || []);
    setExpanded(new Set((a || []).map(x => x.id)));
    setLoading(false);
  };

  const stockOf = (artigo_id) => {
    const ar = artigos.find(a => a.id === artigo_id);
    if (!ar) return 0;
    const vendido = vendas.filter(v => v.artigo_id === artigo_id).reduce((s, v) => s + (v.quantidade || 1), 0);
    return ar.quantidade - vendido;
  };

  const addToCart = (artigo, artista) => {
    setCodeErr("");
    setCarrinho(prev => {
      const existing = prev.find(c => c.id === artigo.id);
      if (existing) {
        if (!artigo.oferta && existing.qty >= stockOf(artigo.id)) {
          setCodeErr("Sem mais stock disponível");
          return prev;
        }
        return prev.map(c => c.id === artigo.id ? { ...c, qty: c.qty + 1 } : c);
      }
      return [...prev, {
        ...artigo, qty: 1,
        artista_nome:   artista.nome_artistico || artista.nome,
        artista_codigo: artista.codigo,
        artista_id:     artista.id,
      }];
    });
    setCodeInput("");
  };

  const addByCode = () => {
    const raw = (codeRef.current?.value || codeInput).trim().toUpperCase();
    setCodeInput(raw);
    const m = raw.match(/^(\d+)-(\d+)$/);
    if (!m) { setCodeErr("Formato inválido — ex: 01-003"); return; }
    const artista = artistas.find(a => a.codigo === parseInt(m[1]));
    if (!artista) { setCodeErr("Artista não encontrado"); return; }
    const artigo  = artigos.find(a => a.artista_id === artista.id && a.codigo === parseInt(m[2]));
    if (!artigo)  { setCodeErr("Artigo não encontrado"); return; }
    const qNoCarr = carrinho.find(c => c.id === artigo.id)?.qty || 0;
    if (!artigo.oferta && (stockOf(artigo.id) - qNoCarr) <= 0) { setCodeErr("Sem stock disponível"); return; }
    addToCart(artigo, artista);
    setTimeout(() => codeRef.current?.focus(), 30);
  };

  const addFromBrowser = (artigo_id, artista_id) => {
    const artigo  = artigos.find(a => a.id === artigo_id);
    const artista = artistas.find(a => a.id === artista_id);
    if (!artigo || !artista) return;
    const qNoCarr = carrinho.find(c => c.id === artigo.id)?.qty || 0;
    if (!artigo.oferta && (stockOf(artigo.id) - qNoCarr) <= 0) { setCodeErr("Sem stock disponível"); return; }
    addToCart(artigo, artista);
  };

  const changeQty = (idx, delta) => {
    setCarrinho(prev => {
      const item = prev[idx];
      if (!item) return prev;
      const nova = item.qty + delta;
      if (nova <= 0) return prev.filter((_, i) => i !== idx);
      if (!item.oferta && nova > stockOf(item.id)) return prev;
      return prev.map((c, i) => i === idx ? { ...c, qty: nova } : c);
    });
  };

  const removeItem = (idx) => setCarrinho(prev => prev.filter((_, i) => i !== idx));

  const confirmarVenda = async () => {
    if (querFatura) {
      const errs = {};
      if (!fatura.nome.trim())  errs.nome  = true;
      if (!fatura.email.trim()) errs.email = true;
      if (!fatura.nif.trim())   errs.nif   = true;
      if (Object.keys(errs).length) { setFatErr(errs); return; }
    }
    setSaving(true);
    const grupo_id = crypto.randomUUID();
    const rows = carrinho.map(item => ({
      artigo_id:       item.id,
      artista_id:      item.artista_id,
      preco:           item.oferta ? 0 : item.preco,
      quantidade:      item.qty,
      oferta:          item.oferta,
      metodo,
      quer_fatura:     querFatura,
      comprador_nome:  querFatura ? fatura.nome  : null,
      comprador_email: querFatura ? fatura.email : null,
      comprador_nif:   querFatura ? fatura.nif   : null,
      grupo_id,
    }));
    const { error } = await supabase.from("vendas").insert(rows);
    setSaving(false);
    if (!error) {
      setVendaOk(true);
      setCarrinho([]);
      setQuerFatura(false);
      setFatura({ nome: "", email: "", nif: "" });
      setFatErr({});
      setMetodo("Dinheiro");
      setBrowserArtista(null);
      await load();
      setTimeout(() => {
        setVendaOk(false);
        setCodeInput("");
        setTimeout(() => codeRef.current?.focus(), 30);
      }, 2500);
    }
  };

  // Group vendas by grupo_id (or by id for legacy rows without grupo_id)
  const vendasAgrupadas = () => {
    const grupos = new Map();
    let seq = 0;
    [...vendas].reverse().forEach(v => {
      const key = v.grupo_id || String(v.id);
      if (!grupos.has(key)) {
        seq++;
        grupos.set(key, {
          key, seq,
          hora:           new Date(v.created_at).toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" }),
          metodo:         v.metodo,
          quer_fatura:    v.quer_fatura,
          comprador_nome: v.comprador_nome,
          comprador_email:v.comprador_email,
          comprador_nif:  v.comprador_nif,
          itens:          [],
          total:          0,
        });
      }
      const g = grupos.get(key);
      const artigo  = artigos.find(a => a.id === v.artigo_id);
      const artista = artistas.find(a => a.id === v.artista_id);
      const qty = v.quantidade || 1;
      g.itens.push({ ...v, artigo, artista, qty,
        titulo:         artigo?.titulo  || "—",
        artista_nome:   artista ? (artista.nome_artistico || artista.nome) : "—",
        artista_codigo: artista?.codigo,
        artigo_codigo:  artigo?.codigo,
      });
      g.total += (v.oferta ? 0 : (v.preco * qty));
    });
    return [...grupos.values()];
  };

  const exportar = () => {
    const wb   = XLSX.utils.book_new();
    const grupos = vendasAgrupadas();

    // Resumo por artista
    const resumo = [
      ["RELATÓRIO FINAL — OBSCENA 2026"], [""],
      ["Cód.", "Artista", "Peças vendidas", "Total público (€)", "Comissão 10% (€)", "A pagar (€)"],
    ];
    artistas.forEach(ar => {
      const vv    = vendas.filter(v => v.artista_id === ar.id && !v.oferta);
      const total = vv.reduce((s, v) => s + (v.preco * (v.quantidade || 1)), 0);
      const pecas = vv.reduce((s, v) => s + (v.quantidade || 1), 0);
      resumo.push([p2(ar.codigo), ar.nome_artistico || ar.nome, pecas, fmt(total), fmt(total * COMM), fmt(total * (1 - COMM))]);
    });
    const gt    = vendas.filter(v => !v.oferta).reduce((s, v) => s + (v.preco * (v.quantidade || 1)), 0);
    const gtPcs = vendas.filter(v => !v.oferta).reduce((s, v) => s + (v.quantidade || 1), 0);
    resumo.push(["", "TOTAL", gtPcs, fmt(gt), fmt(gt * COMM), fmt(gt * (1 - COMM))]);
    const ws0 = XLSX.utils.aoa_to_sheet(resumo);
    ws0["!cols"] = [{wch:6},{wch:26},{wch:16},{wch:20},{wch:18},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws0, "Resumo");

    // Sheet por artista
    artistas.forEach(ar => {
      const vv = vendas.filter(v => v.artista_id === ar.id);
      if (!vv.length) return;
      const rows = [[ar.nome_artistico || ar.nome], [""],
        ["Hora", "Código", "Artigo", "Preço (€)", "Qtd", "Total (€)", "Método", "Fatura"]];
      vv.forEach(v => {
        const ar2 = artigos.find(a => a.id === v.artigo_id);
        const qty = v.quantidade || 1;
        rows.push([
          new Date(v.created_at).toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" }),
          ar2 ? mkCode(ar.codigo, ar2.codigo) : "—",
          ar2?.titulo || "—",
          v.oferta ? "Oferta" : fmt(v.preco),
          qty,
          v.oferta ? "—" : fmt(v.preco * qty),
          v.metodo,
          v.quer_fatura ? `${v.comprador_nome} / ${v.comprador_nif}` : "Não",
        ]);
      });
      const tot = vv.filter(v => !v.oferta).reduce((s, v) => s + (v.preco * (v.quantidade || 1)), 0);
      rows.push(["", "", "Total vendas", "", "", fmt(tot)]);
      rows.push(["", "", "A pagar ao artista", "", "", fmt(tot * (1 - COMM))]);
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = [{wch:8},{wch:10},{wch:30},{wch:12},{wch:6},{wch:12},{wch:14},{wch:30}];
      XLSX.utils.book_append_sheet(wb, ws, (ar.nome_artistico || ar.nome).slice(0,28).replace(/[:\\/?*[\]]/g,""));
    });

    // Faturação
    const fatRows = [["FATURAÇÃO — COMPRADORES"], [""],
      ["Hora", "Nº", "Artigos", "Valor (€)", "Nome", "Email", "NIF"]];
    grupos.filter(g => g.quer_fatura).forEach((g, i) => {
      const artigosStr = g.itens.filter(it => !it.oferta)
        .map(it => `${it.titulo}${it.qty > 1 ? ` ×${it.qty}` : ""}`).join(", ");
      fatRows.push([g.hora, String(i+1).padStart(3,"0"), artigosStr, fmt(g.total),
        g.comprador_nome, g.comprador_email || "", g.comprador_nif]);
    });
    const wsf = XLSX.utils.aoa_to_sheet(fatRows);
    wsf["!cols"] = [{wch:8},{wch:6},{wch:36},{wch:12},{wch:24},{wch:28},{wch:12}];
    XLSX.utils.book_append_sheet(wb, wsf, "Faturação");

    XLSX.writeFile(wb, "obscena_relatorio_final.xlsx");
  };

  // ── LOGIN ──────────────────────────────────────────────────────────────
  if (!auth) return (
    <>
      <style>{FONT}{CSS}</style>
      <div className="login">
        <div className="login-box">
          <div className="login-title">Painel Malandra</div>
          <div className="login-sub">Obscena · Administração</div>
          <label className="lbl">Password</label>
          <input className={`inp${pwErr ? " e" : ""}`} type="password" value={pw}
            onChange={e => { setPw(e.target.value); setPwErr(false); }}
            onKeyDown={e => e.key === "Enter" && login()}
            autoFocus placeholder="••••••••••" />
          {pwErr && <div className="em">Password incorrecta</div>}
          <button className="btn-primary" onClick={login}>Entrar</button>
        </div>
      </div>
    </>
  );

  // ── DERIVED ────────────────────────────────────────────────────────────
  const grupos       = vendasAgrupadas();
  const totalGeral   = grupos.reduce((s, g) => s + g.total, 0);
  const fatPendentes = grupos.filter(g => g.quer_fatura && !faturasProcessadas.has(g.key)).length;
  const carTotal     = carrinho.reduce((s, i) => s + (i.oferta ? 0 : i.preco * i.qty), 0);
  const carItens     = carrinho.reduce((s, i) => s + i.qty, 0);

  const goPage = (p) => {
    setPage(p); setCarrinho([]); setCodeInput(""); setCodeErr("");
    setVendaOk(false); setQuerFatura(false);
    setFatura({ nome:"", email:"", nif:"" }); setFatErr({});
    setBrowserArtista(null);
  };

  const toggleExp = (id) => setExpanded(prev => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });

  const toggleFatProc = (key) => setFatProc(prev => {
    const next = new Set(prev);
    next.has(key) ? next.delete(key) : next.add(key);
    return next;
  });

  // ── RENDER ─────────────────────────────────────────────────────────────
  return (
    <>
      <style>{FONT}{CSS}</style>
      <div className="app">

        {/* ── SIDEBAR ── */}
        <nav className="nav">
          <div className="nav-brand">
            <div className="nav-brand-name">Obscena</div>
            <div className="nav-brand-sub">Malandra · Admin</div>
          </div>
          <div className="nav-items">
            {PAGES.map(p => (
              <div key={p.id} className={`nav-item${page === p.id ? " on" : ""}`} onClick={() => goPage(p.id)}>
                {p.icon} {p.label}
                {p.id === "faturas" && fatPendentes > 0 && <span className="nav-badge">{fatPendentes}</span>}
              </div>
            ))}
          </div>
          <div className="nav-footer">
            <button className="btn-refresh" onClick={load}><Refresh /> Actualizar</button>
            <button className="btn-logout"  onClick={logout}>Sair</button>
          </div>
        </nav>

        {/* ── MAIN ── */}
        <main className="main">
          {loading
            ? <div className="empty">A carregar...</div>
            : <>

            {/* ════════════════════════════════════════
                REGISTO DE VENDA
            ════════════════════════════════════════ */}
            {page === "vendas" && <>
              <div className="page-hdr">
                <div className="page-title">Registo de venda</div>
                {fatPendentes > 0 && (
                  <span style={{fontSize:11,background:"#fff3e0",color:"#7a3300",padding:"3px 10px",borderRadius:8,fontWeight:500}}>
                    {fatPendentes} fatura{fatPendentes > 1 ? "s" : ""} por processar
                  </span>
                )}
              </div>
              <div className="content">
                <div className="vendas-layout">

                  {/* Coluna esquerda */}
                  <div className="vendas-left">
                    <div className="card">
                      <div className="sec-lbl">Código do artigo</div>
                      <div className="code-row">
                        <input ref={codeRef} className="code-inp" placeholder="00-000" maxLength={8}
                          value={codeInput}
                          onChange={e => { setCodeInput(e.target.value.toUpperCase()); setCodeErr(""); }}
                          onKeyDown={e => { if (e.key === "Enter") { addByCode(); } }}
                          autoFocus />
                        <button className="btn-add" onClick={addByCode}><Plus /> Adicionar</button>
                      </div>
                      <div className="hint">Introduz o código e prime Enter</div>
                      {codeErr && <div className="code-err">{codeErr}</div>}
                    </div>

                    <div className="browser-sep">
                      <div className="browser-sep-line" />
                      <div className="browser-sep-txt">ou escolhe por artista</div>
                      <div className="browser-sep-line" />
                    </div>

                    <div className="browser">
                      {artistas.map(ar => {
                        const isOpen = browserArtista === ar.id;
                        return (
                          <div key={ar.id}>
                            <div className={`bar-artista${isOpen ? " on" : ""}`}
                              onClick={() => setBrowserArtista(isOpen ? null : ar.id)}>
                              <span className="ba-code">{p2(ar.codigo)}</span>
                              <span className="ba-nome">{ar.nome}</span>
                              {ar.nome_artistico && <span className="ba-arte">{ar.nome_artistico}</span>}
                              <span className="ba-chev"><ChevR /></span>
                            </div>
                            {isOpen && (
                              <div className="browser-artigos">
                                {artigos.filter(a => a.artista_id === ar.id).map(ar2 => {
                                  const stockDisp = stockOf(ar2.id);
                                  const qNoCarr   = carrinho.find(c => c.id === ar2.id)?.qty || 0;
                                  const podeAdd   = ar2.oferta || (stockDisp - qNoCarr) > 0;
                                  return (
                                    <div key={ar2.id} className={`bar-artigo${!podeAdd ? " esgotado" : ""}`}
                                      onClick={() => podeAdd && addFromBrowser(ar2.id, ar.id)}>
                                      <span className="baa-code">{mkCode(ar.codigo, ar2.codigo)}</span>
                                      <span className="baa-titulo">{ar2.titulo}</span>
                                      {ar2.oferta
                                        ? <span className="oferta-pill">oferta</span>
                                        : <>
                                            <span className="baa-preco">€{fmt(ar2.preco)}</span>
                                            <span className="baa-stock" style={{color: stockDisp===0?"#900":stockDisp<=2?"#7a3300":"#888"}}>
                                              {stockDisp}
                                            </span>
                                          </>}
                                      {qNoCarr > 0 && <span className="in-cart-pill">×{qNoCarr}</span>}
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Coluna direita: carrinho */}
                  <div className="vendas-right">
                    <div className="carr-panel">
                      <div className="carr-hdr">
                        <span className="carr-title">Carrinho</span>
                        <span className="carr-count">{carItens > 0 ? `${carItens} ${carItens === 1 ? "item" : "itens"}` : ""}</span>
                      </div>
                      <div className="carr-items">
                        {carrinho.length === 0
                          ? <div className="carr-empty">Nenhum artigo adicionado</div>
                          : carrinho.map((item, idx) => (
                              <div key={item.id} className="carr-item">
                                <div className="carr-info">
                                  <div className="carr-titulo">{item.titulo}</div>
                                  <div className="carr-sub">{mkCode(item.artista_codigo, item.codigo)} · {item.artista_nome}</div>
                                </div>
                                {item.oferta
                                  ? <span className="oferta-pill">oferta</span>
                                  : <>
                                      <div className="carr-qty">
                                        <div className="qty-btn" onClick={() => changeQty(idx, -1)}>−</div>
                                        <span className="qty-val">{item.qty}</span>
                                        <div className="qty-btn" onClick={() => changeQty(idx, +1)}>+</div>
                                      </div>
                                      <span className="carr-preco">€{fmt(item.preco * item.qty)}</span>
                                    </>}
                                <button className="carr-rm" onClick={() => removeItem(idx)}><X /></button>
                              </div>
                            ))}
                      </div>

                      <div className="carr-footer">
                        {vendaOk && (
                          <div className="venda-ok">
                            <div className="ok-ic"><Check /></div>
                            Venda registada com sucesso!
                          </div>
                        )}
                        {carrinho.length > 0 && !vendaOk ? <>
                          <div className="total-row">
                            <span className="total-lbl">Total</span>
                            <span className="total-val">€{fmt(carTotal)}</span>
                          </div>
                          <div className="sec-lbl">Pagamento</div>
                          <div className="metodos">
                            {METODOS.map(m => (
                              <button key={m} className={`met-btn${metodo === m ? " on" : ""}`} onClick={() => setMetodo(m)}>{m}</button>
                            ))}
                          </div>
                          <div className={`toggle-row`} onClick={() => { setQuerFatura(p => !p); setFatErr({}); }}>
                            <div className={`tbox${querFatura ? " on" : ""}`}>{querFatura && <Check />}</div>
                            <div className="toggle-lbl">Comprador quer fatura</div>
                          </div>
                          {querFatura && (
                            <div className="fat-fields">
                              <div>
                                <div className="fld-lbl">Nome</div>
                                <input className={`fld-inp${fatErr.nome ? " e" : ""}`} value={fatura.nome}
                                  onChange={e => setFatura(p => ({...p, nome: e.target.value}))} placeholder="Nome completo" />
                              </div>
                              <div>
                                <div className="fld-lbl">Email</div>
                                <input className={`fld-inp${fatErr.email ? " e" : ""}`} value={fatura.email}
                                  onChange={e => setFatura(p => ({...p, email: e.target.value}))} placeholder="email@exemplo.com" />
                              </div>
                              <div>
                                <div className="fld-lbl">NIF</div>
                                <input className={`fld-inp${fatErr.nif ? " e" : ""}`} value={fatura.nif}
                                  onChange={e => setFatura(p => ({...p, nif: e.target.value}))} placeholder="123456789" maxLength={9} />
                              </div>
                            </div>
                          )}
                          <button className="btn-confirmar" onClick={confirmarVenda} disabled={saving}>
                            {saving ? "A registar..." : `Confirmar — €${fmt(carTotal)}`}
                          </button>
                        </> : !vendaOk && (
                          <button className="btn-confirmar-empty" disabled>Adiciona artigos para confirmar</button>
                        )}
                      </div>
                    </div>
                  </div>

                </div>
              </div>
            </>}

            {/* ════════════════════════════════════════
                VENDAS REALIZADAS
            ════════════════════════════════════════ */}
            {page === "realizadas" && <>
              <div className="page-hdr"><div className="page-title">Vendas realizadas</div></div>
              <div className="content">
                <div className="real-layout">
                  <div>
                    {grupos.length === 0
                      ? <div className="empty">Nenhuma venda registada ainda.</div>
                      : <div className="venda-list">
                          {grupos.map((g, i) => (
                            <div key={g.key} className="venda-card">
                              <div className="venda-card-hdr">
                                <span className="vh-num">Venda #{String(i + 1).padStart(3, "0")}</span>
                                <span className="vh-hora">{g.hora}</span>
                                <span className="vh-met">{g.metodo}</span>
                                {g.quer_fatura && <span className="vh-fat">Fatura · {g.comprador_nif}</span>}
                                <span className="vh-total">€{fmt(g.total)}</span>
                              </div>
                              {g.itens.map((item, j) => (
                                <div key={j} className="venda-art-row">
                                  <span className="var-code">
                                    {item.artista_codigo && item.artigo_codigo
                                      ? mkCode(item.artista_codigo, item.artigo_codigo)
                                      : "—"}
                                  </span>
                                  <div className="var-info">
                                    <div className="var-titulo">{item.titulo}</div>
                                    <div className="var-artista">{item.artista_nome}</div>
                                  </div>
                                  {item.qty > 1 && <span className="var-qty">×{item.qty}</span>}
                                  {item.oferta
                                    ? <span className="oferta-pill">oferta</span>
                                    : <span className="var-preco">€{fmt(item.preco * item.qty)}</span>}
                                </div>
                              ))}
                            </div>
                          ))}
                        </div>}
                  </div>

                  <div>
                    <div className="stats-list">
                      <div className="stats-list-title">Resumo</div>
                      <div className="stat-row"><span className="stat-lbl">Vendas</span><span className="stat-val">{grupos.length}</span></div>
                      <div className="stat-row"><span className="stat-lbl">Total</span><span className="stat-val">€{fmt(totalGeral)}</span></div>
                      <div className="stat-row"><span className="stat-lbl">Comissão (10%)</span><span className="stat-val">€{fmt(totalGeral * COMM)}</span></div>
                      <div className="stat-row"><span className="stat-lbl">A pagar artistas</span><span className="stat-val">€{fmt(totalGeral * (1 - COMM))}</span></div>
                      <div className="stat-row"><span className="stat-lbl">Dinheiro</span><span className="stat-val">€{fmt(grupos.filter(g=>g.metodo==="Dinheiro").reduce((s,g)=>s+g.total,0))}</span></div>
                      <div className="stat-row"><span className="stat-lbl">MB Way</span><span className="stat-val">€{fmt(grupos.filter(g=>g.metodo==="MB Way").reduce((s,g)=>s+g.total,0))}</span></div>
                      <div className="stat-row"><span className="stat-lbl">Transferência</span><span className="stat-val">€{fmt(grupos.filter(g=>g.metodo==="Transferência").reduce((s,g)=>s+g.total,0))}</span></div>
                      <div className="stat-row"><span className="stat-lbl">Faturas</span><span className="stat-val">{faturasProcessadas.size}/{grupos.filter(g=>g.quer_fatura).length}</span></div>
                    </div>
                  </div>
                </div>
              </div>
            </>}

            {/* ════════════════════════════════════════
                FATURAS
            ════════════════════════════════════════ */}
            {page === "faturas" && (() => {
              const fatGrupos    = grupos.filter(g => g.quer_fatura);
              const pendentes    = fatGrupos.filter(g => !faturasProcessadas.has(g.key)).length;
              const processadas  = faturasProcessadas.size;
              return <>
                <div className="page-hdr"><div className="page-title">Faturas</div></div>
                <div className="content">
                  <div className="fat-layout">
                    <div>
                      {fatGrupos.length === 0
                        ? <div className="empty">Nenhuma fatura solicitada até ao momento.</div>
                        : <div className="fat-table">
                            <div className="fat-thead">
                              <div className="fat-th">
                                <span>#</span><span>Nome</span><span>Email / NIF</span>
                                <span>Artigos</span><span>Valor</span><span>Estado</span>
                              </div>
                            </div>
                            {fatGrupos.map((g, i) => {
                              const isDone      = faturasProcessadas.has(g.key);
                              const artigosStr  = g.itens.filter(it => !it.oferta)
                                .map(it => `${it.titulo}${it.qty > 1 ? ` ×${it.qty}` : ""}`).join(", ");
                              return (
                                <div key={g.key} className={`fat-row-item${isDone ? " done" : ""}`}>
                                  <span className="fat-cell mono">{String(i+1).padStart(3,"0")}</span>
                                  <span className="fat-cell bold">{g.comprador_nome}</span>
                                  <div className="fat-cell" style={{lineHeight:1.6}}>
                                    <div style={{fontSize:12,color:"#444"}}>{g.comprador_email || "—"}</div>
                                    <div style={{fontSize:11,color:"#777",fontWeight:600}}>{g.comprador_nif || "—"}</div>
                                  </div>
                                  <span className="fat-cell" style={{fontSize:11,color:"#555"}}>{artigosStr}</span>
                                  <span className="fat-cell bold">€{fmt(g.total)}</span>
                                  <span className="fat-cell">
                                    {isDone
                                      ? <span className="done-badge"><Check /> Enviada</span>
                                      : <button className="btn-proc" onClick={() => toggleFatProc(g.key)}>Marcar enviada</button>}
                                  </span>
                                </div>
                              );
                            })}
                          </div>}
                    </div>
                    <div>
                      <div className="fat-stats">
                        <div className="fat-stats-title">Resumo</div>
                        <div className="stat-row"><span className="stat-lbl">Total faturas</span><span className="stat-val">{fatGrupos.length}</span></div>
                        <div className="stat-row"><span className="stat-lbl">Por enviar</span><span className="stat-val" style={{color:pendentes>0?"#7a3300":"#111"}}>{pendentes}</span></div>
                        <div className="stat-row"><span className="stat-lbl">Enviadas</span><span className="stat-val" style={{color:processadas>0?"#2a5a1a":"#111"}}>{processadas}</span></div>
                        <div className="stat-row"><span className="stat-lbl">Valor total</span><span className="stat-val">€{fmt(fatGrupos.reduce((s,g)=>s+g.total,0))}</span></div>
                      </div>
                    </div>
                  </div>
                </div>
              </>;
            })()}

            {/* ════════════════════════════════════════
                INVENTÁRIO
            ════════════════════════════════════════ */}
            {page === "inventario" && <>
              <div className="page-hdr"><div className="page-title">Inventário</div></div>
              <div className="content">
                <div className="inv-container">
                  <div className="inv-hdr-row">
                    <span>Código</span>
                    <span>Título</span>
                    <span style={{textAlign:"right"}}>Preço</span>
                    <span style={{textAlign:"center"}}>Stock</span>
                    <span style={{textAlign:"right"}}>Vendido</span>
                  </div>
                  {artistas.map(ar => {
                    const arts    = artigos.filter(a => a.artista_id === ar.id);
                    const isExp   = expanded.has(ar.id);
                    const nVendas = grupos.filter(g => g.itens.some(i => i.artista && i.artista.id === ar.id)).length;
                    return (
                      <div key={ar.id} className="inv-section">
                        <div className="inv-artist-bar" onClick={() => toggleExp(ar.id)}>
                          <span className="inv-acode">{p2(ar.codigo)}</span>
                          <span className="inv-anome">{ar.nome}</span>
                          {ar.nome_artistico && <span className="inv-aarte">&nbsp;· {ar.nome_artistico}</span>}
                          <span style={{flex:1}} />
                          <span className="inv-astats">{arts.length} artigos · {nVendas} vendas</span>
                          <span style={{display:"flex",color:"#bbb",marginLeft:6}}>{isExp ? <ChevU /> : <ChevD />}</span>
                        </div>
                        {isExp && arts.map(ar2 => {
                          const stock   = ar2.oferta ? null : stockOf(ar2.id);
                          const vendido = ar2.oferta ? 0 : (ar2.quantidade - stock);
                          return (
                            <div key={ar2.id} className="inv-artigo-row">
                              <span style={{fontSize:11,fontWeight:600,color:"#888",paddingLeft:28}}>{mkCode(ar.codigo, ar2.codigo)}</span>
                              <span style={{fontWeight:500,color:"#111"}}>{ar2.titulo}</span>
                              <span style={{textAlign:"right",color:"#444"}}>{ar2.oferta ? "—" : `€${fmt(ar2.preco)}`}</span>
                              <span style={{textAlign:"center"}}>
                                {ar2.oferta
                                  ? <span className="oferta-pill">oferta</span>
                                  : <span className={`stk${stock===0?" zero":stock<=2?" low":""}`}>{stock}</span>}
                              </span>
                              <span style={{textAlign:"right",color:vendido>0?"#2a5a1a":"#bbb",fontWeight:vendido>0?500:400}}>
                                {ar2.oferta ? "—" : vendido}
                              </span>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })}
                </div>
              </div>
            </>}

            {/* ════════════════════════════════════════
                ARTISTAS
            ════════════════════════════════════════ */}
            {page === "artistas" && <>
              <div className="page-hdr"><div className="page-title">Artistas</div></div>
              <div className="content">
                <div className="art-list">
                  {artistas.map(ar => {
                    const arts   = artigos.filter(a => a.artista_id === ar.id);
                    const vv     = vendas.filter(v => v.artista_id === ar.id && !v.oferta);
                    const total  = vv.reduce((s, v) => s + (v.preco * (v.quantidade || 1)), 0);
                    const stockI = arts.filter(a => !a.oferta).reduce((s, a) => s + a.quantidade, 0);
                    const stockA = arts.filter(a => !a.oferta).reduce((s, a) => s + stockOf(a.id), 0);
                    const pecas  = vv.reduce((s, v) => s + (v.quantidade || 1), 0);
                    return (
                      <div key={ar.id} className="art-card">
                        <div className="art-top">
                          <span className="art-code">{p2(ar.codigo)}</span>
                          <span className="art-nome">{ar.nome}</span>
                          {ar.nome_artistico && <span className="art-arte">{ar.nome_artistico}</span>}
                        </div>
                        <div className="art-grid">
                          <div className="art-col">
                            <div className="art-f"><span className="art-fl">Email</span><span className="art-fv">{ar.email}</span></div>
                            <div className="art-f"><span className="art-fl">Contacto</span><span className="art-fv">{ar.contacto}</span></div>
                          </div>
                          <div className="art-col">
                            <div className="art-f"><span className="art-fl">Stock inicial</span><span className="art-fv">{stockI} peças</span></div>
                            <div className="art-f"><span className="art-fl">Stock actual</span><span className="art-fv">{stockA} peças</span></div>
                          </div>
                          <div className="art-col">
                            <div className="art-f"><span className="art-fl">Artigos vendidos</span><span className="art-fv">{pecas} peças</span></div>
                            <div className="art-f"><span className="art-fl">A receber</span><span className="art-fv" style={{color:"#2a5a1a"}}>€{fmt(total * (1 - COMM))}</span></div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </>}

            {/* ════════════════════════════════════════
                RELATÓRIO
            ════════════════════════════════════════ */}
            {page === "relatorio" && <>
              <div className="page-hdr"><div className="page-title">Relatório final</div></div>
              <div className="content">
                <div className="rel-card">
                  <div className="rel-title">Exportar acerto final</div>
                  <div className="rel-desc">
                    Gera um ficheiro Excel com um separador por artista (vendas detalhadas e montante a pagar),
                    um resumo geral e a lista de compradores que pediram fatura.
                  </div>
                  <button className="btn-export" onClick={exportar}><Dl /> Descarregar relatório Excel</button>
                </div>
                {artistas.some(ar => vendas.some(v => v.artista_id === ar.id && !v.oferta)) && (
                  <div className="rel-card">
                    <div className="rel-title" style={{marginBottom:14}}>Resumo por artista</div>
                    {artistas.map(ar => {
                      const vv    = vendas.filter(v => v.artista_id === ar.id && !v.oferta);
                      if (!vv.length) return null;
                      const total = vv.reduce((s, v) => s + (v.preco * (v.quantidade || 1)), 0);
                      return (
                        <div key={ar.id} className="rel-row">
                          <div>
                            <span style={{fontSize:11,color:"#aaa",marginRight:10}}>{p2(ar.codigo)}</span>
                            <strong style={{fontWeight:500}}>{ar.nome_artistico || ar.nome}</strong>
                            <span style={{fontSize:12,color:"#888",marginLeft:10}}>
                              {vv.reduce((s,v)=>s+(v.quantidade||1),0)} peças
                            </span>
                          </div>
                          <div style={{textAlign:"right"}}>
                            <div style={{fontSize:15,fontWeight:500}}>€{fmt(total * (1 - COMM))}</div>
                            <div style={{fontSize:11,color:"#888"}}>de €{fmt(total)} públicos</div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </>}

          </>}
        </main>
      </div>
    </>
  );
}
