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

const Ic = {
  Check:   () => <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  X:       () => <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  ChDown:  () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 12 15 18 9"/></svg>,
  ChUp:    () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="18 15 12 9 6 15"/></svg>,
  Dl:      () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  Search:  () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
  Refresh: () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>,
  Gift:    () => <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 12 20 22 4 22 4 12"/><rect x="2" y="7" width="20" height="5"/><line x1="12" y1="22" x2="12" y2="7"/><path d="M12 7H7.5a2.5 2.5 0 0 1 0-5C11 2 12 7 12 7z"/><path d="M12 7h4.5a2.5 2.5 0 0 0 0-5C13 2 12 7 12 7z"/></svg>,
};

const CSS = `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #f4f4f2; }
.adm { min-height: 100vh; font-family: 'Inter', sans-serif; color: #111; font-size: 14px; background: #f4f4f2; }

.login { min-height: 100vh; display: flex; align-items: center; justify-content: center; }
.login-box { background: #fff; border-radius: 16px; padding: 40px; width: 100%; max-width: 360px; }
.login-title { font-size: 18px; font-weight: 500; margin-bottom: 4px; }
.login-sub { font-size: 12px; color: #999; margin-bottom: 28px; }

.lbl { display: block; font-size: 11px; font-weight: 500; color: #444; margin-bottom: 6px; }
.inp { width: 100%; background: #f0f0ee; border: none; outline: none; color: #000; font-family: 'Inter', sans-serif; font-size: 13px; padding: 10px 14px; border-radius: 8px; transition: background .15s; }
.inp:focus { background: #e8e8e5; }
.inp.e { background: #fce8e8; }
.em { font-size: 11px; color: #d00; margin-top: 4px; }

.btn-primary { width: 100%; background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 500; padding: 11px; border-radius: 8px; cursor: pointer; margin-top: 16px; transition: background .12s; }
.btn-primary:hover { background: #333; }

.adm-hdr { background: #fff; border-bottom: 1px solid #ebebeb; padding: 0 28px; display: flex; align-items: center; gap: 14px; height: 50px; }
.adm-brand { font-size: 12px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; }
.adm-sep { width: 1px; height: 14px; background: #ddd; }
.adm-sub { font-size: 11px; color: #999; }
.spacer { flex: 1; }
.btn-sm { background: #f0f0ee; border: none; font-family: 'Inter', sans-serif; font-size: 11px; color: #555; cursor: pointer; padding: 6px 10px; border-radius: 6px; display: flex; align-items: center; gap: 5px; transition: background .12s; }
.btn-sm:hover { background: #e4e4e1; }
.btn-logout { background: none; border: none; font-family: 'Inter', sans-serif; font-size: 11px; color: #bbb; cursor: pointer; padding: 6px 10px; border-radius: 6px; }
.btn-logout:hover { color: #666; }

.adm-tabs { background: #fff; border-bottom: 1px solid #ebebeb; padding: 0 28px; display: flex; }
.tab { padding: 12px 0; margin-right: 24px; font-size: 12px; font-weight: 500; color: #bbb; cursor: pointer; border-bottom: 2px solid transparent; margin-bottom: -1px; transition: color .12s; }
.tab:hover { color: #666; }
.tab.on { color: #111; border-bottom-color: #111; }

.content { max-width: 920px; margin: 0 auto; padding: 28px 20px; }

.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 24px; }
.stat { background: #fff; border-radius: 12px; padding: 16px 18px; }
.stat-lbl { font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #aaa; margin-bottom: 6px; }
.stat-val { font-size: 22px; font-weight: 500; line-height: 1; }
.stat-sub { font-size: 11px; color: #999; margin-top: 4px; }

.card { background: #fff; border-radius: 12px; margin-bottom: 8px; overflow: hidden; }
.artista-hdr { display: flex; align-items: center; gap: 12px; padding: 14px 18px; cursor: pointer; transition: background .12s; user-select: none; }
.artista-hdr:hover { background: #fafaf8; }
.a-code { font-size: 11px; font-weight: 600; color: #aaa; min-width: 24px; font-variant-numeric: tabular-nums; }
.a-nome { font-size: 14px; font-weight: 500; flex: 1; }
.a-arte { font-size: 12px; color: #888; }
.a-stats { font-size: 11px; color: #aaa; text-align: right; line-height: 1.7; min-width: 140px; }
.a-ic { color: #ccc; display: flex; }

.artigo-thead { background: #fafaf8; border-top: 1px solid #f0f0ee; }
.artigo-row { display: grid; grid-template-columns: 72px 1fr 110px 80px 60px 60px; gap: 10px; align-items: center; padding: 9px 18px; font-size: 12px; border-top: 1px solid #f4f4f2; }
.artigo-row:hover { background: #fafaf8; }
.artigo-thead .artigo-row { font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #bbb; padding: 8px 18px; border-top: 1px solid #f0f0ee; }
.artigo-thead .artigo-row:hover { background: transparent; }
.ar-code { font-variant-numeric: tabular-nums; font-weight: 600; color: #555; font-size: 11px; }
.ar-titulo { font-weight: 500; }
.ar-cat { color: #aaa; font-size: 11px; }
.stock { display: inline-flex; align-items: center; justify-content: center; min-width: 26px; height: 20px; border-radius: 5px; font-size: 11px; font-weight: 600; background: #f0f0ee; padding: 0 5px; }
.stock.low { background: #fff3e0; color: #e65100; }
.stock.zero { background: #fce8e8; color: #c00; }
.stock.oferta { background: #e8f4e8; color: #1a6a1a; font-size: 9px; }

.venda-box { background: #fff; border-radius: 12px; padding: 22px 24px; margin-bottom: 14px; }
.venda-title { font-size: 15px; font-weight: 500; margin-bottom: 4px; }
.venda-desc { font-size: 12px; color: #888; margin-bottom: 20px; }
.code-row { display: flex; gap: 10px; align-items: flex-start; }
.code-inp { background: #f0f0ee; border: none; outline: none; color: #000; font-family: 'Inter', sans-serif; font-size: 20px; font-weight: 600; letter-spacing: 3px; text-transform: uppercase; padding: 11px 16px; border-radius: 8px; width: 160px; transition: background .15s; }
.code-inp:focus { background: #e8e8e5; }
.btn-buscar { background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 500; padding: 12px 20px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 8px; height: 46px; transition: background .12s; }
.btn-buscar:hover { background: #333; }
.code-err { font-size: 12px; color: #d00; margin-top: 8px; }

.found-card { background: #fff; border-radius: 12px; padding: 20px 24px; margin-bottom: 14px; border: 1.5px solid #e8e8e8; }
.found-card.ok { border-color: #4caf50; background: #f6fff6; }
.found-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }
.found-fc { font-size: 10px; font-weight: 600; letter-spacing: 1px; color: #aaa; text-transform: uppercase; margin-bottom: 4px; }
.found-titulo { font-size: 18px; font-weight: 500; }
.found-artista { font-size: 13px; color: #666; margin-top: 2px; }
.found-price { font-size: 22px; font-weight: 500; text-align: right; }
.found-share { font-size: 11px; color: #888; text-align: right; margin-top: 2px; }
.found-stock { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: #666; background: #f0f0ee; padding: 4px 10px; border-radius: 6px; margin-bottom: 16px; }

.sec-lbl { font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #aaa; margin-bottom: 8px; }
.metodos { display: flex; gap: 8px; margin-bottom: 14px; }
.metodo-btn { flex: 1; background: #f0f0ee; border: none; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; color: #555; padding: 9px 6px; border-radius: 8px; cursor: pointer; transition: all .12s; }
.metodo-btn.on { background: #111; color: #fff; }

.toggle { display: flex; align-items: center; gap: 10px; padding: 10px 14px; background: #f0f0ee; border-radius: 8px; cursor: pointer; margin-bottom: 10px; transition: background .12s; user-select: none; }
.toggle:hover { background: #e8e8e5; }
.toggle-box { width: 15px; height: 15px; background: #d4d4d1; flex-shrink: 0; border-radius: 4px; display: flex; align-items: center; justify-content: center; transition: all .15s; }
.toggle.on .toggle-box { background: #111; color: #fff; }
.toggle-lbl { font-size: 12px; color: #555; }

.fat-fields { background: #f6f6f4; border-radius: 10px; padding: 14px; margin-bottom: 12px; display: flex; flex-direction: column; gap: 10px; }
.g2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }

.btn-confirmar { background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 14px; font-weight: 500; padding: 14px; border-radius: 10px; cursor: pointer; width: 100%; transition: background .12s; }
.btn-confirmar:hover { background: #333; }
.btn-confirmar:disabled { background: #ccc; cursor: not-allowed; }

.venda-ok { display: flex; align-items: center; gap: 12px; padding: 14px 16px; background: #e6f4e6; border-radius: 10px; font-size: 13px; color: #1a5a1a; font-weight: 500; }
.venda-ok-ic { width: 30px; height: 30px; background: #4caf50; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #fff; flex-shrink: 0; }

.recentes { background: #fff; border-radius: 12px; padding: 20px 24px; }
.recentes-title { font-size: 11px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #aaa; margin-bottom: 14px; }
.venda-row { display: flex; align-items: center; gap: 12px; padding: 9px 0; font-size: 12px; border-top: 1px solid #f4f4f2; }
.vr-code { font-weight: 600; color: #555; font-variant-numeric: tabular-nums; min-width: 58px; font-size: 11px; }
.vr-titulo { flex: 1; }
.vr-metodo { color: #888; min-width: 80px; }
.vr-preco { font-weight: 500; min-width: 50px; text-align: right; }
.vr-hora { color: #bbb; font-size: 11px; min-width: 38px; text-align: right; }
.fat-badge { font-size: 10px; background: #fff3e0; color: #b75000; padding: 2px 6px; border-radius: 4px; font-weight: 600; white-space: nowrap; }

.fat-card { background: #fff; border-radius: 12px; overflow: hidden; }
.fat-row { display: grid; grid-template-columns: 72px 1fr 1fr 120px 70px; gap: 10px; align-items: center; padding: 10px 18px; font-size: 12px; border-top: 1px solid #f4f4f2; }
.fat-row:hover { background: #fafaf8; }
.fat-thead .fat-row { font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #bbb; padding: 9px 18px; background: #fafaf8; border-top: none; }
.fat-thead .fat-row:hover { background: #fafaf8; }

.rel-card { background: #fff; border-radius: 12px; padding: 22px 24px; margin-bottom: 10px; }
.rel-title { font-size: 15px; font-weight: 500; margin-bottom: 4px; }
.rel-desc { font-size: 12px; color: #888; margin-bottom: 18px; line-height: 1.8; }
.btn-export { background: #f0f0ee; border: none; color: #333; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 500; padding: 11px 18px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 8px; transition: background .12s; }
.btn-export:hover { background: #e4e4e1; }
.rel-row { display: flex; justify-content: space-between; align-items: center; padding: 13px 0; border-top: 1px solid #f4f4f2; font-size: 13px; }
.rel-row:first-of-type { border-top: none; }

.empty { text-align: center; padding: 52px 24px; color: #bbb; font-size: 13px; }
.oferta-badge { display: inline-flex; align-items: center; gap: 4px; font-size: 10px; font-weight: 600; color: #1a6a1a; background: #e8f4e8; padding: 2px 7px; border-radius: 4px; }

@media (max-width: 700px) {
  .stats { grid-template-columns: 1fr 1fr; }
  .content { padding: 16px 12px; }
  .adm-hdr, .adm-tabs { padding: 0 14px; }
  .artigo-row { grid-template-columns: 60px 1fr 60px 50px; }
  .artigo-row > :nth-child(3), .artigo-row > :nth-child(6) { display: none; }
}
`;

export default function Admin() {
  const [auth, setAuth]   = useState(() => sessionStorage.getItem("obs_adm") === "1");
  const [pw, setPw]       = useState("");
  const [pwErr, setPwErr] = useState(false);
  const [tab, setTab]     = useState(0);

  const [artistas, setArtistas] = useState([]);
  const [artigos, setArtigos]   = useState([]);
  const [vendas, setVendas]     = useState([]);
  const [loading, setLoading]   = useState(true);
  const [expanded, setExpanded] = useState(new Set());

  const [codigoInput, setCodigoInput] = useState("");
  const [found, setFound]             = useState(null);
  const [foundErr, setFoundErr]       = useState("");
  const [metodo, setMetodo]           = useState("Dinheiro");
  const [querFatura, setQuerFatura]   = useState(false);
  const [fatura, setFatura]           = useState({ nome: "", email: "", nif: "" });
  const [fatErr, setFatErr]           = useState({});
  const [vendaOk, setVendaOk]         = useState(false);
  const [saving, setSaving]           = useState(false);
  const codeRef = useRef(null);

  const login = () => {
    if (pw === PASS) { sessionStorage.setItem("obs_adm", "1"); setAuth(true); }
    else { setPwErr(true); }
  };
  const logout = () => { sessionStorage.removeItem("obs_adm"); setAuth(false); };

  useEffect(() => { if (auth) load(); }, [auth]);

  const load = async () => {
    setLoading(true);
    const [{ data: a }, { data: ar }, { data: v }] = await Promise.all([
      supabase.from("artistas").select("*").order("codigo"),
      supabase.from("artigos").select("*").order("artista_id,codigo"),
      supabase.from("vendas").select("*").order("created_at", { ascending: false }),
    ]);
    setArtistas(a || []);
    setArtigos(ar || []);
    setVendas(v || []);
    setLoading(false);
  };

  const stockOf = id => {
    const ar = artigos.find(a => a.id === id);
    if (!ar) return 0;
    return ar.quantidade - vendas.filter(v => v.artigo_id === id).length;
  };

  const artsOf  = id => artigos.filter(a => a.artista_id === id);
  const vendsOf = id => vendas.filter(v => v.artista_id === id);

  const buscar = () => {
    setFoundErr(""); setFound(null); setVendaOk(false);
    const raw = codigoInput.trim();
    const m = raw.match(/^(\d+)-(\d+)$/);
    if (!m) { setFoundErr("Formato inválido — exemplo: 01-003"); return; }
    const codA = parseInt(m[1]), codAr = parseInt(m[2]);
    const artista = artistas.find(a => a.codigo === codA);
    if (!artista) { setFoundErr("Código de artista não encontrado"); return; }
    const artigo = artigos.find(a => a.artista_id === artista.id && a.codigo === codAr);
    if (!artigo) { setFoundErr("Código de artigo não encontrado"); return; }
    if (artigo.oferta) { setFoundErr("Este artigo é uma oferta — não está à venda"); return; }
    const stock = stockOf(artigo.id);
    if (stock <= 0) { setFoundErr("Sem stock disponível"); return; }
    setFound({ ...artigo, artista, stockAtual: stock });
  };

  const registar = async () => {
    if (!found) return;
    const errs = {};
    if (querFatura) {
      if (!fatura.nome.trim())  errs.nome  = "Obrigatório";
      if (!fatura.email.trim()) errs.email = "Obrigatório";
      if (!fatura.nif.trim())   errs.nif   = "Obrigatório";
    }
    if (Object.keys(errs).length) { setFatErr(errs); return; }
    setSaving(true);
    const { error } = await supabase.from("vendas").insert({
      artigo_id: found.id,
      artista_id: found.artista_id,
      preco: found.preco,
      metodo,
      quer_fatura: querFatura,
      comprador_nome:  querFatura ? fatura.nome  : null,
      comprador_email: querFatura ? fatura.email : null,
      comprador_nif:   querFatura ? fatura.nif   : null,
    });
    setSaving(false);
    if (!error) {
      setVendaOk(true);
      setCodigoInput("");
      setQuerFatura(false);
      setFatura({ nome: "", email: "", nif: "" });
      setFatErr({});
      setMetodo("Dinheiro");
      await load();
      setTimeout(() => { setFound(null); setVendaOk(false); codeRef.current?.focus(); }, 3000);
    }
  };

  const exportar = () => {
    const wb = XLSX.utils.book_new();

    const resumo = [
      ["RELATÓRIO FINAL — OBSCENA"],
      [""],
      ["Cód.", "Artista", "Vendas", "Peças", "Total público (€)", "Comissão (€)", "A pagar (€)"],
    ];
    artistas.forEach(ar => {
      const vv = vendsOf(ar.id);
      const total = vv.reduce((s, v) => s + (+v.preco || 0), 0);
      resumo.push([p2(ar.codigo), ar.nome_artistico || ar.nome,
        [...new Set(vv.map(v => v.artigo_id))].length, vv.length,
        fmt(total), fmt(total * COMM), fmt(total * (1 - COMM))]);
    });
    const gt = vendas.reduce((s, v) => s + (+v.preco || 0), 0);
    resumo.push(["", "TOTAL", "", vendas.length, fmt(gt), fmt(gt * COMM), fmt(gt * (1 - COMM))]);
    const ws0 = XLSX.utils.aoa_to_sheet(resumo);
    ws0["!cols"] = [{wch:6},{wch:24},{wch:8},{wch:8},{wch:18},{wch:14},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws0, "Resumo");

    artistas.forEach(ar => {
      const vv = vendsOf(ar.id);
      if (!vv.length) return;
      const rows = [[ar.nome_artistico || ar.nome], [""],
        ["Hora", "Código", "Artigo", "Preço (€)", "Método", "Fatura"]];
      vv.forEach(v => {
        const ar2 = artigos.find(a => a.id === v.artigo_id);
        rows.push([
          new Date(v.created_at).toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" }),
          ar2 ? mkCode(ar.codigo, ar2.codigo) : "—",
          ar2?.titulo || "—", fmt(v.preco), v.metodo,
          v.quer_fatura ? `${v.comprador_nome} / ${v.comprador_nif}` : "Não",
        ]);
      });
      const tot = vv.reduce((s, v) => s + (+v.preco || 0), 0);
      rows.push(["", "", "Total vendas", fmt(tot)]);
      rows.push(["", "", "A pagar ao artista", fmt(tot * (1 - COMM))]);
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = [{wch:8},{wch:10},{wch:30},{wch:12},{wch:14},{wch:30}];
      XLSX.utils.book_append_sheet(wb, ws, (ar.nome_artistico || ar.nome).slice(0, 28).replace(/[:\\/?*[\]]/g, ""));
    });

    const fatRows = [["FATURAÇÃO — COMPRADORES"], [""],
      ["Hora", "Código", "Artigo", "Artista", "Preço (€)", "Nome", "Email", "NIF"]];
    vendas.filter(v => v.quer_fatura).forEach(v => {
      const ar2 = artigos.find(a => a.id === v.artigo_id);
      const ar  = artistas.find(a => a.id === v.artista_id);
      fatRows.push([
        new Date(v.created_at).toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" }),
        ar2 && ar ? mkCode(ar.codigo, ar2.codigo) : "—",
        ar2?.titulo || "—", ar ? (ar.nome_artistico || ar.nome) : "—",
        fmt(v.preco), v.comprador_nome, v.comprador_email, v.comprador_nif,
      ]);
    });
    const wsf = XLSX.utils.aoa_to_sheet(fatRows);
    wsf["!cols"] = [{wch:8},{wch:10},{wch:28},{wch:20},{wch:12},{wch:24},{wch:28},{wch:12}];
    XLSX.utils.book_append_sheet(wb, wsf, "Faturação");

    XLSX.writeFile(wb, "obscena_relatorio_final.xlsx");
  };

  const fmtH = ts => new Date(ts).toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" });

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

  const totalReceita  = vendas.reduce((s, v) => s + (+v.preco || 0), 0);
  const fatPendente   = vendas.filter(v => v.quer_fatura);
  const TABS = ["Inventário", "Vendas", `Faturação${fatPendente.length ? ` (${fatPendente.length})` : ""}`, "Relatório"];

  return (
    <>
      <style>{FONT}{CSS}</style>
      <div className="adm">

        <div className="adm-hdr">
          <span className="adm-brand">Obscena</span>
          <div className="adm-sep" />
          <span className="adm-sub">Painel de administração</span>
          <div className="spacer" />
          <button className="btn-sm" onClick={load}><Ic.Refresh /> Actualizar</button>
          <button className="btn-logout" onClick={logout}>Sair</button>
        </div>

        <div className="adm-tabs">
          {TABS.map((t, i) => (
            <div key={i} className={`tab${tab === i ? " on" : ""}`} onClick={() => setTab(i)}>{t}</div>
          ))}
        </div>

        <div className="content">
          {loading ? <div className="empty">A carregar...</div> : <>

            {/* ══ INVENTÁRIO ══ */}
            {tab === 0 && <>
              <div className="stats">
                <div className="stat">
                  <div className="stat-lbl">Artistas</div>
                  <div className="stat-val">{artistas.length}</div>
                </div>
                <div className="stat">
                  <div className="stat-lbl">Artigos à venda</div>
                  <div className="stat-val">{artigos.filter(a => !a.oferta).length}</div>
                  {artigos.filter(a => a.oferta).length > 0 &&
                    <div className="stat-sub">+ {artigos.filter(a => a.oferta).length} ofertas</div>}
                </div>
                <div className="stat">
                  <div className="stat-lbl">Vendas</div>
                  <div className="stat-val">{vendas.length}</div>
                  <div className="stat-sub">{fmt(totalReceita)} €</div>
                </div>
                <div className="stat">
                  <div className="stat-lbl">A pagar artistas</div>
                  <div className="stat-val">{fmt(totalReceita * (1 - COMM))} €</div>
                </div>
              </div>

              {artistas.length === 0
                ? <div className="empty">Nenhum artista registado ainda.</div>
                : artistas.map(ar => {
                  const arts = artsOf(ar.id);
                  const vv   = vendsOf(ar.id);
                  const isOpen = expanded.has(ar.id);
                  return (
                    <div key={ar.id} className="card">
                      <div className="artista-hdr" onClick={() => {
                        const next = new Set(expanded);
                        isOpen ? next.delete(ar.id) : next.add(ar.id);
                        setExpanded(next);
                      }}>
                        <span className="a-code">{p2(ar.codigo)}</span>
                        <span className="a-nome">{ar.nome}</span>
                        {ar.nome_artistico && <span className="a-arte">{ar.nome_artistico}</span>}
                        <div className="a-stats">
                          <div>{arts.length} artigos · {arts.reduce((s, a) => s + (+a.quantidade || 0), 0)} peças</div>
                          {vv.length > 0 && <div style={{color:"#2a7a2a"}}>{vv.length} vendas · {fmt(vv.reduce((s,v)=>s+(+v.preco||0),0))} €</div>}
                        </div>
                        <span className="a-ic">{isOpen ? <Ic.ChUp /> : <Ic.ChDown />}</span>
                      </div>

                      {isOpen && <>
                        {arts.length === 0
                          ? <div style={{padding:"12px 18px",fontSize:12,color:"#bbb",borderTop:"1px solid #f4f4f2"}}>Sem artigos</div>
                          : <>
                            <div className="artigo-thead">
                              <div className="artigo-row">
                                <span>Código</span><span>Título</span><span>Categoria</span>
                                <span style={{textAlign:"right"}}>Preço</span>
                                <span style={{textAlign:"center"}}>Stock</span>
                                <span style={{textAlign:"right"}}>Vendas</span>
                              </div>
                            </div>
                            {arts.map(ar2 => {
                              const stock   = ar2.oferta ? null : stockOf(ar2.id);
                              const vendido = ar2.oferta ? 0 : (ar2.quantidade - stock);
                              return (
                                <div key={ar2.id} className="artigo-row">
                                  <span className="ar-code">{mkCode(ar.codigo, ar2.codigo)}</span>
                                  <span className="ar-titulo">{ar2.titulo}</span>
                                  <span className="ar-cat">{ar2.categoria}</span>
                                  <span style={{textAlign:"right"}}>{ar2.oferta ? "—" : `€${fmt(ar2.preco)}`}</span>
                                  <span style={{textAlign:"center"}}>
                                    {ar2.oferta
                                      ? <span className="stock oferta"><Ic.Gift /> oferta</span>
                                      : <span className={`stock${stock===0?" zero":stock<=2?" low":""}`}>{stock}</span>}
                                  </span>
                                  <span style={{textAlign:"right",color:vendido>0?"#2a7a2a":"#ccc"}}>{ar2.oferta?"—":vendido}</span>
                                </div>
                              );
                            })}
                          </>}
                      </>}
                    </div>
                  );
                })
              }
            </>}

            {/* ══ VENDAS ══ */}
            {tab === 1 && <>
              <div className="venda-box">
                <div className="venda-title">Registar venda</div>
                <div className="venda-desc">Introduz o código do artigo (ex: 01-003) e confirma a venda.</div>
                <div className="code-row">
                  <input ref={codeRef} className="code-inp" value={codigoInput}
                    onChange={e => { setCodigoInput(e.target.value); setFoundErr(""); }}
                    onKeyDown={e => e.key === "Enter" && buscar()}
                    placeholder="00-000" maxLength={8} autoFocus />
                  <button className="btn-buscar" onClick={buscar}><Ic.Search /> Procurar</button>
                </div>
                {foundErr && <div className="code-err">{foundErr}</div>}
              </div>

              {found && (
                <div className={`found-card${vendaOk ? " ok" : ""}`}>
                  {vendaOk ? (
                    <div className="venda-ok">
                      <div className="venda-ok-ic"><Ic.Check /></div>
                      Venda registada! A preparar para o próximo...
                    </div>
                  ) : <>
                    <div className="found-top">
                      <div>
                        <div className="found-fc">{mkCode(found.artista.codigo, found.codigo)}</div>
                        <div className="found-titulo">{found.titulo}</div>
                        <div className="found-artista">{found.artista.nome_artistico || found.artista.nome}</div>
                      </div>
                      <div>
                        <div className="found-price">€{fmt(found.preco)}</div>
                        <div className="found-share">Artista recebe €{fmt(found.preco * (1 - COMM))}</div>
                      </div>
                    </div>

                    <div className="found-stock">
                      {found.stockAtual} {found.stockAtual === 1 ? "peça disponível" : "peças disponíveis"}
                      {found.stockAtual <= 2 && <span style={{color:"#e65100",fontWeight:600}}> · Quase esgotado</span>}
                    </div>

                    <div className="sec-lbl">Método de pagamento</div>
                    <div className="metodos">
                      {METODOS.map(m => (
                        <button key={m} className={`metodo-btn${metodo===m?" on":""}`} onClick={() => setMetodo(m)}>{m}</button>
                      ))}
                    </div>

                    <div className={`toggle${querFatura?" on":""}`} onClick={() => { setQuerFatura(p=>!p); setFatErr({}); }}>
                      <div className="toggle-box">{querFatura && <Ic.Check />}</div>
                      <div className="toggle-lbl">Comprador quer fatura</div>
                    </div>

                    {querFatura && (
                      <div className="fat-fields">
                        <div>
                          <label className="lbl">Nome</label>
                          <input className={`inp${fatErr.nome?" e":""}`} value={fatura.nome}
                            onChange={e => setFatura(p=>({...p,nome:e.target.value}))} placeholder="Nome completo" />
                          {fatErr.nome && <div className="em">{fatErr.nome}</div>}
                        </div>
                        <div className="g2">
                          <div>
                            <label className="lbl">Email</label>
                            <input className={`inp${fatErr.email?" e":""}`} value={fatura.email}
                              onChange={e => setFatura(p=>({...p,email:e.target.value}))} placeholder="email@exemplo.com" />
                            {fatErr.email && <div className="em">{fatErr.email}</div>}
                          </div>
                          <div>
                            <label className="lbl">NIF</label>
                            <input className={`inp${fatErr.nif?" e":""}`} value={fatura.nif}
                              onChange={e => setFatura(p=>({...p,nif:e.target.value}))} placeholder="123456789" maxLength={9} />
                            {fatErr.nif && <div className="em">{fatErr.nif}</div>}
                          </div>
                        </div>
                      </div>
                    )}

                    <button className="btn-confirmar" onClick={registar} disabled={saving}>
                      {saving ? "A registar..." : `Confirmar venda — €${fmt(found.preco)}`}
                    </button>
                  </>}
                </div>
              )}

              {vendas.length > 0
                ? <div className="recentes">
                    <div className="recentes-title">Últimas vendas</div>
                    {vendas.slice(0, 12).map(v => {
                      const ar2 = artigos.find(a => a.id === v.artigo_id);
                      const ar  = artistas.find(a => a.id === v.artista_id);
                      return (
                        <div key={v.id} className="venda-row">
                          <span className="vr-code">{ar2&&ar ? mkCode(ar.codigo,ar2.codigo) : "—"}</span>
                          <span className="vr-titulo">{ar2?.titulo || "—"}</span>
                          <span className="vr-metodo">{v.metodo}</span>
                          {v.quer_fatura && <span className="fat-badge">FATURA</span>}
                          <span className="vr-preco">€{fmt(v.preco)}</span>
                          <span className="vr-hora">{fmtH(v.created_at)}</span>
                        </div>
                      );
                    })}
                  </div>
                : <div className="empty">Nenhuma venda registada ainda.</div>
              }
            </>}

            {/* ══ FATURAÇÃO ══ */}
            {tab === 2 && <>
              {fatPendente.length === 0
                ? <div className="empty">Nenhum comprador pediu fatura.</div>
                : <div className="fat-card">
                    <div className="fat-thead">
                      <div className="fat-row">
                        <span>Código</span><span>Nome</span><span>Email / NIF</span><span>Artigo</span><span style={{textAlign:"right"}}>Valor</span>
                      </div>
                    </div>
                    {fatPendente.map(v => {
                      const ar2 = artigos.find(a => a.id === v.artigo_id);
                      const ar  = artistas.find(a => a.id === v.artista_id);
                      return (
                        <div key={v.id} className="fat-row">
                          <span style={{fontSize:11,fontWeight:600,color:"#555"}}>{ar2&&ar ? mkCode(ar.codigo,ar2.codigo) : "—"}</span>
                          <span>{v.comprador_nome}</span>
                          <span style={{fontSize:11,color:"#888",lineHeight:1.7}}>{v.comprador_email}<br/>{v.comprador_nif}</span>
                          <span style={{fontSize:11}}>{ar2?.titulo || "—"}</span>
                          <span style={{textAlign:"right",fontWeight:500}}>€{fmt(v.preco)}</span>
                        </div>
                      );
                    })}
                  </div>
              }
            </>}

            {/* ══ RELATÓRIO ══ */}
            {tab === 3 && <>
              <div className="rel-card">
                <div className="rel-title">Exportar acerto final</div>
                <div className="rel-desc">
                  Gera um ficheiro Excel com um separador por artista (vendas detalhadas e montante a pagar),
                  um resumo geral, e a lista de compradores que pediram fatura.
                </div>
                <button className="btn-export" onClick={exportar}>
                  <Ic.Dl /> Descarregar relatório Excel
                </button>
              </div>

              {artistas.some(ar => vendsOf(ar.id).length > 0) && (
                <div className="rel-card">
                  <div className="rel-title" style={{marginBottom:16}}>Resumo por artista</div>
                  {artistas.map(ar => {
                    const vv = vendsOf(ar.id);
                    if (!vv.length) return null;
                    const total = vv.reduce((s,v) => s+(+v.preco||0), 0);
                    return (
                      <div key={ar.id} className="rel-row">
                        <div>
                          <span style={{fontSize:11,color:"#aaa",marginRight:10}}>{p2(ar.codigo)}</span>
                          <strong style={{fontWeight:500}}>{ar.nome_artistico || ar.nome}</strong>
                          <span style={{fontSize:12,color:"#888",marginLeft:10}}>{vv.length} {vv.length===1?"venda":"vendas"}</span>
                        </div>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:15,fontWeight:500}}>€{fmt(total*(1-COMM))}</div>
                          <div style={{fontSize:11,color:"#888"}}>de €{fmt(total)} públicos</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </>}

          </>}
        </div>
      </div>
    </>
  );
}
