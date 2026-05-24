import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
// Obscena Portal v1.1 - Malandra

/* ── Supabase ── */
const SUPABASE_URL = "https://pxosnmkydmymyzcxsnol.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB4b2FubWt5ZG15bXl6Y3hzbm9sIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk1MzQ2NjIsImV4cCI6MjA5NTExMDY2Mn0.gBL2DqH-8AB4knmoWnwU60Tfaf3TzIXNQSSYlArmb5E";

const sbFetch = async (path, opts = {}) => {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${SUPABASE_KEY}`,
      "Content-Type": "application/json",
      "Prefer": "return=representation",
      ...opts.headers
    },
    ...opts
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(err);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
};

const FONT = `@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');`;

const CATEGORIAS = ["Ilustração","Desenho","Pintura","Design Gráfico","Livro / Publicação","Fotografia","Outra"];
const COMM = 0.10;
const pad = n => String(n).padStart(2, "0");
const mkObra = () => ({ id:`${Date.now()}-${Math.random().toString(36).slice(2)}`, titulo:"", categoria:"", preco:"", quantidade:"1", descricao:"", observacoes:"", open:true });
const artistShare = p => ((parseFloat(p) || 0) * (1 - COMM)).toFixed(2);

const Ic = {
  Edit: () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
  X:    () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Check:() => <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  Plus: () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  Arrow:() => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>,
  ChUp: () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="18 15 12 9 6 15"/></svg>,
  Dl:   () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  Disk: () => <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
};

const STEPS = ["Dados do artista", "Artigos", "Rever & Enviar"];

const CSS = `
* { box-sizing: border-box; margin: 0; padding: 0; }
.portal { background: #fff; min-height: 100vh; font-family: 'Inter', sans-serif; color: #111; font-size: 14px; }

.hdr { padding: 18px 0; display: flex; align-items: center; border-bottom: 1px solid #ebebeb; }
.hdr-inner { max-width: 620px; margin: 0 auto; width: 100%; padding: 0 24px; display: flex; align-items: center; gap: 14px; }
.hdr-brand { font-size: 13px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; }
.hdr-sep { width: 1px; height: 14px; background: #ddd; }
.hdr-sub { font-size: 12px; color: #999; }

.prog { border-bottom: 1px solid #ebebeb; }
.prog-inner { max-width: 620px; margin: 0 auto; padding: 0 24px; display: flex; }
.ps { display: flex; align-items: center; gap: 8px; padding: 12px 0; margin-right: 28px; font-size: 11px; font-weight: 500; letter-spacing: .3px; text-transform: uppercase; color: #ccc; }
.ps.on { color: #111; border-bottom: 2px solid #111; margin-bottom: -1px; }
.ps.done { color: #888; }
.ps-n { width: 17px; height: 17px; border: 1.5px solid currentColor; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 9px; flex-shrink: 0; }
.ps.on .ps-n, .ps.done .ps-n { background: #111; color: #fff; border-color: #111; }

.wrap { max-width: 620px; margin: 0 auto; padding: 48px 24px; }
.stitle { font-size: 22px; font-weight: 500; letter-spacing: -.3px; margin-bottom: 6px; }
.sdesc { font-size: 13px; color: #888; line-height: 1.8; margin-bottom: 36px; }

.fld { margin-bottom: 18px; }
.lbl { display: block; font-size: 11px; font-weight: 500; color: #222; margin-bottom: 6px; }
.req { color: #bbb; font-weight: 400; }

.inp, .sel, .ta {
  width: 100%; background: #f0f0ee; border: none; outline: none;
  color: #000; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 400;
  padding: 10px 14px; border-radius: 8px; appearance: none; transition: background .15s;
}
.inp::placeholder, .ta::placeholder { color: #bbb; }
.inp:focus, .sel:focus, .ta:focus { background: #e8e8e5; }
.inp.e, .sel.e, .ta.e { background: #fce8e8; }
.ta { resize: vertical; min-height: 72px; line-height: 1.6; }
.sel { cursor: pointer; }
.em { font-size: 11px; color: #d00; margin-top: 4px; }
.g2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.g3 { display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 14px; align-items: start; }
.hint { font-size: 12px; color: #666; margin-top: 6px; white-space: nowrap; }
.hint strong { color: #111; font-weight: 500; }

.obra-list { display: flex; flex-direction: column; gap: 6px; }
.obra-item { background: #f0f0ee; border-radius: 12px; overflow: hidden; }

.obra-hdr-closed {
  display: flex; align-items: center; gap: 14px;
  padding: 13px 16px; cursor: pointer; user-select: none;
  border-radius: 12px; transition: background .12s;
}
.obra-hdr-closed:hover { background: #e8e8e5; }
.obra-hdr-open {
  display: flex; align-items: center; gap: 14px;
  padding: 13px 16px; cursor: pointer; user-select: none; transition: background .12s;
}
.obra-hdr-open:hover { background: #eaeae7; }

.or-num { font-size: 11px; color: #999; width: 22px; flex-shrink: 0; font-weight: 500; }
.or-title { font-size: 13px; font-weight: 500; flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.or-cat { font-size: 11px; color: #888; flex-shrink: 0; }
.or-price { font-size: 13px; font-weight: 500; flex-shrink: 0; min-width: 60px; text-align: right; }
.or-ic { flex-shrink: 0; color: #bbb; display: flex; transition: color .12s; }
.obra-hdr-closed:hover .or-ic { color: #666; }
.obra-open-num { font-size: 10px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; color: #888; }
.hdr-spacer { flex: 1; }
.btn-rm { background: none; border: none; color: #bbb; cursor: pointer; font-family: 'Inter', sans-serif; font-size: 11px; font-weight: 500; display: flex; align-items: center; gap: 4px; padding: 4px 8px; border-radius: 6px; transition: color .12s, background .12s; }
.btn-rm:hover { color: #d00; background: rgba(200,0,0,.07); }

.obra-expand { display: grid; grid-template-rows: 0fr; transition: grid-template-rows 0.28s cubic-bezier(0.4,0,0.2,1); }
.obra-item.is-open .obra-expand { grid-template-rows: 1fr; }
.obra-expand-inner { overflow: hidden; }
.obra-expand-content { padding: 4px 16px 16px; opacity: 0; transition: opacity 0.15s ease; }
.obra-item.is-open .obra-expand-content { opacity: 1; transition: opacity 0.22s ease 0.12s; }

.obra-expand-content .inp,
.obra-expand-content .sel,
.obra-expand-content .ta { background: #fff; }
.obra-expand-content .inp:focus,
.obra-expand-content .sel:focus,
.obra-expand-content .ta:focus { background: #f8f8f6; }

.obra-actions { display: flex; gap: 10px; margin-top: 16px; }
.btn-add { flex: 1; background: #e4e4e1; border: none; color: #333; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; padding: 10px 14px; cursor: pointer; border-radius: 8px; display: flex; align-items: center; justify-content: center; gap: 6px; transition: background .12s; }
.btn-add:hover { background: #d8d8d5; }
.btn-review { flex: 1.5; background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; padding: 10px 14px; cursor: pointer; border-radius: 8px; display: flex; align-items: center; justify-content: center; gap: 6px; transition: background .12s; }
.btn-review:hover { background: #333; }

.saved-bar { display: flex; align-items: center; gap: 6px; font-size: 11px; color: #bbb; margin-top: 14px; }

.navrow { display: flex; gap: 10px; margin-top: 40px; }
.bbk { background: #f0f0ee; border: none; color: #555; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; padding: 11px 20px; cursor: pointer; border-radius: 8px; transition: background .12s; }
.bbk:hover { background: #e4e4e1; color: #111; }
.bnxt { flex: 1; background: #111; border: none; color: #fff; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; padding: 11px 20px; cursor: pointer; border-radius: 8px; display: flex; align-items: center; justify-content: center; gap: 8px; transition: background .12s; }
.bnxt:hover { background: #333; }
.btn-dl { background: #f0f0ee; border: none; color: #555; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 500; padding: 11px 16px; cursor: pointer; border-radius: 8px; display: flex; align-items: center; gap: 6px; transition: background .12s; white-space: nowrap; }
.btn-dl:hover { background: #e4e4e1; color: #111; }

.rev-sec { margin-bottom: 32px; }
.rev-sec-title { font-size: 10px; font-weight: 600; letter-spacing: 1.5px; text-transform: uppercase; color: #888; margin-bottom: 12px; }
.rev-grid { background: #f6f6f4; border-radius: 12px; padding: 4px 18px; margin-bottom: 12px; }
.rev-f { display: flex; justify-content: space-between; align-items: baseline; padding: 11px 0; }
.rev-f + .rev-f { border-top: 1px solid #ebebeb; }
.rev-lbl { font-size: 12px; color: #888; }
.rev-val { font-size: 13px; color: #111; }
.rev-totals { display: flex; background: #f0f0ee; border-radius: 12px; margin-bottom: 16px; overflow: hidden; }
.rev-tot { flex: 1; padding: 16px 18px; }
.rev-tot + .rev-tot { border-left: 1px solid #e4e4e1; }
.rev-tot-lbl { font-size: 10px; font-weight: 600; letter-spacing: .5px; text-transform: uppercase; color: #999; margin-bottom: 5px; }
.rev-tot-val { font-size: 18px; font-weight: 500; }
.rev-obra { background: #f6f6f4; padding: 14px 16px; margin-bottom: 8px; border-radius: 10px; }
.rev-obra-top { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 4px; }
.rev-obra-title { font-size: 14px; font-weight: 500; }
.rev-obra-price { font-size: 13px; color: #666; flex-shrink: 0; }
.rev-obra-share { font-size: 11px; color: #888; margin-bottom: 4px; }
.rev-obra-meta { font-size: 11px; color: #999; margin-bottom: 6px; }
.rev-obra-desc { font-size: 13px; color: #555; line-height: 1.7; }
.rev-obra-obs { font-size: 12px; color: #888; margin-top: 5px; font-style: italic; }

.agree { display: flex; align-items: flex-start; gap: 12px; padding: 14px 16px; background: #f0f0ee; border-radius: 10px; cursor: pointer; margin: 20px 0 6px; transition: background .12s; }
.agree:hover { background: #e8e8e5; }
.abox { width: 15px; height: 15px; background: #d4d4d1; flex-shrink: 0; margin-top: 2px; border-radius: 4px; display: flex; align-items: center; justify-content: center; transition: all .15s; }
.abox.chk { background: #111; color: #fff; }
.atxt { font-size: 12px; color: #555; line-height: 1.8; }
.em-agree { font-size: 11px; color: #d00; margin-bottom: 12px; }

.ok { display: flex; flex-direction: column; align-items: center; text-align: center; padding: 72px 24px; }
.ok-icon { width: 44px; height: 44px; background: #f0f0ee; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 24px; }
.ok-title { font-size: 22px; font-weight: 500; margin-bottom: 10px; }
.ok-msg { font-size: 13px; color: #888; max-width: 340px; line-height: 1.9; margin-bottom: 24px; }
.ok-sum { background: #f6f6f4; padding: 4px 20px; border-radius: 12px; width: 100%; max-width: 360px; margin-bottom: 16px; }
.ok-row { display: flex; justify-content: space-between; padding: 11px 0; font-size: 13px; }
.ok-row + .ok-row { border-top: 1px solid #ebebeb; }
.ok-row-lbl { color: #888; }
.ok-actions { display: flex; gap: 10px; flex-wrap: wrap; justify-content: center; }
.btn-reset { background: none; border: none; color: #bbb; font-family: 'Inter', sans-serif; font-size: 11px; padding: 8px 12px; cursor: pointer; border-radius: 6px; transition: color .12s; }
.btn-reset:hover { color: #666; }

.btn-next-circle { background: none; border: none; display: flex; align-items: center; gap: 12px; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 500; color: #111; cursor: pointer; padding: 0; transition: gap .2s; }
.btn-next-circle:hover { gap: 16px; }
.btn-circle { width: 36px; height: 36px; background: #111; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #fff; flex-shrink: 0; transition: background .15s; }
.btn-next-circle:hover .btn-circle { background: #333; }

.or-ic-del { flex-shrink: 0; color: #ccc; display: flex; padding: 4px; border-radius: 4px; transition: color .12s, background .12s; cursor: pointer; }
.or-ic-del:hover { color: #d00; background: rgba(200,0,0,.07); }

@media (max-width: 560px) {
  .wrap { padding: 28px 14px; }
  .g2, .g3 { grid-template-columns: 1fr; }
  .hdr, .prog-inner { padding-left: 14px; padding-right: 14px; }
  .or-cat { display: none; }
  .obra-actions { flex-direction: column; }
  .rev-totals { flex-direction: column; }
  .navrow { flex-wrap: wrap; }
}
`;

export default function ObscenaPortal() {
  const [step, setStep]           = useState(0);
  const [submitted, setSubmitted] = useState(false);
  const [agreed, setAgreed]       = useState(false);
  const [errors, setErrors]       = useState({});
  const [ready, setReady]         = useState(false);
  const [lastSaved, setLastSaved] = useState(null);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState(null);

  const [artista, setArtista] = useState({
    nome:"", nomeArtistico:"", email:"", contacto:"", nif:"", iban:""
  });
  const [obras, setObras] = useState([mkObra()]);

  /* ── Load ── */
  useEffect(() => {
    (async () => {
      try {
        const a   = localStorage.getItem("obs_a");
        const o   = localStorage.getItem("obs_o");
        const s   = localStorage.getItem("obs_s");
        const sub = localStorage.getItem("obs_done");
        if (a) setArtista(JSON.parse(a));
        if (o) { const p = JSON.parse(o); if (p && p.length) setObras(p); }
        if (s) setStep(parseInt(s));
        if (sub === "1") setSubmitted(true);
      } catch {}
      setReady(true);
    })();
  }, []);

  /* ── Auto-save (debounced 1s) ── */
  useEffect(() => {
    if (!ready) return;
    const t = setTimeout(async () => {
      try {
        localStorage.setItem("obs_a", JSON.stringify(artista));
        localStorage.setItem("obs_o", JSON.stringify(obras));
        localStorage.setItem("obs_s", String(step));
        setLastSaved(new Date());
      } catch (e) { console.error(e); }
    }, 1000);
    return () => clearTimeout(t);
  }, [artista, obras, step, ready]);

  /* ── Helpers ── */
  const setA = (k, v) => setArtista(p => ({ ...p, [k]: v }));
  const setO = (id, k, v) => setObras(p => p.map(o => o.id === id ? { ...o, [k]: v } : o));

  const openObra     = id => setObras(p => p.map(o => ({ ...o, open: o.id === id })));
  const collapseObra = id => setObras(p => p.map(o => o.id === id ? { ...o, open: false } : o));

  const removeObra = id => setObras(p => {
    const next = p.filter(o => o.id !== id);
    if (!next.length) return [mkObra()];
    if (!next.some(o => o.open))
      return next.map((o, i) => i === next.length - 1 ? { ...o, open: true } : o);
    return next;
  });

  /* ── Validation ── */
  const validateArtista = () => {
    const e = {};
    if (!artista.nome.trim())                          e.nome     = "Campo obrigatório";
    if (!/\S+@\S+\.\S+/.test(artista.email))           e.email    = "Email inválido";
    if (!artista.contacto.trim())                      e.contacto = "Campo obrigatório";
    if (!/^\d{9}$/.test(artista.nif.replace(/\s/g,""))) e.nif    = "Deve ter 9 dígitos";
    if (!artista.iban.trim())                          e.iban     = "Campo obrigatório";
    setErrors(e); return !Object.keys(e).length;
  };

  const validateOpenObra = () => {
    const open = obras.find(o => o.open);
    if (!open) return true;
    const e = {}, id = open.id;
    if (!open.titulo.trim())                          e[`${id}_t`] = "Obrigatório";
    if (!open.categoria)                              e[`${id}_c`] = "Obrigatório";
    if (!open.preco || +open.preco <= 0)              e[`${id}_p`] = "Preço inválido";
    if (!open.quantidade || +open.quantidade < 1)     e[`${id}_q`] = "Mín. 1";
    if (!open.descricao.trim())                       e[`${id}_d`] = "Obrigatório";
    setErrors(e); return !Object.keys(e).length;
  };

  /* ── Actions ── */
  const handleAddObra = () => {
    if (!validateOpenObra()) return;
    setErrors({});
    setObras(p => [...p.map(o => ({ ...o, open: false })), mkObra()]);
  };

  const handleGoReview = () => {
    if (!validateOpenObra()) return;
    setErrors({});
    setObras(p => p.map(o => ({ ...o, open: false })));
    setStep(2);
  };

  const handleNext0 = () => {
    if (!validateArtista()) return;
    setErrors({}); setStep(1);
  };

  const handleBackToObras = () => {
    setErrors({});
    setObras(p => p.map((o, i) => ({ ...o, open: i === p.length - 1 })));
    setStep(1);
  };

  const handleSubmit = async () => {
    if (!agreed) { setErrors({ agree: "Por favor confirma as informações" }); return; }
    setSubmitting(true);
    setSubmitError(null);
    try {
      // 1. Insert artista
      const token = `${Date.now()}-${Math.random().toString(36).slice(2,10)}`;
      const artistaRows = await sbFetch("artistas", {
        method: "POST",
        body: JSON.stringify({
          nome: artista.nome,
          nome_artistico: artista.nomeArtistico || null,
          email: artista.email,
          contacto: artista.contacto,
          nif: artista.nif,
          iban: artista.iban,
          token
        })
      });
      const artistaId = artistaRows[0].id;

      // 2. Insert artigos
      const artigosPayload = obras.map(o => ({
        artista_id: artistaId,
        titulo: o.titulo,
        categoria: o.categoria,
        preco: parseFloat(o.preco),
        quantidade: parseInt(o.quantidade),
        descricao: o.descricao,
        observacoes: o.observacoes || null
      }));
      await sbFetch("artigos", {
        method: "POST",
        body: JSON.stringify(artigosPayload)
      });

      // 3. Mark as done locally
      try {
        localStorage.setItem("obs_done", "1");
        localStorage.setItem("obs_sub", JSON.stringify({ artista, obras, ts: new Date().toISOString() }));
      } catch {}

      setSubmitted(true);
    } catch (e) {
      setSubmitError("Ocorreu um erro ao enviar o inventário. Verifica a tua ligação e tenta novamente.");
      console.error(e);
    } finally {
      setSubmitting(false);
    }
  };

  const handleReset = async () => {
    try {
      for (const k of ["obs_a","obs_o","obs_s","obs_done","obs_sub"])
        localStorage.removeItem(k);
    } catch {}
    setArtista({ nome:"", nomeArtistico:"", email:"", contacto:"", nif:"", iban:"" });
    setObras([mkObra()]); setStep(0); setSubmitted(false); setAgreed(false); setErrors({});
  };

  /* ── Excel download ── */
  const downloadExcel = () => {
    const rows = [
      ["DADOS DO ARTISTA", ""],
      ["Nome", artista.nome],
      ["Nome artístico", artista.nomeArtistico || "—"],
      ["Email", artista.email],
      ["Contacto", artista.contacto],
      ["NIF", artista.nif],
      ["IBAN", artista.iban],
      ["", ""],
      ["INVENTÁRIO — OBSCENA"],
      ["Nº","Título","Categoria","Preço público (€)",`Valor artista ${(1-COMM)*100}% (€)`,"Quantidade","Total artista (€)","Descrição","Observações"],
      ...obras.map((o,i) => [
        i+1, o.titulo, o.categoria,
        parseFloat(o.preco)||0,
        parseFloat(artistShare(o.preco)),
        parseInt(o.quantidade)||1,
        (parseFloat(artistShare(o.preco))*(parseInt(o.quantidade)||1)).toFixed(2),
        o.descricao, o.observacoes||""
      ]),
      [""],
      ["Nº de obras", obras.length],
      ["Total de peças", obras.reduce((s,o)=>s+(parseInt(o.quantidade)||0),0)],
      ["Valor total público (€)", totalValorPub.toFixed(2)],
      ["Total para artista (€)", totalParaArtista.toFixed(2)],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{wch:20},{wch:32},{wch:20},{wch:18},{wch:22},{wch:12},{wch:18},{wch:40},{wch:30}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventário");
    const slug = (artista.nomeArtistico||artista.nome||"artista").replace(/\s+/g,"_").toLowerCase();
    XLSX.writeFile(wb, `inventario_${slug}_obscena.xlsx`);
  };

  /* ── Totals ── */
  const totalPecas      = obras.reduce((s,o)=>s+(+o.quantidade||0),0);
  const totalValorPub   = obras.reduce((s,o)=>s+((+o.preco||0)*(+o.quantidade||0)),0);
  const totalParaArtista= obras.reduce((s,o)=>s+(parseFloat(artistShare(o.preco))*(+o.quantidade||0)),0);

  const fmtTime = d => d ? `${pad(d.getHours())}:${pad(d.getMinutes())}` : "";

  if (!ready) return (
    <div style={{background:"#fff",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"sans-serif",color:"#ccc",fontSize:"11px",letterSpacing:"2px"}}>
      A CARREGAR
    </div>
  );

  return (
    <>
      <style>{FONT}{CSS}</style>
      <div className="portal">

        {/* ── Header ── */}
        <div className="hdr">
          <div className="hdr-inner">
            <span className="hdr-brand">Obscena</span>
            <div className="hdr-sep" />
            <span className="hdr-sub">Portal do artista · Malandra</span>
          </div>
        </div>

        {submitted ? (
          /* ══ CONFIRMAÇÃO ══ */
          <div className="ok">
            <div className="ok-icon"><Ic.Check /></div>
            <div className="ok-title">Inventário enviado</div>
            <p className="ok-msg">
              Obrigado, {artista.nomeArtistico || artista.nome}. A equipa da Malandra vai rever o teu inventário e entrar em contacto em breve.
            </p>
            <div className="ok-sum">
              {[
                ["Artista",        artista.nomeArtistico || artista.nome],
                ["Email",          artista.email],
                ["Artigos",        `${obras.length}`],
                ["Peças",          `${totalPecas}`],
              ].map(([l,v]) => (
                <div key={l} className="ok-row">
                  <span className="ok-row-lbl">{l}</span>
                  <span>{v}</span>
                </div>
              ))}
            </div>
            <div className="ok-actions">
              <button className="btn-dl" onClick={downloadExcel}>
                <Ic.Dl /> Descarregar inventário Excel
              </button>
              <button className="btn-reset" onClick={handleReset}>Recomeçar do início</button>
            </div>
          </div>
        ) : (
          <>
            {/* ── Progress ── */}
            <div className="prog">
              <div className="prog-inner">
              {STEPS.map((s,i) => (
                <div key={i} className={`ps ${step===i?"on":""} ${step>i?"done":""}`}>
                  <span className="ps-n">{step>i ? <Ic.Check /> : i+1}</span>
                  {s}
                </div>
              ))}
              </div>
            </div>

            <div className="wrap">

              {/* ══ STEP 0: Artista ══ */}
              {step === 0 && <>
                <div className="stitle">Os teus dados</div>
                <p className="sdesc">
                  Informações de contacto e dados necessários para processarmos o teu pagamento após a feira.
                </p>

                <div className="g2">
                  <div className="fld">
                    <label className="lbl">Nome completo <span className="req">*</span></label>
                    <input className={`inp${errors.nome?" e":""}`} value={artista.nome} onChange={e=>setA("nome",e.target.value)} placeholder="Ana Silva" />
                    {errors.nome && <div className="em">{errors.nome}</div>}
                  </div>
                  <div className="fld">
                    <label className="lbl">Nome artístico</label>
                    <input className="inp" value={artista.nomeArtistico} onChange={e=>setA("nomeArtistico",e.target.value)} placeholder="Opcional" />
                  </div>
                </div>
                <div className="g2">
                  <div className="fld">
                    <label className="lbl">Email <span className="req">*</span></label>
                    <input className={`inp${errors.email?" e":""}`} type="email" value={artista.email} onChange={e=>setA("email",e.target.value)} placeholder="ana@email.com" />
                    {errors.email && <div className="em">{errors.email}</div>}
                  </div>
                  <div className="fld">
                    <label className="lbl">Contacto <span className="req">*</span></label>
                    <input className={`inp${errors.contacto?" e":""}`} value={artista.contacto} onChange={e=>setA("contacto",e.target.value)} placeholder="+351 912 345 678" />
                    {errors.contacto && <div className="em">{errors.contacto}</div>}
                  </div>
                </div>
                <div className="g2">
                  <div className="fld">
                    <label className="lbl">NIF <span className="req">*</span></label>
                    <input className={`inp${errors.nif?" e":""}`} value={artista.nif} onChange={e=>setA("nif",e.target.value)} placeholder="123456789" maxLength={9} />
                    {errors.nif && <div className="em">{errors.nif}</div>}
                  </div>
                  <div className="fld">
                    <label className="lbl">IBAN <span className="req">*</span></label>
                    <input className={`inp${errors.iban?" e":""}`} value={artista.iban} onChange={e=>setA("iban",e.target.value)} placeholder="PT50 0000 0000 0000 0000 0000 0" />
                    {errors.iban && <div className="em">{errors.iban}</div>}
                  </div>
                </div>

                <div className="navrow" style={{justifyContent:"flex-end"}}>
                  <button className="btn-next-circle" onClick={handleNext0}>
                    Continuar
                    <span className="btn-circle"><Ic.Arrow /></span>
                  </button>
                </div>
              </>}

              {/* ══ STEP 1: Obras ══ */}
              {step === 1 && <>
                <div className="stitle">Os teus artigos</div>
                <p className="sdesc">
                  Adiciona todos os artigos que vais entregar. Quantidade 1 para peça única, ou o número de exemplares disponíveis para prints e publicações. O teu rascunho é guardado automaticamente.
                </p>

                <div className="obra-list">
                {obras.map((obra, idx) => (
                  <div key={obra.id} className={`obra-item${obra.open?" is-open":""}`}>

                    {/* Header — changes based on open/closed */}
                    {obra.open ? (
                      <div className="obra-hdr-open" onClick={() => collapseObra(obra.id)}>
                        <span className="obra-open-num">Artigo {pad(idx+1)}</span>
                        <span className="hdr-spacer" />
                        {obras.length > 1 && (
                          <button className="btn-rm" onClick={e=>{e.stopPropagation();removeObra(obra.id);}}>
                            <Ic.X /> Remover artigo
                          </button>
                        )}
                        <span className="or-ic" style={{marginLeft:4}}><Ic.ChUp /></span>
                      </div>
                    ) : (
                      <div className="obra-hdr-closed" onClick={()=>openObra(obra.id)}>
                        <span className="or-num">{pad(idx+1)}</span>
                        <span className="or-title">{obra.titulo||"Novo artigo"}</span>
                        <span className="or-cat">{obra.categoria}</span>
                        <span className="or-price">€{parseFloat(obra.preco||0).toFixed(2)}</span>
                        <span className="or-ic"><Ic.Edit /></span>
                        {obras.length > 1 && (
                          <span className="or-ic-del" onClick={e=>{e.stopPropagation();removeObra(obra.id);}}
                            title="Eliminar obra">
                            <Ic.X />
                          </span>
                        )}
                      </div>
                    )}

                    {/* Animated body */}
                    <div className="obra-expand">
                      <div className="obra-expand-inner">
                        <div className="obra-expand-content">

                          <div className="fld">
                            <label className="lbl">Título / Identificação <span className="req">*</span></label>
                            <input className={`inp${errors[`${obra.id}_t`]?" e":""}`} value={obra.titulo} onChange={e=>setO(obra.id,"titulo",e.target.value)} placeholder="" />
                            {errors[`${obra.id}_t`] && <div className="em">{errors[`${obra.id}_t`]}</div>}
                          </div>

                          <div className="g3">
                            <div className="fld" style={{marginBottom:0}}>
                              <label className="lbl">Categoria <span className="req">*</span></label>
                              <select className={`sel${errors[`${obra.id}_c`]?" e":""}`} value={obra.categoria} onChange={e=>setO(obra.id,"categoria",e.target.value)}>
                                <option value="">Selecionar...</option>
                                {CATEGORIAS.map(c=><option key={c} value={c}>{c}</option>)}
                              </select>
                              <div style={{minHeight:18}}>{errors[`${obra.id}_c`] && <div className="em">{errors[`${obra.id}_c`]}</div>}</div>
                            </div>
                            <div className="fld" style={{marginBottom:0}}>
                              <label className="lbl">Preço público (€) <span className="req">*</span></label>
                              <input className={`inp${errors[`${obra.id}_p`]?" e":""}`} type="number" min="0" step="0.5" value={obra.preco} onChange={e=>setO(obra.id,"preco",e.target.value)} placeholder="25.00" />
                              <div className="hint" style={{minHeight:18}}>
                                Recebes <strong>€{artistShare(obra.preco || 0)}</strong> por peça
                              </div>
                              {errors[`${obra.id}_p`] && <div className="em">{errors[`${obra.id}_p`]}</div>}
                            </div>
                            <div className="fld" style={{marginBottom:0}}>
                              <label className="lbl">Quantidade <span className="req">*</span></label>
                              <input className={`inp${errors[`${obra.id}_q`]?" e":""}`} type="number" min="1" step="1" value={obra.quantidade} onChange={e=>setO(obra.id,"quantidade",e.target.value)} placeholder="1" />
                              <div style={{minHeight:18}}>{errors[`${obra.id}_q`] && <div className="em">{errors[`${obra.id}_q`]}</div>}</div>
                            </div>
                          </div>

                          <div className="fld">
                            <label className="lbl">Descrição <span className="req">*</span></label>
                            <textarea className={`ta${errors[`${obra.id}_d`]?" e":""}`} value={obra.descricao} onChange={e=>setO(obra.id,"descricao",e.target.value)} placeholder="Técnica, materiais, dimensões, contexto..." />
                            {errors[`${obra.id}_d`] && <div className="em">{errors[`${obra.id}_d`]}</div>}
                          </div>

                          <div className="fld">
                            <label className="lbl">Observações</label>
                            <textarea className="ta" style={{minHeight:52}} value={obra.observacoes} onChange={e=>setO(obra.id,"observacoes",e.target.value)} placeholder="Notas adicionais para a organização (frágil, não empilhar, etc.)" />
                          </div>

                        </div>
                      </div>
                    </div>
                  </div>
                ))}
                </div>

                {/* Action buttons — outside the article containers */}
                <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:16, paddingTop:20, borderTop:"1px solid #ebebeb"}}>
                  <button className="btn-next-circle" style={{fontSize:12}} onClick={handleAddObra}>
                    <span className="btn-circle" style={{width:30, height:30, background:"#e4e4e1"}}><span style={{color:"#111"}}><Ic.Plus /></span></span>
                    Adicionar outro artigo
                  </button>
                  <button className="btn-next-circle" style={{fontSize:12}} onClick={handleGoReview}>
                    Rever e submeter inventário
                    <span className="btn-circle"><Ic.Arrow /></span>
                  </button>
                </div>

                {/* Saved indicator */}
                {lastSaved && (
                  <div className="saved-bar">
                    <Ic.Disk /> Rascunho guardado às {fmtTime(lastSaved)} — podes fechar e voltar mais tarde
                  </div>
                )}

                <div className="navrow" style={{marginTop:20}}>
                  <button className="bbk" onClick={()=>{setErrors({});setStep(0);}}>← Voltar</button>
                </div>
              </>}

              {/* ══ STEP 2: Rever & Enviar ══ */}
              {step === 2 && <>
                <div className="stitle">Rever & Enviar</div>
                <p className="sdesc">Confirma que tudo está correcto antes de enviares o teu inventário à Malandra.</p>

                {/* Artista */}
                <div className="rev-sec">
                  <div className="rev-sec-title">Dados do artista</div>
                  <div className="rev-grid">
                    {[
                      ["Nome completo",  artista.nome],
                      ["Nome artístico", artista.nomeArtistico||"—"],
                      ["Email",          artista.email],
                      ["Contacto",       artista.contacto],
                      ["NIF",            artista.nif],
                      ["IBAN",           artista.iban],
                    ].map(([l,v])=>(
                      <div key={l} className="rev-f">
                        <span className="rev-lbl">{l}</span>
                        <span className="rev-val">{v}</span>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Artigos */}
                <div className="rev-sec">
                  <div className="rev-sec-title">Artigos</div>
                  <div style={{display:"flex", gap:16, padding:"12px 0", marginBottom:16, borderBottom:"1px solid #ebebeb"}}>
                    <span style={{fontSize:13, color:"#888"}}><strong style={{color:"#111", fontWeight:500}}>{obras.length}</strong> {obras.length===1?"artigo":"artigos"}</span>
                    <span style={{color:"#ddd"}}>·</span>
                    <span style={{fontSize:13, color:"#888"}}><strong style={{color:"#111", fontWeight:500}}>{totalPecas}</strong> peças</span>
                  </div>
                  {obras.map(o=>(
                    <div key={o.id} className="rev-obra">
                      <div className="rev-obra-top">
                        <span className="rev-obra-title">{o.titulo}</span>
                        <span className="rev-obra-price">€{parseFloat(o.preco).toFixed(2)} × {o.quantidade}</span>
                      </div>
                      <div className="rev-obra-share">Recebes €{artistShare(o.preco)} por peça</div>
                      <div className="rev-obra-meta">{o.categoria}</div>
                      <div className="rev-obra-desc">{o.descricao}</div>
                      {o.observacoes && <div className="rev-obra-obs">Obs: {o.observacoes}</div>}
                    </div>
                  ))}
                </div>

                {/* Agreement */}
                <div className="agree" onClick={()=>{setAgreed(p=>!p);setErrors({});}}>
                  <div className={`abox${agreed?" chk":""}`}>{agreed&&<Ic.Check/>}</div>
                  <div className="atxt">
                    Confirmo que as informações são correctas e autorizo a Malandra a vender as artigos listados durante a Obscena, nas condições acordadas (comissão de {COMM*100}% sobre o preço de venda).
                  </div>
                </div>
                {errors.agree && <div className="em-agree">{errors.agree}</div>}
                {submitError && <div className="em-agree" style={{marginTop:8}}>{submitError}</div>}

                <div className="navrow">
                  <button className="bbk" onClick={handleBackToObras} disabled={submitting}>← Voltar</button>
                  <button className="btn-dl" onClick={downloadExcel} disabled={submitting}><Ic.Dl /> Excel</button>
                  <button className="btn-next-circle" onClick={handleSubmit} disabled={submitting} style={{opacity: submitting ? .5 : 1}}>
                    {submitting ? "A enviar..." : "Enviar inventário"}
                    <span className="btn-circle"><Ic.Arrow /></span>
                  </button>
                </div>
              </>}

            </div>
          </>
        )}
      </div>
    </>
  );
}
