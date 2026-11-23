/* CTPL · utilidades comunes */
/* Íconos SVG (trazo) */
  const ICO = {
    qr:'<path d="M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3z"/><path d="M14 14h3v3h-3zM20 14v.01M14 20h.01M17 20h4v-3M20 17v.01"/>',
    casa:'<path d="M3 11l9-8 9 8"/><path d="M5 10v10h14V10"/>',
    carnet:'<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="11" r="2.5"/><path d="M5.5 17c.8-1.8 2-2.7 3.5-2.7s2.7.9 3.5 2.7M15 9h3M15 13h3"/>',
    reloj:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    informe:'<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5M9 13h6M9 17h6"/>',
    user:'<circle cx="12" cy="8" r="4"/><path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6"/>',
    check:'<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    x:'<path d="M6 6l12 12M18 6L6 18"/>',
    descarga:'<path d="M12 3v12M7 10l5 5 5-5M4 20h16"/>',
    camara:'<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
    salir:'<path d="M15 4h4v16h-4M10 8l-4 4 4 4M6 12h10"/>',
    escudo:'<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"/><path d="M8.5 12l2.5 2.5 4.5-5"/>',
    movil:'<rect x="7" y="2.5" width="10" height="19" rx="2"/><path d="M11 18.5h2"/>',
    izq:'<path d="M15 5l-7 7 7 7"/>', der:'<path d="M9 5l7 7-7 7"/>',
    pantalla:'<rect x="3" y="4" width="18" height="12" rx="1.5"/><path d="M8 20h8M12 16v4"/>',
    engrane:'<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
    grupo:'<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c1-3.5 3.5-5 6.5-5s5.5 1.5 6.5 5"/><circle cx="17" cy="9" r="2.5"/><path d="M16.5 14c2.5 0 4.2 1.4 5 4"/>',
    bandeja:'<path d="M4 13l2.5-8h11L20 13v6H4z"/><path d="M4 13h5l1 2h4l1-2h5"/>',
    lista:'<path d="M8 6h13M8 12h13M8 18h13M3.5 6h.01M3.5 12h.01M3.5 18h.01"/>',
    llave:'<circle cx="8" cy="15" r="4"/><path d="M11 12l9-9M17 6l3 3M15 8l2 2"/>',
    mas:'<path d="M12 5v14M5 12h14"/>', recargar:'<path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7"/>',
    expandir:'<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>'
  };
  function ico(n, c){ return '<svg class="ico '+(c||'')+'" viewBox="0 0 24 24">'+(ICO[n]||'')+'</svg>'; }
  function $(id){ return document.getElementById(id); }
  function escH(s){ return String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function toast(t, ms){ const e=$('toast'); if(!e) return; e.textContent=t; e.style.display='block'; clearTimeout(e._t); e._t=setTimeout(()=>e.style.display='none', ms||3200); }
  function cargando(v, t){ const e=$('cargando'); if(!e) return; e.style.display=v?'flex':'none'; e.innerHTML='<span class="spin osc"></span> '+escH(t||'Cargando…'); }
  function abrirM(id){ $(id).classList.add('ver'); } function cerrarM(id){ $(id).classList.remove('ver'); }
  /* Llama al servidor (Apps Script) por la API JSON. Reintenta una vez si Google responde con error temporal. */
  async function llamar(fn, ...args){
    for (let intento = 0; intento < 2; intento++) {
      let res, j = null;
      try {
        res = await fetch(CTPL_API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ fn: fn, args: args }) });
        j = await res.json();
      } catch (e) {
        if (intento === 0) { await new Promise(r => setTimeout(r, 1500)); continue; }
        throw new Error(navigator.onLine === false ? 'Sin conexión a internet.' : 'El servidor no respondió. Intente de nuevo en un momento.');
      }
      if (j.ok) return j.r;
      if (intento === 0 && /DEADLINE_EXCEEDED|server error|INTERNAL/i.test(j.error || '')) { await new Promise(r => setTimeout(r, 1500)); continue; }
      throw new Error(j.error || 'Error del servidor');
    }
  }
  function descargarB64(r){
    const bin = atob(r.b64), u = new Uint8Array(bin.length);
    for (let i=0;i<bin.length;i++) u[i] = bin.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([u], {type:r.mime}));
    const a = document.createElement('a'); a.href = url; a.download = r.nombre; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
  }
  function mayus(t){ t = String(t||''); return t.charAt(0).toUpperCase() + t.slice(1); }
  const DIAS = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  function fechaTxt(f, corto){ const d = new Date(f+'T12:00:00'); return (corto ? DIAS[d.getDay()].slice(0,3) : DIAS[d.getDay()]) + ' ' + d.getDate() + ' ' + (corto ? MESES[d.getMonth()].slice(0,3) : 'de ' + MESES[d.getMonth()]); }
  function fISO(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
  function sumarDias(f, n){ const d = new Date(f+'T12:00:00'); d.setDate(d.getDate()+n); return fISO(d); }
  function lunesDe(f){ const d = new Date(f+'T12:00:00'); const w = d.getDay()||7; d.setDate(d.getDate()-w+1); return fISO(d); }
