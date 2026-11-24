/* CTPL · llave del dispositivo
   La llave se guarda en IndexedDB como CryptoKey NO exportable (ni el funcionario puede copiarla).
   Compatible con Safari/iPhone: abre una conexión nueva en cada operación, reintenta si Safari
   cerró la base ("database connection is closing") y, si IndexedDB no está disponible, usa un respaldo. */
const LLAVERO = {
  MEM: {},                                   // copia en memoria durante la sesión
  _sinIDB: false,
  abrir(){
    return new Promise((ok, mal) => {
      let r;
      try { r = indexedDB.open('ctpl_asistencia', 1); } catch (e) { return mal(e); }
      r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains('k')) r.result.createObjectStore('k'); };
      r.onsuccess = () => { const db = r.result; db.onversionchange = () => { try { db.close(); } catch (e) {} }; ok(db); };
      r.onerror = () => mal(r.error || new Error('IndexedDB no disponible'));
      r.onblocked = () => mal(new Error('IndexedDB bloqueada'));
    });
  },
  async tx(modo, fn){
    let ultimoError;
    for (let intento = 0; intento < 3; intento++) {
      let db = null;
      try {
        db = await this.abrir();
        return await new Promise((ok, mal) => {
          let res, t;
          try { t = db.transaction('k', modo); } catch (e) { return mal(e); }
          const req = fn(t.objectStore('k'));
          req.onsuccess = () => { res = req.result; };
          t.oncomplete = () => ok(res);
          t.onerror = () => mal(t.error || req.error);
          t.onabort = () => mal(t.error || new Error('Transacción cancelada'));
        });
      } catch (e) {
        ultimoError = e;
        if (e && e.name === 'DataCloneError') throw e;           // no se puede guardar ese tipo de dato
        await new Promise(r => setTimeout(r, 150 * (intento + 1)));
      } finally {
        try { if (db) db.close(); } catch (e) {}
      }
    }
    this._sinIDB = true;
    throw ultimoError;
  },
  /* Respaldo en localStorage (solo si IndexedDB falla por completo) */
  _ls(k, v, borrar){
    const n = 'ctpl_' + k;
    try {
      if (borrar) { localStorage.removeItem(n); return; }
      if (v === undefined) { const x = localStorage.getItem(n); return x == null ? undefined : JSON.parse(x); }
      localStorage.setItem(n, JSON.stringify(v));
    } catch (e) {}
  },
  async leer(k){
    if (k in this.MEM) return this.MEM[k];
    let v;
    try { v = await this.tx('readonly', s => s.get(k)); } catch (e) { v = undefined; }
    if (v === undefined) v = this._ls(k);
    if (v !== undefined) this.MEM[k] = v;
    return v;
  },
  async poner(k, v){
    this.MEM[k] = v;
    try { await this.tx('readwrite', s => s.put(v, k)); this._ls(k, null, true); }
    catch (e) {
      if (typeof CryptoKey !== 'undefined' && v instanceof CryptoKey) throw e;   // la llave no se guarda fuera de IndexedDB
      this._ls(k, v);
    }
  },
  async quitar(k){
    delete this.MEM[k];
    this._ls(k, null, true);
    try { await this.tx('readwrite', s => s.delete(k)); } catch (e) {}
  },
  aleatorio(n){ const a = new Uint8Array(n); crypto.getRandomValues(a); return Array.from(a, b => b.toString(16).padStart(2,'0')).join('').toUpperCase(); },
  async idDispositivo(pref){
    const k = (pref||'') + 'dev';
    let d = await this.leer(k).catch(() => null);
    if (!d) { d = (pref ? 'K' : 'D') + this.aleatorio(10); await this.poner(k, d); }
    return d;
  },
  async guardarLlave(b64, pref){
    const k = (pref||'') + 'llave';
    const raw = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    const key = await crypto.subtle.importKey('raw', raw, {name:'HMAC', hash:'SHA-256'}, false, ['sign']);
    try { await this.poner(k, key); }
    catch (e) { await this.poner(k, 'raw:' + b64); }   // navegadores que no permiten guardar CryptoKey
    this.MEM[k] = key;
  },
  async tieneLlave(pref){ return !!(await this.leer((pref||'') + 'llave').catch(() => null)); },
  async firmar(msg, pref){
    const k = (pref||'') + 'llave';
    let key = await this.leer(k);
    if (!key) return '';
    if (typeof key === 'string' && key.indexOf('raw:') === 0) {
      const raw = Uint8Array.from(atob(key.slice(4)), c => c.charCodeAt(0));
      key = await crypto.subtle.importKey('raw', raw, {name:'HMAC', hash:'SHA-256'}, false, ['sign']);
      this.MEM[k] = key;
    }
    const s = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(msg));
    let b = ''; new Uint8Array(s).forEach(x => b += String.fromCharCode(x));
    return btoa(b);
  },
  descripcion(){
    const u = navigator.userAgent;
    const so = /iPhone|iPad/.test(u) ? (/iPad/.test(u) ? 'iPad' : 'iPhone') : /Android/.test(u) ? ((u.match(/Android[^;]*;\s*([^;)]+)/)||[])[1] || 'Android') : /Windows/.test(u) ? 'Windows' : /Mac/.test(u) ? 'Mac' : /Linux/.test(u) ? 'Linux' : 'Otro';
    const nav = /EdgA?\//.test(u) ? 'Edge' : /SamsungBrowser/.test(u) ? 'Samsung Internet' : /CriOS|Chrome/.test(u) ? 'Chrome' : /FxiOS|Firefox/.test(u) ? 'Firefox' : /Safari/.test(u) ? 'Safari' : 'Navegador';
    const app = (window.matchMedia && matchMedia('(display-mode: standalone)').matches) || navigator.standalone ? ' · App' : '';
    return (so + ' · ' + nav + app + ' · ' + screen.width + 'x' + screen.height).replace(/\s+/g,' ').trim();
  }
};
