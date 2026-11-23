/* CTPL · llave del dispositivo */
/* Llave del dispositivo: se guarda en IndexedDB como CryptoKey NO exportable.
     Ni el propio funcionario puede copiarla a otro teléfono. */
  const LLAVERO = {
    _db: null,
    db(){
      if (this._db) return Promise.resolve(this._db);
      return new Promise((ok, mal) => {
        const r = indexedDB.open('ctpl_asistencia', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('k');
        r.onsuccess = () => { this._db = r.result; ok(r.result); };
        r.onerror = () => mal(r.error);
      });
    },
    async leer(k){ const db = await this.db(); return new Promise((ok, mal) => { const t = db.transaction('k').objectStore('k').get(k); t.onsuccess = () => ok(t.result); t.onerror = () => mal(t.error); }); },
    async poner(k, v){ const db = await this.db(); return new Promise((ok, mal) => { const t = db.transaction('k','readwrite').objectStore('k').put(v, k); t.onsuccess = () => ok(); t.onerror = () => mal(t.error); }); },
    async quitar(k){ const db = await this.db(); return new Promise(ok => { const t = db.transaction('k','readwrite').objectStore('k').delete(k); t.onsuccess = t.onerror = () => ok(); }); },
    aleatorio(n){ const a = new Uint8Array(n); crypto.getRandomValues(a); return Array.from(a, b => b.toString(16).padStart(2,'0')).join('').toUpperCase(); },
    async idDispositivo(pref){
      const k = (pref||'') + 'dev';
      let d = await this.leer(k).catch(() => null);
      if (!d) { d = (pref ? 'K' : 'D') + this.aleatorio(10); await this.poner(k, d); }
      return d;
    },
    async guardarLlave(b64, pref){
      const raw = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
      const key = await crypto.subtle.importKey('raw', raw, {name:'HMAC', hash:'SHA-256'}, false, ['sign']);
      await this.poner((pref||'') + 'llave', key);
    },
    async tieneLlave(pref){ return !!(await this.leer((pref||'') + 'llave').catch(() => null)); },
    async firmar(msg, pref){
      const key = await this.leer((pref||'') + 'llave');
      if (!key) return '';
      const s = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(msg));
      let b = ''; new Uint8Array(s).forEach(x => b += String.fromCharCode(x));
      return btoa(b);
    },
    descripcion(){
      const u = navigator.userAgent;
      const so = /iPhone|iPad/.test(u) ? (/iPad/.test(u) ? 'iPad' : 'iPhone') : /Android/.test(u) ? ((u.match(/Android[^;]*;\s*([^;)]+)/)||[])[1] || 'Android') : /Windows/.test(u) ? 'Windows' : /Mac/.test(u) ? 'Mac' : /Linux/.test(u) ? 'Linux' : 'Otro';
      const nav = /EdgA?\//.test(u) ? 'Edge' : /SamsungBrowser/.test(u) ? 'Samsung Internet' : /CriOS|Chrome/.test(u) ? 'Chrome' : /FxiOS|Firefox/.test(u) ? 'Firefox' : /Safari/.test(u) ? 'Safari' : 'Navegador';
      return (so + ' · ' + nav + ' · ' + screen.width + 'x' + screen.height).replace(/\s+/g,' ').trim();
    }
  };
