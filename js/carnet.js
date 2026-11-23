/* CTPL · dibujo del carnet */
const PUESTOS = {
    DOCENTE_TECNICO:   { txt:'Docente técnico',          l:{M:'PERSONAL DOCENTE',F:'PERSONAL DOCENTE'}, r:{M:'TÉCNICO',F:'TÉCNICO'} },
    DOCENTE_ACADEMICO: { txt:'Docente académico',        l:{M:'PERSONAL DOCENTE',F:'PERSONAL DOCENTE'}, r:{M:'ACADÉMICO',F:'ACADÉMICO'} },
    DOCENTE_IDIOMAS:   { txt:'Docente Departamento de Idiomas', l:{M:'PERSONAL DOCENTE',F:'PERSONAL DOCENTE'}, r:{M:'IDIOMAS',F:'IDIOMAS'} },
    ADMINISTRATIVO:    { txt:'Personal administrativo',  l:{M:'PERSONAL',F:'PERSONAL'}, r:{M:'ADMINISTRATIVO',F:'ADMINISTRATIVO'} },
    AUXILIAR_ADMIN:    { txt:'Auxiliar administrativo',  l:{M:'AUXILIAR',F:'AUXILIAR'}, r:{M:'ADMINISTRATIVO',F:'ADMINISTRATIVA'} },
    COORD_TECNICO:     { txt:'Coordinador(a) técnico(a)',l:{M:'COORDINADOR',F:'COORDINADORA'}, r:{M:'TÉCNICO',F:'TÉCNICA'} },
    COORD_EMPRESA:     { txt:'Coordinador(a) con la empresa', l:{M:'COORDINADOR',F:'COORDINADORA'}, r:{M:'CON LA EMPRESA',F:'CON LA EMPRESA'} },
    COORD_ACADEMICO:   { txt:'Coordinador(a) académico(a)', l:{M:'COORDINADOR',F:'COORDINADORA'}, r:{M:'ACADÉMICO',F:'ACADÉMICA'} },
    SUBDIRECTOR:       { txt:'Subdirector(a)',           l:{M:'DIRECCIÓN',F:'DIRECCIÓN'}, r:{M:'SUBDIRECTOR',F:'SUBDIRECTORA'} },
    DIRECTOR:          { txt:'Director(a)',              l:{M:'DIRECCIÓN',F:'DIRECCIÓN'}, r:{M:'DIRECTOR',F:'DIRECTORA'} }
  };
  /* Fondo del carnet: imagen del sitio; el administrador puede reemplazarla desde el panel */
  let PLANTILLA = 'img/plantilla_carnet.jpg';

  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

  /** p = {nombre, cedula, puesto, genero, foto(dataURL), qr(dataURL opcional)} */
  function carnetHTML(p){
    const pu = PUESTOS[p.puesto] || {l:{M:'',F:''}, r:{M:'',F:''}};
    const g = p.genero === 'F' ? 'F' : 'M';
    return '<div class="carnet">' +
      (PLANTILLA ? '<img class="c-bg" src="'+PLANTILLA+'" alt="">' : '<div class="c-bg-vacio"></div>') +
      '<div class="c-foto">' + (p.foto ? '<img src="'+p.foto+'" alt="">' : '') + '</div>' +
      '<div class="c-nombre"><span>' + esc((p.nombre||'NOMBRE COMPLETO').toUpperCase()) + '</span></div>' +
      '<div class="c-ced">' + esc(p.cedula || '000000000') + '</div>' +
      '<div class="c-lbl">' + esc(pu.l[g]) + '</div>' +
      '<div class="c-rol">' + esc(pu.r[g]) + '</div>' +
      '<div class="c-qr">' + (p.qr ? '<img src="'+p.qr+'" alt="">' : '') + '</div>' +
    '</div>';
  }

  /** Ajusta el tamaño de letra para que el nombre quepa en 2 líneas y el puesto en 1 */
  function ajustarCarnets(raiz){
    (raiz || document).querySelectorAll('.carnet').forEach(c => {
      const n = c.querySelector('.c-nombre');
      let f = 5.7; n.style.fontSize = f + 'cqw';
      while (n.firstChild.offsetHeight > n.clientHeight + 1 && f > 3.2){ f -= 0.15; n.style.fontSize = f + 'cqw'; }
      const r = c.querySelector('.c-rol');
      let g = 7.4; r.style.fontSize = g + 'cqw';
      while (r.scrollWidth > r.clientWidth + 1 && g > 4){ g -= 0.15; r.style.fontSize = g + 'cqw'; }
    });
  }
