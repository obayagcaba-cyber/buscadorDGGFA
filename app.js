/* ==========================================================================
   Buscador DGGFA — lógica de la página
   Los datos vienen de datos.js (window.DATOS), generado desde el xlsx del mes.
   ========================================================================== */
(function () {
  'use strict';

  var D = window.DATOS;
  if (!D) { return; }

  /* ---------------------------------------------------------------- modelo */

  // Los textos repetidos vienen como índice de diccionario para achicar el
  // archivo. Esto devuelve el valor real de una celda.
  var idx = {};
  D.claves.forEach(function (clave, i) { idx[clave] = i; });

  function celda(fila, clave) {
    var i = idx[clave];
    var v = fila[i];
    var dic = D.dicts[i];
    return dic ? dic[v] : v;
  }

  // patente normalizada -> fila
  var porPatente = new Map();
  D.filas.forEach(function (fila) {
    porPatente.set(normalizar(celda(fila, 'patente')), fila);
  });

  function normalizar(texto) {
    return String(texto || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  /* ------------------------------------------------------------- utilidades */

  var SIN_DATO = '—';

  function esVacio(v) { return !v || v === '-'; }

  // 'Monto' y 'Valor mensual' llegan como ########## en buena parte del
  // archivo: el número real se perdió al cargarlo, no es un error de lectura.
  function esPerdido(v) { return typeof v === 'string' && v.indexOf('#') !== -1; }

  function aNumero(v) {
    if (esVacio(v) || esPerdido(v)) { return null; }
    var n = parseFloat(String(v).replace(/[$\s ]/g, '').replace(/\./g, '').replace(',', '.'));
    return isNaN(n) ? null : n;
  }

  function pesos(n) {
    return '$ ' + n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function miles(n) { return n.toLocaleString('es-AR'); }

  function texto(el, valor) { el.textContent = valor; return el; }

  function crear(tag, clase, contenido) {
    var el = document.createElement(tag);
    if (clase) { el.className = clase; }
    if (contenido != null) { el.textContent = contenido; }
    return el;
  }

  /* ------------------------------------------------------- ficha por patente */

  // Mismos bloques y campos que la hoja "Buscador" del Excel, con las
  // fórmulas apuntando a la columna correcta.
  var BLOQUES = [
    { num: '01', titulo: 'Identificación del vehículo', campos: [
      { clave: 'patente', rotulo: 'Patente' },
      { clave: 'marca', rotulo: 'Marca' },
      { clave: 'modelo', rotulo: 'Modelo' },
      { clave: 'anio', rotulo: 'Año' },
      { clave: 'tipo', rotulo: 'Tipo de vehículo' },
      { clave: 'detalleUnidad', rotulo: 'Detalle de tipo de unidad' },
      { clave: 'transmision', rotulo: 'Tipo de transmisión' },
      { clave: 'asientos', rotulo: 'Q de asientos' }
    ] },
    { num: '02', titulo: 'Asignación institucional', campos: [
      { clave: 'ministerio', rotulo: 'Ministerio' },
      { clave: 'reparticion', rotulo: 'Repartición' },
      { clave: 'km', rotulo: 'Kilometraje', formato: 'km' }
    ] },
    { num: '03', titulo: 'Servicios activos', campos: [
      { clave: 'telemetria', rotulo: 'Telemetría', formato: 'sino' },
      { clave: 'combustible', rotulo: 'Combustible', formato: 'sino' },
      { clave: 'mantenimiento', rotulo: 'Mantenimiento', formato: 'sino' },
      { clave: 'tipoCombustible', rotulo: 'Tipo de combustible asignado' }
    ] },
    { num: '04', titulo: 'Estado administrativo', campos: [
      { clave: 'anticuacion', rotulo: 'Anticuación', formato: 'marca', si: 'SI - Incluida' },
      { clave: 'poliza', rotulo: 'Póliza de Seguros DGGFA', formato: 'marca', si: 'SI - Figura en base informada' },
      { clave: 'patrimonial', rotulo: 'Inst. Patrimonial sin cobertura', formato: 'marca' },
      { clave: 'fichaEstante', rotulo: 'Ficha estante' },
      { clave: 'observaciones', rotulo: 'Observaciones' }
    ] },
    { num: '05', titulo: 'Seguro y póliza', campos: [
      { clave: 'tipoPoliza', rotulo: 'Tipo de póliza' },
      { clave: 'nroPoliza', rotulo: 'Nro. de póliza' },
      { clave: 'monto', rotulo: 'Monto', formato: 'importe' },
      { clave: 'valorMensual', rotulo: 'Valor mensual', formato: 'importe' },
      { clave: 'cuota', rotulo: 'Costo de la póliza mensual', formato: 'cuota' }
    ] },
    { num: '06', titulo: 'Otros', campos: [
      { clave: 'subasta', rotulo: 'Subasta', formato: 'marca' },
      { clave: 'ausa', rotulo: 'Unidades de AUSA', formato: 'marca' },
      { clave: 'donaciones', rotulo: 'Donaciones', formato: 'marca' },
      { clave: 'ivc', rotulo: 'Flota del IVC', formato: 'marca' }
    ] }
  ];

  function valorMostrable(fila, campo) {
    var bruto = campo.clave === 'cuota' ? null : celda(fila, campo.clave);

    switch (campo.formato) {
      case 'sino':
        // En la base el servicio activo se marca 'SI' (y 'ELÉCTRICO' en combustible).
        if (esVacio(bruto)) { return { texto: 'NO', clase: 'no' }; }
        return { texto: bruto.toUpperCase() === 'SI' ? 'SI' : bruto, clase: 'si' };

      case 'marca':
        // Columnas marcadas con X.
        if (String(bruto).toUpperCase() === 'X') {
          return { texto: campo.si || 'SI', clase: 'si' };
        }
        return { texto: 'NO', clase: 'no' };

      case 'km':
        var km = aNumero(bruto);
        if (km === null) { return { texto: SIN_DATO, clase: 'vacio' }; }
        return { texto: miles(km) + ' km' };

      case 'importe':
        if (esPerdido(bruto)) {
          return { texto: SIN_DATO, clase: 'vacio', nota: 'El archivo de origen no trae el número' };
        }
        if (esVacio(bruto)) { return { texto: SIN_DATO, clase: 'vacio' }; }
        var imp = aNumero(bruto);
        return { texto: imp === null ? bruto : pesos(imp) };

      case 'cuota':
        // Igual que el Excel: monto de la póliza dividido 24.
        var monto = aNumero(celda(fila, 'monto'));
        if (monto === null) { return { texto: SIN_DATO, clase: 'vacio' }; }
        return { texto: pesos(monto / 24), nota: 'Monto ÷ 24' };

      default:
        if (esVacio(bruto)) { return { texto: SIN_DATO, clase: 'vacio' }; }
        return { texto: bruto };
    }
  }

  var contFicha = document.getElementById('ficha');
  var contAviso = document.getElementById('aviso');

  function mostrarFicha(fila) {
    contFicha.innerHTML = '';
    contAviso.className = 'oculto';

    var cab = crear('div', 'ficha-cabecera');
    cab.appendChild(crear('div', 'dominio', celda(fila, 'patente')));

    var marca = celda(fila, 'marca');
    var modelo = celda(fila, 'modelo');
    var anio = celda(fila, 'anio');
    var desc = [marca, modelo].filter(function (t) { return !esVacio(t); }).join(' ');
    if (!esVacio(anio)) { desc += (desc ? ' · ' : '') + anio; }
    cab.appendChild(crear('div', 'vehiculo', desc || SIN_DATO));

    var min = celda(fila, 'ministerio');
    var rep = celda(fila, 'reparticion');
    var org = [min, rep].filter(function (t) { return !esVacio(t); }).join('  ·  ');
    cab.appendChild(crear('div', 'organismo-veh', org || 'Sin asignación registrada'));

    var chips = crear('div', 'etiquetas');
    [['telemetria', 'Telemetría'], ['combustible', 'Combustible'], ['mantenimiento', 'Mantenimiento']]
      .forEach(function (par) {
        var activo = !esVacio(celda(fila, par[0]));
        chips.appendChild(crear('span', 'chip ' + (activo ? 'si' : 'no'), par[1]));
      });
    cab.appendChild(chips);
    contFicha.appendChild(cab);

    BLOQUES.forEach(function (bloque) {
      var sec = crear('section', 'seccion');
      var h3 = crear('h3');
      h3.appendChild(crear('span', 'num', bloque.num));
      h3.appendChild(document.createTextNode(bloque.titulo));
      sec.appendChild(h3);

      var dl = crear('dl', 'datos');
      bloque.campos.forEach(function (campo) {
        var v = valorMostrable(fila, campo);
        var caja = crear('div', 'dato');
        caja.appendChild(crear('dt', null, campo.rotulo));
        var dd = crear('dd', v.clase || null, v.texto);
        if (v.nota) { dd.appendChild(crear('span', 'nota-dato', v.nota)); }
        caja.appendChild(dd);
        dl.appendChild(caja);
      });
      sec.appendChild(dl);
      contFicha.appendChild(sec);
    });
  }

  function mostrarAviso(html, esError) {
    contFicha.innerHTML = '';
    contAviso.className = 'aviso' + (esError ? ' error' : '');
    contAviso.innerHTML = html;
  }

  function estadoInicial() {
    mostrarAviso('Escribí un dominio para ver la ficha completa de la unidad. ' +
      'La base tiene <b>' + miles(D.filas.length) + '</b> unidades cargadas al período ' + D.periodo + '.');
  }

  /* --------------------------------------------------------------- búsqueda */

  var input = document.getElementById('patente');
  var lista = document.getElementById('sugerencias');
  var botonLimpiar = document.getElementById('limpiar');
  var seleccion = -1;

  function sugerir(texto) {
    var q = normalizar(texto);
    lista.innerHTML = '';
    seleccion = -1;
    if (q.length < 2) { lista.className = 'sugerencias oculto'; return; }

    var empiezan = [];
    var contienen = [];
    porPatente.forEach(function (fila, pat) {
      if (empiezan.length + contienen.length > 400) { return; }
      if (pat.indexOf(q) === 0) { empiezan.push(fila); }
      else if (pat.indexOf(q) !== -1) { contienen.push(fila); }
    });

    var res = empiezan.concat(contienen).slice(0, 8);
    if (!res.length) { lista.className = 'sugerencias oculto'; return; }

    res.forEach(function (fila) {
      var li = crear('li');
      li.setAttribute('role', 'option');
      li.appendChild(crear('span', 'pat', celda(fila, 'patente')));
      var det = [celda(fila, 'marca'), celda(fila, 'modelo')]
        .filter(function (t) { return !esVacio(t); }).join(' ');
      li.appendChild(crear('span', 'det', det || celda(fila, 'tipo')));
      li.addEventListener('mousedown', function (e) {
        e.preventDefault();
        elegir(celda(fila, 'patente'));
      });
      lista.appendChild(li);
    });
    lista.className = 'sugerencias';
  }

  function elegir(patente) {
    input.value = patente;
    lista.className = 'sugerencias oculto';
    buscar();
    input.blur();
  }

  function buscar() {
    var q = normalizar(input.value);
    botonLimpiar.className = input.value ? 'limpiar' : 'limpiar oculto';

    if (!q) { estadoInicial(); return; }

    var fila = porPatente.get(q);
    if (fila) {
      mostrarFicha(fila);
      // deja el dominio en la URL para poder compartir el link de una unidad
      if (window.history.replaceState) {
        window.history.replaceState(null, '', '#' + celda(fila, 'patente'));
      }
      return;
    }
    mostrarAviso('No encontramos el dominio <b>' + input.value.toUpperCase() +
      '</b> en la base de ' + D.periodo + '.<br>Revisá que esté bien escrito o probá con las primeras letras.', true);
  }

  input.addEventListener('input', function () {
    sugerir(input.value);
    if (normalizar(input.value).length >= 6) { buscar(); }
    else { botonLimpiar.className = input.value ? 'limpiar' : 'limpiar oculto'; }
  });

  input.addEventListener('keydown', function (e) {
    var items = lista.querySelectorAll('li');
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      if (!items.length) { return; }
      e.preventDefault();
      if (seleccion >= 0) { items[seleccion].removeAttribute('aria-selected'); }
      seleccion += (e.key === 'ArrowDown' ? 1 : -1);
      if (seleccion < 0) { seleccion = items.length - 1; }
      if (seleccion >= items.length) { seleccion = 0; }
      items[seleccion].setAttribute('aria-selected', 'true');
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (seleccion >= 0 && items[seleccion]) { items[seleccion].dispatchEvent(new Event('mousedown')); }
      else { lista.className = 'sugerencias oculto'; buscar(); }
    } else if (e.key === 'Escape') {
      lista.className = 'sugerencias oculto';
    }
  });

  input.addEventListener('blur', function () {
    setTimeout(function () { lista.className = 'sugerencias oculto'; }, 120);
  });

  botonLimpiar.addEventListener('click', function () {
    input.value = '';
    lista.className = 'sugerencias oculto';
    buscar();
    input.focus();
  });

  /* ---------------------------------------------------------------- resumen */

  // Orden de la hoja "Resumen". El texto tiene que coincidir con el de la base.
  var TIPOS = [
    'UTILITARIOS KPB (Kangoo, Partner, Berlingo)',
    'PICK UPS',
    'UTILITARIOS MEDIANOS / PESADOS (NO KPB)',
    'Sedan',
    'Motovehiculo',
    'UTILITARIO ELÉCTRICO',
    'SEDAN ELÉCTRICO',
    'Maquinaria',
    'MOTO ELÉCTRICA'
  ];

  function contar(prueba) {
    var n = 0;
    D.filas.forEach(function (fila) { if (prueba(fila)) { n++; } });
    return n;
  }

  function marcada(fila, clave) { return String(celda(fila, clave)).toUpperCase() === 'X'; }
  function activo(fila, clave) { return !esVacio(celda(fila, clave)); }
  function esTipo(fila, tipo) { return celda(fila, 'tipo') === tipo; }

  // clase: 'por-tipo' cuando la primera columna lista tipos de vehículo,
  // que van siempre en mayúscula.
  function tabla(contenedor, encabezados, filas, pie, clase) {
    var marco = crear('div', 'marco-tabla');
    var t = crear('table', clase || null);

    var thead = crear('thead');
    var tr = crear('tr');
    encabezados.forEach(function (h) { tr.appendChild(crear('th', null, h)); });
    thead.appendChild(tr);
    t.appendChild(thead);

    var tbody = crear('tbody');
    filas.forEach(function (f) {
      var fila = crear('tr');
      f.forEach(function (c, i) {
        fila.appendChild(crear('td', null, i === 0 ? c : miles(c)));
      });
      tbody.appendChild(fila);
    });
    t.appendChild(tbody);

    if (pie) {
      var tfoot = crear('tfoot');
      var fp = crear('tr');
      pie.forEach(function (c, i) { fp.appendChild(crear('td', null, i === 0 ? c : miles(c))); });
      tfoot.appendChild(fp);
      t.appendChild(tfoot);
    }

    marco.appendChild(t);
    contenedor.appendChild(marco);
  }

  function sumaColumnas(filas) {
    var total = [];
    filas.forEach(function (f) {
      for (var i = 1; i < f.length; i++) { total[i] = (total[i] || 0) + f[i]; }
    });
    return total;
  }

  function armarResumen() {
    var cont = document.getElementById('resumen-cuerpo');
    cont.innerHTML = '';

    var totalFlota = D.filas.length;
    var totalPoliza = contar(function (f) { return marcada(f, 'poliza'); });
    var totalAntic = contar(function (f) { return marcada(f, 'anticuacion'); });
    var totalPatrim = contar(function (f) { return marcada(f, 'patrimonial'); });

    // --- KPIs
    var kpis = crear('div', 'kpis');
    function kpi(cifra, rotulo, clase) {
      var c = crear('div', 'kpi' + (clase ? ' ' + clase : ''));
      c.appendChild(crear('div', 'cifra', miles(cifra)));
      c.appendChild(crear('div', 'rotulo', rotulo));
      return c;
    }
    kpis.appendChild(kpi(totalFlota, 'Flota automotor', 'protagonista'));
    kpis.appendChild(kpi(totalPoliza, 'Póliza de Seguros DGGFA informada al 30-3-26'));
    kpis.appendChild(kpi(totalAntic, 'Anticuación ' + D.periodo + ' (activos)'));
    kpis.appendChild(kpi(totalPatrim, 'Inst. Patrimonial sin cobertura', 'menta'));
    cont.appendChild(kpis);

    cont.appendChild(crear('p', 'nota',
      'Referencia — Registro Nacional del Parque Automotor: 3.561 unidades. ' +
      'El total de flota incluye unidades contabilizadas doble por poseer el mismo servicio.'));

    // --- Unidades por tipo
    cont.appendChild(crear('h3', 'subtitulo-tabla principal', 'Unidades por tipo de vehículo'));
    var porTipo = TIPOS.map(function (tipo) {
      return [
        tipo,
        contar(function (f) { return esTipo(f, tipo); }),
        contar(function (f) { return esTipo(f, tipo) && marcada(f, 'poliza'); }),
        contar(function (f) { return esTipo(f, tipo) && marcada(f, 'anticuacion'); }),
        contar(function (f) { return esTipo(f, tipo) && marcada(f, 'patrimonial'); })
      ];
    });
    tabla(cont,
      ['Tipo', 'Total', 'Póliza DGGFA', 'Anticuación', 'Inst. Patrimonial'],
      porTipo,
      ['TOTAL'].concat(sumaColumnas(porTipo).slice(1)),
      'por-tipo');

    // --- Servicios por tipo
    cont.appendChild(crear('h3', 'subtitulo-tabla', 'Servicios por tipo de vehículo'));
    var porServicio = TIPOS.map(function (tipo) {
      return [
        tipo,
        contar(function (f) { return esTipo(f, tipo) && activo(f, 'combustible'); }),
        contar(function (f) { return esTipo(f, tipo) && activo(f, 'telemetria'); }),
        contar(function (f) { return esTipo(f, tipo) && activo(f, 'mantenimiento'); }),
        contar(function (f) {
          return esTipo(f, tipo) && activo(f, 'combustible') && activo(f, 'telemetria') && activo(f, 'mantenimiento');
        })
      ];
    });
    tabla(cont,
      ['Tipo', 'Combustible', 'Telemetría', 'Mantenimiento', 'Los 3 servicios'],
      porServicio,
      ['TOTAL'].concat(sumaColumnas(porServicio).slice(1)),
      'por-tipo');
    cont.appendChild(crear('p', 'nota',
      'Combustible incluye las ' +
      miles(contar(function (f) { return celda(f, 'combustible').toUpperCase() === 'ELÉCTRICO'; })) +
      ' unidades con carga eléctrica asignada.'));

    // --- Kilometraje
    cont.appendChild(crear('h3', 'subtitulo-tabla', 'Kilometraje'));
    function enRango(fila, min, max) {
      var km = aNumero(celda(fila, 'km'));
      if (km === null) { return false; }
      return km > min && km <= max;
    }
    var porKm = TIPOS.map(function (tipo) {
      var a = contar(function (f) { return esTipo(f, tipo) && enRango(f, -1, 50000); });
      var b = contar(function (f) { return esTipo(f, tipo) && enRango(f, 50000, 100000); });
      var c = contar(function (f) { return esTipo(f, tipo) && enRango(f, 100000, Infinity); });
      return [tipo, a, b, c, a + b + c];
    });
    tabla(cont,
      ['Tipo', 'Hasta 50.000 km', 'De 50.000 a 100.000', 'Más de 100.000', 'Total'],
      porKm,
      ['TOTAL'].concat(sumaColumnas(porKm).slice(1)),
      'por-tipo');
    cont.appendChild(crear('p', 'nota',
      'Solo se cuentan las unidades con kilometraje cargado (' +
      miles(contar(function (f) { return aNumero(celda(f, 'km')) !== null; })) + ' de ' + miles(totalFlota) + ').'));

    // --- Ministerios
    cont.appendChild(crear('h3', 'subtitulo-tabla', 'Unidades asignadas por Ministerio'));
    var cuenta = {};
    D.filas.forEach(function (f) {
      var m = celda(f, 'ministerio');
      if (esVacio(m)) { return; }
      cuenta[m] = (cuenta[m] || 0) + 1;
    });
    var porMinisterio = Object.keys(cuenta)
      .map(function (m) { return [m, cuenta[m]]; })
      .sort(function (a, b) { return b[1] - a[1]; });
    tabla(cont,
      ['Ministerio', 'Unidades'],
      porMinisterio,
      ['TOTAL GENERAL'].concat(sumaColumnas(porMinisterio).slice(1)));
    cont.appendChild(crear('p', 'nota',
      'No incluye ' + miles(totalFlota - sumaColumnas(porMinisterio)[1]) +
      ' unidades sin ministerio asignado en la base.'));
  }

  /* --------------------------------------------------------------- solapas */

  var solapas = document.querySelectorAll('.solapa');
  solapas.forEach(function (boton) {
    boton.addEventListener('click', function () {
      solapas.forEach(function (b) {
        var activa = b === boton;
        b.setAttribute('aria-selected', activa ? 'true' : 'false');
        document.getElementById(b.dataset.panel).className = activa ? 'panel' : 'panel oculto';
      });
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });

  /* ----------------------------------------------------------------- inicio */

  document.getElementById('periodo').textContent = D.periodo;
  document.querySelectorAll('.js-periodo').forEach(function (el) { el.textContent = D.periodo; });
  if (D.actualizado) { document.getElementById('actualizado').textContent = D.actualizado; }

  armarResumen();

  var inicial = decodeURIComponent(window.location.hash.replace('#', ''));
  if (inicial && porPatente.has(normalizar(inicial))) {
    input.value = inicial.toUpperCase();
    buscar();
  } else {
    estadoInicial();
  }
})();
