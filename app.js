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

  // Para buscar por nombre: sin acentos ni mayúsculas, así "hacienda"
  // encuentra "MHFGC - Hacienda y Finanzas".
  function normTexto(texto) {
    return String(texto || '').toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  // Ministerios y reparticiones con su cantidad de unidades, para poder
  // buscar por organismo además de por dominio.
  var GRUPOS = [
    { clave: 'ministerio', etiqueta: 'Ministerio',
      desglose: { clave: 'reparticion', titulo: 'Unidades por Repartición', rotulo: 'Repartición' } },
    // En una repartición no hay nivel siguiente: en vez de un desglose va el
    // detalle de las unidades que la componen.
    { clave: 'reparticion', etiqueta: 'Repartición', detalle: true }
  ];

  GRUPOS.forEach(function (g) {
    var cuenta = {};
    D.filas.forEach(function (fila) {
      var v = celda(fila, g.clave);
      if (esVacio(v)) { return; }
      cuenta[v] = (cuenta[v] || 0) + 1;
    });
    g.valores = Object.keys(cuenta)
      .map(function (v) { return { valor: v, norm: normTexto(v), cantidad: cuenta[v] }; })
      .sort(function (a, b) { return b.cantidad - a.cantidad; });
  });

  function filasDe(clave, valor) {
    return D.filas.filter(function (f) { return celda(f, clave) === valor; });
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
      var ajustes = (D.correcciones || {})[celda(fila, 'patente')] || {};

      bloque.campos.forEach(function (campo) {
        var v = valorMostrable(fila, campo);
        var caja = crear('div', 'dato');
        caja.appendChild(crear('dt', null, campo.rotulo));
        var dd = crear('dd', v.clase || null, v.texto);
        if (v.nota) { dd.appendChild(crear('span', 'nota-dato', v.nota)); }
        // Dato ajustado por la DGGFA: se muestra corregido y se dice por qué.
        if (ajustes[campo.clave]) {
          dd.appendChild(crear('span', 'nota-dato ajustado', ajustes[campo.clave]));
        }
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
    mostrarAviso('Escribí un <b>dominio</b> para ver la ficha completa de la unidad, ' +
      'o el nombre de un <b>ministerio</b> o una <b>repartición</b> para ver su resumen de flota.<br>' +
      'La base tiene ' + miles(D.filas.length) + ' unidades cargadas al período ' + D.periodo + '.');
  }

  /* --------------------------------------------------------------- búsqueda */

  var input = document.getElementById('patente');
  var lista = document.getElementById('sugerencias');
  var botonLimpiar = document.getElementById('limpiar');
  var seleccion = -1;

  // Una sugerencia puede ser un dominio o un organismo. Se muestran juntas,
  // cada una con su etiqueta, y el buscador resuelve según lo que se elija.
  function sugerir(texto) {
    lista.innerHTML = '';
    seleccion = -1;

    var crudo = String(texto).trim();
    if (crudo.length < 2) { lista.className = 'sugerencias oculto'; return; }

    var q = normalizar(crudo);
    var qt = normTexto(crudo);
    var opciones = [];

    // Organismos primero: son pocos y quien los busca sabe lo que quiere
    GRUPOS.forEach(function (g) {
      g.valores.forEach(function (v) {
        if (opciones.length > 60 || v.norm.indexOf(qt) === -1) { return; }
        opciones.push({
          grupo: g, valor: v.valor,
          etiqueta: g.etiqueta,
          principal: v.valor,
          detalle: miles(v.cantidad) + (v.cantidad === 1 ? ' unidad' : ' unidades'),
          empieza: v.norm.indexOf(qt) === 0
        });
      });
    });
    opciones.sort(function (a, b) { return (b.empieza ? 1 : 0) - (a.empieza ? 1 : 0); });
    opciones = opciones.slice(0, 5);

    // Dominios
    if (q.length >= 2) {
      var empiezan = [], contienen = [];
      porPatente.forEach(function (fila, pat) {
        if (empiezan.length + contienen.length > 400) { return; }
        if (pat.indexOf(q) === 0) { empiezan.push(fila); }
        else if (pat.indexOf(q) !== -1) { contienen.push(fila); }
      });
      empiezan.concat(contienen).slice(0, opciones.length ? 5 : 8).forEach(function (fila) {
        var det = [celda(fila, 'marca'), celda(fila, 'modelo')]
          .filter(function (t) { return !esVacio(t); }).join(' ');
        opciones.push({
          patente: celda(fila, 'patente'),
          etiqueta: 'Dominio',
          principal: celda(fila, 'patente'),
          detalle: det || celda(fila, 'tipo')
        });
      });
    }

    if (!opciones.length) { lista.className = 'sugerencias oculto'; return; }

    opciones.forEach(function (op) {
      var li = crear('li');
      li.setAttribute('role', 'option');
      li.appendChild(crear('span', 'tipo-sug', op.etiqueta));
      li.appendChild(crear('span', 'pat', op.principal));
      li.appendChild(crear('span', 'det', op.detalle));
      li.addEventListener('mousedown', function (e) {
        e.preventDefault();
        if (op.patente) { elegir(op.patente); }
        else { elegirGrupo(op.grupo, op.valor); }
      });
      lista.appendChild(li);
    });
    lista.className = 'sugerencias';
  }

  function elegir(patente) {
    input.value = patente;
    ajustarCampo();
    lista.className = 'sugerencias oculto';
    buscar();
    input.blur();
  }

  function elegirGrupo(grupo, valor) {
    input.value = valor;
    ajustarCampo();
    lista.className = 'sugerencias oculto';
    input.blur();
    mostrarGrupo(grupo, valor);
  }

  function enlace(hash) {
    if (window.history.replaceState) {
      window.history.replaceState(null, '', '#' + hash);
    }
  }

  // Panel de un ministerio o repartición: los mismos indicadores y cuadros
  // del resumen general, calculados solo sobre sus unidades.
  function mostrarGrupo(grupo, valor) {
    var filas = filasDe(grupo.clave, valor);
    if (!filas.length) {
      mostrarAviso('No hay unidades registradas en <b>' + valor + '</b>.', true);
      return;
    }
    contAviso.className = 'oculto';
    contFicha.innerHTML = '';

    var cab = crear('div', 'ficha-cabecera grupo');
    cab.appendChild(crear('div', 'etiqueta-grupo', grupo.etiqueta));
    cab.appendChild(crear('div', 'nombre-grupo', valor));
    cab.appendChild(crear('div', 'organismo-veh',
      miles(filas.length) + (filas.length === 1 ? ' unidad asignada' : ' unidades asignadas') +
      '  ·  ' + D.periodo));
    contFicha.appendChild(cab);

    var cuerpo = crear('div', 'resumen-grupo');
    contFicha.appendChild(cuerpo);
    armarResumen(cuerpo, filas, { desglose: grupo.desglose, detalle: grupo.detalle });

    enlace(grupo.clave + '=' + encodeURIComponent(valor));
  }

  // Coincidencia exacta de organismo, para que escribir el nombre completo
  // y apretar Enter funcione igual que elegirlo de la lista.
  function grupoExacto(texto) {
    var qt = normTexto(String(texto).trim());
    for (var i = 0; i < GRUPOS.length; i++) {
      var encontrado = null;
      GRUPOS[i].valores.forEach(function (v) {
        if (v.norm === qt) { encontrado = v.valor; }
      });
      if (encontrado) { return { grupo: GRUPOS[i], valor: encontrado }; }
    }
    return null;
  }

  function buscar() {
    var crudo = input.value;
    var q = normalizar(crudo);
    botonLimpiar.className = crudo ? 'limpiar' : 'limpiar oculto';

    if (!crudo.trim()) { estadoInicial(); return; }

    var fila = porPatente.get(q);
    if (fila) {
      mostrarFicha(fila);
      // deja el dominio en la URL para poder compartir el link de una unidad
      enlace(celda(fila, 'patente'));
      return;
    }

    var g = grupoExacto(crudo);
    if (g) { mostrarGrupo(g.grupo, g.valor); return; }

    mostrarAviso('No encontramos <b>' + crudo.toUpperCase() +
      '</b> en la base de ' + D.periodo + '.<br>' +
      'Probá con un dominio, o con el nombre de un ministerio o una repartición.', true);
  }

  // Una patente entra holgada en mayúsculas espaciadas; el nombre de un
  // organismo, no. El campo cambia de registro según lo que se escriba.
  function ajustarCampo() {
    var v = input.value;
    var esNombre = v.indexOf(' ') !== -1 || v.length > 9 || /[a-z]{4,}/.test(v);
    input.className = esNombre ? 'texto-largo' : '';
  }

  input.addEventListener('input', function () {
    ajustarCampo();
    sugerir(input.value);
    if (normalizar(input.value).length >= 6 && porPatente.has(normalizar(input.value))) { buscar(); }
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

  // Cuenta sobre el conjunto que se le pase: la base entera para el resumen
  // general, o solo las unidades de un ministerio o repartición.
  function contar(conjunto, prueba) {
    var n = 0;
    conjunto.forEach(function (fila) { if (prueba(fila)) { n++; } });
    return n;
  }

  function marcada(fila, clave) { return String(celda(fila, clave)).toUpperCase() === 'X'; }
  function activo(fila, clave) { return !esVacio(celda(fila, clave)); }
  function esTipo(fila, tipo) { return celda(fila, 'tipo') === tipo; }

  // Un cuadro = barra de título + encabezados + datos + fila de totales.
  // opciones: {principal: true} para el análisis destacado, {nota: '...'}
  // para la aclaración al pie.
  function cuadro(contenedor, titulo, encabezados, filas, pie, opciones) {
    opciones = opciones || {};
    var sec = crear('section', 'cuadro' + (opciones.principal ? ' principal' : ''));
    sec.appendChild(crear('h3', 'titulo-cuadro', titulo));

    var marco = crear('div', 'marco-tabla');
    var t = crear('table');

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
    sec.appendChild(marco);
    if (opciones.nota) { sec.appendChild(crear('p', 'nota', opciones.nota)); }
    contenedor.appendChild(sec);
  }

  function sumaColumnas(filas) {
    var total = [];
    filas.forEach(function (f) {
      for (var i = 1; i < f.length; i++) { total[i] = (total[i] || 0) + f[i]; }
    });
    return total;
  }

  function enRango(fila, min, max) {
    var km = aNumero(celda(fila, 'km'));
    if (km === null) { return false; }
    return km > min && km <= max;
  }

  /* ---------------------------------------------------------- gráfico anillo

     Paleta de seis tonos suaves generada en OKLCH (L 0.74 / C 0.11) y validada:
     entra en la banda de luminosidad, supera el piso de croma y mantiene una
     separación de ΔE 8.8 bajo daltonismo y 17.3 en visión normal. Los tonos
     tienen poco contraste contra el fondo a propósito, por eso cada segmento
     lleva su valor en la leyenda y el cuadro de detalle repite todo en tabla:
     nunca hay que distinguir un color para leer un dato.

     El orden es fijo. No se recicla ni se agrega un séptimo: a partir del
     quinto valor el resto se agrupa en "Otras". */
  var PALETA = ['#5BC19D', '#B59BE6', '#C4A953', '#DF8DB5', '#6CB2EC', '#E4956D'];
  var MAX_SEGMENTOS = 6;

  /* Antigüedad por tramos. Los cortes salen del año del período del archivo,
     no de una constante: en 2027 se corren solos. */
  var ANIO_BASE = parseInt((D.periodo.match(/(20\d\d)/) || [])[1], 10) || 2026;

  var TRAMOS_ANTIGUEDAD = [
    'Hasta 3 años (' + ANIO_BASE + ' a ' + (ANIO_BASE - 2) + ')',
    'De 3 a 5 años (' + (ANIO_BASE - 3) + ' a ' + (ANIO_BASE - 5) + ')',
    'Más de 5 años (' + (ANIO_BASE - 6) + ' inclusive)',
    'Sin dato'
  ];

  function antiguedad(fila) {
    var a = parseInt(celda(fila, 'anio'), 10);
    if (!a) { return TRAMOS_ANTIGUEDAD[3]; }
    if (a >= ANIO_BASE - 2) { return TRAMOS_ANTIGUEDAD[0]; }
    if (a >= ANIO_BASE - 5) { return TRAMOS_ANTIGUEDAD[1]; }
    return TRAMOS_ANTIGUEDAD[2];
  }

  // dim: {clave} para leer una columna, o {fn} para calcular la categoría.
  // orden: sin definir ordena por cantidad; un arreglo fija el orden.
  function agrupar(filas, dim, orden) {
    var cuenta = {};
    filas.forEach(function (f) {
      var v = dim.fn ? dim.fn(f) : celda(f, dim.clave);
      cuenta[esVacio(v) ? 'Sin dato' : v] = (cuenta[esVacio(v) ? 'Sin dato' : v] || 0) + 1;
    });
    // El recorte va SIEMPRE por volumen: si se recortara por año, los cinco
    // valores mas viejos taparian al resto y "Otras" se comeria el grafico.
    var lista = Object.keys(cuenta)
      .map(function (k) { return [k, cuenta[k]]; })
      .sort(function (a, b) { return b[1] - a[1]; });

    var resto = null;
    if (lista.length > MAX_SEGMENTOS) {
      var sobrantes = lista.slice(MAX_SEGMENTOS - 1);
      lista = lista.slice(0, MAX_SEGMENTOS - 1);
      resto = ['Otras (' + sobrantes.length + ')',
               sobrantes.reduce(function (t, x) { return t + x[1]; }, 0)];
    }

    // Recién ahora, si la dimensión tiene un orden propio, se aplica sobre lo
    // que quedó visible: los tramos de antigüedad van del más nuevo al más
    // viejo, no del más numeroso al menos.
    if (orden && orden.length) {
      lista.sort(function (a, b) { return orden.indexOf(a[0]) - orden.indexOf(b[0]); });
    }
    if (resto) { lista.push(resto); }
    return lista;
  }

  var SVG = 'http://www.w3.org/2000/svg';

  function svgEl(tag, attrs) {
    var el = document.createElementNS(SVG, tag);
    Object.keys(attrs).forEach(function (k) { el.setAttribute(k, attrs[k]); });
    return el;
  }

  // Arco de anillo entre dos ángulos, en grados desde las 12 en punto
  function arco(cx, cy, rExt, rInt, desde, hasta) {
    var p = function (r, a) {
      var rad = (a - 90) * Math.PI / 180;
      return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
    };
    var largo = (hasta - desde) > 180 ? 1 : 0;
    var a1 = p(rExt, desde), a2 = p(rExt, hasta), b1 = p(rInt, hasta), b2 = p(rInt, desde);
    return 'M' + a1 + 'A' + rExt + ',' + rExt + ' 0 ' + largo + ' 1 ' + a2 +
           'L' + b1 + 'A' + rInt + ',' + rInt + ' 0 ' + largo + ' 0 ' + b2 + 'Z';
  }

  function dibujarAnillo(datos, total) {
    var fig = crear('div', 'anillo');
    var svg = svgEl('svg', { viewBox: '0 0 240 240', role: 'img' });
    svg.setAttribute('aria-label', 'Distribución en anillo; los valores están en la leyenda y en el cuadro de detalle.');

    var angulo = 0;
    datos.forEach(function (d, i) {
      var porcion = d[1] / total * 360;
      // 2px de separación entre porciones, que es el fondo asomando
      var recorte = porcion > 6 ? 1.2 : 0;
      var path = svgEl('path', {
        d: arco(120, 120, 100, 64, angulo + recorte, angulo + porcion - recorte),
        fill: PALETA[i % PALETA.length],
        'data-i': i
      });
      path.appendChild(svgEl('title', {})).textContent =
        d[0] + ': ' + miles(d[1]) + ' (' + (d[1] / total * 100).toFixed(1) + '%)';
      svg.appendChild(path);
      angulo += porcion;
    });

    var centro = crear('div', 'anillo-centro');
    centro.appendChild(crear('span', 'anillo-total', miles(total)));
    centro.appendChild(crear('span', 'anillo-rotulo', total === 1 ? 'unidad' : 'unidades'));

    fig.appendChild(svg);
    fig.appendChild(centro);
    return fig;
  }

  function dibujarLeyenda(datos, total) {
    var ul = crear('ul', 'leyenda');
    datos.forEach(function (d, i) {
      var li = crear('li');
      li.appendChild(svgEl('svg', { class: 'muestra', viewBox: '0 0 10 10', 'aria-hidden': 'true' }))
        .appendChild(svgEl('rect', { x: 0, y: 0, width: 10, height: 10, rx: 3, fill: PALETA[i % PALETA.length] }));
      li.appendChild(crear('span', 'leyenda-nombre', d[0]));
      li.appendChild(crear('span', 'leyenda-valor', miles(d[1])));
      li.appendChild(crear('span', 'leyenda-pct', (d[1] / total * 100).toFixed(1) + '%'));
      ul.appendChild(li);
    });
    return ul;
  }

  // Cuadro con el anillo y los botones para cambiar de dimensión
  function cuadroGrafico(cont, filas, dimensiones) {
    var sec = crear('section', 'cuadro grafico');
    sec.appendChild(crear('h3', 'titulo-cuadro', 'Composición de la flota'));

    var cuerpo = crear('div', 'grafico-cuerpo');
    var controles = crear('div', 'dimensiones');
    controles.setAttribute('role', 'tablist');
    controles.setAttribute('aria-label', 'Ver la composición por');
    var lienzo = crear('div', 'grafico-lienzo');

    function pintar(dim) {
      lienzo.innerHTML = '';
      var datos = agrupar(filas, dim, dim.orden);
      var total = datos.reduce(function (t, d) { return t + d[1]; }, 0);
      lienzo.appendChild(dibujarAnillo(datos, total));
      lienzo.appendChild(dibujarLeyenda(datos, total));
    }

    dimensiones.forEach(function (dim, i) {
      var b = crear('button', 'chip-dim', dim.rotulo);
      b.type = 'button';
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', i === 0 ? 'true' : 'false');
      b.addEventListener('click', function () {
        controles.querySelectorAll('.chip-dim').forEach(function (o) {
          o.setAttribute('aria-selected', o === b ? 'true' : 'false');
        });
        pintar(dim);
      });
      controles.appendChild(b);
    });

    cuerpo.appendChild(controles);
    cuerpo.appendChild(lienzo);
    sec.appendChild(cuerpo);
    cont.appendChild(sec);
    pintar(dimensiones[0]);
  }

  /* -------------------------------------------------- detalle de unidades

     Una repartición grande tiene decenas de combinaciones de marca, modelo y
     año: en una tabla plana eso son cuarenta y pico de filas ilegibles. Acá
     se ve una línea por marca, con su barra de proporción, y cada una se
     abre para mostrar sus modelos. Arranca cerrado. */
  var SEP = '\u0001';   // separador que no puede aparecer dentro de un dato

  function cuadroDetalle(cont, filas, total) {
    var marcas = {};
    filas.forEach(function (f) {
      var marca = celda(f, 'marca');
      var modelo = celda(f, 'modelo');
      var anio = celda(f, 'anio');
      marca = esVacio(marca) ? 'Sin marca' : marca;
      if (!marcas[marca]) { marcas[marca] = { total: 0, modelos: {} }; }
      marcas[marca].total++;
      var k = (esVacio(modelo) ? SIN_DATO : modelo) + SEP + (esVacio(anio) ? SIN_DATO : anio);
      marcas[marca].modelos[k] = (marcas[marca].modelos[k] || 0) + 1;
    });

    var lista = Object.keys(marcas)
      .map(function (m) { return { marca: m, datos: marcas[m] }; })
      .sort(function (a, b) { return b.datos.total - a.datos.total; });

    // Modelos distintos, no combinaciones: un mismo modelo en tres años es
    // un modelo, aunque adentro se abra en tres líneas.
    var modelosDistintos = lista.reduce(function (t, x) {
      var vistos = {};
      Object.keys(x.datos.modelos).forEach(function (k) { vistos[k.split(SEP)[0]] = 1; });
      return t + Object.keys(vistos).length;
    }, 0);

    var sec = crear('section', 'cuadro');
    sec.appendChild(crear('h3', 'titulo-cuadro', 'Detalle de unidades'));

    var caja = crear('div', 'detalle');

    var barra = crear('div', 'detalle-barra');
    barra.appendChild(crear('span', 'detalle-resumen',
      miles(lista.length) + (lista.length === 1 ? ' marca' : ' marcas') + '  ·  ' +
      miles(modelosDistintos) + (modelosDistintos === 1 ? ' modelo' : ' modelos')));
    var alternar = crear('button', 'chip-dim', 'Abrir todo');
    alternar.type = 'button';
    barra.appendChild(alternar);
    caja.appendChild(barra);

    // Con muchas marcas la lista se vuelve infinita: se muestran las
    // principales y el resto queda a un clic. Nunca se oculta en silencio.
    var TOPE = 12;
    var ocultas = [];

    var filasMarca = [];
    lista.forEach(function (item, orden) {
      var pct = item.datos.total / total * 100;

      var cabeza = crear('button', 'marca-fila');
      cabeza.type = 'button';
      cabeza.setAttribute('aria-expanded', 'false');

      // La proporción va de fondo de la fila, no en un renglón aparte: así
      // cada marca ocupa una sola línea y quince entran de un vistazo.
      var relleno = crear('span', 'marca-relleno');
      relleno.style.width = Math.max(pct, 1.2) + '%';
      cabeza.appendChild(relleno);

      cabeza.appendChild(crear('span', 'marca-flecha', '›'));
      cabeza.appendChild(crear('span', 'marca-nombre', item.marca));
      cabeza.appendChild(crear('span', 'marca-cant', miles(item.datos.total)));
      cabeza.appendChild(crear('span', 'marca-pct', pct.toFixed(1) + '%'));

      var hijos = crear('div', 'marca-hijos');
      hijos.hidden = true;
      Object.keys(item.datos.modelos)
        .map(function (k) { return [k.split(SEP), item.datos.modelos[k]]; })
        .sort(function (a, b) { return b[1] - a[1]; })
        .forEach(function (m) {
          var h = crear('div', 'hijo');
          h.appendChild(crear('span', 'hijo-modelo', m[0][0]));
          h.appendChild(crear('span', 'hijo-anio', m[0][1]));
          h.appendChild(crear('span', 'hijo-cant', miles(m[1])));
          hijos.appendChild(h);
        });

      cabeza.addEventListener('click', function () {
        var abierto = cabeza.getAttribute('aria-expanded') === 'true';
        cabeza.setAttribute('aria-expanded', abierto ? 'false' : 'true');
        hijos.hidden = abierto;
      });

      filasMarca.push({ cabeza: cabeza, hijos: hijos });
      caja.appendChild(cabeza);
      caja.appendChild(hijos);

      if (orden >= TOPE) {
        cabeza.hidden = true;
        ocultas.push(cabeza);
      }
    });

    var ver = null;
    function revelarOcultas() {
      ocultas.forEach(function (c) { c.hidden = false; });
      ocultas = [];
      if (ver) { ver.remove(); ver = null; }
    }

    if (ocultas.length) {
      ver = crear('button', 'ver-mas',
        'Ver las ' + miles(ocultas.length) + ' marcas restantes');
      ver.type = 'button';
      ver.addEventListener('click', revelarOcultas);
      caja.appendChild(ver);
    }

    alternar.addEventListener('click', function () {
      var abrir = alternar.textContent === 'Abrir todo';
      // Abrir todo revela también las marcas que estaban bajo el tope: si no,
      // aparecerían modelos sueltos sin la marca que los encabeza.
      if (abrir) { revelarOcultas(); }
      filasMarca.forEach(function (f) {
        f.cabeza.setAttribute('aria-expanded', abrir ? 'true' : 'false');
        f.hijos.hidden = !abrir;
      });
      alternar.textContent = abrir ? 'Cerrar todo' : 'Abrir todo';
    });

    sec.appendChild(caja);
    sec.appendChild(crear('p', 'nota', 'Tocá una marca para ver sus modelos y años.'));
    cont.appendChild(sec);
  }

  function kpi(cifra, rotulo, clase) {
    var c = crear('div', 'kpi' + (clase ? ' ' + clase : ''));
    c.appendChild(crear('div', 'cifra', miles(cifra)));
    c.appendChild(crear('div', 'rotulo', rotulo));
    return c;
  }

  /* Arma el juego completo de indicadores y cuadros sobre el conjunto de
     unidades que reciba. Lo usan tanto el resumen general como el de un
     ministerio o una repartición: los números salen del mismo cálculo.

     config.desglose : {clave, titulo, rotulo} del último cuadro
     config.general  : true en la vista de toda la flota */
  function armarResumen(cont, filas, config) {
    config = config || {};
    cont.innerHTML = '';

    var total = filas.length;
    if (!total) {
      cont.appendChild(crear('p', 'nota', 'No hay unidades para mostrar.'));
      return;
    }

    // En una vista filtrada no tiene sentido listar tipos sin ninguna unidad
    var tipos = config.general ? TIPOS : TIPOS.filter(function (t) {
      return contar(filas, function (f) { return esTipo(f, t); }) > 0;
    });

    // --- Indicadores
    var kpis = crear('div', 'kpis');
    kpis.appendChild(kpi(total, config.general ? 'Flota automotor' : 'Unidades asignadas', 'protagonista'));
    kpis.appendChild(kpi(contar(filas, function (f) { return marcada(f, 'poliza'); }),
      'Póliza de Seguros DGGFA informada al 30-3-26'));
    kpis.appendChild(kpi(contar(filas, function (f) { return marcada(f, 'anticuacion'); }),
      'Anticuación ' + D.periodo + ' (activos)'));
    kpis.appendChild(kpi(contar(filas, function (f) { return marcada(f, 'patrimonial'); }),
      'Inst. Patrimonial sin cobertura', 'menta'));
    cont.appendChild(kpis);

    if (config.general) {
      cont.appendChild(crear('p', 'nota',
        'Referencia — Registro Nacional del Parque Automotor: 3.561 unidades. ' +
        'El total de flota incluye unidades contabilizadas doble por poseer el mismo servicio.'));
    }

    // --- Unidades por tipo
    var porTipo = tipos.map(function (tipo) {
      return [
        tipo,
        contar(filas, function (f) { return esTipo(f, tipo); }),
        contar(filas, function (f) { return esTipo(f, tipo) && marcada(f, 'poliza'); }),
        contar(filas, function (f) { return esTipo(f, tipo) && marcada(f, 'anticuacion'); }),
        contar(filas, function (f) { return esTipo(f, tipo) && marcada(f, 'patrimonial'); })
      ];
    });
    // Si la DGGFA ajustó datos respecto de la planilla, el cuadro lo dice en
    // vez de mostrar un número sin explicación.
    var ajustadas = contar(filas, function (f) {
      return !!(D.correcciones || {})[celda(f, 'patente')];
    });
    cuadro(cont, 'Unidades por tipo de vehículo',
      ['Tipo', 'Total', 'Póliza DGGFA', 'Anticuación', 'Inst. Patrimonial'],
      porTipo,
      ['TOTAL'].concat(sumaColumnas(porTipo).slice(1)),
      { principal: true,
        nota: ajustadas
          ? 'Incluye ' + miles(ajustadas) + ' unidades con datos ajustados por la DGGFA ' +
            'respecto de la planilla de origen. El motivo figura en la ficha de cada unidad.'
          : null });

    // --- Servicios por tipo
    var porServicio = tipos.map(function (tipo) {
      return [
        tipo,
        contar(filas, function (f) { return esTipo(f, tipo) && activo(f, 'combustible'); }),
        contar(filas, function (f) { return esTipo(f, tipo) && activo(f, 'telemetria'); }),
        contar(filas, function (f) { return esTipo(f, tipo) && activo(f, 'mantenimiento'); }),
        contar(filas, function (f) {
          return esTipo(f, tipo) && activo(f, 'combustible') && activo(f, 'telemetria') && activo(f, 'mantenimiento');
        })
      ];
    });
    var electricas = contar(filas, function (f) { return celda(f, 'combustible').toUpperCase() === 'ELÉCTRICO'; });
    cuadro(cont, 'Servicios por tipo de vehículo',
      ['Tipo', 'Combustible', 'Telemetría', 'Mantenimiento', 'Los 3 servicios'],
      porServicio,
      ['TOTAL'].concat(sumaColumnas(porServicio).slice(1)),
      { nota: electricas
          ? 'Combustible incluye las ' + miles(electricas) + ' unidades con carga eléctrica asignada.'
          : null });

    // --- Kilometraje
    var porKm = tipos.map(function (tipo) {
      var a = contar(filas, function (f) { return esTipo(f, tipo) && enRango(f, -1, 50000); });
      var b = contar(filas, function (f) { return esTipo(f, tipo) && enRango(f, 50000, 100000); });
      var c = contar(filas, function (f) { return esTipo(f, tipo) && enRango(f, 100000, Infinity); });
      return [tipo, a, b, c, a + b + c];
    });
    cuadro(cont, 'Kilometraje',
      ['Tipo', 'Hasta 50.000 km', 'De 50.000 a 100.000', 'Más de 100.000', 'Total'],
      porKm,
      ['TOTAL'].concat(sumaColumnas(porKm).slice(1)),
      { nota: 'Solo se cuentan las unidades con kilometraje cargado (' +
        miles(contar(filas, function (f) { return aNumero(celda(f, 'km')) !== null; })) +
        ' de ' + miles(total) + ').' });

    // --- Composición: un anillo con la dimensión que se elija
    var dimensiones = [
      { clave: 'tipo', rotulo: 'Tipo' },
      { clave: 'marca', rotulo: 'Marca' },
      { fn: antiguedad, rotulo: 'Antigüedad', orden: TRAMOS_ANTIGUEDAD },
      { clave: 'tipoCombustible', rotulo: 'Combustible' }
    ];
    if (config.desglose) {
      dimensiones.splice(2, 0, { clave: config.desglose.clave, rotulo: config.desglose.rotulo });
    }
    cuadroGrafico(cont, filas, dimensiones);

    // --- Desglose: por ministerio en la vista general, por repartición
    //     dentro de un ministerio. Una repartición ya no tiene nivel
    //     siguiente, así que ahí se pasa directo al detalle.
    if (config.detalle) {
      cuadroDetalle(cont, filas, total);
      return;
    }
    var desglose = config.desglose || { clave: 'ministerio', titulo: 'Unidades asignadas por Ministerio', rotulo: 'Ministerio' };

    var cuenta = {};
    filas.forEach(function (f) {
      var v = celda(f, desglose.clave);
      if (esVacio(v)) { return; }
      cuenta[v] = (cuenta[v] || 0) + 1;
    });
    var porGrupo = Object.keys(cuenta)
      .map(function (v) { return [v, cuenta[v]]; })
      .sort(function (a, b) { return b[1] - a[1]; });

    if (porGrupo.length) {
      var asignadas = sumaColumnas(porGrupo)[1];
      cuadro(cont, desglose.titulo,
        [desglose.rotulo, 'Unidades'],
        porGrupo,
        ['TOTAL'].concat([asignadas]),
        { nota: total - asignadas
            ? 'No incluye ' + miles(total - asignadas) + ' unidades sin ' +
              desglose.rotulo.toLowerCase() + ' asignado en la base.'
            : null });
    }

    // El detalle cierra siempre: es el nivel más fino de todas las vistas.
    cuadroDetalle(cont, filas, total);
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

  armarResumen(document.getElementById('resumen-cuerpo'), D.filas, { general: true });

  // El hash puede traer un dominio (#AB766UL) o un organismo
  // (#ministerio=MHFGC%20-%20Hacienda%20y%20Finanzas)
  var inicial = window.location.hash.replace('#', '');
  var corte = inicial.indexOf('=');
  var claveHash = corte > 0 ? inicial.slice(0, corte) : '';
  var valorHash = corte > 0 ? decodeURIComponent(inicial.slice(corte + 1)) : decodeURIComponent(inicial);

  var grupoInicial = null;
  GRUPOS.forEach(function (g) {
    if (g.clave === claveHash && filasDe(g.clave, valorHash).length) { grupoInicial = g; }
  });

  if (grupoInicial) {
    input.value = valorHash;
    mostrarGrupo(grupoInicial, valorHash);
  } else if (valorHash && porPatente.has(normalizar(valorHash))) {
    input.value = valorHash.toUpperCase();
    buscar();
  } else {
    estadoInicial();
  }
})();
