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

  /* Los niveles intermedios existen solo si jerarquia.xlsx está cargado.
     Sin él, un ministerio se desglosa directo en reparticiones y nada de
     esto aparece. */
  var HAY_JERARQUIA = D.claves.indexOf('secretaria') !== -1;

  // No todas las reparticiones cuelgan de una secretaría: en muchos
  // ministerios el nivel siguiente es directamente una subsecretaría. Esto
  // devuelve el escalón que corresponda en cada caso.
  function dependencia(fila) {
    var sec = celda(fila, 'secretaria');
    if (!esVacio(sec)) { return sec; }
    var sub = celda(fila, 'subsecretaria');
    if (!esVacio(sub)) { return sub; }
    return 'Depende directo del Ministerio';
  }

  // Cada desglose define qué se abre adentro: el escalón siguiente de la
  // cadena. El último nivel se abre en las marcas de sus unidades.
  var SUB_REPARTICION = { clave: 'reparticion', rotulo: 'Repartición', plural: 'reparticiones', vacio: 'Sin repartición' };
  var SUB_MARCA = { clave: 'marca', rotulo: 'Marca', plural: 'marcas', vacio: 'Sin marca' };
  var SUB_DEPENDENCIA = { fn: dependencia, rotulo: 'Secretaría / Subsecretaría', plural: 'secretarías o subsecretarías' };

  var DESGLOSE_REPARTICION = { clave: 'reparticion', titulo: 'Unidades por Repartición', rotulo: 'Repartición', plural: 'reparticiones', hijos: [SUB_MARCA] };
  var DESGLOSE_DEPENDENCIA = { fn: dependencia, titulo: 'Unidades por Secretaría / Subsecretaría', rotulo: 'Secretaría / Subsecretaría', plural: 'secretarías o subsecretarías', hijos: [SUB_REPARTICION] };
  var DESGLOSE_SUBSECRETARIA = { clave: 'subsecretaria', titulo: 'Unidades por Subsecretaría', rotulo: 'Subsecretaría', plural: 'subsecretarías', hijos: [SUB_REPARTICION] };

  // Cada nivel de la cadena es buscable y se desglosa en el siguiente:
  // ministerio > secretaría > subsecretaría > repartición > unidad.
  var GRUPOS = [
    { clave: 'ministerio', etiqueta: 'Ministerio',
      desglose: HAY_JERARQUIA ? DESGLOSE_DEPENDENCIA : DESGLOSE_REPARTICION }
  ];

  if (HAY_JERARQUIA) {
    GRUPOS.push({ clave: 'secretaria', etiqueta: 'Secretaría', desglose: DESGLOSE_SUBSECRETARIA });
    GRUPOS.push({ clave: 'subsecretaria', etiqueta: 'Subsecretaría', desglose: DESGLOSE_REPARTICION });
    GRUPOS.push({ clave: 'funcionario', etiqueta: 'Funcionario', desglose: DESGLOSE_REPARTICION });
  }

  // En una repartición no hay nivel siguiente: en vez de un desglose va el
  // detalle de las unidades que la componen.
  GRUPOS.push({ clave: 'reparticion', etiqueta: 'Repartición', detalle: true });

  // Rango + nombre, como lo anota el organigrama del GCBA
  // ("Ministro Gustavo Arengo Piragine").
  function responsableDe(filas) {
    if (!HAY_JERARQUIA || !filas.length) { return null; }
    var nombre = celda(filas[0], 'funcionario');
    var rango = celda(filas[0], 'rango');
    if (esVacio(nombre)) { return null; }
    // Solo se muestra si todo el conjunto responde a la misma persona.
    for (var i = 1; i < filas.length; i++) {
      if (celda(filas[i], 'funcionario') !== nombre) { return null; }
    }
    return esVacio(rango) ? nombre : rango + '  ·  ' + nombre;
  }

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
      { clave: 'secretaria', rotulo: 'Secretaría', soloConJerarquia: true },
      { clave: 'subsecretaria', rotulo: 'Subsecretaría', soloConJerarquia: true },
      { clave: 'reparticion', rotulo: 'Repartición' },
      { clave: 'responsable', rotulo: 'A cargo de', formato: 'responsable', soloConJerarquia: true },
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

      case 'responsable':
        var quien = celda(fila, 'funcionario');
        if (esVacio(quien)) { return { texto: SIN_DATO, clase: 'vacio' }; }
        var cargo = celda(fila, 'rango');
        return { texto: quien, nota: esVacio(cargo) ? null : cargo };

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
        // Los campos de organigrama solo existen con jerarquia.xlsx cargado
        if (campo.soloConJerarquia && !HAY_JERARQUIA) { return; }
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

    // Quién está a cargo, cuando todo el conjunto responde a la misma
    // persona. En la vista de un funcionario el nombre ya es el título, así
    // que ahí alcanza con su rango.
    if (grupo.clave === 'funcionario') {
      var rango = celda(filas[0], 'rango');
      var mismo = filas.every(function (f) { return celda(f, 'rango') === rango; });
      if (mismo && !esVacio(rango)) {
        cab.appendChild(crear('div', 'responsable-grupo', rango));
      }
    } else {
      var quien = responsableDe(filas);
      if (quien) { cab.appendChild(crear('div', 'responsable-grupo', quien)); }
    }
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

  // Niveles reutilizables para el detalle que se abre al tocar una porción
  var NIVEL_MARCA = { clave: 'marca', vacio: 'Sin marca', plural: 'marcas' };
  var NIVEL_MODELO = {
    fn: function (f) {
      var modelo = celda(f, 'modelo');
      var anio = celda(f, 'anio');
      return (esVacio(modelo) ? SIN_DATO : modelo) + SEP + (esVacio(anio) ? SIN_DATO : anio);
    },
    etiqueta: function (v) { return v.split(SEP)[0]; },
    apostillaTexto: function (v) { return v.split(SEP)[1]; },
    plural: 'modelos'
  };
  var NIVEL_DETALLE_UNIDAD = { clave: 'detalleUnidad', vacio: 'Sin detalle de unidad', plural: 'detalles' };
  var NIVEL_TIPO = { clave: 'tipo', vacio: 'Sin tipo', plural: 'tipos' };
  var NIVEL_REPARTICION = { clave: 'reparticion', vacio: 'Sin repartición', plural: 'reparticiones' };

  /* Antigüedad por tramos. Los cortes salen del año del período del archivo,
     no de una constante: en 2027 se corren solos. */
  var ANIO_BASE = parseInt((D.periodo.match(/(20\d\d)/) || [])[1], 10) || 2026;

  var TRAMOS_ANTIGUEDAD = [
    'Hasta 3 años (' + ANIO_BASE + ' a ' + (ANIO_BASE - 2) + ')',
    'De 3 a 5 años (' + (ANIO_BASE - 3) + ' a ' + (ANIO_BASE - 5) + ')',
    'De 6 a 10 años (' + (ANIO_BASE - 6) + ' a ' + (ANIO_BASE - 10) + ')',
    'Más de 10 años (' + (ANIO_BASE - 11) + ' inclusive)'
  ];

  // Devuelve vacío cuando no hay año, para que quede fuera del anillo junto
  // al resto de los datos faltantes en vez de ser un tramo más.
  function antiguedad(fila) {
    var a = parseInt(celda(fila, 'anio'), 10);
    if (!a) { return ''; }
    if (a >= ANIO_BASE - 2) { return TRAMOS_ANTIGUEDAD[0]; }
    if (a >= ANIO_BASE - 5) { return TRAMOS_ANTIGUEDAD[1]; }
    if (a >= ANIO_BASE - 10) { return TRAMOS_ANTIGUEDAD[2]; }
    return TRAMOS_ANTIGUEDAD[3];
  }

  // dim: {clave} para leer una columna, o {fn} para calcular la categoría.
  // orden: sin definir ordena por cantidad; un arreglo fija el orden.
  function agrupar(filas, dim, orden) {
    var cuenta = {};
    filas.forEach(function (f) {
      var v = dim.fn ? dim.fn(f) : celda(f, dim.clave);
      // Algunas dimensiones no admiten un "Sin dato": una unidad sin
      // ministerio no es un ministerio más, y con 1.468 casos se come el
      // gráfico y aplasta a los doce reales.
      if (esVacio(v)) {
        if (dim.omitirVacios) { return; }
        v = 'Sin dato';
      }
      cuenta[v] = (cuenta[v] || 0) + 1;
    });
    // El recorte va SIEMPRE por volumen: si se recortara por año, los cinco
    // valores mas viejos taparian al resto y "Otras" se comeria el grafico.
    // Cada porción recuerda qué valores la componen, para poder filtrar por
    // ella cuando se la toca — incluida "Otras", que son varios.
    var lista = Object.keys(cuenta)
      .map(function (k) { return { nombre: k, cantidad: cuenta[k], valores: [k] }; })
      .sort(function (a, b) { return b.cantidad - a.cantidad; });

    var resto = null;
    if (lista.length > MAX_SEGMENTOS) {
      var sobrantes = lista.slice(MAX_SEGMENTOS - 1);
      lista = lista.slice(0, MAX_SEGMENTOS - 1);
      resto = {
        nombre: 'Otras (' + sobrantes.length + ')',
        cantidad: sobrantes.reduce(function (t, x) { return t + x.cantidad; }, 0),
        valores: sobrantes.map(function (x) { return x.nombre; })
      };
    }

    // Recién ahora, si la dimensión tiene un orden propio, se aplica sobre lo
    // que quedó visible: los tramos de antigüedad van del más nuevo al más
    // viejo, no del más numeroso al menos.
    if (orden && orden.length) {
      lista.sort(function (a, b) { return orden.indexOf(a.nombre) - orden.indexOf(b.nombre); });
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

  function dibujarAnillo(datos, total, alElegir) {
    var fig = crear('div', 'anillo');
    var svg = svgEl('svg', { viewBox: '0 0 240 240', role: 'img' });
    svg.setAttribute('aria-label', 'Distribución en anillo; los valores están en la leyenda y en el cuadro de detalle.');

    var angulo = 0;
    datos.forEach(function (d, i) {
      var porcion = d.cantidad / total * 360;
      // 2px de separación entre porciones, que es el fondo asomando
      var recorte = porcion > 6 ? 1.2 : 0;
      var path = svgEl('path', {
        d: arco(120, 120, 100, 64, angulo + recorte, angulo + porcion - recorte),
        fill: PALETA[i % PALETA.length],
        'data-i': i
      });
      path.appendChild(svgEl('title', {})).textContent =
        d.nombre + ': ' + miles(d.cantidad) + ' (' + (d.cantidad / total * 100).toFixed(1) + '%)';
      if (alElegir) {
        path.setAttribute('class', 'porcion');
        path.addEventListener('click', function () { alElegir(i); });
      }
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

  function dibujarLeyenda(datos, total, alElegir) {
    var ul = crear('ul', 'leyenda');
    datos.forEach(function (d, i) {
      var li = crear('li');
      if (alElegir) {
        li.setAttribute('role', 'button');
        li.setAttribute('tabindex', '0');
        li.addEventListener('click', function () { alElegir(i); });
        li.addEventListener('keydown', function (e) {
          if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); alElegir(i); }
        });
      }
      li.appendChild(svgEl('svg', { class: 'muestra', viewBox: '0 0 10 10', 'aria-hidden': 'true' }))
        .appendChild(svgEl('rect', { x: 0, y: 0, width: 10, height: 10, rx: 3, fill: PALETA[i % PALETA.length] }));
      li.appendChild(crear('span', 'leyenda-nombre', d.nombre));
      li.appendChild(crear('span', 'leyenda-valor', miles(d.cantidad)));
      li.appendChild(crear('span', 'leyenda-pct', (d.cantidad / total * 100).toFixed(1) + '%'));
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

    var zoom = crear('div', 'zoom-porcion');

    function pintar(dim) {
      lienzo.innerHTML = '';
      zoom.innerHTML = '';
      var datos = agrupar(filas, dim, dim.orden);
      var total = datos.reduce(function (t, d) { return t + d.cantidad; }, 0);

      // Tocar una porción abre debajo el detalle de esas unidades. Vuelve a
      // tocarse para cerrarla, y elegir otra la reemplaza.
      var elegida = -1;
      function elegir(i) {
        var mismo = elegida === i;
        elegida = mismo ? -1 : i;
        lienzo.querySelectorAll('.porcion').forEach(function (p, k) {
          p.classList.toggle('apagada', elegida >= 0 && k !== elegida);
        });
        lienzo.querySelectorAll('.leyenda li').forEach(function (l, k) {
          l.classList.toggle('elegida', k === elegida);
        });
        zoom.innerHTML = '';
        if (elegida < 0) { return; }

        var d = datos[elegida];
        var dentro = filas.filter(function (f) {
          var v = dim.fn ? dim.fn(f) : celda(f, dim.clave);
          return d.valores.indexOf(esVacio(v) ? 'Sin dato' : v) !== -1;
        });
        cuadroPlegable(zoom, {
          titulo: d.nombre,
          filas: dentro,
          niveles: dim.detalle || [NIVEL_MARCA, NIVEL_MODELO],
          resumen: function (c) {
            var pct = (d.cantidad / total * 100).toFixed(1).replace('.', ',');
            return miles(dentro.length) + (dentro.length === 1 ? ' unidad' : ' unidades') +
              '  ·  ' + pct + ' % del gráfico';
          },
          masTexto: 'más',
          nota: 'Detalle de la porción elegida. Tocala de nuevo en el gráfico para cerrarla.'
        });
        var cerrar = crear('button', 'cerrar-zoom', 'Cerrar');
        cerrar.type = 'button';
        cerrar.addEventListener('click', function () { elegir(elegida); });
        zoom.appendChild(cerrar);
      }

      lienzo.appendChild(dibujarAnillo(datos, total, elegir));

      var panel = crear('div', 'panel-leyenda');
      panel.appendChild(dibujarLeyenda(datos, total, elegir));

      /* Lo que no tiene el dato no entra al anillo, pero tampoco se esconde:
         va debajo de la leyenda, separado y en gris. Gris y no un color de
         la paleta, porque no es una categoría más: es la ausencia del dato
         con el que se está cortando. */
      var fuera = filas.length - total;
      if (fuera > 0) {
        var falta = crear('div', 'sin-dato');
        falta.appendChild(crear('span', 'sin-dato-marca'));
        falta.appendChild(crear('span', 'sin-dato-nombre', dim.sinDato || 'Sin dato'));
        falta.appendChild(crear('span', 'sin-dato-valor', miles(fuera)));
        falta.appendChild(crear('span', 'sin-dato-pct',
          (fuera / filas.length * 100).toFixed(1) + '%'));
        panel.appendChild(falta);
        panel.appendChild(crear('p', 'sin-dato-pie',
          'Queda fuera del gráfico: el anillo compara las ' + miles(total) +
          ' unidades que sí tienen el dato.'));
      }

      lienzo.appendChild(panel);
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
    cuerpo.appendChild(zoom);
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

  /* Lista plegable genérica: una línea por grupo, con la proporción de fondo,
     y adentro el nivel siguiente. La usan el desglose por organismo y el
     detalle de unidades, que son el mismo gesto sobre datos distintos.

     padre : {fn|clave, vacio}          qué encabeza cada línea
     hijo  : {fn}  -> [texto, apostilla]  qué aparece al abrirla (opcional)
     resumen(nPadres, nHijos) -> texto de la barra superior              */
  /* Lista plegable de N niveles: una línea por grupo con la proporción de
     fondo, y adentro el nivel siguiente, hasta llegar a la hoja. La usan el
     desglose por organismo y el detalle de unidades, que son el mismo gesto
     sobre datos distintos.

     niveles : [{clave|fn, vacio, plural, apostilla}] de mayor a menor.
               El último se dibuja como línea simple, los demás se abren.
     resumen(conteos[]) -> texto de la barra superior                      */
  function cuadroPlegable(cont, opciones) {
    var niveles = opciones.niveles;
    var TOPE = 12;

    function valorDe(fila, nivel) {
      var v = nivel.fn ? nivel.fn(fila) : celda(fila, nivel.clave);
      return esVacio(v) ? (nivel.vacio || SIN_DATO) : v;
    }

    function agruparPor(filas, nivel) {
      var mapa = {};
      var orden = [];
      filas.forEach(function (f) {
        var v = valorDe(f, nivel);
        if (!mapa[v]) { mapa[v] = []; orden.push(v); }
        mapa[v].push(f);
      });
      return orden
        .map(function (v) { return { nombre: v, filas: mapa[v] }; })
        .sort(function (a, b) { return b.filas.length - a.filas.length; });
    }

    // Cuántos valores distintos hay en cada nivel, para la barra de arriba.
    // Si la hoja combina dos datos, cuenta por el principal: un mismo modelo
    // en tres años es un modelo, aunque se abra en tres líneas.
    var conteos = niveles.map(function (nivel) {
      var vistos = {};
      opciones.filas.forEach(function (f) {
        var v = valorDe(f, nivel);
        vistos[nivel.etiqueta ? nivel.etiqueta(v) : v] = 1;
      });
      return Object.keys(vistos).length;
    });

    var plegables = [];   // todo lo que "Abrir todo" debe alcanzar

    // Dibuja un nivel dentro de un contenedor y se llama a sí misma para el
    // siguiente, hasta la hoja.
    function dibujar(caja, filas, profundidad, totalReferencia, alOcultar) {
      var nivel = niveles[profundidad];
      var esHoja = profundidad === niveles.length - 1;

      agruparPor(filas, nivel).forEach(function (grupo, orden) {
        var cantidad = grupo.filas.length;
        var pct = cantidad / totalReferencia * 100;

        if (esHoja) {
          var hoja = crear('div', 'hijo nivel-' + profundidad);
          hoja.appendChild(crear('span', 'hijo-modelo',
            nivel.etiqueta ? nivel.etiqueta(grupo.nombre) : grupo.nombre));
          hoja.appendChild(crear('span', 'hijo-anio',
            nivel.apostillaTexto ? nivel.apostillaTexto(grupo.nombre) : ''));
          hoja.appendChild(crear('span', 'hijo-cant', miles(cantidad)));
          caja.appendChild(hoja);
          return;
        }

        var cabeza = crear('button', 'marca-fila nivel-' + profundidad);
        cabeza.type = 'button';
        cabeza.setAttribute('aria-expanded', 'false');

        // La proporción va de fondo de la fila, no en un renglón aparte
        var relleno = crear('span', 'marca-relleno');
        relleno.style.width = Math.max(pct, 1.2) + '%';
        cabeza.appendChild(relleno);
        cabeza.appendChild(crear('span', 'marca-flecha', '›'));
        cabeza.appendChild(crear('span', 'marca-nombre', grupo.nombre));
        cabeza.appendChild(crear('span', 'marca-cant', miles(cantidad)));
        cabeza.appendChild(crear('span', 'marca-pct', pct.toFixed(1) + '%'));

        var hijos = crear('div', 'marca-hijos');
        hijos.hidden = true;
        dibujar(hijos, grupo.filas, profundidad + 1, cantidad, null);

        cabeza.addEventListener('click', function () {
          var abierto = cabeza.getAttribute('aria-expanded') === 'true';
          cabeza.setAttribute('aria-expanded', abierto ? 'false' : 'true');
          hijos.hidden = abierto;
        });

        plegables.push({ cabeza: cabeza, hijos: hijos });
        caja.appendChild(cabeza);
        caja.appendChild(hijos);

        // El tope solo aplica al primer nivel: adentro no hay tantas líneas
        if (profundidad === 0 && orden >= TOPE && alOcultar) { alOcultar(cabeza); }
      });
    }

    if (!opciones.filas.length) { return; }

    var sec = crear('section', 'cuadro');
    sec.appendChild(crear('h3', 'titulo-cuadro', opciones.titulo));

    var caja = crear('div', 'detalle');
    var barra = crear('div', 'detalle-barra');
    barra.appendChild(crear('span', 'detalle-resumen', opciones.resumen(conteos)));
    var alternar = crear('button', 'chip-dim', 'Abrir todo');
    alternar.type = 'button';
    barra.appendChild(alternar);
    caja.appendChild(barra);

    var ocultas = [];
    dibujar(caja, opciones.filas, 0, opciones.filas.length, function (el) {
      el.hidden = true;
      ocultas.push(el);
    });

    var ver = null;
    function revelarOcultas() {
      ocultas.forEach(function (c) { c.hidden = false; });
      ocultas = [];
      if (ver) { ver.remove(); ver = null; }
    }

    if (ocultas.length) {
      ver = crear('button', 'ver-mas', 'Ver ' + miles(ocultas.length) + ' ' + (opciones.masTexto || 'más'));
      ver.type = 'button';
      ver.addEventListener('click', revelarOcultas);
      caja.appendChild(ver);
    }

    alternar.addEventListener('click', function () {
      var abrir = alternar.textContent === 'Abrir todo';
      // Abrir todo revela también lo que estaba bajo el tope: si no,
      // aparecerían hijos sueltos sin la línea que los encabeza.
      if (abrir) { revelarOcultas(); }
      plegables.forEach(function (f) {
        f.cabeza.setAttribute('aria-expanded', abrir ? 'true' : 'false');
        f.hijos.hidden = !abrir;
      });
      alternar.textContent = abrir ? 'Cerrar todo' : 'Abrir todo';
    });

    sec.appendChild(caja);
    if (opciones.nota) { sec.appendChild(crear('p', 'nota', opciones.nota)); }
    cont.appendChild(sec);
  }

  function cuadroDetalle(cont, filas, total) {
    cuadroPlegable(cont, {
      titulo: 'Detalle de unidades',
      filas: filas,
      niveles: [
        { clave: 'marca', vacio: 'Sin marca' },
        // La hoja combina modelo y año: son el mismo dato de la unidad
        { fn: function (f) {
            var modelo = celda(f, 'modelo');
            var anio = celda(f, 'anio');
            return (esVacio(modelo) ? SIN_DATO : modelo) + SEP + (esVacio(anio) ? SIN_DATO : anio);
          },
          etiqueta: function (v) { return v.split(SEP)[0]; },
          apostillaTexto: function (v) { return v.split(SEP)[1]; } }
      ],
      resumen: function (c) {
        return miles(c[0]) + (c[0] === 1 ? ' marca' : ' marcas') + '  ·  ' +
               miles(c[1]) + (c[1] === 1 ? ' modelo' : ' modelos');
      },
      masTexto: 'marcas más',
      nota: 'Tocá una marca para ver sus modelos y años.'
    });
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
    // Todas omiten los faltantes: el anillo compara lo que hay, y lo que
    // falta se muestra aparte, debajo de la leyenda.
    // 'detalle' es la cadena que se abre al tocar una porción: el corte que
    // mejor explica ese grupo, no siempre el mismo.
    var dimensiones = [
      { clave: 'tipo', rotulo: 'Tipo', omitirVacios: true, sinDato: 'Sin tipo registrado',
        detalle: [NIVEL_DETALLE_UNIDAD, NIVEL_MARCA, NIVEL_MODELO] },
      { clave: 'marca', rotulo: 'Marca', omitirVacios: true, sinDato: 'Sin marca registrada',
        detalle: [NIVEL_MODELO] },
      { fn: antiguedad, rotulo: 'Antigüedad', orden: TRAMOS_ANTIGUEDAD, omitirVacios: true,
        sinDato: 'Sin año registrado', detalle: [NIVEL_TIPO, NIVEL_MARCA, NIVEL_MODELO] },
      { clave: 'tipoCombustible', rotulo: 'Combustible', omitirVacios: true,
        sinDato: 'Sin combustible asignado', detalle: [NIVEL_TIPO, NIVEL_MARCA] }
    ];
    /* Dimensión organizacional del anillo. En una vista de organismo es la
       repartición: la secretaría ya la desglosa el cuadro de abajo, y como
       ese desglose se calcula con una función y no con una columna, acá
       terminaba agrupando todo en "Sin dato". En la vista general la
       repartición no dice nada (son 248), así que va el ministerio. */
    if (config.general) {
      dimensiones.splice(2, 0, { clave: 'ministerio', rotulo: 'Ministerio',
        omitirVacios: true, sinDato: 'Sin ministerio asignado',
        detalle: [NIVEL_REPARTICION, NIVEL_TIPO, NIVEL_MARCA] });
    } else if (config.desglose) {
      dimensiones.splice(2, 0, { clave: 'reparticion', rotulo: 'Repartición',
        omitirVacios: true, sinDato: 'Sin repartición asignada',
        detalle: [NIVEL_TIPO, NIVEL_MARCA, NIVEL_MODELO] });
    }
    cuadroGrafico(cont, filas, dimensiones);

    // --- Desglose: por ministerio en la vista general, por repartición
    //     dentro de un ministerio. Una repartición ya no tiene nivel
    //     siguiente, así que ahí se pasa directo al detalle.
    if (config.detalle) {
      cuadroDetalle(cont, filas, total);
      return;
    }
    var desglose = config.desglose || {
      clave: 'ministerio',
      titulo: 'Unidades asignadas por Ministerio',
      rotulo: 'Ministerio',
      plural: 'ministerios',
      // En la vista general se baja el organigrama completo dentro del
      // cuadro: ministerio, secretaría y repartición.
      hijos: HAY_JERARQUIA ? [SUB_DEPENDENCIA, SUB_REPARTICION] : [SUB_REPARTICION]
    };

    // Cuántas unidades quedan fuera del desglose por no tener el dato
    var asignadas = contar(filas, function (f) {
      return !esVacio(desglose.fn ? desglose.fn(f) : celda(f, desglose.clave));
    });

    /* Cada línea se abre en el escalón siguiente, como una mamushka. En el
       resumen general eso baja dos veces: ministerio, secretaría y
       repartición, que es el organigrama completo dentro de un cuadro. */
    var cadena = [{ fn: desglose.fn, clave: desglose.clave, plural: desglose.plural, rotulo: desglose.rotulo }]
      .concat(desglose.hijos || []);

    cuadroPlegable(cont, {
      titulo: desglose.titulo,
      // El desglose corre sobre las que tienen el dato: las que no, ya se
      // cuentan en la nota y en su propio cuadro.
      filas: filas.filter(function (f) {
        return !esVacio(desglose.fn ? desglose.fn(f) : celda(f, desglose.clave));
      }),
      niveles: cadena,
      resumen: function (c) {
        var partes = [miles(asignadas) + (asignadas === 1 ? ' unidad' : ' unidades')];
        cadena.forEach(function (nivel, i) {
          var nombre = c[i] === 1 && nivel.rotulo ? nivel.rotulo.toLowerCase() : nivel.plural;
          partes.push(miles(c[i]) + ' ' + nombre);
        });
        return partes.join('  ·  ');
      },
      masTexto: 'más',
      nota: (total - asignadas
        ? 'No incluye ' + miles(total - asignadas) + ' unidades sin ' +
          desglose.rotulo.toLowerCase() + ' asignado en la base. '
        : '') + (cadena.length > 1 ? 'Tocá una línea para bajar al nivel siguiente.' : '')
    });

    /* Las unidades sin ministerio no son un ministerio más: casi ninguna
       tiene servicios ni anticuación, y las observaciones muestran que son
       bajas, subastas, paraderos desconocidos o flota informada por
       terceros. Van en su propio cuadro, agrupadas por tipo y abiertas en
       el motivo que trae la base. */
    if (config.general) {
      var huerfanas = filas.filter(function (f) { return esVacio(celda(f, 'ministerio')); });
      if (huerfanas.length) {
        var activas = contar(huerfanas, function (f) { return marcada(f, 'anticuacion'); });
        cuadroPlegable(cont, {
          titulo: 'Unidades sin ministerio asignado',
          filas: huerfanas,
          total: huerfanas.length,
          niveles: [
            { clave: 'tipo', vacio: 'Sin tipo' },
            { clave: 'observaciones', vacio: 'Sin observaciones en la base' }
          ],
          resumen: function (c) {
            return miles(huerfanas.length) + ' unidades  ·  ' + miles(c[0]) +
              (c[0] === 1 ? ' tipo' : ' tipos') + '  ·  ' + miles(c[1]) +
              (c[1] === 1 ? ' observación' : ' observaciones');
          },
          masTexto: 'tipos más',
          nota: 'Son las unidades que el anillo deja fuera del corte por ministerio. ' +
                'Solo ' + miles(activas) + ' figura' + (activas === 1 ? '' : 'n') +
                ' en la anticuación de ' + D.periodo + ', así que en su enorme mayoría no ' +
                'están operativas. Tocá un tipo para ver qué dice la base de cada una.'
        });
      }
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
