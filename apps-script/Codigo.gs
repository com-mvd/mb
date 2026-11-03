/**
 * Aplicación web que lee la planilla TAREAS SUPERVISION 2026.
 * Devuelve las filas violetas, azules y naranjas de la hoja de un día.
 *
 *   GET ?fecha=0311   → filas de la hoja 03/11
 *   GET               → filas de hoy (hora de Montevideo)
 *
 * Las hojas diarias se llaman DDMM ("0311"); las de agosto se llamaban
 * D+M o DD+M ("98" = 9/8, "128" = 12/8). Se aceptan también con barra.
 */

const COLORES = {
  '#b4a7d6': 'violeta',
  '#a4c2f4': 'azul',
  '#f9cb9c': 'naranja',
};
const ZONA = 'America/Montevideo';
// Proyecto independiente (no ligado a la planilla): desde la planilla el editor no abre.
const PLANILLA = '1H_6i-89kIRAc0GAVM42eNTC6uSLIiLSMdiKVzf44fbo';

function doGet(e) {
  const libro = SpreadsheetApp.openById(PLANILLA);
  const fechas = listarFechas_(libro);
  const pedida = (e && e.parameter && e.parameter.fecha) ||
    Utilities.formatDate(new Date(), ZONA, 'ddMM');

  const elegida = fechas.find(f => f.id === pedida);
  const resultado = { fechas: fechas, fecha: pedida, filas: [] };
  if (elegida) {
    resultado.filas = leerHoja_(libro.getSheetByName(elegida.hoja));
  } else {
    resultado.error = 'No hay hoja para ' + pedida.slice(0, 2) + '/' + pedida.slice(2);
  }
  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Hojas diarias como [{id:'DDMM', hoja:'nombre real'}], de la más nueva a la más vieja. */
function listarFechas_(libro) {
  const vistas = {};
  libro.getSheets().forEach(h => {
    const nombre = h.getName().trim();
    let d, m;
    let p = nombre.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (p) { d = +p[1]; m = +p[2]; }
    else if (/^\d{4}$/.test(nombre)) { d = +nombre.slice(0, 2); m = +nombre.slice(2); }
    else if (/^\d{3}$/.test(nombre)) { d = +nombre.slice(0, 2); m = +nombre.slice(2); }
    else if (/^\d{2}$/.test(nombre)) { d = +nombre[0]; m = +nombre[1]; }
    else return;
    if (d < 1 || d > 31 || m < 1 || m > 12) return;
    const id = String(d).padStart(2, '0') + String(m).padStart(2, '0');
    if (!vistas[id]) vistas[id] = { id: id, hoja: h.getName(), orden: m * 100 + d };
  });
  return Object.values(vistas)
    .sort((a, b) => b.orden - a.orden)
    .map(f => ({ id: f.id, hoja: f.hoja }));
}

function leerHoja_(hoja) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  const rango = hoja.getRange(2, 1, ultima - 1, 6);
  const valores = rango.getDisplayValues();
  const fondos = rango.getBackgrounds();

  const filas = [];
  let turno = '';
  valores.forEach((v, i) => {
    const a = v[0].trim();
    if (/^TURNO\s/i.test(a)) { turno = a.replace(/^TURNO\s+/i, ''); return; }
    const color = COLORES[(fondos[i][1] || '').toLowerCase()];
    if (!color || !v[1].trim()) return;
    filas.push({
      turno: turno,
      color: color,
      horario: a,
      tarea: v[1].trim(),
      detalle: v[2].trim(),
      supervisor: v[3].trim(),
      estado: v[4].trim(),
      observaciones: v[5].trim(),
    });
  });
  return filas;
}
