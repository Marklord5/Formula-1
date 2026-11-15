// Archivos que lee la página (carpeta data/). Para actualizar, reemplaza el archivo con el MISMO nombre.
// `kind` decide qué gráficas y totales se calculan para cada uno.
export const FILES = [
  { kind: 'final',   name: 'Corte Final',  url: 'data/corte-final.xlsx'  },
  { kind: 'daily',   name: 'Corte Diario', url: 'data/corte-diario.xlsx' },
  { kind: 'returns', name: 'Devolución',   url: 'data/devolucion.xlsx'   },
  { kind: 'hospitality', name: 'Corte Hospitalidades', url: 'data/corte-hospitalidades.xlsx' },
  { kind: 'ecos', name: 'Corte ECOS', url: 'data/corte-ecos.xlsx' }
];

// Reporte del sistema de acreditación (solicitudes, asignaciones y escaneos). Se ve arriba como panorama global.
export const GLOBAL = { name: 'Reporte Global', url: 'data/reporte-global.xlsx' };

// Control de entregas en paquete: solo se lee su hoja REPORTE (franja debajo de los totales); el Excel completo se descarga
export const PACKAGE = { name: 'Control Entregas Paquete', url: 'data/control-entregas-paquete.xlsx', sheet: /^REPORTE$/ };

// Qué significa cada columna (aparece al pasar el mouse por el encabezado de la tabla)
export const GLOSSARY = {
  'PRODUCIDO': 'Cuántas se imprimieron',
  'TOTAL PRODUCIDO': 'Cuántas se imprimieron',
  'ENTREGADO PAQUETE': 'Entregadas en paquete directo al solicitante',
  'ACREDITADO MODULO': 'Entregadas en los módulos de acreditación',
  'TOTAL ACREDITACIONES': 'Total entregado (paquete + módulo)',
  'SOBRANTE BODEGA': 'Lo que quedó sin salir de bodega',
  'SOBRANTE MODULO': 'Lo que regresó o quedó en los módulos',
  'TOTAL SOBRANTE': 'Todo lo que sobró (bodega + módulo)',
  'EXISTENCIA EN BODEGA': 'Lo que quedaba en bodega después de ese movimiento',
  'CANTIDAD': 'Número de acreditaciones del movimiento',
  'FOLIO INICIAL': 'Primer folio entregado',
  'FOLIO FINAL': 'Último folio entregado',
  'SOBRANTE ENTREGADO': 'Lo que sobró de ese genérico',
  'TOTAL ENTREGADO': 'Cuántas se entregaron',
  'RESTANTE': 'Lo que quedó sin entregar',
  'CANTIDAD SOLICITADA': 'Cuántas pidió el solicitante',
  'SOBRANTE REINTEGRO': 'Lo que se regresó después de entregar',
  'CANTIDAD DEVUELTA': 'Piezas que regresaron después del evento',
  'CANTIDAD ENTREGADA': 'Cuántos pases o candados se entregaron',
  'CANTIDAD PRODUCIDA': 'Cuántas se produjeron de esa categoría y día'
};

// Tipos de acreditación, como los trabaja la oficina. Una categoría u hoja cae en el primer tipo cuyo patrón
// coincida; todo lo demás es Operación (el último, sin patrón).
// Un tipo con `book` toma sus números de ese Excel en vez del Corte Final (Hospitalidades → hoja CORTE FINAL del
// Corte Hospitalidades); sus categorías del Corte Final se ignoran para no contarlas dos veces.
export const GROUPS = [
  { name: 'Montaje y Desmontaje', pattern: /MONTAJE/, rule: 'las que dicen MONTAJE en la columna CATEGORIA' },
  { name: 'Hospitalidades', pattern: /(^|STAFF )HOSPITALIDADES|BANORTE|SANTANDER|BANAMEX|TSB|EXPERIENCES|PLATINO|ZONA PLUS|SPORTS? LOUNGE/, book: 'hospitality', rule: 'las filas de su tabla, menos staff, talento e invitados zona verde' },
  { name: 'Operación', rule: 'el resto de la columna CATEGORIA (incluye staff, talento e invitados zona verde de los patrocinadores)' }
];

// Staff, talento e invitados zona verde de los patrocinadores: aunque digan BANORTE, SANTANDER, TSB… son Operación
export const OPERATIVE = /STAFF|TALENTO|INVITADOS? ZONA VERDE/;

export const TOP_HOSPITALITY = 12; // categorías del Corte Hospitalidades antes de "Ver todas"
export const TOP_PER_GROUP = 6; // categorías por tipo en la gráfica del Corte Final antes de "Ver todas"
