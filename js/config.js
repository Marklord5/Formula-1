// Archivos que lee la página (carpeta data/). Para actualizar, reemplaza el archivo con el MISMO nombre.
// `kind` decide qué gráficas y totales se calculan para cada uno.
export const FILES = [
  { kind: 'final',   name: 'Corte Final',  url: 'data/corte-final.xlsx'  },
  { kind: 'daily',   name: 'Corte Diario', url: 'data/corte-diario.xlsx' },
  { kind: 'returns', name: 'Devolución',   url: 'data/devolucion.xlsx'   }
];

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
  'SOBRANTE ENTREGADO': 'Lo que sobró de ese genérico'
};

// Tipos de acreditación, como los trabaja la oficina. Una categoría u hoja cae en el primer tipo cuyo patrón
// coincida; todo lo demás es Operación (el último, sin patrón).
// Hospitalidades sigue la hoja FOLIOS STAFF HOSPITALIDADES del Corte Final; bórralo para contarlo dentro de Operación.
export const GROUPS = [
  { name: 'Montaje y Desmontaje', pattern: /MONTAJE/ },
  { name: 'Hospitalidades', pattern: /(^|STAFF )HOSPITALIDADES|BANORTE|SANTANDER|BANAMEX|TSB|EXPERIENCES|PLATINO|ZONA PLUS|SPORTS? LOUNGE|INVITADO/ },
  { name: 'Operación' }
];

export const TOP_PER_GROUP = 6; // categorías por tipo en la gráfica del Corte Final antes de "Ver todas"
