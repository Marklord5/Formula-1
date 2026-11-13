// Abierto con doble clic (file://) el navegador bloquea los módulos y la carpeta data/: pide usar el lanzador.
if (location.protocol === 'file:') document.getElementById('status').dataset.state = 'file';
