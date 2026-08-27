/* ================================================
   Credenciales App — script.js
   ================================================
   Toda la configuracion de posicion, tipografia y
   colores esta en el objeto CONFIG de abajo.
   Ajusta las coordenadas a ojo segun el resultado.
   ================================================ */

const CONFIG = {
  // Dimensiones del canvas (deben coincidir con plantilla.png)
  canvasWidth:  1063,
  canvasHeight: 1417,

  fields: {
    // Nombre completo — banda crema superior, centrado
    nombre: {
      x:           531,   // centro horizontal del canvas
      y:           500,   // posicion vertical (baseline del texto)
      maxWidth:    770,   // ancho maximo antes de reducir fuente
      align:       'center',
      maxFontSize: 70,    // tamano de fuente maximo (px)
      minFontSize: 20,    // tamano de fuente minimo (px)
    },

    // Usuario — dentro del recuadro blanco, despues del icono persona
    usuario: {
      x:           280,
      y:           840,
      maxWidth:    640,
      align:       'left',
      maxFontSize: 50,
      minFontSize: 18,
    },

    // Contrasena — dentro del recuadro blanco, despues del icono candado
    contrasena: {
      x:           280,
      y:           1100,
      maxWidth:    640,
      align:       'left',
      maxFontSize: 50,
      minFontSize: 18,
    },
  },

  fontFamily:  'Montserrat',
  fontWeight:  700,          // 700 = bold
  textColor:   '#005987',

  // Nombre del PDF descargado
  pdfFileName: 'credenciales.pdf',

  // Prefijo del nombre de la hoja a leer (case-insensitive, busca la primera hoja que EMPIECE con este texto).
  // Ejemplos que funcionan sin cambiar código: "Listado Alumnos 2026", "Listado Alumnos 2027", etc.
  // Si el archivo tiene una sola hoja, se usa esa sin importar el nombre.
  sheetName: 'Listado Alumnos',

  // Columnas requeridas en el Excel (case-insensitive)
  requiredColumns: ['nombre', 'apellido', 'usuario academia', 'contraseña academia'],
};

/* ── Referencias al DOM ── */
const fileInput          = document.getElementById('file-input');
const dropZone           = document.getElementById('drop-zone');
const dropIcon           = document.getElementById('drop-icon');
const dropLabel          = document.getElementById('drop-label');

const tabLocal           = document.getElementById('tab-local');
const tabLink            = document.getElementById('tab-link');
const zoneLocal          = document.getElementById('zone-local');
const zoneLink           = document.getElementById('zone-link');
const inputLink          = document.getElementById('input-link');
const btnLoadLink        = document.getElementById('btn-load-link');
const sourceLoadedBadge  = document.getElementById('source-loaded-badge');
const sourceLoadedName   = document.getElementById('source-loaded-name');
const btnClearSource     = document.getElementById('btn-clear-source');

const btnGenerate        = document.getElementById('btn-generate');
const progressSec        = document.getElementById('progress-section');
const progressBar        = document.getElementById('progress-bar');
const progressText       = document.getElementById('progress-text');
const progressLabelText  = document.getElementById('progress-label-text');
const messageBox         = document.getElementById('message-box');
const privacyBadge       = document.getElementById('privacy-badge');
const groupFilter        = document.getElementById('group-filter');
const selectGrupo        = document.getElementById('select-grupo');
const spinner            = document.getElementById('spinner');

let selectedFile = null;
let cachedRows = null;

/* ── Cambio de pestañas (Tabs) ── */
tabLocal.addEventListener('click', () => {
  if (tabLocal.classList.contains('active') && sourceLoadedBadge.style.display === 'none') return;
  clearFile();
  tabLocal.classList.add('active');
  tabLink.classList.remove('active');
  zoneLocal.style.display = 'block';
  zoneLink.style.display = 'none';
});

tabLink.addEventListener('click', () => {
  if (tabLink.classList.contains('active') && sourceLoadedBadge.style.display === 'none') return;
  clearFile();
  tabLink.classList.add('active');
  tabLocal.classList.remove('active');
  zoneLink.style.display = 'block';
  zoneLocal.style.display = 'none';
});

/* ── Drag & drop ── */
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) setFile(f);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

function populateUIWithRows(rows) {
  cachedRows = rows;
  if (cachedRows.length === 0) {
    throw new Error('El archivo no contiene filas de datos.');
  }
  
  // Validamos apenas se carga para dar feedback rapido
  const colError = validateColumns(cachedRows[0]);
  if (colError) {
    throw new Error(colError);
  }
  
  // Buscar grupos unicos
  const grupos = new Set();
  let hasGroupColumn = false;
  cachedRows.forEach(row => {
    if ('grupo' in row) hasGroupColumn = true;
    const g = trim(row['grupo']);
    if (g) grupos.add(g);
  });
  
  // Mostrar select si hay grupos
  if (hasGroupColumn && grupos.size > 0) {
    selectGrupo.innerHTML = '<option value="todos">Todos los grupos</option>';
    Array.from(grupos).sort().forEach(g => {
      const opt = document.createElement('option');
      opt.value = g;
      opt.textContent = g;
      selectGrupo.appendChild(opt);
    });
    groupFilter.style.display = 'block';
  } else {
    groupFilter.style.display = 'none';
  }
  
  btnGenerate.disabled = false;
}

async function setFile(f) {
  if (!f.name.match(/\.xlsx$/i)) {
    showMessage('Solo se admiten archivos .xlsx exportados desde Excel o Google Sheets.', 'error');
    btnGenerate.disabled = true;
    return;
  }
  selectedFile = f;
  hideMessage();
  progressSec.classList.remove('visible', 'completed');
  
  setBusy(true);
  try {
    const rows = await readExcel(selectedFile);
    populateUIWithRows(rows);

    // Mostrar pastilla limpia y ocultar contenedor pesado
    sourceLoadedName.textContent = '✓ ' + f.name;
    sourceLoadedBadge.style.display = 'flex';
    zoneLocal.style.display = 'none';
    zoneLink.style.display = 'none';
  } catch (err) {
    showMessage(err.message, 'error');
    btnGenerate.disabled = true;
  }
  setBusy(false);
}

/* ── Carga por URL ── */
btnLoadLink.addEventListener('click', async () => {
  const url = inputLink.value.trim();
  if (!url) return;
  hideMessage();
  progressSec.classList.remove('visible', 'completed');
  setBusy(true);
  try {
    const rows = await readExcelFromUrl(url);
    populateUIWithRows(rows);

    // Mostrar pastilla limpia y ocultar contenedor pesado
    sourceLoadedName.textContent = '✓ Planilla de Google Sheets cargada';
    sourceLoadedBadge.style.display = 'flex';
    zoneLocal.style.display = 'none';
    zoneLink.style.display = 'none';
  } catch (err) {
    showMessage(err.message, 'error');
    btnGenerate.disabled = true;
  }
  setBusy(false);
});

inputLink.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    btnLoadLink.click();
  }
});

/* ── Boton × limpiar planilla ── */
btnClearSource.addEventListener('click', clearFile);

function clearFile() {
  selectedFile = null;
  cachedRows = null;
  
  // Limpiar inputs
  fileInput.value = '';             
  inputLink.value = '';

  // Ocultar badge limpio
  sourceLoadedBadge.style.display = 'none';
  sourceLoadedName.textContent = '';

  // Restaurar el contenedor correspondiente a la pestaña activa
  if (tabLocal.classList.contains('active')) {
    zoneLocal.style.display = 'block';
    zoneLink.style.display = 'none';
  } else {
    zoneLink.style.display = 'block';
    zoneLocal.style.display = 'none';
  }

  groupFilter.style.display = 'none';
  selectGrupo.innerHTML = '<option value="todos">Todos los grupos</option>';
  btnGenerate.disabled = true;
  hideMessage();
  progressSec.classList.remove('visible', 'completed');
  progressBar.style.width = '0%';
}

/* ── Boton principal ── */
btnGenerate.addEventListener('click', async () => {
  if (!cachedRows) {
    showMessage('Primero selecciona un archivo .xlsx o carga un enlace de Google Sheets.', 'error');
    return;
  }
  await generate();
});

/* ── Logica principal ── */
async function generate() {
  hideMessage();
  setBusy(true);

  try {
    if (!cachedRows) throw new Error('No hay datos cargados.');
    
    // Filtrar segun lo que haya en el select
    const selectedGroup = selectGrupo.value;
    const rows = selectedGroup === 'todos' 
      ? cachedRows 
      : cachedRows.filter(r => trim(r['grupo']) === selectedGroup);

    if (rows.length === 0) {
      showMessage('No hay alumnos en el grupo seleccionado.', 'error');
      setBusy(false);
      return;
    }

    // 3. Esperar carga de fuente
    await loadFont();

    // 4. Cargar imagen plantilla
    const templateImg = await loadImage('./plantilla.png');

    // 5. Generar paginas
    const { jsPDF } = window.jspdf;
    // Calcula el tamano de pagina en mm manteniendo la proporcion de la imagen
    const mmW = 210; // A4-ish ancho en mm (aprox)
    const ratio = CONFIG.canvasHeight / CONFIG.canvasWidth;
    const mmH = mmW * ratio;

    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: [mmW, mmH],
    });

    let skipped = 0;
    const total = rows.length;

    setProgress(0, total);
    progressLabelText.textContent = 'Generando credenciales...';
    progressSec.classList.remove('completed');
    progressSec.classList.add('visible');

    // Canvas oculto reutilizable
    const canvas = document.createElement('canvas');
    canvas.width  = CONFIG.canvasWidth;
    canvas.height = CONFIG.canvasHeight;
    const ctx = canvas.getContext('2d');

    for (let i = 0; i < total; i++) {
      const row = rows[i];
      const nombre     = trim(row['nombre']);
      const apellido   = trim(row['apellido']);
      const usuario    = trim(row['usuario academia']);
      const contrasena = trim(row['contraseña academia']);

      // Validar fila
      if (!nombre && !apellido && !usuario && !contrasena) {
        skipped++;
        setProgress(i + 1, total);
        continue;
      }

      const nombreCompleto = [nombre, apellido].filter(Boolean).join(' ');

      // Dibujar en canvas
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(templateImg, 0, 0, CONFIG.canvasWidth, CONFIG.canvasHeight);

      drawField(ctx, nombreCompleto, CONFIG.fields.nombre);
      drawField(ctx, usuario,        CONFIG.fields.usuario);
      drawField(ctx, contrasena,     CONFIG.fields.contrasena);

      // Agregar pagina al PDF
      if (i > 0) pdf.addPage([mmW, mmH]);
      const imgData = canvas.toDataURL('image/jpeg', 0.92);
      pdf.addImage(imgData, 'JPEG', 0, 0, mmW, mmH);

      // Yield al event loop cada 5 paginas para no bloquear la UI
      if ((i + 1) % 5 === 0) {
        setProgress(i + 1, total);
        await yieldFrame();
      }
    }

    setProgress(total, total);
    progressLabelText.textContent = 'Credenciales generadas ✔';
    progressText.textContent = `${total - skipped} de ${total} generada(s)`;
    progressSec.classList.add('completed');

    // 6. Descargar PDF
    pdf.save(CONFIG.pdfFileName);

    if (skipped > 0) {
      showMessage(`Se omitieron ${skipped} fila(s) con campos vacíos.`, 'warn');
    }

  } catch (err) {
    console.error(err);
    progressSec.classList.remove('visible', 'completed');
    showMessage('Error inesperado: ' + err.message, 'error');
  }

  setBusy(false);
}

/* ── Helpers ── */

function trim(v) {
  return v !== undefined && v !== null ? String(v).trim() : '';
}

function yieldFrame() {
  return new Promise(r => requestAnimationFrame(r));
}

function processWorkbookData(data) {
  const workbook = XLSX.read(data, { type: 'array' });
  const sheetNames = workbook.SheetNames;

  let targetSheetName;
  if (sheetNames.length === 1) {
    targetSheetName = sheetNames[0];
  } else {
    const prefix = CONFIG.sheetName.trim().toLowerCase();
    targetSheetName = sheetNames.find(n => n.trim().toLowerCase().startsWith(prefix));
    if (!targetSheetName) {
      throw new Error(
        `No se encontró ninguna hoja que empiece con "${CONFIG.sheetName}" en el archivo.\n` +
        `Verifica que la pestaña del listado de alumnos empiece con ese texto.`
      );
    }
  }

  const sheet = workbook.Sheets[targetSheetName];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  const normalized = rows.map(r => {
    const out = {};
    for (const k of Object.keys(r)) {
      out[k.trim().toLowerCase()] = r[k];
    }
    return out;
  });
  return normalized;
}

async function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        resolve(processWorkbookData(data));
      } catch (err) {
        reject(new Error('No se pudo leer el archivo Excel: ' + err.message));
      }
    };
    reader.onerror = () => reject(new Error('Error al leer el archivo local.'));
    reader.readAsArrayBuffer(file);
  });
}

async function readExcelFromUrl(url) {
  try {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) throw new Error('Enlace de Google Sheets no válido.');
    const id = match[1];
    
    // Endpoint nativo de exportacion de Google
    const exportUrl = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
    const res = await fetch(exportUrl);
    
    if (!res.ok) {
      throw new Error('No se pudo descargar la planilla. Asegúrate de que el acceso esté configurado como "Cualquier persona con el enlace puede leer" y verifica tu conexión a internet.');
    }
    
    const buffer = await res.arrayBuffer();
    const data = new Uint8Array(buffer);
    return processWorkbookData(data);
  } catch (err) {
    throw new Error(err.message);
  }
}

function validateColumns(firstRow) {
  const keys = Object.keys(firstRow);
  const missing = CONFIG.requiredColumns.filter(c => !keys.includes(c));
  if (missing.length > 0) {
    return `Faltan las siguientes columnas en el Excel: ${missing.join(', ')}.\nColumnas encontradas: ${keys.join(', ')}`;
  }
  return null;
}

async function loadFont() {
  // Carga la fuente Montserrat y espera que este disponible para Canvas
  const fontSpec = `${CONFIG.fontWeight} 40px '${CONFIG.fontFamily}'`;
  await document.fonts.load(fontSpec);
  await document.fonts.ready;
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload  = () => resolve(img);
    img.onerror = () => reject(new Error('No se pudo cargar la imagen: ' + src));
    img.src     = src;
  });
}

function drawField(ctx, text, field) {
  if (!text) return;

  let fontSize = field.maxFontSize;
  ctx.textAlign    = field.align;
  ctx.textBaseline = 'alphabetic';

  // Ajuste automatico de fuente
  while (fontSize > field.minFontSize) {
    ctx.font = `${CONFIG.fontWeight} ${fontSize}px '${CONFIG.fontFamily}'`;
    const w = ctx.measureText(text).width;
    if (w <= field.maxWidth) break;
    fontSize--;
  }

  ctx.fillStyle = CONFIG.textColor;
  ctx.font      = `${CONFIG.fontWeight} ${fontSize}px '${CONFIG.fontFamily}'`;
  ctx.fillText(text, field.x, field.y);
}

function setProgress(current, total) {
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  progressBar.style.width = pct + '%';
  progressText.textContent = `Procesando ${current} de ${total}`;
}

function setBusy(busy) {
  btnGenerate.disabled = busy || !cachedRows;
  spinner.classList.toggle('visible', busy);
  if (!busy) {
    // Mantener la barra visible pero dejar de girar el spinner
  }
}

function showMessage(msg, type) {
  messageBox.textContent = msg;
  messageBox.className   = `message-box ${type} visible`;
  // Ocultar el badge de privacidad cuando hay cualquier mensaje (error, exito o advertencia)
  privacyBadge.style.display = 'none';
}

function hideMessage() {
  messageBox.className = 'message-box';
  privacyBadge.style.display = '';  // restaurar badge al limpiar mensaje
}
