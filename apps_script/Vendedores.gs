/**
 * Alta de vendedores: crea el archivo personal de cada uno y le da acceso.
 *
 * COMO SE INSTALA (una sola vez):
 *   1. Abrir la planilla maestra
 *   2. Extensiones > Apps Script
 *   3. Pegar todo este archivo y guardar
 *   4. Recargar la planilla: aparece el menu "Ventas"
 *   5. La primera vez que se usa, Google pide autorizacion. Hay que darsela:
 *      el script crea archivos y los comparte, y para eso necesita permiso.
 *
 * POR QUE ESTO VIVE EN APPS SCRIPT Y NO EN PYTHON, COMO EL RESTO:
 * La cuenta de servicio que usa el pipeline no tiene espacio en Drive (Google
 * les da cuota cero), asi que no puede crear archivos: devuelve
 * 403 storageQuotaExceeded. Apps Script corre con la cuenta del dueño, que si
 * puede. Crear el archivo y repartir accesos es LO UNICO que hace este script.
 * Asignar leads, reponer y sincronizar lo sigue haciendo Python, que se puede
 * testear y ya corre solo dos veces por dia.
 *
 * La hoja "Mi panel" del archivo NO la crea este script: la agrega Python en la
 * primera corrida (sumar una hoja a un archivo que ya existe no consume cuota
 * de Drive, crear el archivo si). Aca solo se crea "Mis clientes".
 *
 * QUE ESCRIBE EN Config (las columnas A, C y E no se tocan):
 *   A  Vendedor           (ya existia)
 *   G  Email
 *   H  ID del archivo personal
 *   I  URL del archivo personal
 *   J  Estado del acceso
 */

var CONFIG = '⚙️ Config';
var FILA_VENDEDORES = 3;      // en Config los nombres arrancan en A3

var COL_NOMBRE = 1;           // A
var COL_EMAIL = 7;            // G
var COL_ID = 8;               // H
var COL_URL = 9;              // I
var COL_ESTADO = 10;          // J
var TOTAL_COLUMNAS = 10;

var HOJA_VENDEDOR = 'Mis clientes';

// Mismas columnas que ve el vendedor en su archivo. Las descriptivas las
// llena Python; el vendedor solo toca Estado y Motivo.
var COLUMNAS_VENDEDOR = ['Negocio', 'Teléfono', 'Estado', 'Motivo',
                         'Observaciones', 'Última gestión', 'Ciudad',
                         'Categoría', 'Link en Maps', 'Prioridad'];

var COL_ESTADO_VENDEDOR = 3;  // C
var COL_MOTIVO_VENDEDOR = 4;  // D


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ventas')
    .addItem('Agregar vendedor', 'agregarVendedor')
    .addItem('Reparar accesos de todos', 'repararAccesos')
    .addItem('Mandar el link por mail a todos', 'mandarLinks')
    .addSeparator()
    .addItem('Ver estado del sistema', 'estadoDelSistema')
    .addToUi();
}


/** Config, creando las columnas nuevas si la hoja es angosta. */
function _config() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG);
  if (!hoja) throw new Error('No encuentro la hoja ' + CONFIG);
  if (hoja.getMaxColumns() < TOTAL_COLUMNAS) {
    hoja.insertColumnsAfter(hoja.getMaxColumns(),
                            TOTAL_COLUMNAS - hoja.getMaxColumns());
  }
  // Encabezados de las columnas nuevas, solo si estan vacios.
  var enc = [['EMAIL', 'ID ARCHIVO', 'URL ARCHIVO', 'ESTADO']];
  var rango = hoja.getRange(1, COL_EMAIL, 1, 4);
  if (rango.getValues()[0].join('') === '') rango.setValues(enc).setFontWeight('bold');
  return hoja;
}


/** [{fila, nombre, email, id, url}] de todos los vendedores cargados. */
function _vendedores() {
  var hoja = _config();
  var ultima = hoja.getLastRow();
  if (ultima < FILA_VENDEDORES) return [];
  var datos = hoja.getRange(FILA_VENDEDORES, 1,
                            ultima - FILA_VENDEDORES + 1, TOTAL_COLUMNAS).getValues();
  var salida = [];
  for (var i = 0; i < datos.length; i++) {
    var nombre = String(datos[i][COL_NOMBRE - 1] || '').trim();
    if (!nombre) continue;
    salida.push({
      fila: FILA_VENDEDORES + i,
      nombre: nombre,
      email: String(datos[i][COL_EMAIL - 1] || '').trim(),
      id: String(datos[i][COL_ID - 1] || '').trim(),
      url: String(datos[i][COL_URL - 1] || '').trim()
    });
  }
  return salida;
}


function _emailValido(email) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email);
}


/**
 * Crea el archivo personal de un vendedor y lo deja listo para que Python le
 * cargue los leads. Idempotente: si ya tiene archivo, no crea otro.
 */
function _crearArchivo(nombre, email) {
  var libro = SpreadsheetApp.create('Ventas — ' + nombre);
  var hoja = libro.getSheets()[0];
  hoja.setName(HOJA_VENDEDOR);

  hoja.getRange(1, 1, 1, COLUMNAS_VENDEDOR.length)
      .setValues([COLUMNAS_VENDEDOR])
      .setFontWeight('bold')
      .setBackground('#1E5C41')
      .setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);
  hoja.setColumnWidth(1, 260);   // Negocio
  hoja.setColumnWidth(2, 130);   // Teléfono
  hoja.setColumnWidth(3, 140);   // Estado
  hoja.setColumnWidth(4, 170);   // Motivo
  hoja.setColumnWidth(5, 320);   // Observaciones

  _aplicarDesplegables(hoja);

  // El archivo se comparte con el vendedor como editor, y con la cuenta de
  // servicio tambien: es la que despues le carga los leads y le lee el
  // trabajo para llevarlo al master.
  var archivo = DriveApp.getFileById(libro.getId());
  archivo.addEditor(email);
  archivo.addEditor(_cuentaDeServicio());

  return libro;
}


/**
 * El mail de la cuenta de servicio sale de mirar quien tiene acceso al master:
 * es el unico editor que termina en gserviceaccount.com. Asi no hay que
 * hardcodearlo (y este archivo puede vivir en un repo publico).
 */
function _cuentaDeServicio() {
  var editores = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getEditors();
  for (var i = 0; i < editores.length; i++) {
    var mail = editores[i].getEmail();
    if (mail.indexOf('gserviceaccount.com') !== -1) return mail;
  }
  throw new Error('No encuentro la cuenta de servicio entre los editores del master. ' +
                  'Tiene que tener acceso para poder cargar los leads.');
}


/** Desplegables de Estado y Motivo, con la misma lista que usa el master. */
function _aplicarDesplegables(hoja) {
  var cfg = SpreadsheetApp.getActive().getSheetByName(CONFIG);
  var estados = _columnaNoVacia(cfg, 3);   // C
  var motivos = _columnaNoVacia(cfg, 5);   // E
  var filas = Math.max(hoja.getMaxRows() - 1, 1);

  if (estados.length) {
    hoja.getRange(2, COL_ESTADO_VENDEDOR, filas, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(estados, true).build());
  }
  if (motivos.length) {
    hoja.getRange(2, COL_MOTIVO_VENDEDOR, filas, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(motivos, true).build());
  }
}


function _columnaNoVacia(hoja, col) {
  var ultima = hoja.getLastRow();
  if (ultima < FILA_VENDEDORES) return [];
  var vals = hoja.getRange(FILA_VENDEDORES, col, ultima - FILA_VENDEDORES + 1, 1).getValues();
  var salida = [];
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] || '').trim();
    if (v) salida.push(v);
  }
  return salida;
}


/**
 * Le manda al vendedor el link de su archivo.
 *
 * DriveApp.addEditor() da el permiso pero NO avisa: el archivo le aparece en
 * "Compartido conmigo" y listo, y si nadie le dijo que vaya a buscarlo ahi, no
 * lo encuentra nunca. Paso con los primeros 10. El mail sale de la cuenta del
 * dueño de la planilla, porque Apps Script corre con esa cuenta, asi que al
 * vendedor le llega de una direccion que conoce y no de un robot.
 */
function _avisar(nombre, email, url) {
  MailApp.sendEmail({
    to: email,
    subject: 'Tu archivo de ventas — Organizate',
    htmlBody:
      '<p>Hola ' + nombre + ', este es tu archivo de ventas. Es solo tuyo:</p>' +
      '<p><a href="' + url + '">' + url + '</a></p>' +
      '<p>Entra con este mismo mail. Se abre en <b>Mi panel</b>: ahi ves que rubro te ' +
      'toco, cuantos negocios tenes y cuanto te falta para que te entren nuevos. ' +
      'En <b>Mis clientes</b> estan los negocios para llamar.</p>' +
      '<p>Dos cosas que conviene saber:</p>' +
      '<ul>' +
      '<li>Carga siempre el <b>Motivo</b> cuando hablas con alguien. El motivo vacio ' +
      'no cuenta como conversacion, y sin eso no te entran negocios nuevos.</li>' +
      '<li>Podes cambiar de rubro desde el panel, hasta la llamada 10 o cada vez que ' +
      'te entra una tanda nueva.</li>' +
      '</ul>' +
      '<p>Se actualiza solo varias veces por dia. Guardalo en favoritos.</p>'
  });
}


/** Reenvia el link a todos los que ya tienen archivo. */
function mandarLinks() {
  var ui = SpreadsheetApp.getUi();
  var vendedores = _vendedores().filter(function (v) { return v.email && v.url; });
  if (!vendedores.length) return ui.alert('Todavia no hay ningun archivo creado.');

  var nombres = vendedores.map(function (v) { return v.nombre; }).join(', ');
  if (ui.alert('Mandar el link por mail a ' + vendedores.length + ' vendedores?\n\n' +
               nombres, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var ok = [], fallados = [];
  for (var i = 0; i < vendedores.length; i++) {
    try {
      _avisar(vendedores[i].nombre, vendedores[i].email, vendedores[i].url);
      ok.push(vendedores[i].nombre);
    } catch (e) {
      fallados.push(vendedores[i].nombre + ' (' + e.message + ')');
    }
  }
  ui.alert('Mail enviado a: ' + (ok.join(', ') || 'ninguno') +
           (fallados.length ? '\n\nCON PROBLEMAS: ' + fallados.join(', ') : ''));
}


/** Alta completa de un vendedor, desde el menu. */
function agregarVendedor() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt('Agregar vendedor', 'Nombre y apellido:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var nombre = r1.getResponseText().trim();
  if (!nombre) return ui.alert('Falta el nombre.');

  var r2 = ui.prompt('Agregar vendedor', 'Email de Google de ' + nombre + ':',
                     ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var email = r2.getResponseText().trim().toLowerCase();
  if (!_emailValido(email)) return ui.alert('Ese email no parece valido: ' + email);

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return ui.alert('Hay otra operacion corriendo. Proba de nuevo.');

  try {
    var hoja = _config();
    var vendedores = _vendedores();

    for (var i = 0; i < vendedores.length; i++) {
      if (vendedores[i].email && vendedores[i].email.toLowerCase() === email) {
        return ui.alert('Ese email ya es de ' + vendedores[i].nombre + '.');
      }
    }

    // Si el nombre ya esta en Config se reusa esa fila, asi no quedan dos
    // filas para la misma persona. Si ademas ya tiene archivo, no se crea otro.
    var fila = null;
    for (var j = 0; j < vendedores.length; j++) {
      if (vendedores[j].nombre.toLowerCase() === nombre.toLowerCase()) {
        if (vendedores[j].id) {
          return ui.alert(nombre + ' ya tiene archivo:\n' + vendedores[j].url);
        }
        fila = vendedores[j].fila;
        break;
      }
    }
    if (fila === null) {
      fila = vendedores.length ? vendedores[vendedores.length - 1].fila + 1 : FILA_VENDEDORES;
      hoja.getRange(fila, COL_NOMBRE).setValue(nombre);
    }

    var libro = _crearArchivo(nombre, email);
    hoja.getRange(fila, COL_EMAIL, 1, 4).setValues(
      [[email, libro.getId(), libro.getUrl(), 'ACTIVO']]);
    SpreadsheetApp.flush();

    ui.alert('Listo: ' + nombre + '\n\n' + libro.getUrl() +
             '\n\nLos leads se cargan solos en la proxima corrida (00:00 o 12:00).');
  } finally {
    lock.releaseLock();
  }
}


/**
 * Repasa Config y arregla lo que falte: crea el archivo del que no lo tiene y
 * reaplica el acceso del que lo perdio. Sirve para dar de alta en lote a los
 * vendedores cuyos mails se cargaron a mano en la columna G.
 */
function repararAccesos() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return ui.alert('Hay otra operacion corriendo.');

  try {
    var hoja = _config();
    var vendedores = _vendedores();
    var creados = [], arreglados = [], sinEmail = [], fallados = [];

    for (var i = 0; i < vendedores.length; i++) {
      var v = vendedores[i];
      if (!v.email) { sinEmail.push(v.nombre); continue; }
      if (!_emailValido(v.email)) { fallados.push(v.nombre + ' (email invalido)'); continue; }

      try {
        if (!v.id) {
          var libro = _crearArchivo(v.nombre, v.email);
          hoja.getRange(v.fila, COL_EMAIL, 1, 4).setValues(
            [[v.email, libro.getId(), libro.getUrl(), 'ACTIVO']]);
          // Solo a los recien creados: reenviarle el link a todos en cada
          // reparacion seria spamear al equipo entero. Para eso esta
          // "Mandar el link por mail a todos".
          _avisar(v.nombre, v.email, libro.getUrl());
          creados.push(v.nombre);
        } else {
          var archivo = DriveApp.getFileById(v.id);
          archivo.addEditor(v.email);
          archivo.addEditor(_cuentaDeServicio());
          hoja.getRange(v.fila, COL_ESTADO).setValue('ACTIVO');
          arreglados.push(v.nombre);
        }
      } catch (e) {
        fallados.push(v.nombre + ' (' + e.message + ')');
        hoja.getRange(v.fila, COL_ESTADO).setValue('ERROR');
      }
    }
    SpreadsheetApp.flush();

    var msg = 'Archivos creados: ' + (creados.join(', ') || 'ninguno') +
              '\nAccesos revisados: ' + (arreglados.join(', ') || 'ninguno');
    if (sinEmail.length) msg += '\n\nSin email cargado (columna G): ' + sinEmail.join(', ');
    if (fallados.length) msg += '\n\nCON PROBLEMAS: ' + fallados.join(', ');
    ui.alert(msg);
  } finally {
    lock.releaseLock();
  }
}


/** Un vistazo rapido: quien tiene archivo, quien no, y quien queda afuera. */
function estadoDelSistema() {
  var vendedores = _vendedores();
  var listos = [], faltan = [];
  for (var i = 0; i < vendedores.length; i++) {
    (vendedores[i].id && vendedores[i].email ? listos : faltan).push(vendedores[i].nombre);
  }
  SpreadsheetApp.getUi().alert(
    'Vendedores en Config: ' + vendedores.length +
    '\n\nCon archivo y acceso (' + listos.length + '): ' + (listos.join(', ') || '-') +
    '\n\nLes falta email o archivo (' + faltan.length + '): ' + (faltan.join(', ') || '-') +
    '\n\nPara los que faltan: cargales el email en la columna G de Config y corré ' +
    '"Reparar accesos de todos".');
}
