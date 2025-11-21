// ========================================
// CONFIGURACIÓN GLOBAL

const CONFIG_DEFAULTS = {
  sendWindow: { startHour: 8, endHour: 14 },
  workDays: [1, 2, 3, 4, 5],
  dailyLimit: 20,
  senderAlias: 'bhinternational@bhterra.es',
  senderName: 'equipo bhterrainational',
  sheets: {
    empresas: 'Empresas',
    contactos: 'Contactos',
    clientes: 'Clientes',
    prospectosFase1: 'Prospectos_Fase1',
    prospectosFase2: 'Prospectos_Fase2',
    outbox: 'Outbox',
    logs: 'Logs_Sistema'
  }
};
// ==================================================================
//               ↓↓↓ PEGA ESTE BLOQUE COMPLETO ↓↓↓
// ==================================================================

/**
 * Obtiene la configuración completa del sistema.
 * --- MODIFICADO: AHORA UTILIZA CACHESERVICE PARA MEJORAR EL RENDIMIENTO ---
 * Utiliza CacheService para un acceso rápido. Si no está en caché, lee
 * desde PropertiesService y la guarda en caché por 5 minutos.
 * @returns {Object} El objeto de configuración completo.
 */
function getSystemConfiguration() {
  const cache = CacheService.getScriptCache();
  const cachedConfig = cache.get('system_config');
  if (cachedConfig) {
    return JSON.parse(cachedConfig);
  }

  const properties = PropertiesService.getScriptProperties();
  const automationStatus = obtenerEstadoTriggers().automatizacion ? 'activo' : 'inactivo';
  
  const config = {
    automationStatus: automationStatus,
    dailyLimit: parseInt(properties.getProperty('dailyLimit')) || CONFIG_DEFAULTS.dailyLimit,
    startHour: parseInt(properties.getProperty('startHour')) || CONFIG_DEFAULTS.sendWindow.startHour,
    endHour: parseInt(properties.getProperty('endHour')) || CONFIG_DEFAULTS.sendWindow.endHour,
    workDays: JSON.parse(properties.getProperty('workDays') || JSON.stringify(CONFIG_DEFAULTS.workDays)),
    sheets: CONFIG_DEFAULTS.sheets, 
    senderAlias: CONFIG_DEFAULTS.senderAlias,
    senderName: CONFIG_DEFAULTS.senderName
  };

  cache.put('system_config', JSON.stringify(config), 300); // Cachear por 5 minutos
  return config;
}


/**
 * Actualiza las propiedades, limpia la caché de configuración y devuelve la nueva configuración.
 * --- MODIFICADO: AÑADIDO LOCKSERVICE Y GESTIÓN DE CACHÉ ---
 * Protegido con LockService para evitar condiciones de carrera.
 * @param {Object} newConfig Un objeto con las nuevas configuraciones.
 * @returns {Object} Un objeto con el estado del éxito y la nueva configuración.
 */
function updateSystemConfiguration(newConfig) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'El sistema está ocupado. Inténtalo de nuevo en unos segundos.' };
  }

  try {
    const properties = PropertiesService.getScriptProperties();
    
    if (newConfig.hasOwnProperty('dailyLimit')) {
      const limit = parseInt(newConfig.dailyLimit);
      if (!isNaN(limit) && limit >= 0) properties.setProperty('dailyLimit', String(limit));
    }
    if (newConfig.hasOwnProperty('startHour')) {
      const hour = parseInt(newConfig.startHour);
      if (!isNaN(hour) && hour >= 0 && hour <= 23) properties.setProperty('startHour', String(hour));
    }
    if (newConfig.hasOwnProperty('endHour')) {
      const hour = parseInt(newConfig.endHour);
      if (!isNaN(hour) && hour >= 0 && hour <= 23) properties.setProperty('endHour', String(hour));
    }
    if (newConfig.hasOwnProperty('workDays')) {
      properties.setProperty('workDays', JSON.stringify(newConfig.workDays));
    }
    
    // --- NUEVO: Limpiamos la caché para que la próxima llamada a getSystemConfiguration obtenga los nuevos valores ---
    CacheService.getScriptCache().remove('system_config');
    
    log('INFO', 'Configuración del sistema actualizada desde el panel de administración.');
    
    return { 
      success: true, 
      message: 'Configuración guardada.', 
      newConfig: getSystemConfiguration() 
    };

  } catch (e) {
    log('ERROR', `Error al actualizar la configuración: ${e.message}`);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * FUNCIÓN DE COMPATIBILIDAD SIMPLE Y ROBUSTA.
 * Reemplaza el objeto CONFIG para mantener compatibilidad con el código existente.
 * Esta función devuelve la configuración actual del sistema.
 * 
 * NOTA: Para acceder a propiedades, usa: CONFIG().sheets.empresas
 * O mejor aún, usa directamente: getSystemConfiguration().sheets.empresas
 * 
 * @returns {Object} El objeto de configuración completo.
 */
function CONFIG() {
  return getSystemConfiguration();
}

function sanitizarFilaParaRespuesta(row) {
  if (!row || typeof row !== 'object') {
    return row;
  }

  const sanitized = {};
  Object.keys(row).forEach(key => {
    const value = row[key];
    if (value instanceof Date) {
      sanitized[key] = value.toISOString();
    } else if (Array.isArray(value)) {
      sanitized[key] = value.map(item => (item instanceof Date ? item.toISOString() : item));
    } else {
      sanitized[key] = value;
    }
  });
  return sanitized;
}



// ========================================
// MENÚ Y ACTIVADORES PRINCIPALES (onOpen, onEdit)
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sistema CRM')
    .addItem('🖥️ Lanzar CRM', 'lanzarCrm')
    .addSeparator()
    .addItem('⚙️ Admin: Crear/Verificar Esquema', 'crearEsquema')
    .addItem('🎨 Admin: Aplicar Formatos y Menús', 'aplicarFormatosYValidaciones')
    // --- INICIO DE LA LÍNEA AÑADIDA ---
    .addItem('🔧 Admin: Instalar Activador de Limpieza', 'instalarActivadorDeLimpieza')
    // --- FIN DE LA LÍNEA AÑADIDA ---
    .addSeparator()
    .addItem('🔄 Admin: Sincronizar Nuevas Empresas', 'sincronizarNuevasEmpresas')
    .addToUi();
}
/**
 * Se ejecuta al editar una celda.
 * --- VERSIÓN CORREGIDA Y OPTIMIZADA ---
 * 1. Elimina `Utilities.sleep()` para mayor eficiencia en la hoja 'Contactos'.
 * 2. Integra la acción "Reactivar Secuencia" con la nueva hoja 'Conversaciones'.
 * 3. Refactoriza la lógica de actualización de prospectos para mayor claridad.
 */
// ==================================================================
//               ↓↓↓ REEMPLAZA ESTA FUNCIÓN COMPLETA ↓↓↓
// ==================================================================

function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    if (row <= 1) return; // Ignorar ediciones en el encabezado

    const config = CONFIG();

    // --- LÓGICA DE COLA PARA LA HOJA 'Empresas' ---
    if (sheet.getName() === config.sheets.empresas) {
      const empresaData = obtenerDatosEmpresa(sheet, row);
      if (empresaData && empresaData.ID_Empresa && empresaData.NombreEmpresa) {
        // En lugar de crear el prospecto aquí, creamos una tarea
        const payload = {
          empresaId: empresaData.ID_Empresa,
          // Incluimos solo los datos necesarios para minimizar el tamaño del payload
          data: {
            ID_Empresa: empresaData.ID_Empresa,
            NombreEmpresa: empresaData.NombreEmpresa,
            PAIS: empresaData.PAIS,
            EmailGeneral: empresaData.EmailGeneral,
            EmailContacto: empresaData.EmailContacto,
          }
        };
        agregarTareaALaCola('CREATE_PROSPECT_F1', payload);
      }
      return;
    }

    // --- LÓGICA DE COLA PARA LA HOJA 'Contactos' ---
    if (sheet.getName() === config.sheets.contactos) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const datosFilaCompleta = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
      const contactoData = crearObjetoDesdeArray(headers, datosFilaCompleta);

      // Si se cumplen las condiciones para crear un prospecto de Fase 2
      if (contactoData.ContactoId && contactoData.EmpresaId && contactoData.NombreContacto && contactoData.EmailContacto) {
        const payload = { 
          contactoId: contactoData.ContactoId, 
          data: contactoData // Aquí podemos pasar el objeto completo
        };
        agregarTareaALaCola('CREATE_PROSPECT_F2', payload);
      }
      return;
    }
    
    // La lógica de los menús desplegables es una acción del usuario que espera una respuesta
    // visual inmediata, por lo que la mantenemos síncrona aquí. Esta parte es rápida y no justifica el uso de la cola.
    const esHojaProspectosF1 = sheet.getName() === config.sheets.prospectosF1;
    const esHojaProspectosF2 = sheet.getName() === config.sheets.prospectosF2;

    if (esHojaProspectosF1 || esHojaProspectosF2) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const nombreColumna = headers[range.getColumn() - 1];
      const valorCelda = range.getValue().toString();
      
      if (nombreColumna === 'Gestionar Prospecto' && valorCelda !== '') {
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(15000)) return;
        try {
          const prospectoData = crearObjetoDesdeArray(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
          procesarAccionMenuGestionar(sheet, range, row, valorCelda, prospectoData);
        } finally {
          lock.releaseLock();
        }
      }
      if (esHojaProspectosF1 && nombreColumna === 'IniciarFase2' && valorCelda.toUpperCase() === 'SÍ') {
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(15000)) return;
        try {
          range.clearContent();
          const prospectoData = crearObjetoDesdeArray(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
          const nuevoIdContactoGenerado = iniciarFase2DesdeProspecto(prospectoData);
          if (nuevoIdContactoGenerado) {
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'Estado')).setValue('Convertido');
            const timestamp = `[${new Date().toLocaleString('es-ES')}]`;
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'Última Acción Manual')).setValue(`Convertido a ${nuevoIdContactoGenerado} ${timestamp}`);
          }
        } finally {
          lock.releaseLock();
        }
      }
    }
  } catch (error) {
    log('ERROR', `Error crítico en onEdit: ${error.toString()}`);
  }
}

function agregarTareaALaCola(tipo, payload) {
  try {
    const queueSheet = obtenerHoja('Queue_System');
    if (!queueSheet) {
      log('ERROR_CRITICO', 'No se encontró la hoja de sistema "Queue_System".');
      return;
    }
    const taskId = Utilities.getUuid();
    const payloadString = JSON.stringify(payload);

    // --- INICIO DE LA CORRECCIÓN LÓGICA ---
    const lastRow = queueSheet.getLastRow();
    
    // Solo verificamos duplicados si hay al menos una tarea existente (lastRow > 1)
    if (lastRow > 1) {
      const numRowsToCheck = Math.min(10, lastRow - 1); // Comprobamos 10 o menos si no hay tantas
      const startRow = lastRow - numRowsToCheck + 1;
      
      const lastTasks = queueSheet.getRange(startRow, 3, numRowsToCheck, 2).getValues();
      const isDuplicate = lastTasks.some(row => row[0] === tipo && row[1] === payloadString);

      if (isDuplicate) {
        log('INFO', `Se omitió la adición de una tarea duplicada a la cola. Tipo: ${tipo}`);
        return; // Salimos de la función si es un duplicado
      }
    }
    // --- FIN DE LA CORRECCIÓN LÓGICA ---

    // Si no es duplicado (o si la hoja está vacía), añadimos la nueva tarea.
    queueSheet.appendRow([new Date(), taskId, tipo, payloadString, 'pending', '']);

  } catch (e) {
    log('ERROR', `No se pudo agregar la tarea a la cola: ${e.message}`);
  }
}

/**
 * Función auxiliar que encapsula la lógica síncrona del menú "Gestionar Prospecto".
 * Se llama desde onEdit para mantener el código principal más limpio.
 */
function procesarAccionMenuGestionar(sheet, range, row, valorCelda, prospectoData) {
    const estadoCol = obtenerIndiceColumna(sheet, 'Estado');
    const esHojaProspectosF1 = sheet.getName() === CONFIG().sheets.prospectosF1;
    range.clearContent();
    const timestamp = `[${new Date().toLocaleString('es-ES')}]`;
    const accionCol = obtenerIndiceColumna(sheet, 'Última Acción Manual');

    const actualizarAccion = (mensaje) => {
        if (accionCol > 0) sheet.getRange(row, accionCol).setValue(`${mensaje} ${timestamp}`);
    };

    switch (valorCelda) {
        case 'NO CONTACTAR':
            sheet.getRange(row, estadoCol).setValue('Baja');
            actualizarAccion('Baja Manual');
            break;
        case 'Pausar (Manual)':
            sheet.getRange(row, estadoCol).setValue('pausado');
            actualizarAccion('Pausado');
            break;
        case 'Reactivar Secuencia':
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'FechaUltimoEnvio')).clearContent();
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'SeguimientoActual')).setValue(0);
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'ThreadId')).clearContent();
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'RespuestaRecibida')).clearContent();
            sheet.getRange(row, obtenerIndiceColumna(sheet, 'ContenidoRespuesta')).clearContent();
            sheet.getRange(row, estadoCol).setValue('nuevo');
            actualizarAccion('Reactivado');
            
            const tipoEntidad = esHojaProspectosF1 ? 'fase1' : 'fase2';
            const entidadId = esHojaProspectosF1 ? prospectoData.ID_Empresa : prospectoData.ContactoId;
            if (entidadId) {
                actualizarEstadoConversacion(entidadId, tipoEntidad, 'archivado');
            }
            break;
        case '✅ CONVERTIR A CLIENTE':
            const faseActual = esHojaProspectosF1 ? 'fase1' : 'fase2';
            const exito = convertirProspectoACliente(prospectoData, faseActual);
            if (exito) {
                sheet.getRange(row, estadoCol).setValue('Cliente');
                actualizarAccion('Convertido a Cliente');
            } else {
                actualizarAccion('ERROR de conversión');
            }
            break;
    }
}
// ========================================
// LÓGICA PRINCIPAL DE PROCESAMIENTO
// ========================================

function ejecutarAutomatizacion() {
  try {
    log('INFO', 'Iniciando ciclo de automatización');
    if (!esHorarioLaboral()) { return; }
    verificarRespuestas();
    procesarFase1();
    procesarFase2();
    procesarOutbox();
   
    log('INFO', 'Ciclo de automatización completado');
  } catch (error) {
    log('ERROR', `Error en automatización: ${error.toString()}`);
  }
}




function procesarFase1() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.prospectosFase1);
  if (!sheet || sheet.getLastRow() < 2) return;

  const rangoCompleto = sheet.getDataRange();
  const todosLosDatos = rangoCompleto.getValues();
  const headers = todosLosDatos.shift(); // Saca los encabezados

  let cambiosRealizados = false;

  const datosModificados = todosLosDatos.map((row, index) => {
    const prospecto = crearObjetoDesdeArray(headers, row);
    
    // Solo procesamos prospectos que tienen ID y están en un estado "enviable"
    const esEnviable = prospecto.ID_Empresa && (prospecto.Estado === 'nuevo' || prospecto.Estado === 'activo');
    if (!esEnviable) {
      return row; // Si no es enviable, devolvemos la fila sin cambios
    }

    const resultadoProceso = procesarProspecto(prospecto, 'fase1');

    if (resultadoProceso.enviar) {
      const correo = resultadoProceso.correo;
      encolarCorreo({
        prospectoId: prospecto.ID_Empresa,
        contactoId: null,
        toEmail: resultadoProceso.emailDestino,
        asunto: correo.asunto,
        html: correo.html,
        threadId: prospecto.ThreadId || null
      });
      
      // Actualizamos la fila (el array 'row') en memoria
      row[headers.indexOf('FechaUltimoEnvio')] = new Date();
      row[headers.indexOf('SeguimientoActual')] = correo.tipoSeguimiento;
      if (prospecto.Estado === 'nuevo') {
        row[headers.indexOf('Estado')] = 'activo';
      }
      cambiosRealizados = true;
    }
    
    return row; // Devolvemos la fila (modificada o no)
  });

  // Si se realizó al menos un cambio, escribimos todo de vuelta a la hoja
  if (cambiosRealizados) {
    sheet.getRange(2, 1, datosModificados.length, headers.length).setValues(datosModificados);
    log('INFO', 'Lote de cambios para Prospectos_Fase1 escrito en la hoja.');
  }
}

function procesarFase2() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.prospectosFase2);
  if (!sheet || sheet.getLastRow() < 2) return;

  const rangoCompleto = sheet.getDataRange();
  const todosLosDatos = rangoCompleto.getValues();
  const headers = todosLosDatos.shift();

  let cambiosRealizados = false;

  const datosModificados = todosLosDatos.map((row, index) => {
    const prospecto = crearObjetoDesdeArray(headers, row);

    // Solo procesamos prospectos que tienen ID, email y están en estado "enviable"
    const esEnviable = prospecto.ID && prospecto.EmailContacto && (prospecto.Estado === 'nuevo' || prospecto.Estado === 'activo');
    if (!esEnviable) {
      return row;
    }

    const resultadoProceso = procesarProspecto(prospecto, 'fase2');

    if (resultadoProceso.enviar) {
      const correo = resultadoProceso.correo;
      encolarCorreo({
        prospectoId: prospecto.ID,
        contactoId: prospecto.ContactoId,
        toEmail: resultadoProceso.emailDestino,
        asunto: correo.asunto,
        html: correo.html,
        threadId: prospecto.ThreadId || null
      });
      
      // Actualizamos la fila (el array 'row') en memoria
      row[headers.indexOf('FechaUltimoEnvio')] = new Date();
      row[headers.indexOf('SeguimientoActual')] = correo.tipoSeguimiento;
      if (prospecto.Estado === 'nuevo') {
        row[headers.indexOf('Estado')] = 'activo';
      }
      cambiosRealizados = true;
    }
    
    return row;
  });

  if (cambiosRealizados) {
    sheet.getRange(2, 1, datosModificados.length, headers.length).setValues(datosModificados);
    log('INFO', 'Lote de cambios para Prospectos_Fase2 escrito en la hoja.');
  }
}

// ==================================================================
//               ↓↓↓ REEMPLAZA ESTA FUNCIÓN COMPLETA ↓↓↓
// ==================================================================
/**
 * Lógica central para procesar un único prospecto.
 * --- ARQUITECTURA MEJORADA: Ahora es una función "pura" ---
 * No modifica la hoja directamente. Recibe datos y devuelve una decisión
 * ('enviar' o no 'enviar') y la información necesaria.
 * @param {object} prospect El objeto de datos del prospecto.
 * @param {string} fase La fase del prospecto ('fase1' o 'fase2').
 * @returns {object} Un objeto con la decisión: { enviar: boolean, emailDestino?: string, correo?: object }
 */
function procesarProspecto(prospect, fase) {
  // La guarda de robustez se movió a las funciones principales (procesarFase1/Fase2)
  const correo = obtenerSiguienteCorreoParaProspecto(prospect, fase);

  if (!correo || !correo.asunto || !correo.html) {
    return { enviar: false }; // No hay correo que enviar
  }

  let emailDestino;
  if (fase === 'fase1') {
    emailDestino = (prospect.EmailContacto && validarEmailBasico(prospect.EmailContacto)) 
                   ? prospect.EmailContacto 
                   : prospect.EmailGeneral;
  } else {
    emailDestino = prospect.EmailContacto;
  }

  if (emailDestino && validarEmailBasico(emailDestino)) {
    return { 
      enviar: true, 
      emailDestino: emailDestino, 
      correo: correo 
    };
  } else {
    log('WARN', `No se encontró un email de destino válido para el prospecto ID: ${prospect.ID_Empresa || prospect.ID}`);
    return { enviar: false }; // No hay un email válido a donde enviar
  }
}
// ========================================
// GESTIÓN DE OUTBOX Y ENVÍO DE CORREOS
// ========================================

function encolarCorreo(params) {
  if (!validarEmailBasico(params.toEmail)) { log('ERROR', `Email inválido: ${params.toEmail}`); return false; }
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.outbox);
  if (!sheet) return false;
  const newRow = [Utilities.getUuid(), params.prospectoId || '', params.contactoId || '', params.toEmail, params.asunto, params.html, 'pending', 0, calcularProximoEnvio(), '', params.threadId || '', ''];
  sheet.appendRow(newRow);
  return true;
}
/**
 * Procesa el Outbox de manera eficiente usando procesamiento por lotes.
 * --- ARQUITECTURA MEJORADA: PATRÓN RECOLECTAR-Y-ACTUAR ---
 * 1. Lee todos los datos del Outbox en memoria UNA SOLA VEZ.
 * 2. Itera sobre los correos pendientes, llamando a una función de envío "pura".
 * 3. Modifica la copia de los datos en memoria con los resultados.
 * 4. Escribe TODOS los cambios de vuelta a la hoja en UNA SOLA OPERACIÓN.
 */
function procesarOutbox() {
  const config = getSystemConfiguration();
  const sheet = obtenerHoja(config.sheets.outbox);
  if (!sheet || sheet.getLastRow() < 2) return;
  const ahora = new Date();

  // --- LÓGICA DE LÍMITES Y RITMO (sin cambios, sigue siendo segura) ---
  const enviadosHoy = contarEnviosHoy();
  if (enviadosHoy >= config.dailyLimit) {
    log('INFO', `Límite diario de ${config.dailyLimit} alcanzado.`);
    return;
  }
  const horaActual = parseInt(Utilities.formatDate(ahora, "Europe/Madrid", "H"), 10);
  const totalHorasJornada = config.endHour - config.startHour;
  const horasTranscurridas = Math.max(0, horaActual - config.startHour);
  const horasRestantes = totalHorasJornada - horasTranscurridas;
  const correosRestantes = config.dailyLimit - enviadosHoy;
  const dynamicHourlyLimit = (horasRestantes > 0) ? Math.ceil(correosRestantes / horasRestantes) : correosRestantes;
  const enviadosUltimaHora = contarEnviosUltimaHora();
  if (enviadosUltimaHora >= dynamicHourlyLimit) {
    log('INFO', `Ritmo de envío para la hora actual (${dynamicHourlyLimit}) alcanzado.`);
    return;
  }
  const maxEnviosPorCiclo = Math.min(5, dynamicHourlyLimit - enviadosUltimaHora);
  if (maxEnviosPorCiclo <= 0) return;

  // --- INICIO DE LA LÓGICA DE BATCH PROCESSING ---
  // PASO 1: LEER UNA VEZ (Toda la "libreta" a memoria)
  const rangoCompleto = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = rangoCompleto.getValues(); // Esta es nuestra copia en memoria.
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let enviadosEnEsteCiclo = 0;
  let cambiosRealizados = false;
  
  // PASO 2: PROCESAR EN MEMORIA (Iteramos y "anotamos" los cambios en nuestra copia `data`)
  for (let i = 0; i < data.length && enviadosEnEsteCiclo < maxEnviosPorCiclo; i++) {
    const rowData = data[i]; // La fila actual de nuestra copia en memoria.
    const correo = crearObjetoDesdeArray(headers, rowData);

    // Solo procesamos los que están pendientes y programados
    if (correo.estado === 'pending' && new Date(correo.scheduledAt) <= ahora) {
      // Marcamos como 'sending' en memoria para evitar que otro ciclo lo tome si este es largo
      rowData[headers.indexOf('estado')] = 'sending'; 
      
      const resultadoEnvio = enviarCorreo(correo); // Llamamos al "mensajero".
      
      if (resultadoEnvio.success) {
        enviadosEnEsteCiclo++;
      }
      
      // Actualizamos nuestra copia en memoria (`rowData`) con el "reporte" del mensajero
      const updateData = resultadoEnvio.updateData;
      rowData[headers.indexOf('estado')] = updateData.estado;
      rowData[headers.indexOf('intentos')] = updateData.intentos;
      rowData[headers.indexOf('sentAt')] = updateData.sentAt || rowData[headers.indexOf('sentAt')];
      rowData[headers.indexOf('messageApiId')] = updateData.messageApiId || rowData[headers.indexOf('messageApiId')];
      rowData[headers.indexOf('threadId')] = updateData.threadId || rowData[headers.indexOf('threadId')];
      
      cambiosRealizados = true; // Marcamos que nuestra "libreta" tiene cambios.
    }
  }

  // PASO 3: ESCRIBIR UNA VEZ (Si hay cambios, actualizamos la hoja de golpe)
  if (cambiosRealizados) {
    rangoCompleto.setValues(data);
    log('INFO', `Lote de Outbox procesado. ${enviadosEnEsteCiclo} correos enviados.`);
  }
}


function enviarCorreo(correo) {
  const config = CONFIG();
  // Preparamos el "reporte" con los datos por defecto
  const updateData = {
    estado: correo.estado,
    intentos: parseInt(correo.intentos || 0),
    sentAt: null,
    messageApiId: null,
    threadId: correo.threadId // Mantenemos el threadId si ya existe
  };

  try {
    const alias = config.senderAlias || Session.getActiveUser().getEmail();
    const fromAddressHeader = buildFromHeader(config.senderName || '', alias);
    
    let messageResource = {};
    if (correo.threadId && String(correo.threadId).trim() !== '') {
      messageResource.threadId = correo.threadId;
    }

    const emailRaw =
        `From: ${fromAddressHeader}\r\n` +
        `To: ${correo.toEmail}\r\n` +
        `Subject: ${encodeHeaderUtf8(String(correo.asunto || ''))}\r\n` +
        `Content-Type: text/html; charset=UTF-8\r\n\r\n` +
        `${correo.html}`;
    
    messageResource.raw = Utilities.base64Encode(emailRaw, Utilities.Charset.UTF_8).replace(/\+/g, '-').replace(/\//g, '_');
    
    // Realizamos la única acción crítica: el envío
    const sentMessage = Gmail.Users.Messages.send(messageResource, 'me');

    // Llenamos el "reporte" con los datos del éxito
    updateData.estado = 'sent';
    updateData.sentAt = new Date();
    if (sentMessage && sentMessage.id) {
      updateData.messageApiId = sentMessage.id;
    }
    if (sentMessage && sentMessage.threadId) {
      updateData.threadId = sentMessage.threadId;
      // Si era un correo nuevo (sin threadId), hacemos las actualizaciones relacionadas
      if (!correo.threadId) {
        const esFase2 = correo.contactoId && correo.contactoId.toString().trim() !== '';
        const tipoEntidad = esFase2 ? 'fase2' : 'fase1';
        const entidadId = esFase2 ? correo.contactoId : correo.prospectoId;
        crearOActualizarConversacion(entidadId, tipoEntidad, sentMessage.threadId, correo.asunto);
        actualizarThreadIdEnProspecto(correo.prospectoId, correo.contactoId, sentMessage.threadId);
      }
    }
    return { success: true, updateData: updateData };

  } catch (error) {
    const errorString = error.toString();
    log('ERROR', `Correo falló para ${correo.toEmail}: ${errorString}`);
    updateData.intentos += 1;

    // Llenamos el "reporte" con los datos del fallo
    if (esErrorPermanente(errorString)) {
      updateData.estado = 'failed';
      log('WARN', `Fallo permanente para ${correo.toEmail}.`);
      actualizarEstadoProspectoPorFallo(correo.prospectoId, correo.contactoId, correo.toEmail);
    } else {
      updateData.estado = (updateData.intentos >= 3) ? 'failed' : 'pending';
    }
    return { success: false, updateData: updateData };
  }
}


function construirCorreoBase(from, correo) {
  return (
    `MIME-Version: 1.0\r\n` +
    `From: ${from}\r\n` +
    `To: ${correo.toEmail}\r\n` +
    `Subject: ${correo.asunto}\r\n` + // ya viene codificado
    `Content-Type: text/html; charset=UTF-8\r\n\r\n` +
    `${correo.html}`
  );
}


function obtenerUltimoMessageIdDeLaOutbox(threadId) {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.outbox);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const colThreadId = headers.indexOf('threadId');
  const colMessageApiId = headers.indexOf('messageApiId');
  const colEstado = headers.indexOf('estado');
  if (colThreadId === -1 || colMessageApiId === -1 || colEstado === -1) return null;
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (row[colThreadId] === threadId && row[colEstado] === 'sent' && row[colMessageApiId]) {
      return row[colMessageApiId].toString().trim();
    }
  }
  return null;
}

// ========================================
// SISTEMA DE DETECCIÓN DE RESPUESTAS
// ========================================

function verificarRespuestas() {
  const convSheet = obtenerHoja('Conversaciones');
  if (!convSheet || convSheet.getLastRow() < 2) return;

  const data = convSheet.getDataRange().getValues();
  const headers = data.shift();
  const threadIdCol = headers.indexOf('ThreadId');
  const estadoCol = headers.indexOf('Estado');
  const estadosProcesables = ['nuevo', 'activo', 'pausado', 'Pospuesto'];

  const convUpdates = [];
  const prospectoF1Updates = [];
  const prospectoF2Updates = [];
  
  const config = CONFIG();
  const prospectosF1Sheet = obtenerHoja(config.sheets.prospectosF1);
  const prospectosF1Data = prospectosF1Sheet.getDataRange().getValues();
  const prospectosF1Headers = prospectosF1Data.shift();
  const f1Map = new Map(prospectosF1Data.map(row => [row[prospectosF1Headers.indexOf('ID_Empresa')], crearObjetoDesdeArray(prospectosF1Headers, row)]));

  const prospectosF2Sheet = obtenerHoja(config.sheets.prospectosF2);
  const prospectosF2Data = prospectosF2Sheet.getDataRange().getValues();
  const prospectosF2Headers = prospectosF2Data.shift();
  const f2Map = new Map(prospectosF2Data.map(row => [row[prospectosF2Headers.indexOf('ContactoId')], crearObjetoDesdeArray(prospectosF2Headers, row)]));

  data.forEach((row, index) => {
    const estadoActualConv = row[estadoCol];
    if (estadoActualConv === 'activo') {
      const threadId = row[threadIdCol];
      const mensajeRespuesta = tieneRespuesta(threadId);

      if (mensajeRespuesta) {
        const entidadId = row[headers.indexOf('EntidadID')];
        const tipoEntidad = row[headers.indexOf('TipoEntidad')];

        let prospectoActual;
        let estadoProspectoActual;
        if (tipoEntidad === 'fase1' && f1Map.has(entidadId)) {
            prospectoActual = f1Map.get(entidadId);
            estadoProspectoActual = prospectoActual.Estado;
        } else if (tipoEntidad === 'fase2' && f2Map.has(entidadId)) {
            prospectoActual = f2Map.get(entidadId);
            estadoProspectoActual = prospectoActual.Estado;
        }
        
        if (!estadoProspectoActual || !estadosProcesables.includes(String(estadoProspectoActual).toLowerCase())) {
            log('INFO', `Respuesta detectada para ${entidadId} pero su estado es '${estadoProspectoActual}'. Se omite la actualización.`);
            return;
        }

        convUpdates.push({
          id: entidadId,
          updates: { 'Estado': 'respondido', 'FechaUltimaActividad': new Date() }
        });

        const notaActual = prospectoActual.Notas || '';
        const timestamp = new Date().toLocaleString('es-ES');
        const nuevaNota = `${notaActual}${notaActual ? '\n' : ''}[${timestamp}] Respuesta detectada (Sistema de Conversaciones).`;

        const prospectoUpdate = {
          id: entidadId,
          updates: {
            'Estado': 'respondido',
            'RespuestaRecibida': new Date(),
            'ContenidoRespuesta': mensajeRespuesta.getPlainBody().trim(),
            'Notas': nuevaNota
          }
        };

        if (tipoEntidad === 'fase1') {
          prospectoF1Updates.push(prospectoUpdate);
        } else if (tipoEntidad === 'fase2') {
          prospectoF2Updates.push(prospectoUpdate);
        }
      }
    }
  });

  if (convUpdates.length > 0) {
    log('INFO', `Detectadas ${convUpdates.length} nuevas respuestas. Actualizando sistemas...`);
    batchUpdateSheet('Conversaciones', 'EntidadID', convUpdates);
  }
  if (prospectoF1Updates.length > 0) {
    batchUpdateSheet(config.sheets.prospectosF1, 'ID_Empresa', prospectoF1Updates);
  }
  if (prospectoF2Updates.length > 0) {
    batchUpdateSheet(config.sheets.prospectosF2, 'ContactoId', prospectoF2Updates);
  }
  
  if (convUpdates.length > 0) {
      invalidarCacheVistas();
  }
}


function tieneRespuesta(threadId) {
  try {
    if (!threadId) return null;
    const thread = GmailApp.getThreadById(threadId);
    if (!thread || thread.getMessageCount() <= 1) return null;
    const messages = thread.getMessages();
    const ultimoMensaje = messages[messages.length - 1];
    const remitenteUltimo = ultimoMensaje.getFrom().toLowerCase();
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const config = CONFIG();
    const senderAlias = config.senderAlias ? config.senderAlias.toLowerCase() : null;
    let esMiMensaje = remitenteUltimo.includes(userEmail);
    if (senderAlias && !esMiMensaje) { esMiMensaje = remitenteUltimo.includes(senderAlias); }
    if (!esMiMensaje) { return ultimoMensaje; }
    return null;
  } catch (error) {
    log('ERROR', `Error en thread ${threadId}: ${error.toString()}`);
    return null;
  }
}



// --- Helpers para cabeceras UTF-8 (RFC 2047) ---
function encodeHeaderUtf8(s) {
  if (!s) return '';
  return /[^\x00-\x7F]/.test(s)
    ? `=?UTF-8?B?${Utilities.base64Encode(s, Utilities.Charset.UTF_8)}?=`
    : s;
}

function buildFromHeader(name, email) {
  const encName = (name || '').trim() ? encodeHeaderUtf8(name.trim()) : '';
  return encName ? `${encName} <${email}>` : `<${email}>`;
}

// ========================================
// FUNCIONES DE UTILIDAD Y GESTIÓN
// ========================================

// Reemplaza la función antigua con esta
function esHorarioLaboral() {
  const config = getSystemConfiguration();
  try {
    const ahora = new Date();
    const diaSemanaMadrid = parseInt(Utilities.formatDate(ahora, "Europe/Madrid", "u"), 10);
    const horaDelDiaMadrid = parseInt(Utilities.formatDate(ahora, "Europe/Madrid", "H"), 10);

    const esDiaLaboral = config.workDays.includes(diaSemanaMadrid);
    const esHoraLaboral = horaDelDiaMadrid >= config.startHour && horaDelDiaMadrid < config.endHour;

    if (esDiaLaboral && esHoraLaboral) {
      return true;
    } else {
      log('INFO', `Fuera de horario. Día en Madrid: ${diaSemanaMadrid}, Hora en Madrid: ${horaDelDiaMadrid}. No se procesará.`);
      return false;
    }
  } catch (e) {
    log('ERROR', `Error crítico al verificar el horario laboral: ${e.toString()}`);
    return false;
  }
}

// ==================================================================
//               ↓↓↓ REEMPLAZA ESTA FUNCIÓN COMPLETA ↓↓↓
// ==================================================================
/**
 * Calcula la próxima ventana de envío disponible basándose en la configuración del sistema.
 * --- VERSIÓN CORREGIDA Y ROBUSTA ---
 * Ahora lee la configuración desde getSystemConfiguration() en lugar de usar valores fijos.
 */
function calcularProximoEnvio() {
  const ahora = new Date();
  
  // ANTERIOR: La función usaba valores fijos (hardcodeados) para horarios y días.
  // NUEVO: Obtenemos la configuración actual del sistema para usarla en los cálculos.
  const config = getSystemConfiguration();

  // Si ya estamos en horario laboral, la fecha de envío es inmediata.
  if (esHorarioLaboral()) {
    return ahora;
  }

  // Si no, calculamos la próxima ventana disponible.
  let proximaFecha = new Date();

  // Forzamos la zona horaria a Madrid para todos los cálculos
  let [year, month, day] = Utilities.formatDate(proximaFecha, "Europe/Madrid", "yyyy,MM,dd").split(',');
  
  // Creamos la fecha candidata para la hora de inicio de hoy en Madrid
  proximaFecha = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), config.startHour, 0, 0);

  // Si la hora de inicio de hoy ya pasó O hoy no es un día laboral, empezamos a buscar desde mañana.
  const hoyEnNumero = parseInt(Utilities.formatDate(new Date(), "Europe/Madrid", "u"), 10);
  if (new Date() > proximaFecha || !config.workDays.includes(hoyEnNumero)) {
      proximaFecha.setDate(proximaFecha.getDate() + 1);
      // Reiniciamos la hora a la de inicio del nuevo día para ser precisos
      proximaFecha.setHours(config.startHour, 0, 0, 0);
  }

  // Avanzamos día por día hasta encontrar el próximo día laboral configurado.
  while (!config.workDays.includes(parseInt(Utilities.formatDate(proximaFecha, "Europe/Madrid", "u"), 10))) {
    proximaFecha.setDate(proximaFecha.getDate() + 1);
  }

  log('INFO', `Próximo envío calculado para: ${proximaFecha.toLocaleString('es-ES')}`);
  return proximaFecha;
}

function contarEnviosHoy() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.outbox);
  if (!sheet) return 0;
  const data = sheet.getRange('J2:J').getValues();
  const hoy = new Date().setHours(0, 0, 0, 0);
  return data.filter(d => d[0] && new Date(d[0]).setHours(0, 0, 0, 0) === hoy).length;
}

function contarEnviosUltimaHora() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.outbox);
  if (!sheet) return 0;
  const data = sheet.getRange('J2:J').getValues();
  const unaHoraAtras = Date.now() - 3600000;
  return data.filter(d => d[0] && new Date(d[0]).getTime() > unaHoraAtras).length;
}

function obtenerHoja(nombreHoja) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
}

function obtenerIndiceColumna(sheet, nombreColumna) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(nombreColumna) + 1;
}

function crearObjetoDesdeArray(headers, row, rowIndex) {
  const obj = {};
  headers.forEach((header, index) => { obj[header] = row[index]; });
  if (rowIndex && !isNaN(rowIndex) && rowIndex > 0) {
    obj.__rowNumber__ = rowIndex;
  }
  return obj;
}

function log(tipo, mensaje) {
  console.log(`[${tipo}] ${mensaje}`);
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.logs);
  if (sheet) { sheet.appendRow([new Date(), tipo, mensaje]); }
}

function obtenerDatosEmpresa(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return crearObjetoDesdeArray(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
}

function obtenerDatosContacto(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return crearObjetoDesdeArray(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
}




function crearProspectoFase2DesdeContacto(contactoData) {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.prospectosFase2);
  if (!sheet) { log('ERROR', `No se encontró hoja: ${config.sheets.prospectosFase2}`); return false; }
  
  const headers = ESQUEMAS.Prospectos_Fase2;
  const prospectoData = new Array(headers.length).fill('');

  // Mapeo de datos desde el contacto al nuevo prospecto
  prospectoData[headers.indexOf('ID')] = contactoData.EmpresaId.toString().trim();
  prospectoData[headers.indexOf('EmpresaId')] = contactoData.EmpresaId || '';
  prospectoData[headers.indexOf('ContactoId')] = contactoData.ContactoId || '';
  prospectoData[headers.indexOf('NombreEmpresa')] = contactoData.NombreEmpresa || '';
  prospectoData[headers.indexOf('PAIS')] = contactoData.PAIS || ''; // --- LÍNEA AÑADIDA ---
  prospectoData[headers.indexOf('NombreContacto')] = contactoData.NombreContacto || '';
  prospectoData[headers.indexOf('EmailContacto')] = contactoData.EmailContacto || '';
  prospectoData[headers.indexOf('SeguimientoActual')] = 0;
  prospectoData[headers.indexOf('Estado')] = 'nuevo';
  
  sheet.appendRow(prospectoData);
  return true;
}
function crearProspectoDesdeEmpresa(empresaData, sheet) { // Acepta la hoja como parámetro
  if (!sheet) return false;
  
  const headers = ESQUEMAS.Prospectos_Fase1;
  const prospectoData = new Array(headers.length).fill('');

  prospectoData[headers.indexOf('ID_Empresa')] = empresaData.ID_Empresa.toString().trim();
  prospectoData[headers.indexOf('NombreEmpresa')] = empresaData.NombreEmpresa || '';
  prospectoData[headers.indexOf('PAIS')] = empresaData.PAIS || '';
  prospectoData[headers.indexOf('EmailGeneral')] = (validarEmailBasico(empresaData.EmailGeneral)) ? empresaData.EmailGeneral : '';
  prospectoData[headers.indexOf('EmailContacto')] = empresaData.EmailContacto || '';
  prospectoData[headers.indexOf('SeguimientoActual')] = 0;
  prospectoData[headers.indexOf('Estado')] = 'nuevo';
  
  sheet.appendRow(prospectoData);
  return true;
}

// ==================================================================
//        ↓↓↓ REEMPLAZA TU crearProspectoFase2DesdeContacto CON ESTO ↓↓↓
// ==================================================================
function crearProspectoFase2DesdeContacto(contactoData, sheet) { // Acepta la hoja como parámetro
  if (!sheet) { 
    log('ERROR', `No se encontró hoja para crear prospecto F2`); 
    return false; 
  }
  
  const headers = ESQUEMAS.Prospectos_Fase2;
  const prospectoData = new Array(headers.length).fill('');

  prospectoData[headers.indexOf('ID')] = contactoData.EmpresaId.toString().trim();
  prospectoData[headers.indexOf('EmpresaId')] = contactoData.EmpresaId || '';
  prospectoData[headers.indexOf('ContactoId')] = contactoData.ContactoId || '';
  prospectoData[headers.indexOf('NombreEmpresa')] = contactoData.NombreEmpresa || '';
  prospectoData[headers.indexOf('PAIS')] = contactoData.PAIS || '';
  prospectoData[headers.indexOf('NombreContacto')] = contactoData.NombreContacto || '';
  prospectoData[headers.indexOf('EmailContacto')] = contactoData.EmailContacto || '';
  prospectoData[headers.indexOf('SeguimientoActual')] = 0;
  prospectoData[headers.indexOf('Estado')] = 'nuevo';
  
  sheet.appendRow(prospectoData);
  return true;
}



function validarEmailBasico(email) {
  return email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

// ==================================================================
//               ↓↓↓ REEMPLAZA ESTA FUNCIÓN COMPLETA ↓↓↓
// ==================================================================
/**
 * Actualiza el ThreadId en la hoja de un prospecto específico.
 * --- MODIFICADO: Ahora es una función de apoyo, no la principal. ---
 * El sistema principal de respuestas ahora usa la hoja 'Conversaciones'.
 */
function actualizarThreadIdEnProspecto(prospectoId, contactoId, threadId) {
  const config = CONFIG();
  const esFase2 = contactoId && contactoId.toString().trim() !== '';
  const sheet = obtenerHoja(esFase2 ? config.sheets.prospectosFase2 : config.sheets.prospectosFase1);
  if (!sheet) return false;

  const clave = esFase2 ? String(contactoId).trim() : String(prospectoId).trim();
  const nombreColClave = esFase2 ? 'ContactoId' : 'ID_Empresa';
  
  const colClave = obtenerIndiceColumna(sheet, nombreColClave);
  const colThread = obtenerIndiceColumna(sheet, 'ThreadId');
  if (colClave < 1 || colThread < 1) return false;

  const last = sheet.getLastRow();
  if (last < 2) return false;

  const vals = sheet.getRange(2, colClave, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === clave) {
      sheet.getRange(i + 2, colThread).setValue(threadId);
      return true;
    }
  }
  return false;
}

/**
 * Crea un nuevo contacto en la hoja 'Contactos' a partir de un prospecto de Fase 1.
 * Si el prospecto tiene una respuesta, la usa como contexto.
 * VERSIÓN MODIFICADA: Se ha eliminado la lógica que dependía de la columna "Gancho".
 */
function iniciarFase2DesdeProspecto(prospectoFase1Data) {
  try {
    const config = CONFIG();
    const contactosSheet = obtenerHoja(config.sheets.contactos);
    if (!contactosSheet) {
      log('ERROR', 'Faltan la hoja Contactos.');
      return null;
    }
    // --- LÓGICA MODIFICADA ---
    // Ahora toma el ID de la empresa desde la columna 'ID_Empresa'.
    const empresaId = prospectoFase1Data.ID_Empresa;
    // --- FIN DE LA LÓGICA MODIFICADA ---
    let valorParaNuevaColumna = '';
    const contenidoRespuesta = prospectoFase1Data.ContenidoRespuesta ? prospectoFase1Data.ContenidoRespuesta.toString().trim() : '';
    // ... (el resto de la función no cambia) ...
    // ...
    newRow[headersContactos.indexOf('EmpresaId')] = empresaId;
    // ...
    // ... (el resto de la función no cambia) ...
  } catch (e) {
    // ...
  }
}

function generarSiguienteContactoId() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.contactos);
  if (!sheet) return null;
  const todosLosIds = sheet.getRange('A2:A').getValues().flat().filter(String);
  if (todosLosIds.length === 0) { return 'C-001'; }
  let maxNum = 0;
  todosLosIds.forEach(id => {
    const numero = parseInt(id.split('-')[1]);
    if (numero > maxNum) { maxNum = numero; }
  });
  const nuevoNumero = maxNum + 1;
  const numeroFormateado = String(nuevoNumero).padStart(3, '0');
  return `C-${numeroFormateado}`;
}

// ========================================
// SINCRONIZACIÓN
// ========================================


// =======================================================
// --- LÓGICA DEL IMPORTADOR INTELIGENTE DE EXCEL ---
// =======================================================

function mostrarSidebarImportador() {
  const html = HtmlService.createHtmlOutputFromFile('ImportarSidebar').setTitle('Importador de Empresas');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Importa empresas desde un archivo Excel, omitiendo duplicados y reportando el resultado.
 * VERSIÓN MEJORADA Y ROBUSTA.
 */
/**
 * Importa empresas desde un archivo Excel, omitiendo duplicados y reportando el resultado.
 * VERSIÓN MEJORADA Y ROBUSTA.
 */
function importarArchivoExcel(fileInfo) {
  let tempFileId = null, tempSheetId = null;
  try {
    const decodedBytes = Utilities.base64Decode(fileInfo.base64Data);
    const blob = Utilities.newBlob(decodedBytes, fileInfo.mimeType, fileInfo.fileName);
    const tempFile = DriveApp.getRootFolder().createFile(blob);
    tempFileId = tempFile.getId();
    const convertedFile = Drive.Files.copy({ title: `[TEMP] ${fileInfo.fileName}`, mimeType: MimeType.GOOGLE_SHEETS }, tempFileId);
    tempSheetId = convertedFile.id;
    
    const datosAImportar = SpreadsheetApp.openById(tempSheetId).getSheets()[0].getDataRange().getValues();
    if (datosAImportar.length <= 1) {
      return { error: true, message: 'El archivo está vacío o solo contiene encabezados.' };
    }
    const encabezadosImportados = datosAImportar.shift();

    const config = CONFIG();
    const empresasSheet = obtenerHoja(config.sheets.empresas);
    const datosExistentes = empresasSheet.getDataRange().getValues();
    const encabezadosExistentes = datosExistentes.shift() || [];
    
    const indiceCif = encabezadosExistentes.indexOf('CIF_NIF');
    const indiceWeb = encabezadosExistentes.indexOf('Web');
    const indiceEmail = encabezadosExistentes.indexOf('EmailGeneral');
    
    const cifsExistentes = new Set(datosExistentes.map(row => row[indiceCif]).filter(String));
    const websExistentes = new Set(datosExistentes.map(row => row[indiceWeb]).filter(String));
    const emailsExistentes = new Set(datosExistentes.map(row => row[indiceEmail]).filter(String));

    const indiceCifImportado = encabezadosImportados.indexOf('CIF_NIF');
    const indiceWebImportado = encabezadosImportados.indexOf('Web');
    const indiceEmailImportado = encabezadosImportados.indexOf('EmailGeneral');
    
    const nuevasFilas = [];
    let siguienteIdNum = generarSiguienteEmpresaId(true);
    let duplicadosOmitidos = 0;

    for (const filaAImportar of datosAImportar) {
      const cif = filaAImportar[indiceCifImportado];
      const web = filaAImportar[indiceWebImportado];
      const email = filaAImportar[indiceEmailImportado];

      if ((cif && cifsExistentes.has(cif)) || (web && websExistentes.has(web)) || (email && emailsExistentes.has(email))) {
        duplicadosOmitidos++;
        continue; // Omite esta iteración y pasa a la siguiente fila
      }

      const nuevaFila = [];
      const nuevoId = `EMP-${String(siguienteIdNum).padStart(3, '0')}`;
      ESQUEMAS.Empresas.forEach(header => {
        if (header === 'ID_Empresa') {
          nuevaFila.push(nuevoId);
        } else {
          const indiceEnImportado = encabezadosImportados.indexOf(header);
          nuevaFila.push(indiceEnImportado !== -1 ? filaAImportar[indiceEnImportado] : '');
        }
      });
      nuevasFilas.push(nuevaFila);
      siguienteIdNum++;
    }

    if (nuevasFilas.length > 0) {
      empresasSheet.getRange(empresasSheet.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
      SpreadsheetApp.flush();
      Utilities.sleep(1500);
      sincronizarNuevasEmpresas(); // Esta función ya invalida la caché
    }

    const totalProcesado = datosAImportar.length;
    const nuevasImportadas = nuevasFilas.length;
    let mensajeFinal = `¡Proceso completado! \n- Filas procesadas: ${totalProcesado} \n- Empresas nuevas importadas: ${nuevasImportadas} \n- Empresas omitidas (duplicados): ${duplicadosOmitidos}`;

    return { error: false, message: mensajeFinal };
    
  } catch (e) {
    return { error: true, message: `Error inesperado: ${e.message}` };
  } finally {
    if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
    if (tempSheetId) DriveApp.getFileById(tempSheetId).setTrashed(true);
  }
}

function generarSiguienteEmpresaId(soloNumero = false) {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.empresas);
  if (!sheet) return soloNumero ? 1 : 'EMP-001';
  const todosLosIds = sheet.getRange('A2:A').getValues().flat().filter(String);
  if (todosLosIds.length === 0) { return soloNumero ? 1 : 'EMP-001'; }
  let maxNum = 0;
  todosLosIds.forEach(id => {
    const numero = parseInt(id.split('-')[1]);
    if (!isNaN(numero) && numero > maxNum) { maxNum = numero; }
  });
  const nuevoNumero = maxNum + 1;
  if (soloNumero) { return nuevoNumero; }
  return `EMP-${String(nuevoNumero).padStart(3, '0')}`;
}

// =======================================================
// --- LÓGICA DEL IMPORTADOR DE PLANTILLAS (Fase 1 y 2) ---
// =======================================================

function mostrarSidebarImportadorF1() {
  const t = HtmlService.createTemplateFromFile('PlantillasSidebar');
  t.titulo = 'Importador Plantillas (Fase 1)';
  t.funcionServidor = 'importarPlantillasFase1';
  const html = t.evaluate().setTitle('Importador Fase 1');
  SpreadsheetApp.getUi().showSidebar(html);
}

function mostrarSidebarImportadorF2() {
  const t = HtmlService.createTemplateFromFile('PlantillasSidebar');
  t.titulo = 'Importador Plantillas (Fase 2)';
  t.funcionServidor = 'importarPlantillasFase2';
  const html = t.evaluate().setTitle('Importador Fase 2');
  SpreadsheetApp.getUi().showSidebar(html);
}

function importarPlantillasGenerico(fileInfo, fase) {
  let tempFileId = null, tempSheetId = null;
  try {
    const decodedBytes = Utilities.base64Decode(fileInfo.base64Data);
    const blob = Utilities.newBlob(decodedBytes, fileInfo.mimeType, fileInfo.fileName);
    const tempFile = DriveApp.getRootFolder().createFile(blob);
    tempFileId = tempFile.getId();
    const convertedFile = Drive.Files.copy({ title: `[TEMP] ${fileInfo.fileName}`, mimeType: MimeType.GOOGLE_SHEETS }, tempFileId);
    tempSheetId = convertedFile.id;
    const datosAImportar = SpreadsheetApp.openById(tempSheetId).getSheets()[0].getDataRange().getValues();
    const encabezadosImportados = datosAImportar.shift();
    const config = CONFIG();
    const nombreHojaProspectos = (fase === 'fase1') ? config.sheets.prospectosFase1 : config.sheets.prospectosFase2;
    const esquemaEsperado = (fase === 'fase1') ? ESQUEMAS.Prospectos_Fase1 : ESQUEMAS.Prospectos_Fase2;
    const columnasPlantillas = esquemaEsperado.filter(h => h.includes('Asunto') || h.includes('HtmlCuerpo'));
    
    // --- LÓGICA MODIFICADA ---
    // Las columnas que identifican a un prospecto de Fase 1 ahora usan 'ID_Empresa'.
    const columnasIdentificadoras = (fase === 'fase1') ? ['ID_Empresa', 'NombreEmpresa'] : ['ID', 'EmpresaId', 'ContactoId', 'NombreEmpresa'];
    // --- FIN DE LA LÓGICA MODIFICADA ---

    const encabezadosEsperados = [...columnasIdentificadoras, ...columnasPlantillas];

    const columnasFaltantes = encabezadosEsperados.filter(h => !encabezadosImportados.includes(h));

    if (columnasFaltantes.length > 0) {
      return { error: true, message: `Estructura de archivo incorrecta. Faltan las siguientes columnas obligatorias: ${columnasFaltantes.join(', ')}` };
    }
    
    const prospectosSheet = obtenerHoja(nombreHojaProspectos);
    const prospectosData = prospectosSheet.getDataRange().getValues();
    const prospectosHeaders = prospectosData.shift();

    // --- LÓGICA MODIFICADA ---
    // El mapa de prospectos ahora se crea usando la clave 'ID_Empresa' para la Fase 1.
    const idColumnName = (fase === 'fase1') ? 'ID_Empresa' : 'ID';
    const prospectosMap = new Map(prospectosData.map((row, index) => [row[prospectosHeaders.indexOf(idColumnName)], { ...crearObjetoDesdeArray(prospectosHeaders, row), __rowNumber__: index + 2 }]));
    // --- FIN DE LA LÓGICA MODIFICADA ---
    
    const datosParaActualizar = [];
    for (let i = 0; i < datosAImportar.length; i++) {
      const filaImportada = crearObjetoDesdeArray(encabezadosImportados, datosAImportar[i]);
      
      // --- LÓGICA MODIFICADA ---
      // Buscamos en el mapa usando la clave correcta.
      const idDeBusqueda = (fase === 'fase1') ? filaImportada.ID_Empresa : filaImportada.ID;
      const prospectoExistente = prospectosMap.get(idDeBusqueda);
      // --- FIN DE LA LÓGICA MODIFICADA ---

      if (!prospectoExistente || !columnasIdentificadoras.every(idCol => String(prospectoExistente[idCol]) === String(filaImportada[idCol]))) {
        return { error: true, message: `Error en Fila ${i + 2}: El prospecto con ID "${idDeBusqueda}" no se encontró o los datos de identificación no coinciden.` };
      }
      datosParaActualizar.push({ rowIndex: prospectoExistente.__rowNumber__, data: filaImportada });
    }
    
    if (datosParaActualizar.length > 0) {
      const prospectosHeadersOriginales = prospectosSheet.getRange(1, 1, 1, prospectosSheet.getLastColumn()).getValues()[0];
      datosParaActualizar.forEach(item => {
        columnasPlantillas.forEach(nombreColumna => {
          const colIndex = prospectosHeadersOriginales.indexOf(nombreColumna) + 1;
          if (colIndex > 0 && item.data[nombreColumna] !== undefined) {
            prospectosSheet.getRange(item.rowIndex, colIndex).setValue(item.data[nombreColumna]);
          }
        });
      });
    }
    return { error: false, message: `Plantillas actualizadas para ${datosParaActualizar.length} prospectos.` };
  } catch (e) {
    return { error: true, message: `Error inesperado: ${e.message}` };
  } finally {
    if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
    if (tempSheetId) DriveApp.getFileById(tempSheetId).setTrashed(true);
  }
}

function importarPlantillasFase1(fileInfo) {
  return importarPlantillasGenerico(fileInfo, 'fase1');
}

function importarPlantillasFase2(fileInfo) {
  return importarPlantillasGenerico(fileInfo, 'fase2');
}

// ========================================
// FUNCIONES DE MENÚ Y DIAGNÓSTICO
// ========================================

function iniciarAutomatizacionCompleta() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('🚀 AUTOMATIZACIÓN COMPLETA', 'Esto activará el sistema en modo PRODUCCIÓN. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (respuesta === ui.Button.YES) {
    activarAutomatizacion();
    ui.alert('Sistema Activado. La automatización se ejecutará cada 15 minutos.');
  }
}








function generarDiagnosticoCompleto() {
  const config = getSystemConfiguration();
  const estado = {};
  estado.hojas = {};
  Object.keys(config.sheets).forEach(key => {
    const nombreHoja = config.sheets[key];
    estado.hojas[nombreHoja] = obtenerHoja(nombreHoja) ? '✅ OK' : '❌ FALTANTE';
  });
  estado.triggers = obtenerEstadoTriggers();
  estado.limites = { hoy: `${contarEnviosHoy()} / ${config.dailyLimit}` };
  estado.recuentos = {
    empresas: (obtenerHoja(config.sheets.empresas)?.getLastRow() - 1) || 0,
    contactos: (obtenerHoja(config.sheets.contactos)?.getLastRow() - 1) || 0,
    outboxPendientes: obtenerHoja(config.sheets.outbox)?.getRange('G2:G').getValues().filter(v => v[0] === 'pending').length || 0,
  };
  return estado;
}
function formatearDiagnosticoCompleto(diagnostico) {
  let mensaje = "--- ESTADO DEL SISTEMA ---\n\n";
  mensaje += "▪️ Hojas Requeridas:\n";
  Object.keys(diagnostico.hojas).forEach(nombre => { mensaje += `  - ${nombre}: ${diagnostico.hojas[nombre]}\n`; });
  mensaje += "\n▪️ Automatización (Triggers):\n";
  mensaje += `  - Principal: ${diagnostico.triggers.automatizacion ? '✅ Activado' : '❌ Desactivado'}\n`;
  mensaje += `  - Limpieza de Logs: ${diagnostico.triggers.limpiezaLogs ? '✅ Activado' : '❌ Desactivado'}\n`;
  mensaje += "\n▪️ Límite de Envío Diario:\n";
  mensaje += `  - Progreso hoy: ${diagnostico.limites.hoy}\n`;
  mensaje += "\n▪️ Recuento de Datos:\n";
  mensaje += `  - Empresas: ${diagnostico.recuentos.empresas}\n`;
  mensaje += `  - Contactos: ${diagnostico.recuentos.contactos}\n`;
  mensaje += `  - Correos pendientes: ${diagnostico.recuentos.outboxPendientes}\n`;
  return mensaje;
}

// ========================================
// AUTOMATIZACIÓN, ESQUEMAS Y OTROS
// ========================================
// ==================================================================
//               ↓↓↓ AÑADE LA LÍNEA 'Conversaciones' A TUS ESQUEMAS ↓↓↓
// ==================================================================
const ESQUEMAS = {
  Empresas: [
    'ID_Empresa', 'NombreEmpresa', 'PAIS', 'Region_Estado', 'Subdivision_Menor', 'Ciudad', 'Dirección', 'CIF_NIF', 'Sector', 'Subsector', 'Web', 'EmailGeneral', 'TelefonoGeneral', 'informe', 'Notas', 'MarcasExtranjeras', 'ContactoClave', 'EmailContacto', 'FuenteContacto'
  ],
  Contactos: [
    'ContactoId', 'EmpresaId', 'NombreEmpresa', 'PAIS', 'NombreContacto', 'Cargo', 'Departamento', 'EmailContacto', 'TelefonoContacto', 'Respuesta Hilo Completo', 'Estado', 'Notas'
  ],
  Clientes: [
    'ClienteID', 'ID_Empresa_Original', 'NombreEmpresa', 'PAIS', 'ContactoPrincipal', 'EmailPrincipal', 'TelefonoPrincipal', 'FechaConversion', 'ServicioContratado', 'NotasCliente', 'ThreadId'
  ],
  // --- INICIO DE LA LÍNEA AÑADIDA ---
  Conversaciones: ['ConvID', 'EntidadID', 'TipoEntidad', 'ThreadId', 'Estado', 'AsuntoUltimoCorreo', 'FechaUltimaActividad'],
  // --- FIN DE LA LÍNEA AÑADIDA ---
  Prospectos_Fase1: [
    'ID_Empresa', 'NombreEmpresa', 'PAIS', 'EmailGeneral', 'EmailContacto', 'AsuntoCorreoInicial', 'HtmlCuerpoCorreoInicial', 'AsuntoSeguimiento1', 'HtmlCuerpoSeguimiento1', 'AsuntoSeguimiento2', 'HtmlCuerpoSeguimiento2', 'AsuntoSeguimiento3', 'HtmlCuerpoSeguimiento3', 'FechaUltimoEnvio', 'SeguimientoActual', 'Estado', 'ThreadId', 'RespuestaRecibida', 'Notas', 'ContenidoRespuesta', 'IniciarFase2', 'Gestionar Prospecto', 'Última Acción Manual', 'FechaReactivacion'
  ],
  Prospectos_Fase2: [
    'ID', 'EmpresaId', 'ContactoId', 'NombreEmpresa', 'PAIS', 'NombreContacto', 'EmailContacto', 'AsuntoCorreoInicial', 'HtmlCuerpoCorreoInicial', 'AsuntoSeguimientoF2_1', 'HtmlCuerpoSeguimientoF2_1', 'AsuntoSeguimientoF2_2', 'HtmlCuerpoSeguimientoF2_2', 'AsuntoSeguimientoF2_3', 'HtmlCuerpoSeguimientoF2_3', 'FechaUltimoEnvio', 'SeguimientoActual', 'Estado', 'ThreadId', 'RespuestaRecibida', 'Notas', 'ContenidoRespuesta', 'Gestionar Prospecto', 'Última Acción Manual', 'FechaReactivacion'
  ],
  Outbox: ['envioId', 'prospectoId', 'contactoId', 'toEmail', 'asunto', 'html', 'estado', 'intentos', 'scheduledAt', 'sentAt', 'threadId', 'messageApiId'],
  Logs_Sistema: ['Timestamp', 'Tipo', 'Mensaje'],
  Queue_System: ['Timestamp', 'TaskId', 'Type', 'Payload', 'Status', 'ResultMessage']
};
/**
 * Instala el activador diario para la limpieza de archivos temporales.
 * La función es idempotente: busca y elimina cualquier activador antiguo
 * para la misma función antes de crear uno nuevo.
 */
function instalarActivadorDeLimpieza() {
  const nombreFuncionLimpieza = 'limpiarArchivosTemporales';
  
  // 1. Hacerla Idempotente: Buscar y eliminar activadores existentes.
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === nombreFuncionLimpieza) {
      ScriptApp.deleteTrigger(trigger);
      log('INFO', `Se ha eliminado un activador de limpieza antiguo para evitar duplicados.`);
    }
  }

  // 2. Crear el Nuevo Activador diario.
  ScriptApp.newTrigger(nombreFuncionLimpieza)
    .timeBased()
    .everyDays(1)
    .atHour(3) // Se ejecutará entre las 3 y 4 a.m.
    .create();

  log('SUCCESS', `Activador de limpieza diaria instalado correctamente para la función '${nombreFuncionLimpieza}'.`);

  // 3. Proporcionar Feedback al Usuario.
  SpreadsheetApp.getUi().alert(
    'Configuración Completada',
    'El activador de limpieza automática ha sido instalado correctamente. Se ejecutará una vez al día para eliminar archivos de exportación antiguos.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}



// Reemplaza la función antigua con esta
function desactivarAutomatizacion() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (['ejecutarAutomatizacion'].includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // --- NUEVO: Invalidar caché de configuración para actualizar automationStatus ---
  CacheService.getScriptCache().remove('system_config');
  log('INFO', 'Automatización desactivada.');
  return { success: true, status: 'inactivo' };
}

function obtenerEstadoTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const estado = { automatizacion: false, limpiezaLogs: false };
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'ejecutarAutomatizacion') estado.automatizacion = true;
    if (funcName === 'limpiarLogsViejos') estado.limpiezaLogs = true;
  });
  return estado;
}

function limpiarLogsViejos() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.logs);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange('A2:A').getValues();
  const hace30Dias = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  let filasAEliminar = data.filter(d => d[0] && new Date(d[0]) < hace30Dias).length;
  if (filasAEliminar > 0) {
    sheet.deleteRows(2, filasAEliminar);
  }
}
function crearEsquema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // Definimos aquí las hojas que son dinámicas y cuyos encabezados no deben ser alterados.
    const hojasSistemaProtegidas = ['Queue_System', 'Logs_Sistema'];

    Object.keys(ESQUEMAS).forEach(nombreHoja => {
      let sheet = ss.getSheetByName(nombreHoja);
      
      if (!sheet) {
        // La hoja no existe, la creamos sin importar de qué tipo sea.
        sheet = ss.insertSheet(nombreHoja);
        log('INFO', `Hoja '${nombreHoja}' creada.`);
        if (ESQUEMAS[nombreHoja]) {
          const columnas = ESQUEMAS[nombreHoja];
          sheet.getRange(1, 1, 1, columnas.length).setValues([columnas]);
          formatearEncabezados(sheet, columnas.length);
        }
      } else {
        // La hoja ya existe, ahora decidimos si actualizamos los encabezados.
        
        // --- INICIO DE LA LÓGICA DE PROTECCIÓN ---
        // Si es una hoja protegida, no hacemos nada más y pasamos a la siguiente.
        if (hojasSistemaProtegidas.includes(nombreHoja)) {
          log('INFO', `Se omite la actualización de encabezados para la hoja de sistema protegida: '${nombreHoja}'.`);
          return; 
        }
        // --- FIN DE LA LÓGICA DE PROTECCIÓN ---

        // Si NO es una hoja protegida, verificamos y actualizamos si es necesario.
        if (ESQUEMAS[nombreHoja]) {
          const columnas = ESQUEMAS[nombreHoja];
          const encabezadosActuales = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          let necesitaActualizacion = columnas.length > encabezadosActuales.length || columnas.some((col, i) => col !== encabezadosActuales[i]);
          
          if (necesitaActualizacion) {
            sheet.getRange(1, 1, 1, columnas.length).setValues([columnas]);
            formatearEncabezados(sheet, columnas.length);
            log('INFO', `Encabezados actualizados para '${nombreHoja}'.`);
          }
        }
      }
    });

    configurarFormatosGenerales();
    ui.alert('Esquema verificado. Todas las hojas están creadas y configuradas correctamente.');

  } catch (e) {
    log('ERROR', `Error crítico en crearEsquema: ${e.toString()}`);
    ui.alert(`Error Crítico al crear el esquema: ${e.message}. Revisa los logs.`);
  }
}



function configurarFormatosGenerales() {
  const config = CONFIG();
  const outboxSheet = obtenerHoja(config.sheets.outbox);
  if (outboxSheet) {
    const range = outboxSheet.getRange('G2:G');
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('sent').setBackground('#d9ead3').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('failed').setBackground('#f4cccc').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pending').setBackground('#fff2cc').setRanges([range]).build()
    ];
    outboxSheet.setConditionalFormatRules(rules);
  }
  const logsSheet = obtenerHoja(config.sheets.logs);
  if (logsSheet) {
    logsSheet.setColumnWidth(1, 150);
    logsSheet.setColumnWidth(2, 80);
    logsSheet.setColumnWidth(3, 400);
  }
}
/** 
function crearOVerificarHoja(nombreHoja, columnas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) { sheet = ss.insertSheet(nombreHoja); }
  const encabezadosActuales = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let necesitaActualizacion = columnas.length > encabezadosActuales.length || columnas.some((col, i) => col !== encabezadosActuales[i]);
  if (necesitaActualizacion) {
    sheet.getRange(1, 1, 1, columnas.length).setValues([columnas]);
    formatearEncabezados(sheet, columnas.length);
  }
}*/

function formatearEncabezados(sheet, numColumnas) {
  sheet.setFrozenRows(1);
  const headerRange = sheet.getRange(1, 1, 1, numColumnas);
  headerRange.setFontWeight('bold').setBackground('#e0e0e0');
}

function aplicarFormatosYValidaciones() {
  const ui = SpreadsheetApp.getUi();
  try {
    configurarFormatosGenerales();
    const config = CONFIG();
    const prospectosF1Sheet = obtenerHoja(config.sheets.prospectosFase1);
    const prospectosF2Sheet = obtenerHoja(config.sheets.prospectosFase2);
    // --- LÍNEA MODIFICADA: Añadimos la nueva opción al menú ---
    const opcionesGestionar = ['', 'NO CONTACTAR', 'Reactivar Secuencia', 'Pausar (Manual)', '✅ CONVERTIR A CLIENTE'];
    const reglaGestionar = SpreadsheetApp.newDataValidation().requireValueInList(opcionesGestionar, true).build();
    const reglaIniciarF2 = SpreadsheetApp.newDataValidation().requireValueInList(['', 'SÍ'], true).build();

    if (prospectosF1Sheet) {
      const colGestionarF1 = obtenerIndiceColumna(prospectosF1Sheet, 'Gestionar Prospecto');
      const colIniciarF2 = obtenerIndiceColumna(prospectosF1Sheet, 'IniciarFase2');
      if (colGestionarF1 > 0) prospectosF1Sheet.getRange(2, colGestionarF1, prospectosF1Sheet.getMaxRows() - 1).setDataValidation(reglaGestionar);
      if (colIniciarF2 > 0) prospectosF1Sheet.getRange(2, colIniciarF2, prospectosF1Sheet.getMaxRows() - 1).setDataValidation(reglaIniciarF2);
      
      aplicarReglasDeColorPorEstado(prospectosF1Sheet);
    }
    
    if (prospectosF2Sheet) {
      const colGestionarF2 = obtenerIndiceColumna(prospectosF2Sheet, 'Gestionar Prospecto');
      if (colGestionarF2 > 0) prospectosF2Sheet.getRange(2, colGestionarF2, prospectosF2Sheet.getMaxRows() - 1).setDataValidation(reglaGestionar);
      
      aplicarReglasDeColorPorEstado(prospectosF2Sheet);
    }
    ui.alert('Éxito', 'Formatos, menús y reglas de color automáticas aplicadas.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Ocurrió un error: ${e.toString()}`, ui.ButtonSet.OK);
  }
}




    /**
 * Protege la primera fila (encabezados) de todas las hojas definidas en CONFIG.
 * Solo el propietario de la hoja de cálculo podrá editar los encabezados.
 */




function activarProteccionModoSeguro() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = CONFIG();
    const sheetNames = Object.values(config.sheets);
    let protectedCount = 0;

    ui.alert('Activando Modo Seguro. Se añadirá una advertencia al intentar editar los encabezados.');

    sheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        // Primero, eliminamos protecciones viejas para no duplicarlas
        sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
          if (p.getRange().getRow() === 1 && p.getRange().getNumRows() === 1) p.remove();
        });
        
        // Creamos la nueva protección
        const protection = sheet.getRange('1:1').protect();
        
        // --- LA LÍNEA CLAVE ---
        // En lugar de bloquear, solo mostramos una advertencia.
        protection.setWarningOnly(true);
        
        protection.setDescription(`Protección de encabezado (Modo Seguro) para: ${sheetName}`);
        protectedCount++;
      }
    });

    log('INFO', `Se han protegido con advertencia ${protectedCount} encabezados.`);
    ui.alert('Modo Seguro Activado', `Se ha activado la protección con advertencia para ${protectedCount} hojas.`, ui.ButtonSet.OK);

  } catch (e) {
    log('ERROR', `Error al activar la protección: ${e.toString()}`);
    ui.alert('Error', `No se pudo activar la protección: ${e.message}.`);
  }
}

/**
 * Elimina la protección de los encabezados para permitir la edición libre.
 */
function desactivarProteccionModoEdicion() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = CONFIG();
    const sheetNames = Object.values(config.sheets);
    let unprotectedCount = 0;

    const respuesta = ui.alert(
      '🔓 Desactivar Protección (Modo Edición)', 
      'Esto eliminará la advertencia de los encabezados, permitiendo que se editen libremente. ¿Continuar?', 
      ui.ButtonSet.YES_NO
    );

    if (respuesta !== ui.Button.YES) return;

    sheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
          if (p.getRange().getRow() === 1 && p.getRange().getNumRows() === 1) {
            p.remove();
            unprotectedCount++;
          }
        });
      }
    });

    log('INFO', `Se ha desactivado la protección de ${unprotectedCount} encabezados.`);
    ui.alert('Modo Edición Activado', `Se ha eliminado la protección de los encabezados. Ahora son editables sin advertencia.`, ui.ButtonSet.OK);

  } catch (e) {
    log('ERROR', `Error al desactivar la protección: ${e.toString()}`);
    ui.alert('Error', `No se pudo desactivar la protección: ${e.message}.`);
  }
}



// En el editor de Apps Script, ejecuta:
function verificarListoParaEnvio() {
  const sheet = obtenerHoja('Prospectos_Fase1');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  let listos = 0;
  data.forEach(row => {
    const prospect = crearObjetoDesdeArray(headers, row);
    if (prospect.Estado === 'nuevo' && prospect.EmailGeneral && 
        prospect.AsuntoCorreoInicial && prospect.HtmlCuerpoCorreoInicial) {
      listos++;
    }
  });
  
  console.log(`Prospectos listos para enviar: ${listos}`);
  return listos;
}


/**
 * Se ejecuta al hacer clic en el botón de la hoja "Contactos".
 * Lee la fila seleccionada, la convierte a JSON y la muestra en un sidebar.
 */
/**
 * Pide al usuario un ContactoId a través de una ventana emergente,
 * busca al contacto en la hoja correspondiente y muestra sus datos en un sidebar JSON.
 * VERSIÓN ROBUSTA: No depende de la selección de celda.
 */
function mostrarJsonDelContacto() {
  const ui = SpreadsheetApp.getUi();
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.contactos);

  // 1. Validar que la hoja 'Contactos' exista.
  if (!sheet) {
    ui.alert('Error Crítico', `No se encontró la hoja "${config.sheets.contactos}".`, ui.ButtonSet.OK);
    return;
  }

  // 2. Pedir al usuario el ContactoId a través de una ventana emergente.
  const response = ui.prompt(
    'Obtener JSON de Contacto',
    'Por favor, introduce el ContactoId (ej: C-001):',
    ui.ButtonSet.OK_CANCEL
  );

  // 3. Procesar la respuesta del usuario.
  if (response.getSelectedButton() !== ui.Button.OK) {
    return; // El usuario presionó Cancelar o cerró la ventana.
  }

  const contactoIdBuscado = response.getResponseText().trim().toUpperCase();
  if (!contactoIdBuscado) {
    ui.alert('Aviso', 'No se introdujo ningún ID.', ui.ButtonSet.OK);
    return; // El usuario presionó OK sin escribir nada.
  }

  // 4. Buscar la fila correspondiente al ID introducido.
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Saca la primera fila (encabezados)
  const idColIndex = headers.indexOf('ContactoId');

  if (idColIndex === -1) {
    ui.alert('Error de Configuración', 'No se encontró la columna "ContactoId" en la hoja de Contactos.', ui.ButtonSet.OK);
    return;
  }

  let filaEncontrada = null;
  for (const row of data) {
    // Comparamos de forma robusta (ignorando espacios y mayúsculas/minúsculas)
    if (row[idColIndex] && row[idColIndex].toString().trim().toUpperCase() === contactoIdBuscado) {
      filaEncontrada = row;
      break; // Detenemos la búsqueda en cuanto encontramos la primera coincidencia
    }
  }

  // 5. Si no se encuentra el ID, mostrar un error.
  if (!filaEncontrada) {
    ui.alert('Error', `No se encontró ningún contacto con el ID "${contactoIdBuscado}".`, ui.ButtonSet.OK);
    return;
  }

  // 6. Si se encuentra, construir el objeto JSON y mostrarlo en el sidebar (lógica anterior).
  const contactoJson = {};
  headers.forEach((header, index) => {
    let key = header;
    if (header === 'Respuesta Hilo Completo') {
      key = 'Hilo';
    }
    contactoJson[key] = filaEncontrada[index];
  });

  const jsonString = JSON.stringify(contactoJson, null, 2);

  const template = HtmlService.createTemplateFromFile('JsonSidebar');
  template.jsonString = jsonString;

  const htmlOutput = template.evaluate()
    .setWidth(400)
    .setTitle(`JSON del Contacto: ${contactoJson.NombreContacto || contactoIdBuscado}`);
  
  ui.showSidebar(htmlOutput);
}


// =================================================================================
// INICIO DEL BLOQUE DE REFACTORIZACIÓN DE BÚSQUEDA
// Contiene la nueva herramienta _checkIdExistsWithTextFinder y las funciones optimizadas.
// =================================================================================

/**
 * --- NUEVA FUNCIÓN DE UTILIDAD (ÓPTIMA) ---
 * Realiza una búsqueda de existencia de ID altamente eficiente y escalable usando TextFinder.
 * Es una función "privada" para uso interno del script.
 * @param {string} sheetName El nombre de la hoja donde buscar.
 * @param {string} idColumnLetter La letra de la columna que contiene los IDs (ej: 'A', 'C').
 * @param {string|number} idValue El valor del ID que se está buscando.
 * @returns {boolean} True si el ID se encuentra, false en caso contrario.
 * @private
 */
function _checkIdExistsWithTextFinder(sheetName, idColumnLetter, idValue) {
  // Guarda de robustez: si no hay un ID para buscar, no existe.
  if (!idValue) {
    return false;
  }

  try {
    const sheet = obtenerHoja(sheetName);
    // Si la hoja no existe o está vacía, el ID no puede existir.
    if (!sheet || sheet.getLastRow() < 2) {
      return false;
    }

    // Definimos el rango de búsqueda a una sola columna para máxima eficiencia.
    const searchRange = sheet.getRange(`${idColumnLetter}2:${idColumnLetter}`);
    
    // Usamos la API TextFinder, la forma más rápida de buscar un valor en Sheets.
    const textFinder = searchRange.createTextFinder(idValue.toString().trim());
    
    // Coincidencia exacta para evitar que "C-1" coincida con "C-10".
    textFinder.matchEntireCell(true);
    
    // findNext() devuelve el Rango si lo encuentra, o null si no.
    // Devolvemos true si el resultado NO es null.
    return textFinder.findNext() !== null;

  } catch (e) {
    log('ERROR', `Error en _checkIdExistsWithTextFinder para ID ${idValue} en hoja ${sheetName}: ${e.message}`);
    // En caso de un error inesperado, es más seguro devolver 'false' para no bloquear operaciones.
    return false;
  }
}

/**
 * --- FUNCIÓN REFACTORIZADA ---
 * Verifica si ya existe un prospecto para un ID de empresa dado.
 * Ahora delega la lógica a la función optimizada _checkIdExistsWithTextFinder.
 */
function existeProspectoParaEmpresa(idEmpresa) {
  return _checkIdExistsWithTextFinder(CONFIG().sheets.prospectosFase1, 'A', idEmpresa);
}

/**
 * --- FUNCIÓN REFACTORIZADA ---
 * Verifica si ya existe un prospecto de Fase 2 para un ID de contacto dado.
 * Ahora delega la lógica a la función optimizada _checkIdExistsWithTextFinder.
 */
function existeProspectoFase2ParaContacto(contactoId) {
  // Nota: La columna ContactoId está en la 'C' según tu esquema ESQUEMAS.Prospectos_Fase2
  return _checkIdExistsWithTextFinder(CONFIG().sheets.prospectosFase2, 'C', contactoId);
}

// =================================================================================
// FIN DEL BLOQUE DE REFACTORIZACIÓN DE BÚSQUEDA
// =================================================================================


/**
 * Se ejecuta semanalmente (los viernes) para limpiar la hoja de logs.
 * SOLO limpia la hoja si no se ha registrado ningún log de tipo 'ERROR' durante la semana (Lunes-Viernes).
 */
function limpiarLogsSiNoHayErroresSemanales() {
  const config = getSystemConfiguration();
  const sheet = obtenerHoja(config.sheets.logs);
  if (!sheet || sheet.getLastRow() < 2) {
    log('INFO', 'Limpieza de logs omitida: la hoja está vacía.');
    return; // No hay nada que limpiar
  }

  // 1. Definir el inicio de la semana (Lunes a las 00:00)
  const ahora = new Date();
  const diaDeHoy = parseInt(Utilities.formatDate(ahora, "Europe/Madrid", "u"), 10); // 1=Lunes, 7=Domingo
  const diasARestar = diaDeHoy - 1;
  const inicioSemana = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() - diasARestar);
  inicioSemana.setHours(0, 0, 0, 0);

  // 2. Leer los logs y buscar errores en la semana actual
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const timestampCol = headers.indexOf('Timestamp');
  const tipoCol = headers.indexOf('Tipo');

  let hayErroresEstaSemana = false;
  for (const row of data) {
    const timestamp = new Date(row[timestampCol]);
    // Si el log es de esta semana y es de tipo ERROR
    if (timestamp >= inicioSemana && row[tipoCol] === 'ERROR') {
      hayErroresEstaSemana = true;
      break; // Encontramos un error, no necesitamos seguir buscando
    }
  }

  // 3. Decidir si limpiar la hoja o no
  if (hayErroresEstaSemana) {
    log('INFO', 'Limpieza de logs omitida: Se encontraron errores esta semana.');
  } else {
    // No hay errores, procedemos a limpiar SOLO las filas de datos, conservando los encabezados.
    const rangoALimpiar = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    rangoALimpiar.clearContent();
    log('INFO', 'Limpieza semanal de logs completada. No se encontraron errores esta semana.');
  }
}

// ==================================================================
// --- MÓDULO DE SINCRONIZACIÓN INSTANTÁNEA Y CONFIGURACIÓN ---
// ==================================================================

/**
 * EL "TRABAJADOR": Compara empresas con prospectos y crea las entradas faltantes.
 */
function sincronizarNuevasEmpresas() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Iniciando sincronización. Se crearán los prospectos para las empresas nuevas. Por favor, espera.');

  log('INFO', 'Iniciando sincronización de nuevas empresas...');
  const config = CONFIG();
  const empresasSheet = obtenerHoja(config.sheets.empresas);
  const prospectosSheet = obtenerHoja(config.sheets.prospectosFase1);

  if (!empresasSheet || !prospectosSheet) {
    log('ERROR', 'Sincronización abortada: Faltan hojas de Empresas o Prospectos_Fase1.');
    ui.alert('Error: Faltan las hojas de Empresas o Prospectos_Fase1.');
    return;
  }

  const empresasData = empresasSheet.getDataRange().getValues();
  const prospectosData = prospectosSheet.getDataRange().getValues();
  const empresasHeaders = empresasData.shift();
  const prospectosHeaders = prospectosData.shift() || [];
  
  const indiceIdProspecto = prospectosHeaders.indexOf('ID_Empresa');
  if (indiceIdProspecto === -1 && prospectosSheet.getLastRow() > 1) {
    log('ERROR', 'No se encontró la columna "ID_Empresa" en Prospectos_Fase1.');
    ui.alert('Error: No se encontró la columna "ID_Empresa" en Prospectos_Fase1.');
    return;
  }
  
  const idsProspectosExistentes = new Set(prospectosData.map(row => row[indiceIdProspecto]));
  let prospectosCreados = 0;
  
  empresasData.forEach(filaEmpresa => {
    const empresaObjeto = crearObjetoDesdeArray(empresasHeaders, filaEmpresa);
    if (empresaObjeto.ID_Empresa && !idsProspectosExistentes.has(empresaObjeto.ID_Empresa)) {
      if (crearProspectoDesdeEmpresa(empresaObjeto)) {
        prospectosCreados++;
        idsProspectosExistentes.add(empresaObjeto.ID_Empresa); 
      }
    }
  });

  log('INFO', `Sincronización completada. Se crearon ${prospectosCreados} nuevos prospectos.`);
  // --- NUEVO: Invalidar caché de vistas después de sincronizar ---
  invalidarCacheVistas();
  ui.alert(`Sincronización Completa: Se han creado ${prospectosCreados} nuevos prospectos.`); // Alerta Segura
}

// ==================================================================
//               ↓↓↓ BLOQUE CONSOLIDADO Y CORREGIDO ↓↓↓
// ==================================================================

/**
 * EL "CONFIGURADOR" - VERSIÓN CONSOLIDADA Y CORREGIDA: Idempotente
 * Instala el activador "Al cambiar" para la sincronización automática.
 * Primero limpia cualquier activador antiguo que apunte a la misma función para evitar duplicados.
 */
function configurarActivadorAutomatico() {
  const ss = SpreadsheetApp.getActive();
  const nombreFuncionActivador = 'responderAlCambio';

  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === nombreFuncionActivador) {
      ScriptApp.deleteTrigger(trigger);
      log('INFO', 'Se ha eliminado un activador "onChange" antiguo para evitar duplicados.');
    }
  }

  ScriptApp.newTrigger(nombreFuncionActivador)
      .forSpreadsheet(ss)
      .onChange()
      .create();

  log('SUCCESS', 'Activador "onChange" para sincronización en tiempo real instalado correctamente.');
  SpreadsheetApp.getUi().alert('Configuración Completa', 'El activador automático ha sido instalado. El sistema ya está funcionando en tiempo real.', SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * EL "SENSOR" - VERSIÓN CONSOLIDADA
 * Se activa INSTANTÁNEAMENTE cuando Google Sheets detecta un cambio.
 * @param {Object} e El objeto de evento que proporciona Google.
 */
function responderAlCambio(e) {
  const lock = LockService.getScriptLock();
  // Un tiempo de bloqueo corto es suficiente
  if (!lock.tryLock(15000)) {
    log('WARN', 'responderAlCambio omitido por bloqueo de concurrencia.');
    return;
  }
  
  try {
    // Nos interesa solo cuando se insertan filas o se edita la hoja de Empresas.
    if (e.changeType === 'INSERT_ROW' || e.changeType === 'EDIT') {
      const hojaAfectada = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
      
      if (hojaAfectada === CONFIG().sheets.empresas) {
        log('INFO', `Cambio detectado (${e.changeType}) en 'Empresas'. Iniciando sincronización...`);
        // Pequeña pausa para permitir que Sheets procese completamente el cambio antes de leer.
        Utilities.sleep(2000); 
        sincronizarNuevasEmpresas();
      }
    }
  } catch (error) {
    log('ERROR', `Error en el activador responderAlCambio: ${error.toString()}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * VERSIÓN CONSOLIDADA
 */
function iniciarModoPruebaPasoAPaso() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('🧪 MODO PRUEBA MANUAL', 'Esto desactivará la automatización y te permitirá probar paso a paso. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (respuesta === ui.Button.YES) {
    desactivarAutomatizacion();
    ui.alert('Modo Prueba Iniciado', 'La automatización está desactivada.', ui.ButtonSet.OK);
  }
}

/**
 * VERSIÓN CONSOLIDADA
 */
function mostrarDiagnosticoCompleto() {
  const ui = SpreadsheetApp.getUi();
  try {
    const diagnostico = generarDiagnosticoCompleto();
    const mensaje = formatearDiagnosticoCompleto(diagnostico);
    ui.alert('📊 DIAGNÓSTICO DEL SISTEMA', mensaje, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error en Diagnóstico', `No se pudo completar el diagnóstico: ${e.toString()}`, ui.ButtonSet.OK);
  }
}


/**
 * Analiza un mensaje de error de envío para determinar si es un fallo permanente.
 * @param {string} errorString El mensaje de error.
 * @returns {boolean} True si el error se considera permanente, false en caso contrario.
 */
function esErrorPermanente(errorString) {
  const errorLower = errorString.toLowerCase();
  const palabrasClavePermanentes = [
    'host or domain name not found', // El dominio no existe (DNS)
    'user unknown',                  // El usuario no existe en el servidor
    'no such user',                  // Variante de "usuario no existe"
    'mailbox unavailable',           // El buzón no está disponible o no existe
    'recipient address rejected',    // Dirección rechazada por el servidor
    'unrouteable address',           // Dirección no enrutable
    '550',                           // Código SMTP estándar para "Buzón no disponible"
    '553'                            // Código SMTP para "Nombre de buzón no permitido"
  ];

  return palabrasClavePermanentes.some(clave => errorLower.includes(clave));
}

/**
 * Actualiza el estado de un prospecto a "Baja (Email Inválido)" tras un fallo de envío permanente.
 * @param {string} prospectoId El ID del prospecto (ID_Empresa en Fase 1).
 * @param {string} contactoId El ID del contacto (solo para Fase 2).
 * @param {string} emailFallido El email que falló.
 */
function actualizarEstadoProspectoPorFallo(prospectoId, contactoId, emailFallido) {
  const esFase2 = contactoId && contactoId.toString().trim() !== '';
  const { sheets } = CONFIG();
  const nombreHoja = esFase2 ? sheets.prospectosFase2 : sheets.prospectosFase1;
  const sheet = obtenerHoja(nombreHoja);
  if (!sheet) return;

  const claveBusqueda = esFase2 ? String(contactoId).trim() : String(prospectoId).trim();
  const nombreColClave = esFase2 ? 'ContactoId' : 'ID_Empresa';
  
  const colClaveIndex = obtenerIndiceColumna(sheet, nombreColClave) -1;
  const colEstadoIndex = obtenerIndiceColumna(sheet, 'Estado') -1;
  const colNotasIndex = obtenerIndiceColumna(sheet, 'Notas') -1;
  
  if (colClaveIndex < 0 || colEstadoIndex < 0) return;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[colClaveIndex]).trim() === claveBusqueda) {
      const rowIndex = i + 2; // +1 por los headers, +1 por el índice base 0
      sheet.getRange(rowIndex, colEstadoIndex + 1).setValue('Baja (Email Inválido)');
      if (colNotasIndex >= 0) {
        const notaActual = row[colNotasIndex] || '';
        const timestamp = new Date().toLocaleString('es-ES');
        const nuevaNota = `${notaActual}${notaActual ? '\n' : ''}[${timestamp}] Fallo de envío permanente a ${emailFallido}.`;
        sheet.getRange(rowIndex, colNotasIndex + 1).setValue(nuevaNota);
      }
      log('INFO', `Prospecto ${claveBusqueda} marcado como 'Baja' por email inválido: ${emailFallido}`);
      break; 
    }
  }
}

/**
 * Crea y aplica reglas de formato condicional a una hoja de prospectos para colorear
 * las filas enteras según el valor de la columna 'Estado'.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a la que se aplicarán las reglas.
 */
function aplicarReglasDeColorPorEstado(sheet) {
  const colEstadoIndex = obtenerIndiceColumna(sheet, 'Estado');
  if (colEstadoIndex <= 0) return; // No se encontró la columna 'Estado'

  const colEstadoLetra = sheet.getRange(1, colEstadoIndex).getA1Notation().replace("1", "");
  const rangoAplicacion = sheet.getRange(`A2:${sheet.getLastColumn()}`);
  
  // Limpiamos reglas anteriores para evitar duplicados
  sheet.clearConditionalFormatRules();

  // Regla para BAJAS (rojo)
  const reglaBaja = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR($${colEstadoLetra}2="Baja", $${colEstadoLetra}2="Baja (Email Inválido)")`)
    .setBackground("#ea9999") // Un rojo suave
    .setRanges([rangoAplicacion])
    .build();

  // Regla para RESPONDIDOS (amarillo)
  const reglaRespondido = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${colEstadoLetra}2="respondido"`)
    .setBackground("#fff2cc") // Un amarillo suave
    .setRanges([rangoAplicacion])
    .build();
    
  // Regla para CONVERTIDOS (verde)
  const reglaConvertido = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${colEstadoLetra}2="Convertido"`)
    .setBackground("#d9ead3") // Un verde suave
    .setRanges([rangoAplicacion])
    .build();

  // Regla para PAUSADOS (gris)
  const reglaPausado = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${colEstadoLetra}2="pausado"`)
    .setBackground("#f3f3f3") // Un gris claro
    .setRanges([rangoAplicacion])
    .build();

  const rules = [reglaBaja, reglaRespondido, reglaConvertido, reglaPausado];
  sheet.setConditionalFormatRules(rules);
  log('INFO', `Reglas de color aplicadas a la hoja '${sheet.getName()}'.`);
}





/**
 * Procesa la conversión de un prospecto a cliente.
 * @param {object} prospectoData Los datos de la fila del prospecto.
 * @param {string} fase La fase del prospecto ('fase1' o 'fase2').
 */
function convertirProspectoACliente(prospectoData, fase) {
  try {
    const ui = SpreadsheetApp.getUi();
    const config = CONFIG();
    const clientesSheet = obtenerHoja(config.sheets.clientes);
    if (!clientesSheet) {
      log('ERROR', 'No se encontró la hoja Clientes para la conversión.');
      ui.toast('Error: No se encuentra la hoja "Clientes".');
      return false;
    }
    const idEmpresa = (fase === 'fase1' ? prospectoData.ID_Empresa : prospectoData.EmpresaId);
    const emailPrincipal = prospectoData.EmailContacto || prospectoData.EmailGeneral;
    if (!idEmpresa || !emailPrincipal) {
      log('ERROR', `Conversión fallida: Faltan datos críticos. ID: ${idEmpresa}, Email: ${emailPrincipal}`);
      ui.toast('Conversión Fallida: Faltan ID o Email.');
      return false;
    }
    const nuevoClienteId = generarSiguienteClienteId();
    const idsExistentes = clientesSheet.getRange('B2:B').getValues().flat();
    if (idsExistentes.includes(idEmpresa)) {
      log('WARN', `Intento de convertir un prospecto que ya es cliente. ID: ${idEmpresa}`);
      ui.toast('Este prospecto ya ha sido convertido anteriormente.');
      return true;
    }
    const nuevaFilaCliente = [
      nuevoClienteId, idEmpresa, prospectoData.NombreEmpresa || '', prospectoData.PAIS || '',
      prospectoData.NombreContacto || '', emailPrincipal, prospectoData.TelefonoContacto || '',
      new Date(), '', prospectoData.Notas || '', prospectoData.ThreadId || ''
    ];
    clientesSheet.appendRow(nuevaFilaCliente);
    // --- NUEVO: Invalidar caché de clientes después de convertir ---
    invalidarCacheVistas();
    log('INFO', `Prospecto ${idEmpresa} convertido a Cliente con ID ${nuevoClienteId}.`);
    ui.toast(`¡Éxito! ${prospectoData.NombreEmpresa} ha sido añadido a Clientes.`);
    return true;
  } catch (e) {
    log('ERROR', `Error en convertirProspectoACliente: ${e.toString()}`);
    return false;
  }
}
/**
 * Genera el siguiente ID de cliente disponible (ej: CL-001, CL-002).
 * @returns {string} El nuevo ID de cliente.
 */
function generarSiguienteClienteId() {
  const config = CONFIG();
  const sheet = obtenerHoja(config.sheets.clientes);
  if (!sheet) return 'CL-001';
  const todosLosIds = sheet.getRange('A2:A').getValues().flat().filter(String);
  if (todosLosIds.length === 0) { return 'CL-001'; }
  let maxNum = 0;
  todosLosIds.forEach(id => {
    const numero = parseInt(id.split('-')[1]);
    if (!isNaN(numero) && numero > maxNum) { maxNum = numero; }
  });
  const nuevoNumero = maxNum + 1;
  const numeroFormateado = String(nuevoNumero).padStart(3, '0');
  return `CL-${numeroFormateado}`;
}


/**
 * Determina qué correo (inicial o de seguimiento) corresponde enviar a un prospecto.
 * @param {object} prospect El objeto de datos del prospecto.
 * @param {string} fase La fase del prospecto ('fase1' o 'fase2').
 * @returns {object|null} Un objeto con el asunto y el html del correo, o null si no hay nada que enviar.
 */
function obtenerSiguienteCorreoParaProspecto(prospect, fase) {
    const seguimiento = parseInt(prospect.SeguimientoActual) || 0;
    let asunto, html, tipoSeguimiento;

    if (fase === 'fase1') {
        switch (seguimiento) {
            case 0: asunto = prospect.AsuntoCorreoInicial; html = prospect.HtmlCuerpoCorreoInicial; tipoSeguimiento = 1; break;
            case 1: asunto = prospect.AsuntoSeguimiento1; html = prospect.HtmlCuerpoSeguimiento1; tipoSeguimiento = 2; break;
            case 2: asunto = prospect.AsuntoSeguimiento2; html = prospect.HtmlCuerpoSeguimiento2; tipoSeguimiento = 3; break;
            case 3: asunto = prospect.AsuntoSeguimiento3; html = prospect.HtmlCuerpoSeguimiento3; tipoSeguimiento = 4; break;
            default: return null;
        }
    } else {
        switch (seguimiento) {
            case 0: asunto = prospect.AsuntoCorreoInicial; html = prospect.HtmlCuerpoCorreoInicial; tipoSeguimiento = 1; break;
            case 1: asunto = prospect.AsuntoSeguimientoF2_1; html = prospect.HtmlCuerpoSeguimientoF2_1; tipoSeguimiento = 2; break;
            case 2: asunto = prospect.AsuntoSeguimientoF2_2; html = prospect.HtmlCuerpoSeguimientoF2_2; tipoSeguimiento = 3; break;
            case 3: asunto = prospect.AsuntoSeguimientoF2_3; html = prospect.HtmlCuerpoSeguimientoF2_3; tipoSeguimiento = 4; break;
            default: return null;
        }
    }

    if (asunto && html) {
        return { asunto: asunto, html: html, tipoSeguimiento: tipoSeguimiento };
    }
    return null;
}

/**
 * Lanza la aplicación web del CRM en una ventana modal a pantalla completa.
 */
function lanzarCrm() {
  // La línea clave es createTemplateFromFile y .evaluate()
  const html = HtmlService.createTemplateFromFile('WebApp')
      .evaluate() // Esto procesa los <?!= ... ?>
      .setWidth(1200)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'B2B Prospección CRM');
}
// ==================================================================
//               ↓↓↓ REEMPLAZA ESTA FUNCIÓN COMPLETA ↓↓↓
// ==================================================================
// ==================================================================
//               ↓↓↓ REEMPLAZA TU FUNCIÓN CON ESTA VERSIÓN DEFINITIVA ↓↓↓
// ==================================================================
/**
 * FUNCIÓN MAESTRA - VERSIÓN DEFINITIVA Y CORREGIDA
 * Soluciona el "error silencioso" de serialización de fechas que congela la UI.
 * Convierte manualmente todos los objetos Date a strings ISO 8601 antes de enviar la respuesta.
 * Mantiene la lógica de filtrado y paginación server-side.
 *
 * @param {string} fuenteDeDatos - El nombre de la fuente ('fase1', 'fase2', 'clientes').
 * @param {number} pagina - El número de página a devolver.
 * @param {number} limite - El número de registros por página.
 * @param {string} filtro - El estado por el cual filtrar (ej: 'activo', 'respondido', 'todos').
 * @returns {Object} Un objeto con la configuración, los datos y la paginación, 100% seguro para JSON.
 */
function obtenerDatosParaVista(fuenteDeDatos, pagina = 1, limite = 50, filtro = 'todos') {
  try {
    const systemConfig = getSystemConfiguration();
    let sheetName;
    
    switch (fuenteDeDatos) {
      case 'fase1': sheetName = systemConfig.sheets.prospectosFase1; break;
      case 'fase2': sheetName = systemConfig.sheets.prospectosFase2; break;
      case 'clientes': sheetName = systemConfig.sheets.clientes; break;
      default: throw new Error(`Fuente de datos no válida: ${fuenteDeDatos}`);
    }

    const sheet = obtenerHoja(sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { 
        config: obtenerConfiguracionDeVista(fuenteDeDatos), 
        data: [], 
        pagination: { currentPage: 1, perPage: limite, totalPages: 0, totalRecords: 0 }
      };
    }

    const allDataValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let allDataObjects = allDataValues.map(row => crearObjetoDesdeArray(headers, row));

    const datosFiltrados = (filtro === 'todos' || !filtro)
      ? allDataObjects
      : allDataObjects.filter(item => item.Estado && String(item.Estado).toLowerCase() === filtro.toLowerCase());

    const totalRegistros = datosFiltrados.length;
    const totalPaginas = Math.ceil(totalRegistros / limite) || 1;
    const paginaActual = Math.max(1, Math.min(parseInt(pagina), totalPaginas));
    
    const offset = (paginaActual - 1) * limite;
    let datosPaginados = datosFiltrados.slice(offset, offset + limite);

    // =========================================================================
    //               ↓↓↓ ESTA ES LA CORRECCIÓN MÁS IMPORTANTE ↓↓↓
    // Iteramos sobre los datos que vamos a enviar y convertimos cualquier objeto
    // de tipo Fecha a un string estándar (ISO 8601). Esto evita el error silencioso.
    // =========================================================================
    datosPaginados = datosPaginados.map(sanitizarFilaParaRespuesta);


    const config = obtenerConfiguracionDeVista(fuenteDeDatos);

    return { 
      config: config, 
      data: datosPaginados,
      pagination: {
        currentPage: paginaActual,
        perPage: limite,
        totalPages: totalPaginas,
        totalRecords: totalRegistros
      }
    };

  } catch (e) {
    log('ERROR', `Error crítico en obtenerDatosParaVista(${fuenteDeDatos}, filtro: ${filtro}): ${e.message}`);
    throw new Error(`No se pudieron cargar los datos. Detalles: ${e.message}`);
  }
}

/**
 * Función auxiliar para obtener la configuración de una vista específica.
 * @param {string} fuenteDeDatos - El nombre de la fuente.
 * @returns {Object} El objeto de configuración de la vista.
 */
function obtenerConfiguracionDeVista(fuenteDeDatos) {
  switch (fuenteDeDatos) {
    case 'fase1': 
      return { 
        titulo: 'Leads (Fase 1)', 
        descripcion: 'Centro de mando de prospectos en la primera fase de contacto.', 
        idColumn: 'ID_Empresa', 
        columnasMostradas: ['ID_Empresa', 'NombreEmpresa', 'PAIS', 'EmailGeneral', 'FechaUltimoEnvio', 'SeguimientoActual', 'Estado'] 
      };
    case 'fase2': 
      return { 
        titulo: 'Leads (Fase 2)', 
        descripcion: 'Prospectos que han avanzado a la segunda fase de contacto.', 
        idColumn: 'ContactoId', 
        columnasMostradas: ['ContactoId', 'NombreEmpresa', 'PAIS', 'EmailContacto', 'FechaUltimoEnvio', 'SeguimientoActual', 'Estado'] 
      };
    case 'clientes': 
      return { 
        titulo: 'Clientes', 
        descripcion: 'Listado de todas las empresas convertidas en clientes.', 
        idColumn: 'ClienteID', 
        columnasMostradas: ['ClienteID', 'NombreEmpresa', 'PAIS', 'EmailPrincipal', 'FechaConversion', 'ServicioContratado', 'ID_Empresa_Original'] 
      };
    default: 
      return {};
  }
}

/**
 * Permite incluir el contenido de otros archivos HTML en el principal.
 * Usado para separar CSS y JS.
 * @param {string} filename El nombre del archivo a incluir.
 * @returns {string} El contenido del archivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Calcula y devuelve un objeto completo de datos para el Dashboard.
 * Incluye KPIs, datos para gráficos y listas de acción.
 * @returns {Object} Un objeto completo con los datos del Dashboard.
 */
function obtenerDatosDelDashboard() {
  try {
    const hoy = new Date();
    const inicioHoy = new Date(new Date().setHours(0, 0, 0, 0));

    // --- CALCULAR KPIs ---
    const config = getSystemConfiguration();
    const enviadosHoy = contarEnviosHoy();
    const objetivoDiario = config.dailyLimit;
    
    const prospectosF1 = obtenerHoja(config.sheets.prospectosFase1)?.getDataRange().getValues() || [];
    const prospectosF2 = obtenerHoja(config.sheets.prospectosFase2)?.getDataRange().getValues() || [];

    const todosLosProspectos = [];
    if (prospectosF1.length > 1) {
      const headers = prospectosF1.shift();
      prospectosF1.forEach(row => todosLosProspectos.push(crearObjetoDesdeArray(headers, row)));
    }
    if (prospectosF2.length > 1) {
      const headers = prospectosF2.shift();
      prospectosF2.forEach(row => todosLosProspectos.push(crearObjetoDesdeArray(headers, row)));
    }

    let leadsActivos = 0;
    let respuestasHoy = 0;
    let conversionesHoy = 0;
    
    todosLosProspectos.forEach(p => {
      if (p.Estado === 'activo') leadsActivos++;
      
      if (p.RespuestaRecibida && new Date(p.RespuestaRecibida) >= inicioHoy) {
        respuestasHoy++;
      }
      
      if ((p.Estado === 'Cliente' || p.Estado === 'Convertido') && p['Última Acción Manual']) {
        const accion = p['Última Acción Manual'];
        const fechaAccionMatch = accion.match(/\[(.*?)\]/);
        if (fechaAccionMatch) {
            const [fechaStr, horaStr] = fechaAccionMatch[1].split(', ');
            if(fechaStr && horaStr) {
              const [dia, mes, anio] = fechaStr.split('/');
              const fechaAccion = new Date(`${anio}-${mes}-${dia}T${horaStr}`);
              if (fechaAccion >= inicioHoy) conversionesHoy++;
            }
        }
      }
    });

    // NOTA: Los datos para el gráfico y los paneles de acción se añadirán en pasos futuros.
    return {
      kpis: {
        enviadosHoy: enviadosHoy,
        objetivoDiario: objetivoDiario,
        respuestasHoy: respuestasHoy,
        leadsActivos: leadsActivos,
        conversionesHoy: conversionesHoy
      }
    };

  } catch (e) {
    log('ERROR', `Error en obtenerDatosDelDashboard: ${e.message}`);
    throw new Error('No se pudieron calcular los datos del Dashboard.');
  }
}

/**
 * Obtiene los correos enviados hoy desde la hoja Outbox.
 * @returns {Array<Object>} Un array de objetos, cada uno representando un envío.
 */
function obtenerActividadDeHoy() {
  try {
    const config = getSystemConfiguration();
    const outboxSheet = obtenerHoja(config.sheets.outbox);
    if (!outboxSheet || outboxSheet.getLastRow() < 2) return [];

    const data = outboxSheet.getDataRange().getValues();
    const headers = data.shift();
    const sentAtCol = headers.indexOf('sentAt');
    
    if (sentAtCol === -1) throw new Error('No se encontró la columna "sentAt" en la hoja Outbox.');
    
    const inicioHoy = new Date(new Date().setHours(0, 0, 0, 0));
    const enviosDeHoy = [];

    data.forEach(row => {
      const envio = crearObjetoDesdeArray(headers, row);
      if (envio.estado === 'sent' && envio.sentAt && new Date(envio.sentAt) >= inicioHoy) {
        enviosDeHoy.push({
          toEmail: envio.toEmail,
          asunto: envio.asunto,
          estado: envio.estado,
          sentAt: new Date(envio.sentAt).toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' })
        });
      }
    });
    
    // Ordenamos por hora de envío, del más reciente al más antiguo
    return enviosDeHoy.sort((a, b) => b.sentAt.localeCompare(a.sentAt));

  } catch (e) {
    log('ERROR', `Error en obtenerActividadDeHoy: ${e.message}`);
    throw new Error(e.message);
  }
}




/**
 * Obtiene todos los prospectos que tienen una respuesta, de ambas fases.
 * VERSIÓN CORREGIDA: Maneja y ordena las fechas de forma robusta.
 * @returns {Array<Object>} Un array de prospectos que han respondido, ordenados por fecha.
 */
function obtenerProspectosConRespuesta() {
  try {
    const config = getSystemConfiguration();
    const prospectosF1 = obtenerHoja(config.sheets.prospectosFase1)?.getDataRange().getValues() || [];
    const prospectosF2 = obtenerHoja(config.sheets.prospectosFase2)?.getDataRange().getValues() || [];
    
    let todosLosProspectos = [];
    if (prospectosF1.length > 1) {
      const headers = prospectosF1.shift();
      prospectosF1.forEach(row => todosLosProspectos.push({ ...crearObjetoDesdeArray(headers, row), fase: 'fase1' }));
    }
    if (prospectosF2.length > 1) {
      const headers = prospectosF2.shift();
      prospectosF2.forEach(row => todosLosProspectos.push({ ...crearObjetoDesdeArray(headers, row), fase: 'fase2' }));
    }

    const prospectosConRespuesta = todosLosProspectos
      .filter(p => p.Estado === 'respondido' && p.RespuestaRecibida instanceof Date) // Nos aseguramos de que sea una fecha válida
      .map(p => {
        return {
          id: p.ID_Empresa || p.ID,
          nombre: p.NombreEmpresa,
          fechaRespuestaObj: p.RespuestaRecibida, // Guardamos el objeto Date para ordenar
          fechaRespuestaStr: p.RespuestaRecibida.toLocaleString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
          contenidoRespuesta: p.ContenidoRespuesta || '(Sin contenido)',
          fase: p.fase
        };
      });
      
    // Ordenar por el objeto Date, de más reciente a más antigua
    prospectosConRespuesta.sort((a, b) => b.fechaRespuestaObj - a.fechaRespuestaObj);

    // Devolvemos un objeto limpio, sin el objeto Date auxiliar
    return prospectosConRespuesta.map(({ fechaRespuestaObj, ...resto }) => ({
      ...resto,
      fechaRespuesta: resto.fechaRespuestaStr
    }));

  } catch (e) {
    log('ERROR', `Error en obtenerProspectosConRespuesta: ${e.message}`);
    throw new Error('No se pudieron obtener los prospectos con respuesta.');
  }
}

/**
 * --- NUEVA FUNCIÓN AUXILIAR PARA LA VISTA DE RESPUESTAS ---
 * Añade una nueva nota a un prospecto específico.
 * @param {string} prospectoId El ID del prospecto.
 * @param {string} fase La fase del prospecto ('fase1' o 'fase2').
 * @param {string} nuevaNota El texto de la nota a añadir.
 */
function agregarNotaProspecto(prospectoId, fase, nuevaNota) {
  if (!nuevaNota || nuevaNota.trim() === '') {
    throw new Error('La nota no puede estar vacía.');
  }
  
  const config = CONFIG();
  const nombreHoja = (fase === 'fase1') ? config.sheets.prospectosFase1 : config.sheets.prospectosFase2;
  const sheet = obtenerHoja(nombreHoja);
  if (!sheet) throw new Error(`Hoja no encontrada: ${nombreHoja}`);

  const colIdNombre = (fase === 'fase1') ? 'ID_Empresa' : 'ID';
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idColIndex = headers.indexOf(colIdNombre);
  const notasColIndex = headers.indexOf('Notas');

  if (notasColIndex === -1) throw new Error('No se encontró la columna "Notas".');
  
  const rowIndex = data.findIndex(row => row[idColIndex] == prospectoId);
  
  if (rowIndex === -1) throw new Error(`No se encontró el prospecto con ID ${prospectoId}.`);
  
  const FilaRealSheet = rowIndex + 2;
  const notaActual = sheet.getRange(FilaRealSheet, notasColIndex + 1).getValue() || '';
  const timestamp = `[${new Date().toLocaleString('es-ES')}]`;
  const notaFinal = `${notaActual}${notaActual ? '\n' : ''}${timestamp} ${nuevaNota.trim()}`;
  
  sheet.getRange(FilaRealSheet, notasColIndex + 1).setValue(notaFinal);
}

/**
 * --- NUEVA FUNCIÓN HELPER ---
 * Invalida todas las cachés de vistas para forzar recarga de datos.
 */
function invalidarCacheVistas() {
  const cache = CacheService.getScriptCache();
  cache.remove('data_view_fase1');
  cache.remove('data_view_fase2');
  cache.remove('data_view_clientes');
  log('INFO', 'Cachés de vistas invalidadas.');
}

/**
 * Función central que se ejecuta desde el frontend para manejar acciones en la vista de Respuestas.
 * --- MODIFICADO: AÑADIDO LOCKSERVICE Y NUEVO CASE PARA 'agregarNota' ---
 * Protegido con LockService para evitar condiciones de carrera.
 */
function ejecutarAccionRespuesta(prospectoId, accion, fase, datosExtra) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'El sistema está ocupado. Inténtalo de nuevo en unos segundos.' };
  }
  
  try {
    const config = CONFIG();
    const nombreHoja = (fase === 'fase1') ? config.sheets.prospectosFase1 : config.sheets.prospectosFase2;
    const sheet = obtenerHoja(nombreHoja);
    if (!sheet) throw new Error(`No se encontró la hoja de prospectos: ${nombreHoja}`);
    
    // --- NUEVO: Limpiamos la caché de esta vista para que los cambios se reflejen inmediatamente ---
    invalidarCacheVistas();
    
    // En Fase 1 el identificador único es 'ID_Empresa', en Fase 2 usamos 'ID' que es una copia del 'EmpresaId'
    const colIdNombre = (fase === 'fase1') ? 'ID_Empresa' : 'ID';
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColIndex = headers.indexOf(colIdNombre);
    
    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      // Usamos '==' en lugar de '===' porque los IDs pueden venir como número o texto
      if (data[i][idColIndex] == prospectoId) {
        rowIndex = i + 2; // +1 por los headers, +1 por el índice base 0
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error(`No se encontró el prospecto con ID ${prospectoId} en la hoja ${nombreHoja}`);
    
    // Obtenemos un objeto con los datos del prospecto para reutilizarlo
    const prospectoData = crearObjetoDesdeArray(headers, data[rowIndex - 2]);
    const estadoCol = headers.indexOf('Estado') + 1;
    const accionCol = headers.indexOf('Última Acción Manual') + 1;
    const fechaReactivacionCol = headers.indexOf('FechaReactivacion') + 1;
    const timestamp = `[${new Date().toLocaleString('es-ES')}]`;

    switch (accion) {
      // --- NUEVO CASE ---
      case 'agregarNota':
        agregarNotaProspecto(prospectoId, fase, datosExtra.nota);
        if (accionCol > 0) sheet.getRange(rowIndex, accionCol).setValue(`Nota añadida desde CRM ${timestamp}`);
        break;
        
      case 'baja':
        sheet.getRange(rowIndex, estadoCol).setValue('Baja');
        if (accionCol > 0) sheet.getRange(rowIndex, accionCol).setValue(`Baja desde CRM ${timestamp}`);
        break;
        
      case 'pausar':
        sheet.getRange(rowIndex, estadoCol).setValue('pausado');
        if (accionCol > 0) sheet.getRange(rowIndex, accionCol).setValue(`Pausado desde CRM ${timestamp}`);
        break;
        
      case 'posponer':
        // datosExtra.meses nos lo envía el frontend
        const meses = parseInt(datosExtra.meses, 10);
        const fechaReactivacion = new Date();
        fechaReactivacion.setMonth(fechaReactivacion.getMonth() + meses);
        
        sheet.getRange(rowIndex, estadoCol).setValue('Pospuesto');
        if (fechaReactivacionCol > 0) sheet.getRange(rowIndex, fechaReactivacionCol).setValue(fechaReactivacion);
        if (accionCol > 0) sheet.getRange(rowIndex, accionCol).setValue(`Pospuesto por ${meses} meses desde CRM ${timestamp}`);
        break;
        
      case 'iniciarFase2':
        // Creamos un nuevo contacto en la hoja 'Contactos'
        const nuevoContacto = {
          EmpresaId: prospectoId,
          NombreEmpresa: prospectoData.NombreEmpresa,
          PAIS: prospectoData.PAIS,
          NombreContacto: datosExtra.nuevoNombre,
          EmailContacto: datosExtra.nuevoEmail,
          ContactoId: generarSiguienteContactoId()
        };
        const contactosSheet = obtenerHoja(config.sheets.contactos);
        const contactosHeaders = contactosSheet.getRange(1, 1, 1, contactosSheet.getLastColumn()).getValues()[0];
        const nuevaFilaContacto = contactosHeaders.map(header => nuevoContacto[header] || '');
        contactosSheet.appendRow(nuevaFilaContacto);
        
        Utilities.sleep(500); // Pequeña pausa para asegurar la escritura antes de crear el prospecto
        
        // Reutilizamos nuestra lógica para crear el prospecto de Fase 2 a partir del nuevo contacto
        crearProspectoFase2DesdeContacto(nuevoContacto);
        
        // Actualizamos el estado del prospecto original de Fase 1
        sheet.getRange(rowIndex, estadoCol).setValue('Convertido');
        if (accionCol > 0) sheet.getRange(rowIndex, accionCol).setValue(`Convertido a Fase 2 (Contacto: ${nuevoContacto.ContactoId}) ${timestamp}`);
        break;
        
      case 'convertirCliente':
        // Reutilizamos la función que ya teníamos
        convertirProspectoACliente(prospectoData, fase);
        
        // Actualizamos el estado del prospecto
        sheet.getRange(rowIndex, estadoCol).setValue('Cliente');
        if (accionCol > 0) sheet.getRange(rowIndex, accionCol).setValue(`Convertido a Cliente desde CRM ${timestamp}`);
        invalidarCacheVistas(); // Invalidar caché de clientes también
        break;
        
      default:
        throw new Error(`Acción desconocida: ${accion}`);
    }
    
    return { success: true, message: 'Acción ejecutada con éxito.' };

  } catch (e) {
    log('ERROR', `Error en ejecutarAccionRespuesta: ${e.message}`);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}



function exportarVistaAExcel(datos, columnas, titulo) {
  try {
    const nombreArchivo = `[CRM Export] ${titulo} - ${new Date().toLocaleString('es-ES')}`;
    const ss = SpreadsheetApp.create(nombreArchivo);
    const sheet = ss.getSheets()[0];

    const datosParaEscribir = [columnas];
    datos.forEach(objeto => {
      const fila = columnas.map(col => objeto[col] || '');
      datosParaEscribir.push(fila);
    });

    sheet.getRange(1, 1, datosParaEscribir.length, columnas.length).setValues(datosParaEscribir);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
    SpreadApp.flush(); // Aseguramos que los cambios se escriban antes de generar la URL.

    const fileId = ss.getId();
    
    // La lógica de limpieza ahora está centralizada en la función `limpiarArchivosTemporales`
    // y la gestión de la carpeta temporal, por lo que esta función es ahora más simple y directa.
    const archivo = DriveApp.getFileById(fileId);
    const carpetaTemporal = obtenerCarpetaTemporal();
    archivo.moveTo(carpetaTemporal);
    
    const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
    return url;

  } catch (e) {
    log('ERROR', `Error en exportarVistaAExcel: ${e.message}`);
    throw new Error('No se pudo generar el archivo de exportación.');
  }
}

/**
 * Función de utilidad para eliminar archivos temporales.
 * @param {GoogleAppsScript.Events.TimeDriven} e - Objeto de evento del trigger.
 */
function eliminarArchivoTemporal(e) {
    const fileId = e.args[0];
    if (fileId) {
        try {
            DriveApp.getFileById(fileId).setTrashed(true);
            log('INFO', `Archivo temporal ${fileId} eliminado.`);
        } catch (err) {
            log('WARN', `No se pudo eliminar el archivo temporal ${fileId}: ${err.message}`);
        }
    }
}


/**
 * Busca un único registro por su ID en una fuente de datos específica y lo devuelve como JSON.
 * @param {string} dataSource - La fuente de datos ('fase1', 'fase2', 'clientes').
 * @param {string} itemId - El ID del registro a buscar.
 * @returns {Object} El objeto JSON del registro encontrado.
 */
function obtenerJsonDeItem(dataSource, itemId) {
  try {
    const { sheets } = CONFIG();
    let sheetName;
    let idColumnName;

    switch (dataSource) {
        case 'fase1':
          sheetName = sheets.prospectosFase1;
        idColumnName = 'ID_Empresa';
        break;
        case 'fase2':
          sheetName = sheets.prospectosFase2;
        idColumnName = 'ContactoId';
        break;
        case 'clientes':
          sheetName = sheets.clientes;
        idColumnName = 'ClienteID';
        break;
      default:
        throw new Error('Fuente de datos no válida.');
    }

    const sheet = obtenerHoja(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idColIndex = headers.indexOf(idColumnName);

    if (idColIndex === -1) throw new Error(`Columna ID '${idColumnName}' no encontrada.`);

    for (const row of data) {
      if (row[idColIndex] == itemId) { // Usamos '==' para comparar texto y número
        return sanitizarFilaParaRespuesta(crearObjetoDesdeArray(headers, row));
      }
    }
    
    throw new Error(`Registro con ID '${itemId}' no encontrado.`);

  } catch (e) {
    log('ERROR', `Error en obtenerJsonDeItem: ${e.message}`);
    throw new Error(e.message);
  }
}


/**
 * NUEVA FUNCIÓN MAESTRA: Genera una exportación avanzada basada en las opciones del usuario.
 * @param {string} dataSource La fuente de datos ('fase1', 'fase2', 'clientes').
 * @param {Object} exportOptions Un objeto con las opciones de filtrado.
 * @returns {string} La URL de descarga del archivo Excel.
 */
function generarExportacionAvanzada(dataSource, exportOptions) {
  try {
    const { config, data: allData } = obtenerDatosParaVista(dataSource);
    if (!allData || allData.length === 0) {
      throw new Error('No hay datos para exportar en la fuente seleccionada.');
    }

    let filteredData;

    switch (exportOptions.type) {
      case 'currentView':
        // Para esta opción, el frontend ya nos envía los datos filtrados.
        filteredData = exportOptions.data;
        break;

      case 'idRange':
        const idColumn = config.idColumn;
        // Función auxiliar para extraer el número del ID (ej: "EMP-050" -> 50)
        const getNum = (id) => parseInt(String(id).split('-')[1], 10) || 0;
        const startNum = getNum(exportOptions.startId);
        const endNum = getNum(exportOptions.endId);
        
        if (startNum === 0 || endNum === 0 || startNum > endNum) {
          throw new Error('Rango de ID inválido. Asegúrate de que los IDs tengan el formato correcto (ej: EMP-001) y que el ID inicial sea menor o igual al final.');
        }

        filteredData = allData.filter(item => {
          const itemNum = getNum(item[idColumn]);
          return itemNum >= startNum && itemNum <= endNum;
        });
        break;

      case 'status':
        filteredData = allData.filter(item => item.Estado === exportOptions.status);
        break;
        
      default:
        throw new Error('Tipo de exportación no válido.');
    }

    if (filteredData.length === 0) {
      throw new Error('La selección no produjo ningún resultado. No se generó el archivo.');
    }
    
    // Reutilizamos la lógica de creación del archivo
    return crearArchivoExcel(filteredData, config.columnasMostradas, config.titulo);

  } catch (e) {
    log('ERROR', `Error en generarExportacionAvanzada: ${e.message}`);
    throw new Error(e.message);
  }
}

/**
 * FUNCIÓN REFACTORIZADA: Crea el archivo de Excel y lo guarda en una carpeta temporal.
 * Se ha eliminado por completo la creación de activadores dinámicos.
 */
function crearArchivoExcel(datos, columnas, titulo) {
  const nombreArchivo = `[CRM Export] ${titulo} - ${new Date().toLocaleString('es-ES').replace(/[\/:]/g, '-')}`;
  
  // 1. Obtiene o crea la carpeta temporal dedicada para las exportaciones.
  const carpetaTemporal = obtenerCarpetaTemporal();
  
  // 2. Crea el archivo de hoja de cálculo.
  const ss = SpreadsheetApp.create(nombreArchivo);
  const sheet = ss.getSheets()[0];

  const datosParaEscribir = [columnas];
  datos.forEach(objeto => {
    const fila = columnas.map(col => objeto[col] || '');
    datosParaEscribir.push(fila);
  });

  sheet.getRange(1, 1, datosParaEscribir.length, columnas.length).setValues(datosParaEscribir);
  sheet.setFrozenRows(1);
  sheet.getRange("1:1").setFontWeight("bold");
  sheet.autoResizeColumns(1, columnas.length);
  SpreadsheetApp.flush(); // Asegura que todos los cambios se escriban.

  // 3. Mueve el archivo recién creado a la carpeta temporal.
  const fileId = ss.getId();
  const archivo = DriveApp.getFileById(fileId);
  archivo.moveTo(carpetaTemporal);

  // 4. Devuelve la URL de descarga, que sigue siendo válida.
  const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
  log('INFO', `Exportación generada: ${nombreArchivo} en la carpeta temporal.`);
  return url;
}

/**
 * NUEVA FUNCIÓN AUXILIAR: Obtiene una referencia a la carpeta "CRM_Exports_Temp".
 * Si la carpeta no existe, la crea.
 * @returns {DriveApp.Folder} El objeto de la carpeta.
 */
function obtenerCarpetaTemporal() {
  const nombreCarpeta = "CRM_Exports_Temp";
  const carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  
  if (carpetas.hasNext()) {
    // La carpeta ya existe, la devolvemos.
    return carpetas.next();
  } else {
    // La carpeta no existe, la creamos en la raíz del Drive.
    log('INFO', `Creando carpeta temporal para exportaciones: ${nombreCarpeta}`);
    return DriveApp.createFolder(nombreCarpeta);
  }
}

/**
 * NUEVA FUNCIÓN DE LIMPIEZA: Revisa la carpeta temporal y elimina archivos antiguos.
 * Esta función será ejecutada por un único activador diario.
 */
function limpiarArchivosTemporales() {
  log('INFO', 'Iniciando rutina de limpieza de archivos de exportación temporales.');
  const carpeta = obtenerCarpetaTemporal();
  const archivos = carpeta.getFiles();
  const umbralTiempo = new Date(Date.now() - 24 * 60 * 60 * 1000); // Archivos de más de 24 horas
  let archivosEliminados = 0;

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (archivo.getDateCreated() < umbralTiempo) {
      try {
        archivo.setTrashed(true);
        archivosEliminados++;
      } catch (e) {
        log('ERROR', `No se pudo eliminar el archivo temporal ${archivo.getName()}: ${e.message}`);
      }
    }
  }

  if (archivosEliminados > 0) {
    log('SUCCESS', `Limpieza completada. Se eliminaron ${archivosEliminados} archivos antiguos.`);
  } else {
    log('INFO', 'Limpieza completada. No se encontraron archivos antiguos para eliminar.');
  }
}


function activarAutomatizacion() {
  desactivarAutomatizacion(); 
  ScriptApp.newTrigger('ejecutarAutomatizacion').timeBased().everyMinutes(15).create();
  // --- NUEVO: Invalidar caché de configuración para actualizar automationStatus ---
  CacheService.getScriptCache().remove('system_config');
  log('INFO', 'Automatización activada desde el panel de administración.');
  return { success: true, status: 'activo' };
}

// ==================================================================
// --- MÓDULO DE ENVÍO MANUAL DIRECTO ---
// ==================================================================

// ==================================================================
//               ↓↓↓ REEMPLAZA ESTA FUNCIÓN COMPLETA ↓↓↓
// ==================================================================
/**
 * Envía un correo electrónico de forma inmediata desde la interfaz web.
 * --- ARQUITECTURA MEJORADA: Ahora utiliza la hoja "Conversaciones". ---
 * @param {object} mailData - Objeto con los datos del correo.
 * @returns {object} - Un objeto con el resultado de la operación.
 */
function enviarCorreoManualDirecto(mailData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'El sistema está ocupado. Inténtalo de nuevo.' };
  }

  try {
    if (!validarEmailBasico(mailData.to)) {
      throw new Error('La dirección de correo electrónico no es válida.');
    }

    const config = getSystemConfiguration();
    const alias = config.senderAlias || Session.getActiveUser().getEmail();
    const fromAddressHeader = buildFromHeader(config.senderName || '', alias);

    const messageResource = {
      raw: Utilities.base64Encode(
        `From: ${fromAddressHeader}\r\n` +
        `To: ${mailData.to}\r\n` +
        `Subject: ${encodeHeaderUtf8(mailData.subject)}\r\n` +
        `Content-Type: text/html; charset=UTF-8\r\n\r\n` +
        `${mailData.body}`,
        Utilities.Charset.UTF_8
      ).replace(/\+/g, '-').replace(/\//g, '_')
    };

    if (mailData.useThread && mailData.threadId) {
      messageResource.threadId = mailData.threadId;
    }

    const sentMessage = Gmail.Users.Messages.send(messageResource, 'me');
    
    // --- LÓGICA DE ARQUITECTURA NUEVA ---
    if (sentMessage && sentMessage.threadId) {
        // En Fase 1 el ID es ID_Empresa, en Fase 2 es ContactoId. El frontend nos lo pasa como 'prospectoId'.
        crearOActualizarConversacion(mailData.prospectoId, mailData.fase, sentMessage.threadId, mailData.subject);
    }

    // El registro en Outbox para el historial sigue siendo útil.
    const outboxSheet = obtenerHoja(config.sheets.outbox);
    if (outboxSheet) {
      const esFase2 = mailData.fase === 'fase2';
      const newRow = [
        Utilities.getUuid(),
        esFase2 ? '' : mailData.prospectoId,
        esFase2 ? mailData.prospectoId : '',
        mailData.to,
        mailData.subject,
        mailData.body,
        'sent', 0, new Date(), new Date(),
        sentMessage.threadId || '', sentMessage.id || ''
      ];
      outboxSheet.appendRow(newRow);
    }
    
    agregarNotaProspecto(mailData.prospectoId, mailData.fase, 'Correo manual enviado desde CRM.');

    log('INFO', `Correo manual enviado a ${mailData.to} vía nuevo sistema de conversaciones.`);
    return { success: true, message: 'Correo enviado y conversación registrada.' };

  } catch (e) {
    log('ERROR', `Error en enviarCorreoManualDirecto: ${e.message}`);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ==================================================================
// --- MÓDULO DE GESTIÓN DE CONVERSACIONES ---
// ==================================================================

/**
 * Busca una conversación activa para una entidad y la actualiza, o crea una nueva si no existe.
 * Esta función se convierte en el punto central para gestionar los hilos de conversación.
 * @param {string} entidadId - El ID del prospecto (ej: 'EMP-001') o cliente.
 * @param {string} tipoEntidad - El tipo ('prospecto-f1', 'prospecto-f2', 'cliente').
 * @param {string} threadId - El ID del hilo de Gmail.
 * @param {string} asunto - El asunto del correo enviado.
 * @returns {boolean} - True si la operación fue exitosa.
 */
function crearOActualizarConversacion(entidadId, tipoEntidad, threadId, asunto) {
  const config = CONFIG();
  // Asumimos que aún no tenemos una hoja de Conversaciones en CONFIG, usamos el nombre directamente.
  const convSheet = obtenerHoja('Conversaciones'); 
  if (!convSheet) {
    log('ERROR', 'No se encontró la hoja "Conversaciones" para registrar el hilo.');
    return false;
  }

  const data = convSheet.getDataRange().getValues();
  const headers = data.shift() || ESQUEMAS.Conversaciones;
  const entidadIdCol = headers.indexOf('EntidadID');
  const tipoEntidadCol = headers.indexOf('TipoEntidad');

  let rowIndex = -1;
  // Buscamos si ya existe una conversación para esta entidad
  for (let i = 0; i < data.length; i++) {
    if (data[i][entidadIdCol] == entidadId && data[i][tipoEntidadCol] == tipoEntidad) {
      rowIndex = i + 2; // +1 por headers, +1 por índice base 0
      break;
    }
  }

  const ahora = new Date();
  if (rowIndex !== -1) {
    // La conversación ya existe, la actualizamos
    convSheet.getRange(rowIndex, headers.indexOf('ThreadId') + 1).setValue(threadId);
    convSheet.getRange(rowIndex, headers.indexOf('Estado') + 1).setValue('activo');
    convSheet.getRange(rowIndex, headers.indexOf('AsuntoUltimoCorreo') + 1).setValue(asunto);
    convSheet.getRange(rowIndex, headers.indexOf('FechaUltimaActividad') + 1).setValue(ahora);
    log('INFO', `Conversación actualizada para ${tipoEntidad} ${entidadId} con nuevo ThreadId.`);
  } else {
    // No existe, creamos una nueva fila
    const convId = `CONV-${Utilities.getUuid()}`;
    const nuevaFila = [
      convId,
      entidadId,
      tipoEntidad,
      threadId,
      'activo',
      asunto,
      ahora
    ];
    convSheet.appendRow(nuevaFila);
    log('INFO', `Nueva conversación creada para ${tipoEntidad} ${entidadId}.`);
  }
  return true;
}

/**
 * Actualiza la entidad original (prospecto, cliente, etc.) tras detectar una respuesta.
 * Esta función es llamada por el nuevo sistema de verificación de respuestas.
 * @param {string} entidadId - El ID de la entidad a actualizar.
 * @param {string} tipoEntidad - El tipo de entidad ('fase1', 'fase2').
 * @param {GoogleAppsScript.Gmail.GmailMessage} mensajeRespuesta - El objeto del mensaje de respuesta de Gmail.
 */
function actualizarEntidadConRespuesta(entidadId, tipoEntidad, mensajeRespuesta) {
  let sheetName, idColumnName;
  const config = CONFIG();

  switch (tipoEntidad) {
    case 'fase1':
      sheetName = config.sheets.prospectosFase1;
      idColumnName = 'ID_Empresa';
      break;
    case 'fase2':
      sheetName = config.sheets.prospectosFase2;
      idColumnName = 'ContactoId'; // El identificador único en fase 2 es el ContactoId
      break;
    default:
      log('WARN', `Tipo de entidad desconocido '${tipoEntidad}' al actualizar con respuesta.`);
      return;
  }

  const sheet = obtenerHoja(sheetName);
  if (!sheet) {
    log('ERROR', `No se encontró la hoja '${sheetName}' para actualizar la entidad ${entidadId}.`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idColIndex = headers.indexOf(idColumnName);
  
  if (idColIndex === -1) {
      log('ERROR', `No se encontró la columna ID '${idColumnName}' en la hoja '${sheetName}'.`);
      return;
  }

  const rowIndex = data.findIndex(row => String(row[idColIndex]).trim() == String(entidadId).trim());
  
  if (rowIndex === -1) {
    log('WARN', `No se encontró la entidad con ID ${entidadId} en la hoja ${sheetName}.`);
    return;
  }
  
  const FilaRealSheet = rowIndex + 2;
  
  try {
    const ahora = new Date();
    sheet.getRange(FilaRealSheet, obtenerIndiceColumna(sheet, 'Estado')).setValue('respondido');
    sheet.getRange(FilaRealSheet, obtenerIndiceColumna(sheet, 'RespuestaRecibida')).setValue(ahora);
    
    const notasActuales = sheet.getRange(FilaRealSheet, obtenerIndiceColumna(sheet, 'Notas')).getValue() || '';
    const nuevaNota = `${notasActuales}${notasActuales ? '\n' : ''}[${ahora.toLocaleString('es-ES')}] Respuesta detectada (Sistema de Conversaciones).`;
    sheet.getRange(FilaRealSheet, obtenerIndiceColumna(sheet, 'Notas')).setValue(nuevaNota);
    
    const colContenidoRespuesta = obtenerIndiceColumna(sheet, 'ContenidoRespuesta');
    if (colContenidoRespuesta > 0) {
      const hiloCompleto = mensajeRespuesta.getPlainBody();
      sheet.getRange(FilaRealSheet, colContenidoRespuesta).setValue(hiloCompleto.trim());
    }
    
    invalidarCacheVistas(); // Forzamos la actualización de la interfaz
    log('INFO', `Entidad ${entidadId} (${tipoEntidad}) actualizada con éxito a 'respondido'.`);

  } catch (error) {
    log('ERROR', `Error actualizando la entidad ${entidadId}: ${error.toString()}`);
  }
}


/**
 * Busca una conversación por su EntidadID y TipoEntidad y actualiza su estado.
 * Se utiliza principalmente para archivar conversaciones cuando se reactiva una secuencia.
 * @param {string} entidadId - El ID de la entidad (ID_Empresa o ContactoId).
 * @param {string} tipoEntidad - El tipo de entidad ('fase1' o 'fase2').
 * @param {string} nuevoEstado - El nuevo estado para la conversación (ej: 'archivado').
 * @returns {boolean} - True si se encontró y actualizó la conversación.
 */
function actualizarEstadoConversacion(entidadId, tipoEntidad, nuevoEstado) {
  const convSheet = obtenerHoja('Conversaciones');
  if (!convSheet || convSheet.getLastRow() < 2) return false;

  const data = convSheet.getRange('B2:E').getValues(); // EntidadID, TipoEntidad, ThreadId, Estado
  const headers = ['EntidadID', 'TipoEntidad', 'ThreadId', 'Estado'];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() === String(entidadId).trim() && row[1] === tipoEntidad) {
      const rowIndex = i + 2;
      const estadoCol = headers.indexOf('Estado') + 2; // +2 porque nuestro rango empieza en B
      convSheet.getRange(rowIndex, estadoCol).setValue(nuevoEstado);
      log('INFO', `Conversación para ${tipoEntidad} ${entidadId} actualizada a estado '${nuevoEstado}'.`);
      return true;
    }
  }
  return false;
}

// ==================================================================
//          ↓↓↓ PEGA ESTE BLOQUE COMPLETO AL FINAL DE TU SCRIPT ↓↓↓
// ==================================================================

/**
 * NUEVA UTILIDAD: Actualiza múltiples filas de una hoja en lote de forma ultra eficiente.
 * @param {string} sheetName El nombre de la hoja a actualizar.
 * @param {string} idColumnName El nombre de la columna que contiene el ID único.
 * @param {Array<Object>} updates Un array de objetos con los cambios a aplicar.
 * Formato de updates: [{ id: 'valorId', updates: { 'Columna': 'nuevoValor' } }]
 */
function batchUpdateSheet(sheetName, idColumnName, updates) {
  const sheet = obtenerHoja(sheetName);
  if (!sheet || updates.length === 0) return;

  try {
    const range = sheet.getDataRange();
    const data = range.getValues();
    const headers = data.shift();
    const idColIndex = headers.indexOf(idColumnName);
    
    if (idColIndex === -1) {
      log('ERROR', `[batchUpdateSheet] No se encontró la columna ID '${idColumnName}' en la hoja '${sheetName}'.`);
      return;
    }
    
    const idToRowIndexMap = new Map(data.map((row, index) => [String(row[idColIndex]).trim(), index]));
    let cambiosRealizados = false;

    updates.forEach(updateInfo => {
      const id = String(updateInfo.id).trim();
      if (idToRowIndexMap.has(id)) {
        const rowIndex = idToRowIndexMap.get(id);
        for (const colName in updateInfo.updates) {
          const colIndex = headers.indexOf(colName);
          if (colIndex !== -1) {
            data[rowIndex][colIndex] = updateInfo.updates[colName];
            cambiosRealizados = true;
          }
        }
      } else {
         log('WARN', `[batchUpdateSheet] No se encontró el ID '${id}' para actualizar en la hoja '${sheetName}'.`);
      }
    });

    if (cambiosRealizados) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      log('INFO', `[batchUpdateSheet] Actualización por lotes completada para la hoja '${sheetName}'.`);
    }
  } catch (e) {
    log('ERROR', `Error crítico en batchUpdateSheet para la hoja '${sheetName}': ${e.message}`);
  }
}

function procesarColaDeTareas() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { 
    log('INFO', 'Procesamiento de cola omitido, ya hay un proceso en ejecución.');
    return;
  }
  
  try {
    const config = CONFIG();

    // --- PROGRAMACIÓN DEFENSIVA: VERIFICACIÓN DE HOJAS ---
    const queueSheet = obtenerHoja('Queue_System');
    if (!queueSheet || queueSheet.getLastRow() < 2) {
      lock.releaseLock();
      return;
    }

    const prospectosF1Sheet = obtenerHoja(config.sheets.prospectosFase1);
    const prospectosF2Sheet = obtenerHoja(config.sheets.prospectosFase2);

    if (!prospectosF1Sheet) {
      log('ERROR_CRITICO', `No se encontró la hoja '${config.sheets.prospectosF1}'. El procesamiento de la cola se detiene. Verifica el nombre de la pestaña.`);
      lock.releaseLock();
      return;
    }
    if (!prospectosF2Sheet) {
      log('ERROR_CRITICO', `No se encontró la hoja '${config.sheets.prospectosF2}'. El procesamiento de la cola se detiene. Verifica el nombre de la pestaña.`);
      lock.releaseLock();
      return;
    }
    // --- FIN DE LA VERIFICACIÓN ---

    const rangoCompleto = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, ESQUEMAS.Queue_System.length);
    const tareas = rangoCompleto.getValues();
    const tareasPendientes = [];
    
    tareas.forEach((tarea, index) => {
      if (tarea[4] === 'pending') {
        tareasPendientes.push({ data: tarea, rowIndex: index });
      }
    });

    if (tareasPendientes.length === 0) {
      lock.releaseLock();
      return;
    }
    
    log('INFO', `Procesando ${tareasPendientes.length} tareas de la cola.`);

    const idsProspectosF1Existentes = new Set(prospectosF1Sheet.getRange('A2:A').getValues().flat().map(String));
    const idsProspectosF2Existentes = new Set(prospectosF2Sheet.getRange('C2:C').getValues().flat().map(String));
    
    let cambiosRealizados = false;

    for (const tareaInfo of tareasPendientes) {
      const tarea = tareaInfo.data;
      const tipo = tarea[2];
      const payload = JSON.parse(tarea[3]);

      try {
        let resultadoMsg = 'Tarea completada.';
        switch (tipo) {
          case 'CREATE_PROSPECT_F1':
            if (!idsProspectosF1Existentes.has(String(payload.empresaId))) {
              if (crearProspectoDesdeEmpresa(payload.data, prospectosF1Sheet)) {
                 idsProspectosF1Existentes.add(String(payload.empresaId));
                 resultadoMsg = `Prospecto ${payload.empresaId} creado.`;
              } else {
                throw new Error('La función crearProspectoDesdeEmpresa devolvió false.');
              }
            } else {
              resultadoMsg = `Omitido: Prospecto ${payload.empresaId} ya existe.`;
            }
            break;
            
          case 'CREATE_PROSPECT_F2':
            if (!idsProspectosF2Existentes.has(String(payload.contactoId))) {
               if (crearProspectoFase2DesdeContacto(payload.data, prospectosF2Sheet)) {
                  idsProspectosF2Existentes.add(String(payload.contactoId));
                  resultadoMsg = `Prospecto F2 para ${payload.contactoId} creado.`;
               } else {
                 throw new Error('La función crearProspectoFase2DesdeContacto devolvió false.');
               }
            } else {
              resultadoMsg = `Omitido: Prospecto F2 para ${payload.contactoId} ya existe.`;
            }
            break;
        }
        
        tareas[tareaInfo.rowIndex][4] = 'completed';
        tareas[tareaInfo.rowIndex][5] = resultadoMsg;
        cambiosRealizados = true;

      } catch (e) {
        log('ERROR', `Fallo al procesar tarea ${tarea[1]}: ${e.message}`);
        tareas[tareaInfo.rowIndex][4] = 'failed';
        tareas[tareaInfo.rowIndex][5] = e.message.substring(0, 500);
        cambiosRealizados = true;
      }
    }
    
    if (cambiosRealizados) {
       rangoCompleto.setValues(tareas);
    }

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// Función de utilidad (debe existir en tu código)
function crearNuevoProspecto(payload, sheet) {
  // Ahora esta función recibe un objeto 'sheet' que ya ha sido validado.
  sheet.appendRow([
    payload.companyId,
    payload.companyName,
    // ... resto de los datos
  ]);
}

    
/**
 * Instala el activador para el procesador de cola. Debe ejecutarse UNA SOLA VEZ.
 */
function instalarActivadorProcesadorCola() {
  const nombreFuncion = 'procesarColaDeTareas';
  
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === nombreFuncion) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger(nombreFuncion)
    .timeBased()
    .everyMinutes(1)
    .create();
    
  log('SUCCESS', `Activador para '${nombreFuncion}' instalado correctamente.`);
  SpreadsheetApp.getUi().alert('¡Éxito!', 'El activador para el procesador de tareas ha sido instalado. El sistema ahora procesará las ediciones en segundo plano.', SpreadsheetApp.getUi().ButtonSet.OK);
}




