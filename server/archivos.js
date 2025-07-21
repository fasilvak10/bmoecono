function crearInforme(empresa) {
  /*   empresa = {
      "id": "EX-2021-21518258-APN-SSEC#MDP",
      "tipoArchivo": "Notificacion",
      "tipoTramite": "regulares-outlined",
      "carpeta": "1jqKSYBWw6olm36myMIhy7sutIXkGZzW8",
      "tramite": "Inscripción al Régimen"
    }
    console.log('Empresa:', empresa); */

  const idNewIngreso = "1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg";
  const idCierres = "1EtfdMDuLgLr5KZVnjxGduUWH31ABUjgtmiAMUZCzW9c";

  const hojaModelos = SpreadsheetApp.openById(idNewIngreso).getSheetByName("modelos");
  const ss = empresa.tipoTramite === "regulares-outlined"
    ? SpreadsheetApp.openById(idNewIngreso).getSheetByName("main")
    : SpreadsheetApp.openById(idCierres).getSheetByName("principal_cierres");

  const expedienteIndex = empresa.tipoTramite === "regulares-outlined" ? 2 : 7;

  const data = ss.getDataRange().getValues().filter(row => row[expedienteIndex] === empresa.id);
  console.log('Datos encontrados:', data);

  if (data.length === 0) {
    console.warn('No se encontraron datos para ese expediente');
    return null;
  }

  const buscarFila = (id) => {
    const ids = ss.getRange(2, 1, ss.getLastRow() - 1, 8).getValues().map(row => row[expedienteIndex]);
    const index = ids.indexOf(id);
    return index !== -1 ? index + 2 : -1;
  };

  const fila = buscarFila(empresa.id);
  console.log(fila)

  // Buscar modelo y carpeta
  let modeloId = '';
  let carpetaTramite = '';

  hojaModelos.getDataRange().getValues().forEach(row => {
    if (row[3] === empresa.tramite) {
      carpetaTramite = row[2];
    }
    if (row[0] === empresa.tipoArchivo) {
      modeloId = row[1];
    }
  });

  console.log('Modelo:', modeloId, 'Carpeta trámite:', carpetaTramite);

  const indice = indice_tramites_lec.getIndice(empresa.tipoTramite);
  console.log('Índice:', indice);

  const campos = (estructuraCampos[empresa.tipoTramite]
    && estructuraCampos[empresa.tipoTramite][empresa.tipoArchivo])
    || [];

  if (campos.length === 0) {
    console.warn('⚠️ No se definieron campos para esta combinación de trámite y archivo');
    return null;
  }

  // Construcción del objeto empresa
  let dataEmpresa = data.map(row => {
    const obj = {};

    campos.forEach(campo => {
      obj[campo] = row[indice[campo]] || '';
    });

    return obj;
  });

  //console.log('Resultado:', (dataEmpresa.length === 1 ? dataEmpresa[0] : dataEmpresa));

  if (dataEmpresa.length === 1) {
    dataEmpresa = dataEmpresa[0];
    dataEmpresa.indiceCarpetaId = indice.idcarpetaEmpresa;
    dataEmpresa.indiceCarpetalink = indice.linkcarpetaEmpresa;

    // ⬇️ Agregar esta línea
    Object.assign(dataEmpresa, obtenerDatosAdicionales(empresa));
  }

  console.log('Resultado:', dataEmpresa)

  return generarDocumentoOCarpeta(empresa, dataEmpresa, modeloId, fila, ss, carpetaTramite)
}

function buscarFilaPorID(sheet, id, columna = 0) {
  const data = sheet.getDataRange().getValues();
  return data.find(row => row[columna] === id) || null;
}

function obtenerDatosAdicionales(empresa) {
  let hoja, filaEncontrada;

  if (empresa.tipoTramite === "regulares-outlined") {
    const archivo = SpreadsheetApp.openById('14ADAMj9-BhYZ_xCGU1MzCWeGzjCoE2UxQDOSjZTh0EI');

    if (empresa.tipoArchivo === 'Informe Inscripción') {
      hoja = archivo.getSheetByName("Mod Negocios Inscrip");
    } else if (empresa.tipoArchivo === 'Informe Revalidación Bienal') {
      hoja = archivo.getSheetByName("Mod Negocios Reval");
    }

    if (hoja) filaEncontrada = buscarFilaPorID(hoja, empresa.id, 0);

  } else {
    if (empresa.tipoArchivo === 'Informe Cierre Auditoría' || empresa.tipoArchivo === 'Informe Cierre Auditoría Simplificado') {
      const archivo = SpreadsheetApp.openById('1yZJKtLxp-LIBDJLzAvmOxNIMD4YM_4eqnogNRzMiSPI'); // reemplazar
      hoja = archivo.getSheetByName('Auditoria');
      if (hoja) filaEncontrada = buscarFilaPorID(hoja, empresa.id, 18);
    }
  }

  if (!filaEncontrada) {
    console.warn("No se encontraron datos adicionales para:", empresa.id);
    return {};
  }

  // Armás acá los datos adicionales a agregar

  return {
    objeto: filaEncontrada[2] || '',
    modelo: filaEncontrada[3] || '',
    valor: filaEncontrada[4] || '',
    modelo: filaEncontrada[5] || '',
    comercializacion: filaEncontrada[6] || '',
    descripcion: filaEncontrada[7] || '',
    // Agregá los campos que necesites con su índice
  };
}

const estructuraCampos = {
  "regulares-outlined": {
    "Informe Inscripción": [
      'nombre', 'cuit', 'expediente', 'tramite', 'tipoDeEmpresa',
      'periodoInicio', 'periodoFinalizacion', 'objeto', 'modelo', 'valor', 'comercializacion', 'descripcion', 'porcentajePromo',
      'ventasTotal', 'expoPromovida', 'porcentajeExpo', 'empleados',
      'porcentajeIn vestigacion', 'masaPromovida', 'porcentajeCapacitacion',
      'calidad', 'estadoCalidad', 'actividad', 'ventasPromo', 'idcarpetaEmpresa', 'fechaPresentacionTramite', 'fechaAltaIva'
    ],
    "Informe Acreditación Anual": [
      'nombre', 'cuit', 'expediente', 'tramite', 'tipoDeEmpresa', 'actividad', 'periodoInicio', 'periodoFinalizacion', 'fechaInscripcion', 'empleados', 'expoPromovida', 'ventasPromo', 'porcentajeExpo', 'idcarpetaEmpresa', 'fechaPresentacionTramite'],
    "Informe Revalidación Bienal": [
      'nombre', 'cuit', 'expediente', 'tramite', 'tipoDeEmpresa', 'actividad', 'periodoInicio', 'periodoFinalizacion', 'fechaInscripcion', 'empleados', 'porcentajePromo', 'ventasTotal', 'expoPromovida', 'ventasPromo', 'ventasPromoDos', 'expoPromovidaDos', 'porcentajeExpoDos', 'idcarpetaEmpresa', 'fechaPresentacionTramite'],
    "Informe Baja": [
      'nombre', 'cuit', 'expediente', 'tramite', 'tipoDeActo', 'numeroActo', 'fechaInscripcion', 'universidad', 'periodoInicio', 'periodoFinalizacion', 'ifTasa', 'ifFonpec', 'idcarpetaEmpresa', 'fechaPresentacionTramite'],
    "Notificacion": [
      'nombre', 'cuit', 'expediente', 'tramite', 'idcarpetaEmpresa'],
    "Notificacion TASA": [
      'nombre', 'cuit', 'expediente', 'idcarpetaEmpresa']
  },
  "cierres-outlined": {
    'Informe Cierre Auditoría Simplificado': [
      'nombre', 'cuit', 'expediente', 'periodoInicio', 'periodoFinalizacion', 'universidad', 'ifTasa', 'ifFonpec', 'idcarpetaEmpresa'],
    'Informe Cierre Auditoría': [
      'nombre', 'cuit', 'expediente', 'periodoInicio', 'periodoFinalizacion', 'universidad', 'ifTasa', 'ifFonpec', 'idcarpetaEmpresa'],
    "Notificacion": [
      'nombre', 'cuit', 'expediente', 'tramite', 'idcarpetaEmpresa'],
    "Notificacion TASA": [
      'nombre', 'cuit', 'expediente', 'idcarpetaEmpresa']
  }
};

function generarDocumentoOCarpeta(empresa, dataEmpresa, modeloId, fila, ss, carpetaTramite) {

  if (empresa.id === '') return;

  const main = ss;
  let urlResultado = '';

  // 👉 Crear Carpeta si no existe
  if (empresa.carpeta === "") {
    const carpeta = DriveApp.getFolderById(carpetaTramite);
    const carpetaCopia = carpeta.createFolder(`${dataEmpresa.nombre} - ${dataEmpresa.cuit}`);
    const carpetaId = carpetaCopia.getId();
    const carpetaUrl = carpetaCopia.getUrl();

    // ✅ Guardamos los datos de la carpeta nueva en la hoja
    main.getRange(fila, (dataEmpresa.indiceCarpetalink) + 1).setValue(carpetaUrl);
    main.getRange(fila, (dataEmpresa.indiceCarpetaId) + 1).setValue(carpetaId);

    // ✅ También actualizamos el ID de carpeta para que el resto del código pueda usarlo
    dataEmpresa.idcarpetaEmpresa = carpetaId;

    urlResultado = carpetaUrl;
  }

  // ✅ Mover la creación de documento fuera del "else", para que siempre se ejecute
  const archivoDoc = DriveApp.getFileById(modeloId);
  const carpeta = DriveApp.getFolderById(dataEmpresa.idcarpetaEmpresa);
  console.log(dataEmpresa.tipoDeEmpresa);

  const copiaArchivo = archivoDoc.makeCopy(carpeta);
  const copiaId = copiaArchivo.getId();
  const copiaUrl = copiaArchivo.getUrl();

  let nombreArchivo = '';

  // 👉 Nombres por tipo
  if (empresa.tipoArchivo === "Informe Inscripción") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - ${empresa.tipoArchivo.toUpperCase()}`;
  } else if (empresa.tipoArchivo === "Notificacion") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - NOTIFICACIÓN LEC`;
  } else if (empresa.tipoArchivo === "Informe Acreditación Anual") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - ${empresa.tipoArchivo.toUpperCase()}`;
  } else if (empresa.tipoArchivo === "Informe Revalidación Bienal") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - ${empresa.tipoArchivo.toUpperCase()}`;
  } else if (empresa.tipoArchivo === "Informe Baja") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - ${empresa.tipoArchivo.toUpperCase()}`;
  } else if (empresa.tipoArchivo === "Notificación TASA") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} ${empresa.tipoArchivo.toUpperCase()}`;
  } else if (empresa.tipoArchivo === "Informe Cierre Auditoría Simplificado") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - ${empresa.tipoArchivo.toUpperCase()}`
  } else if (empresa.tipoArchivo === "Informe Cierre Auditoría") {
    nombreArchivo = `${dataEmpresa.nombre} - Cuit ${dataEmpresa.cuit} - ${empresa.tipoArchivo.toUpperCase()}`
  }

  copiaArchivo.setName(nombreArchivo);

  const doc = DocumentApp.openById(copiaId);
  const body = doc.getBody();

  // Reemplazos
  const reemplazosBasicos = {
    '{{nombre}}': dataEmpresa.nombre || '',
    '{{cuit}}': dataEmpresa.cuit || '',
    '{{expediente}}': dataEmpresa.expediente || '',
    '{{tramite}}': dataEmpresa.tramite || ''
  };

  const reemplazosExtendidos = {
    '{{tipoDeEmpresa}}': dataEmpresa.tipoDeEmpresa || '',
    '{{periodoInicio}}': dataEmpresa.periodoInicio instanceof Date
      ? dataEmpresa.periodoInicio.toLocaleDateString('es-AR', { month: 'long', year: 'numeric' }).replace(' de ', ' ')
      : '',
    '{{periodoFinalizacion}}': dataEmpresa.periodoFinalizacion instanceof Date
      ? dataEmpresa.periodoFinalizacion.toLocaleDateString('es-AR', { month: 'long', year: 'numeric' }).replace(' de ', ' ')
      : '',
    '{{porcentajePromo}}': isFinite(dataEmpresa.porcentajePromo) ? `${dataEmpresa.porcentajePromo * 100}%` : '',
    '{{ventasTotal}}': isFinite(dataEmpresa.ventasTotal) ? Number(dataEmpresa.ventasTotal).toLocaleString('es-AR', { style: 'currency', currency: 'ARS' }) : '',
    '{{expoPromovida}}': isFinite(dataEmpresa.expoPromovida) ? Number(dataEmpresa.expoPromovida).toLocaleString('es-AR', { style: 'currency', currency: 'ARS' }) : '',
    '{{porcentajeExpo}}': isFinite(dataEmpresa.porcentajeExpo) ? `${dataEmpresa.porcentajeExpo * 100}%` : '',
    '{{empleados}}': dataEmpresa.empleados ?? '',
    '{{porcentajeInvestigacion}}': isFinite(dataEmpresa.porcentajeInvestigacion) ? `${dataEmpresa.porcentajeInvestigacion * 100}%` : '',
    '{{masaPromovida}}': isFinite(dataEmpresa.masaPromovida) ? Number(dataEmpresa.masaPromovida).toLocaleString('es-AR', { style: 'currency', currency: 'ARS' }) : '',
    '{{porcentajeCapacitacion}}': isFinite(dataEmpresa.porcentajeCapacitacion) ? `${dataEmpresa.porcentajeCapacitacion * 100}%` : '',
    '{{calidad}}': dataEmpresa.calidad || '',
    '{{estadoCalidad}}': dataEmpresa.estadoCalidad || '',
    '{{actividad}}': dataEmpresa.actividad || '',
    '{{ventasPromo}}': isFinite(dataEmpresa.ventasPromo) ? Number(dataEmpresa.ventasPromo).toLocaleString('es-AR', { style: 'currency', currency: 'ARS' }) : '',
    '{{micro}}': dataEmpresa.micro || '',
    '{{articuloNueve}}': dataEmpresa.articuloNueve || '',
    '{{modificaciones}}': dataEmpresa.modificaciones || '',
    '{{fechaInscripcion}}': dataEmpresa.fechaInscripcion instanceof Date ? dataEmpresa.fechaInscripcion.toLocaleDateString('es-AR') : '',
    '{{tipoDeActo}}': dataEmpresa.tipoDeActo || '',
    '{{numeroActo}}': dataEmpresa.numeroActo || '',
    '{{ventasPromoDos}}': dataEmpresa.ventasPromoDos === '' ? '' : isFinite(dataEmpresa.ventasPromoDos) ? Number(dataEmpresa.ventasPromoDos).toLocaleString('es-AR', { style: 'currency', currency: 'ARS' }) : '',
    '{{expoPromovidaDos}}': isFinite(dataEmpresa.expoPromovidaDos) ? Number(dataEmpresa.expoPromovidaDos).toLocaleString('es-AR', { style: 'currency', currency: 'ARS' }) : '',
    '{{porcentajeExpoDos}}': isFinite(dataEmpresa.porcentajeExpoDos) ? `${dataEmpresa.porcentajeExpoDos * 100}%` : '',
    '{{fechaPresentacionTramite}}': dataEmpresa.fechaPresentacionTramite instanceof Date ? dataEmpresa.fechaPresentacionTramite.toLocaleDateString('es-AR') : '',
    '{{fechaAltaIva}}': dataEmpresa.fechaAltaIva instanceof Date ? dataEmpresa.fechaAltaIva.toLocaleDateString('es-AR') : '',
    '{{objeto}}': dataEmpresa.objeto || '',
    '{{modelo}}': dataEmpresa.modelo || '',
    '{{valor}}': dataEmpresa.valorAgregado || '',
    '{{comercializacion}}': dataEmpresa.comercializacion || '',
    '{{descripcion}}': dataEmpresa.descripcion || ''
  };

  const reemplazos = (empresa.tipoArchivo !== "Notificacion")
    ? { ...reemplazosBasicos, ...reemplazosExtendidos }
    : reemplazosBasicos;

  for (let clave in reemplazos) {
    body.replaceText(clave, reemplazos[clave]);
  }

  doc.saveAndClose();
  urlResultado = copiaUrl;

  SpreadsheetApp.flush();
  console.log("resultado: ", urlResultado);
  return urlResultado;
}
