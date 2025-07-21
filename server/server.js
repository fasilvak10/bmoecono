function getData() {
    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    //const sheet = ss.getSheetByName('main');
    const sheet = ss.getSheetByName('main');
    const usuariosSheet = ss.getSheetByName("user");

    const allData = sheet.getDataRange().getDisplayValues();
    const usuariosData = usuariosSheet.getDataRange().getDisplayValues();

    // Crea un mapa de usuarios basado en los datos de la hoja de usuarios
    const usuariosMap = usuariosData.reduce((map, row) => {
        map[row[0]] = row[7];
        return map;
    }, {});

    // Obtiene cuits y los nombres asignados
    const cuits = allData.slice(1).map(row => row[4]);
    const paresEmpresaNombre = AsignarUsuarios.devolverNombreEmpresas(cuits);

    // Define los encabezados para la tabla
    const headers = [
        'Analista Asignado (evaluador)',
        'Razón social',
        'Cuit',
        'Número de expediente',
        'Tipo de Trámite',
        'Fecha Cambio Estado',
        'Estado',
        'Sub Estado',
        'Tareas Pendientes',
        'Informado 1266'
    ];

    // Mapea los datos para AG Grid y también prepara datos adicionales
    const data = [];
    const additionalData = [];

    allData.slice(1).forEach(row => {
        const cuit = row[4];
        const name = paresEmpresaNombre[cuit] || row[3];
        const estado = row[38];
        const ulMovimiento = row[39];
        const tramite = row[5];
        const expediente = row[2];

        const asesorLegal = row[43];
        const fechaRevision = row[45];
        const fecharevifinal = row[48];
        const analista = row[41];
        const revisorFinal = row[46];

        const tamanio = row[7];
        let fechapresentacion = row[6];
        let actividad = row[18];
        let observaciones = row[49];
        const conclusion = row[55];
        const informetecnico = row[55];
        const pvcierre = row[56];
        const numeroActo = row[57];
        const fechaacto = row[58];
        const fechapvcierre = row[59];
        const fechaInscripcion = row[60];
        const conciTasa = row[67];
        const conciFonpec = row[70];
        const periodoIncio = row[8];
        const periodoFinal = row[9];
        const fechaSubsanacion = row[52];
        const plazoSubsanacion = row[54];
        const diasPlazo = row[74];

        const disponibleEmision = row[79];
        const informadoAfip = row[62];

        /* REQUISITOS */

        const ventasTotal = row[19];
        const ventasPromo = row[20];
        const empleados = row[24];
        const masaSalarial = row[25];
        const imasd = row[26];
        const capacitacion = row[28];
        const exportaciones = row[30];
        const calidadTipo = row[32];
        const calidadEstado = row[33];
        const tareaPendiente = row[50];

        const ventasTotales2 = row[75]
        const ventasPromo2 = row[76]
        const expor2 = row[77];

        const carpetaEmpresa = row[63];

        const microMenor = row[22];

        const subestado = row[40];

        const revisorRealLegal = row[80];
        const revisorRealFinal = row[81];

        let responsable;
        if (asesorLegal &&
            ((estado === 'En Revisión' && !fechaRevision) ||
                (estado === 'Informe Técnico Firmado' && fechaRevision) ||
                ['Con Observaciones', 'Confeccionando Acto', 'A la Firma SSEC', 'En Despacho',
                    'Prefirma', 'En Jurídicos', 'Revisión Cierre Bienal', 'Informe Técnico Firmado Cierre Bienal'].includes(estado))) {
            responsable = asesorLegal;
        } else if (asesorLegal && estado === 'En Revisión' && fechaRevision && !fecharevifinal) {
            responsable = revisorFinal;
        } else {
            responsable = analista;
        }

        // Datos principales para AG Grid
        data.push([
            responsable,
            name,
            cuit,
            expediente,
            tramite,
            ulMovimiento,
            estado,
            subestado,
            tareaPendiente,
            informadoAfip

        ]);

        // Datos adicionales que se van a usar en el formulario
        additionalData.push({
            empresa: name,
            cuit: cuit,
            expediente: expediente,
            fechapresentacion: fechapresentacion,
            tramite: tramite,
            actividad: actividad,
            analista: analista,
            ulMovimiento: ulMovimiento,

            estadoEmpresa: estado,
            tamanio: tamanio,
            fechaSubsanacion: fechaSubsanacion.split('/').reverse().join('-'),
            plazoSubsanacion: plazoSubsanacion,
            cantidadDias: diasPlazo,

            periodoIncial: periodoIncio,
            periodoFinal: periodoFinal,
            masaSalarial: masaSalarial,
            nomina: empleados,

            ventasTotales: ventasTotal,
            ventasPromo: ventasPromo,
            exportaciones: exportaciones,

            ventasanio02: ventasTotales2,
            ventasPromoanio2: ventasPromo2,
            exportacionesanio2: expor2,

            microMenor: microMenor,
            capacitacion: capacitacion,
            imasd: imasd,
            calidadTipo: calidadTipo,
            calidadEstado: calidadEstado,

            fechaRevisionLegal: fechaRevision.split('/').reverse().join('-'),
            asesorLegal: asesorLegal,
            fechaRevisionFinal: fecharevifinal.split('/').reverse().join('-'),
            revisorFinal: revisorFinal,

            revisorRealLegal: revisorRealLegal,
            revisorRealFinal: revisorRealFinal,

            iftecnico: informetecnico,
            provi: pvcierre,
            numacto: numeroActo,

            fechaActo: fechaacto.split('/').reverse().join('-'),
            fechaProvi: fechapvcierre.split('/').reverse().join('-'),
            fechaInscripcion: fechaInscripcion.split('/').reverse().join('-'),

            disponibleEmision: disponibleEmision,
            informadoAfip: informadoAfip,
            observaciones: observaciones,
            carpetaEmpresa: carpetaEmpresa,

        });
    });

    return { headers: headers, data: data, additionalData: additionalData };
}

/* PARA MELA */

function getCierres() {
    const ssex = SpreadsheetApp.openById('1EtfdMDuLgLr5KZVnjxGduUWH31ABUjgtmiAMUZCzW9c');
    const bienal = ssex.getSheetByName('principal_cierres');
    const rango = bienal.getRange(1, 1, bienal.getLastRow(), bienal.getLastColumn()).getDisplayValues();

    let additionalData = [];

    function formatFecha(fecha) {
        return fecha && fecha.trim() !== '' ? fecha.split('/').reverse().join('-') : '';
    }

    const [header, ...data] = rango.map(i => {
        let cuit = i[1];
        let razon = i[2];
        let periodoDesde = i[4];
        let periodoHasta = i[5];
        let observacionAuditoria = i[3];

        let cuerpoAuditor = i[6];
        let expediente = i[7];
        let estado = i[8];
        const fechamodificacion = i[9];
        const subestado = i[10];
        let analista = i[11];
        let fechapresentacion = i[12];
        const asesorLegal = i[13];
        const fechaAsignacionLegal = i[14];
        const fechaRevisionLegal = i[15];
        const revisorFinal = i[16];
        const fechaRevisionfinal = i[17];
        let observaciones = i[18];
        const actividadPendiente = i[19];

        const revisorRealLegal = i[38];
        const revisorRealFinal = i[39];

        const conciFonpec = i[20];
        const conciTasa = i[21];
        const informetecnico = i[23];
        const pvcierre = i[24];
        const fechapvcierre = i[25];
        const diasPlazo = i[26];
        const fechaVencimientoPlazo = i[27];
        const conclusionTecnico = i[28];
        const montoAfavor = i[29];
        const montoEnContra = i[30];
        const montoConsolidadoAjuste = i[31];
        const estadoFinalCierre = i[32];

        const cierreEjercicio = i[35];

        let carpeta = i[33];
        let idCarpeta = i[34];

        // Lógica optimizada para responsable
        let responsable = analista; // valor por defecto

        const tieneAsesorLegal = asesorLegal && asesorLegal.trim() !== '';
        const sinFechaRevisionLegal = !fechaRevisionLegal || fechaRevisionLegal.trim() === '';
        const conFechaRevisionLegal = fechaRevisionLegal && fechaRevisionLegal.trim() !== '';
        const sinFechaRevisionFinal = !fechaRevisionfinal || fechaRevisionfinal.trim() === '';

        if (
            tieneAsesorLegal &&
            (
                (estado === 'En Revisión' && sinFechaRevisionLegal) ||
                (estado === 'Informe Técnico Firmado' && conFechaRevisionLegal)
            )
        ) {
            responsable = asesorLegal;
        } else if (
            tieneAsesorLegal &&
            estado === 'En Revisión' &&
            conFechaRevisionLegal &&
            sinFechaRevisionFinal
        ) {
            responsable = revisorFinal;
        }

        additionalData.push({
            empresa: razon,
            cuit: cuit,
            expediente: expediente,
            observacionAuditoria: observacionAuditoria,
            cuerpoAuditor: cuerpoAuditor,
            analista: analista,
            ulMovimiento: fechamodificacion,
            tramite: 'Cierre Auditoría',

            estadoEmpresa: estado,
            plazoSubsanacion: fechaVencimientoPlazo,
            cantidadDiasProvi: diasPlazo,

            periodoIncial: periodoDesde,
            periodoFinal: periodoHasta,

            fechaRevisionLegal: formatFecha(fechaRevisionLegal),
            asesorLegal: asesorLegal,
            fechaRevisionFinal: formatFecha(fechaRevisionfinal),
            revisorFinal: revisorFinal,

            revisorRealLegal: revisorRealLegal,
            revisorRealFinal: revisorRealFinal,

            iftecnico: informetecnico,
            provi: pvcierre,
            fechaProvi: formatFecha(fechapvcierre),
            conclusionTecnico: conclusionTecnico,

            observaciones: observaciones,
            carpetaEmpresa: carpeta,

            conciFonpec: conciFonpec,
            conciTasa: conciTasa,
            fechaVencimientoPlazo: fechaVencimientoPlazo,
            montoAfavor: montoAfavor,
            montoEnContra: montoEnContra,
            montoConsolidadoAjuste: montoConsolidadoAjuste,
            estadoFinalCierre: estadoFinalCierre,
        });

        return [responsable, razon, cuit, expediente, observacionAuditoria, fechamodificacion, estado, subestado, actividadPendiente];
    });

    //console.log({ headers: header, data: data, additionalData: additionalData });
    return { headers: header, data: data, additionalData: additionalData };
}



function traerUsuarios() {

    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const userLogin = ss.getSheetByName('user')


    const [header, ...data] = userLogin.getDataRange().getDisplayValues().map(i => {
        let nombre = i[3];
        let mail = i[1];
        let funcion = i[4];
        let rol = i[6];
        let equipo = i[7];
        let estado = i[5];
        let linkCharacter = i[15];
        let timeSession = i[9]



        return [nombre, mail, funcion, rol, equipo, estado, linkCharacter, timeSession]
    })

    return { headers: header, data: data }

}

function verificarSesionEnServidor() {
  // Suponiendo que tenés el valor válido en una celda de una hoja
  const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg').getSheetByName('user');
  const sessionActual = ss.getRange("Q2").getValue(); // o de donde corresponda

  return sessionActual;
}

function verificarPassword(form) {
    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const userSpread = ss.getSheetByName('user');
    const userData = userSpread.getDataRange().getValues();

    const checkPasword = userData.filter(i => i[5] === 'VERDADERO' && i[1] === form.exampleInputEmail1).map(i => {
        const correo = i[1];
        const password = i[2];
        const name = i[0];
        const rol = i[6];
        const tarea = i[4];
        const equipo = i[7];
        const completeName = i[3];
        let session = i[8];
        let timeSession = i[9];
        let e_click = i[10];
        let access_day = i[11];
        let timers = i[12];
        let checkin = i[13];
        let character_type = i[14];
        let linkCharacter = i[15];

        if (password === form.exampleInputPassword1) {
            const usuario = {
                name,
                rol,
                tarea,
                equipo,
                completeName,
                session, timeSession, e_click, access_day, timers, checkin, character_type, linkCharacter
            }
            return usuario;
        } else {
            throw new Error('Correo o contraseña inválidos');
        }
    });

    if (checkPasword.length > 0) {
        return checkPasword;
    } else {
        throw new Error('Correo o contrasela inválidos');
    }
}


/* ENVÍOS Y REGISTRO EN LA BASE DE DATOS */

function guardarFormulario(empresa) {
    console.log(empresa)
    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const main = ss.getSheetByName("main");
    const rb = ss.getSheetByName('RB-25');
    const tableName = 'Anuales LEC';

    const buscarFila = (id) => {
        const ids = main.getRange(2, 1, main.getLastRow() - 1, 4).getValues().map(row => row[2]);
        const index = ids.indexOf(id);
        return index !== -1 ? index + 2 : -1;
    };

    const fila = buscarFila(empresa.id);
    console.log(fila)


    // Cache cell values to reduce API calls
    const currentValues = {
        estado: main.getRange(fila, 39).getValue(),
        revisorlegal: main.getRange(fila, 44).getValue(),
        subestado: main.getRange(fila, 41).getValue(),
        valorRevisado: main.getRange(fila, 46).getValue(),
        revisorfinal: main.getRange(fila, 47).getValue(),
        valorRevision: main.getRange(fila, 49).getValue(),
        valor1266: main.getRange(fila, 63).getValue()
    };

    // Prepare batch updates
    const updates = [];

    // Estado handling
    if (empresa.estado !== currentValues.estado) {
        updates.push({ row: fila, col: 39, value: empresa.estado });
        updates.push({ row: fila, col: 40, value: new Date() });

        if (empresa.estado !== 'Aprobado' && empresa.estado !== 'Providencia Notificada' && empresa.estado !== 'Solicitud de Baja Otorgada') {
            if (empresa.estado === 'En Revisión' && !currentValues.revisorlegal) {
                updates.push({ row: fila, col: 44, value: AsignarUsuarios.asignacionLec('Legal', '', tableName) });
                updates.push({ row: fila, col: 45, value: new Date() });
            } else if (empresa.estado === 'Subsanación') {
                updates.push({ row: fila, col: 53, value: empresa.fechaSubsanacion });
                updates.push({ row: fila, col: 75, value: empresa.cantidadDias });
            } else {
                updates.push({ row: fila, col: 53, value: '' });
                updates.push({ row: fila, col: 75, value: '' });
            }
        } else {
            updates.push({ row: fila, col: 41, value: 'cerrada' });
        }
    }

    // Revisado handling
    if (empresa.revisionLegal && !currentValues.valorRevisado && !currentValues.revisorfinal) {
        updates.push({ row: fila, col: 46, value: new Date() });
        updates.push({ row: fila, col: 81, value: empresa.revisorRealLegal });
        updates.push({ row: fila, col: 47, value: AsignarUsuarios.asignacionLec('Revisor', '', tableName) });
    }

    // RevisionFinal handling
    if ((empresa.revisionFinal && empresa.revisionFinal !== 'FALSE') && !currentValues.valorRevision) {
        updates.push({ row: fila, col: 49, value: new Date() });
        updates.push({ row: fila, col: 82, value: empresa.revisorRealFinal });
    }

    // Revi1266 handling
    if ((empresa.revi1266 && empresa.revi1266 !== 'FALSE' && empresa.revi1266 !== false) && !currentValues.valor1266) {
        updates.push({ row: fila, col: 63, value: 'TRUE' });
    }

    // Map of field data to update
    const fieldUpdates = [
        { value: empresa.actividad, col: 19, current: main.getRange(fila, 19).getValue() },
        { value: empresa.periodoIncial, col: 9, current: main.getRange(fila, 9).getValue() },
        { value: empresa.periodoFinal, col: 10, current: main.getRange(fila, 10).getValue() },
        { value: empresa.tamanio, col: 8, current: main.getRange(fila, 8).getValue() },
        { value: empresa.masaSalarial, col: 26, current: main.getRange(fila, 26).getValue() },
        { value: empresa.nomina, col: 25, current: main.getRange(fila, 25).getValue() },
        { value: empresa.ventasTotales, col: 20, current: main.getRange(fila, 20).getValue() },
        { value: empresa.ventasPromo, col: 21, current: main.getRange(fila, 21).getValue() },
        { value: empresa.exportaciones, col: 31, current: main.getRange(fila, 31).getValue() },
        { value: empresa.ventasTotalesanio2, col: 76, current: main.getRange(fila, 76).getValue() },
        { value: empresa.ventasPromoanio2, col: 77, current: main.getRange(fila, 77).getValue() },
        { value: empresa.exportacionesanio2, col: 78, current: main.getRange(fila, 78).getValue() },
        { value: empresa.capacitacion, col: 29, current: main.getRange(fila, 29).getValue() },
        { value: empresa.imasd, col: 27, current: main.getRange(fila, 27).getValue() },
        { value: empresa.calidadTipo, col: 33, current: main.getRange(fila, 33).getValue() },
        { value: empresa.calidadEstado, col: 34, current: main.getRange(fila, 34).getValue() },
        { value: empresa.observaciones, col: 50, current: main.getRange(fila, 50).getValue() },
        { value: empresa.iftecnico, col: 56, current: main.getRange(fila, 56).getValue() },
        { value: empresa.actoOprovi, col: 57, current: main.getRange(fila, 57).getValue() },
        { value: empresa.numacto, col: 58, current: main.getRange(fila, 58).getValue() },
        { value: empresa.fechaActo, col: 59, current: main.getRange(fila, 59).getValue() },
        { value: empresa.fechaProvi, col: 60, current: main.getRange(fila, 60).getValue() },
        { value: empresa.fechaInscripcion, col: 61, current: main.getRange(fila, 61).getValue() },
        { value: empresa.reviEmision, col: 80, current: main.getRange(fila, 80).getValue() }
    ];

    // Only update if values have changed
    fieldUpdates.forEach(field => {
        if (field.value !== field.current) {
            updates.push({ row: fila, col: field.col, value: field.value });
        }
    });

    // Apply all updates in batch
    updates.forEach(update => {
        main.getRange(update.row, update.col).setValue(update.value);
    });

    // Add record to RB-25 sheet
    rb.appendRow([
        empresa.id,
        empresa.empresa,
        empresa.cuit,
        empresa.estado,
        empresa.dataUser?.completeName || '',
        empresa.dataUser?.rol || '',
        empresa.dataUser?.tarea || '',
        empresa.dataUser?.equipo || '',
        new Date()
    ]);

    SpreadsheetApp.flush();
    return true;
}

function guardarFormularioCierre(empresa) {
    console.log(empresa);
    const ss = SpreadsheetApp.openById('1EtfdMDuLgLr5KZVnjxGduUWH31ABUjgtmiAMUZCzW9c');
    const main = ss.getSheetByName("principal_cierres");
    const rb = ss.getSheetByName('RB-25');
    const tableName = 'Anuales LEC';

    const buscarFila = (id) => {
        const ids = main.getRange(2, 8, main.getLastRow() - 1, 1).getValues().flat(); // Columna H
        const index = ids.indexOf(id);
        return index !== -1 ? index + 2 : -1;
    };

    const fila = buscarFila(empresa.id);
    if (fila === -1) {
        console.error("ID no encontrado en la hoja.");
        return false;
    }

    const valoresActuales = main.getRange(fila, 1, 1, main.getLastColumn()).getValues()[0];
    const updates = [];

    // Campos comunes
    const campos = {
        periodoIncial: 5,
        periodoFinal: 6,
        estado: 9,
        revisionLegal: 16,
        revisionFinal: 18,
        iftecnico: 24,
        actoOprovi: 25,
        fechaProvi: 26,
        observaciones: 19,
        revisorLegal: 14,
        revisorFinal: 17,
        cuerpoAuditor: 7,
        observacionAuditoria: 4,
        cantidadDiasProvi: 27,
        conclusionTecnico: 29,
        montoAfavor: 30,
        montoEnContra: 31,
        estadoFinalCierre: 33,
        revisorRealLegal: 39,
        revisorRealFinal: 40,
    };

    // Estado y lógica especial
    /* aca está el error en mi opinión fs.  */
    const estadoActual = valoresActuales[campos.estado - 1];
    const revisorLegalActual = valoresActuales[campos.revisorLegal - 1];
    const revisorFinalActual = valoresActuales[campos.revisorFinal - 1];
    const valorRevisado = valoresActuales[campos.revisionLegal - 1];
    const valorRevision = valoresActuales[campos.revisionFinal - 1];

    if (empresa.estado && empresa.estado !== estadoActual) {
        updates.push({ row: fila, col: campos.estado, value: empresa.estado });
        updates.push({ row: fila, col: 10, value: new Date() }); // Asumimos columna 10 = fecha estado

        // Estado lógico
        if (!['Aprobado', 'Solicitud de Baja Otorgada'].includes(empresa.estado)) {
            if (empresa.estado === 'En Revisión' && !revisorLegalActual) {
                const nuevoRevisor = AsignarUsuarios.asignacionLec('Legal', '', tableName);
                updates.push({ row: fila, col: campos.revisorLegal, value: nuevoRevisor });
                updates.push({ row: fila, col: 15, value: new Date() }); // fecha asignación legal, col 17
                updates.push({ row: fila, col: 10, value: new Date() });
            } else if (empresa.estado === 'Providencia Notificada') {
                updates.push({ row: fila, col: campos.fechaProvi, value: empresa.fechaProvi });
                updates.push({ row: fila, col: 27, value: empresa.cantidadDiasProvi });
                updates.push({ row: fila, col: 10, value: new Date() });
            }

        }

    }

    updates.push({ row: fila, col: campos.estadoFinalCierre, value: empresa.estadoFinalCierre });



    // Asignación revisión legal
    if (empresa.revisionLegal && !valorRevisado && !revisorFinalActual) {
        updates.push({ row: fila, col: campos.revisionLegal, value: new Date() });
        updates.push({ row: fila, col: campos.revisorFinal, value: AsignarUsuarios.asignacionLec('Revisor', '', tableName) });
        updates.push({ row: fila, col: campos.revisorRealLegal, value: empresa.revisorRealLegal });
    }

    // Revisión final
    if (empresa.revisionFinal && !valorRevision) {
        updates.push({ row: fila, col: campos.revisionFinal, value: new Date() });
        updates.push({ row: fila, col: campos.revisorRealFinal, value: empresa.revisorRealFinal });
    }

    // Campos generales
    for (const campo in campos) {
        if (['estado', 'revisionLegal', 'revisionFinal', 'revisorLegal', 'revisorFinal', 'estadoFinalCierre', 'fechaProvi', 'cantidadDiasProvi'].includes(campo)) continue;

        const col = campos[campo];
        const nuevoValor = empresa[campo];
        const actual = valoresActuales[col - 1];
        const valorFormateado = typeof nuevoValor === 'boolean' ? nuevoValor.toString() : nuevoValor;

        if (valorFormateado !== '' && valorFormateado !== actual.toString()) {
            updates.push({ row: fila, col, value: nuevoValor });
        }
    }

    // Aplicar updates
    updates.forEach(update => {
        main.getRange(update.row, update.col).setValue(update.value);
    });

    // Registro en hoja RB-25
    rb.appendRow([
        empresa.id,
        empresa.empresa,
        empresa.cuit,
        empresa.estado,
        empresa.dataUser?.completeName || '',
        empresa.dataUser?.rol || '',
        empresa.dataUser?.tarea || '',
        empresa.dataUser?.equipo || '',
        new Date()
    ]);

    SpreadsheetApp.flush();
    return true;
}


function getRegistro() {

    const ss = SpreadsheetApp.openById("1iGmEB8l0NYsoUuGED1F1Mj8uglg3Wpoah4BiBygLHIU");
    const hoja = ss.getSheetByName("Registro Empresas LEC");
    const rango = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();

    const arreglo = rango.map(i => {
        const cuit = i[1];
        const rl = i[8];
        const periodo = (i[21]).replace(/[^a-zA-Z ]/g, "").trim().replace(" ", " - ");
        let ifActo = (i[15]).slice(0, 2) === "RS" ? "Resolución" : "Disposición";
        const numActo = i[16];
        const fechaInscripcion = i[19];
        const provincia = i[12];
        const correo = i[6];
        const actividad = i[11];



        return [cuit, rl, periodo, ifActo, numActo, fechaInscripcion, provincia, correo, actividad]
    });

    return arreglo;

}

function obtenerArchivosDesdeUrl(urlCarpeta) {

    let ubic = urlCarpeta.split('/').pop();
    const folder = DriveApp.getFolderById(ubic);
    const files = folder.getFiles();
    const resultados = [];

    while (files.hasNext()) {
        const file = files.next();
        resultados.push({
            nombre: file.getName(),
            url: file.getUrl(),
            tipo: file.getMimeType()
        });
    }

    return (resultados);
}

function calculadoraDiasHabiles(fechaInicial, agregarDias) {
    try {

        if (!fechaInicial || isNaN(new Date(fechaInicial))) {
            throw new Error("Fecha de inicio no válida: " + fechaInicial);
        }
        if (isNaN(agregarDias)) {
            throw new Error("Número de días no válido: " + agregarDias);
        }

        var sheet = SpreadsheetApp.openById('1RSkpO0MhsaSjaITbMoRr-9vLdLT3Z7ws4FBCmdajM-Q').getSheetByName('Hoja 1');
        var feriados = sheet.getRange('B:B').getValues().flat().filter(String).map(date => new Date(date));

        var currentDate = new Date(fechaInicial);
        if (isNaN(currentDate)) {
            throw new Error("Fecha de inicio no válida: " + fechaInicial);
        }

        var diasAgregados = 0;

        // Mover la fecha al siguiente día hábil antes de empezar a contar
        currentDate.setDate(currentDate.getDate() + 1);
        while (currentDate.getDay() == 0 || currentDate.getDay() == 6 || esFeriado(currentDate, feriados)) {
            currentDate.setDate(currentDate.getDate() + 1);
        }

        while (diasAgregados < agregarDias) {
            currentDate.setDate(currentDate.getDate() + 1);

            if (currentDate.getDay() != 0 && currentDate.getDay() != 6 && !esFeriado(currentDate, feriados)) {
                diasAgregados++;
            }
        }
        return currentDate.toISOString();
    } catch (error) {
        throw error;
    }
}

function esFeriado(date, feriados) {
    for (var i = 0; i < feriados.length; i++) {
        if (feriados[i].toDateString() == date.toDateString()) {
            return true;
        }
    }
    return false;
}


function getModels() {
    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const modelos = ss.getSheetByName("modelos");
    const dataModelos = modelos.getDataRange().getValues().filter((row) => row[0] !== "");


    return dataModelos;
}

function selectedCharacter(userName, character, imgSrc){

    console.log(userName)
    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const user = ss.getSheetByName("user");

    const buscarFila = (userName) => {
        const ids = user.getRange(2, 1, user.getLastRow() - 1, 4).getValues().map(row => row[0]);
        const index = ids.indexOf(userName);
        return index !== -1 ? index + 2 : -1;
    };

    const fila = buscarFila(userName);
    console.log(fila)

    user.getRange(fila, 14).setValue(true);
    user.getRange(fila, 15).setValue(character);
    user.getRange(fila, 16).setValue(imgSrc);

    return true;
}