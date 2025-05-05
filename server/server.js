function getData() {
    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    //const sheet = ss.getSheetByName('main');
    const sheet = ss.getSheetByName('Copia de main');
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
        'Tareas Pendientes'
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
            tareaPendiente
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



function getCierres() {

    const ssex = SpreadsheetApp.openById('1EtfdMDuLgLr5KZVnjxGduUWH31ABUjgtmiAMUZCzW9c');
    const bienal = ssex.getSheetByName('principal_cierres');
    const rango = bienal.getRange(1, 1, bienal.getLastRow(), bienal.getLastColumn()).getDisplayValues();

    const [header, ...data] = rango.map(i => {
        let id = i[6];
        let analista = i[7];
        let razon = i[2];
        let expediente = i[6];
        let estado = i[19];
        let fechapresentacion = i[4];
        let cuit = i[1];
        let actividad = "";
        let observaciones = i[25];
        const conclusion = i[3];
        const ulrevisor = i[10];
        const fechaRevision = i[11];
        const informetecnico = i[12];
        const pvcierre = i[13];
        const fechapvcierre = i[14];
        const fecharevifinal = i[17];
        const revisorFinal = i[20];
        const conciTasa = i[16];
        const conciFonpec = i[15];
        const subestado = i[18];
        const fechamodificacion = i[22];
        const actividadPendiente = i[27];
        let responsable;
        const diasPlazo = i[28];
        const montoConsolidadoAjuste = i[31];
        const estadoFinalCierre = i[32];
        const fechaVencimientoPlazo = i[33];
        const conclusionTecnico = i[34];

        const montoAfavor = i[29];
        const montoEnContra = i[30];

        if ((ulrevisor != '' && estado == 'Revisión Cierre Bienal' && fechaRevision == '') || (ulrevisor != '' && estado == 'Informe Técnico Firmado Cierre Bienal' && fechaRevision != '')) {
            responsable = ulrevisor;
        } else if (ulrevisor != '' && estado == 'Revisión Cierre Bienal' && fechaRevision != '' && fecharevifinal == '') {
            responsable = revisorFinal;
        } else {
            responsable = analista;
        }

        let carpeta = i[35];

        /*         const findteam = dataUsuarios.find(r => r[0] === responsable);
                const equipo = findteam ? findteam[7] : null; */

        return [responsable, razon, cuit, expediente, conclusion, fechamodificacion, estado, subestado, actividadPendiente];
    });

    //console.log({ headers: header, data: data }) 
    return { headers: header, data: data }

}


function traerUsuarios() {

    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const userLogin = ss.getSheetByName('user')


    const [header, ...data] = userLogin.getDataRange().getDisplayValues().map(i => {
        let nombre = i[3];
        let mail = i[1];
        let funcion = i[4];
        let rol = i[6];
        let equipo = i[7]
        let estado = i[5]

        return [nombre, mail, funcion, rol, equipo, estado]
    })
    return { headers: header, data: data }

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

        if (password === form.exampleInputPassword1) {
            const usuario = {
                name,
                rol,
                tarea,
                equipo,
                completeName
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
    const main = ss.getSheetByName("Copia de main");
    const rb = ss.getSheetByName('RB-23');
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
    if (empresa.revisado && !currentValues.valorRevisado && !currentValues.revisorfinal) {
        updates.push({ row: fila, col: 46, value: new Date() });
        updates.push({ row: fila, col: 47, value: AsignarUsuarios.asignacionLec('Revisor', '', tableName) });
    }

    // RevisionFinal handling
    if ((empresa.revisionFinal && empresa.revisionFinal !== 'FALSE') && !currentValues.valorRevision) {
        updates.push({ row: fila, col: 49, value: new Date() });
    }

    // Revi1266 handling
    if ((empresa.revi1266 && empresa.revi1266 !== 'FALSE' && empresa.revi1266 !== false) && !currentValues.valor1266) {
        updates.push({ row: fila, col: 63, value: empresa.revi1266 });
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

    // Add record to RB-23 sheet
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





