function getDataRepo() {
    const sheet = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg').getSheetByName('main');
    const usuarios = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg').getSheetByName("user");

    const rango = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
    const dataUsuarios = usuarios.getDataRange().getDisplayValues();
    const cuits = rango.map(i => { return i[4] });
    const paresEmpresaNombre = AsignarUsuarios.devolverNombreEmpresas(cuits);

    const [header, ...data] = sheet.getDataRange().getDisplayValues().map(i => {
        let id = i[2];
        let cuit = i[4];
        let analista = i[41];
        let empresa = i[3];

        const name = paresEmpresaNombre[cuit] !== undefined ? paresEmpresaNombre[cuit] : empresa;

        let expediente = i[2];
        const tamanio = i[7];
        const tramite = i[5];
        let estado = i[38];
        let fechapresentacion = i[6];
        let actividad = i[18];
        let observaciones = i[49];
        const conclusion = i[55];
        const asesorLegal = i[43];
        const fechaRevision = i[45];
        const informetecnico = i[55];
        const pvcierre = i[56];
        const numeroActo = i[57];
        const fechaacto = i[58];
        const fechapvcierre = i[59];
        const fechaInscripción = i[60];
        const fecharevifinal = i[48];
        const revisorFinal = i[46];
        const conciTasa = i[67];
        const conciFonpec = i[70];
        const periodoIncio = i[8];
        const periodoFinal = i[9];
        const fechaSubsanacion = i[52];
        const plazoSubsanacion = i[54];
        const diasPlazo = i[74];

        const informadoAfip = i[62];

        /* REQUISITOS */

        const ventasTotal = i[19];
        const ventasPromo = i[20];
        const empleados = i[24];
        const masaSalarial = i[25];
        const imasd = i[26];
        const capacitacion = i[28];
        const exportaciones = i[30];
        const calidadTipo = i[32];
        const calidadEstado = i[33];
        const tareaPendiente = i[50];

        const ventasTotales2 = i[75]
        const ventasPromo2 = i[76]
        const expor2 = i[77];

        const microMenor = i[22]

        const ulMovimiento = i[39];

        let subestado = (i[40] == 'abierta') ? true : false;
        let responsable;

        if (estado === 'Subsanación' && plazoSubsanacion === 'TRUE') {
            subestado = false;
        } else {
            subestado
        }

        if ((asesorLegal != '' && estado == 'En Revisión' && fechaRevision == '') || (asesorLegal != '' && estado == 'Informe Técnico Firmado' && fechaRevision != '') || (estado === 'Con Observaciones' || estado === 'Confeccionando Acto' || estado === 'A la Firma SSEC' || estado === 'En Despacho' || estado === 'Prefirma' || estado === 'En Jurídicos' || estado === 'Revisión Cierre Bienal' || estado === 'Informe Técnico Firmado Cierre Bienal')) {
            responsable = asesorLegal;
        } else if (asesorLegal != '' && estado == 'En Revisión' && fechaRevision != '' && fecharevifinal == '') {
            responsable = revisorFinal;
        } else {
            responsable = analista;
        }

        const findteam = dataUsuarios.find(r => r[0] === responsable);
        const equipo = findteam ? findteam[7] : null;

        let carpeta = i[63];
        const datoEmision = i[79];

        return [responsable, name, cuit, expediente, tramite, ulMovimiento, estado]
    })

    console.log({ headers: header, data: data })
    //return { headers: header, data: data };
}

function traerUsuariosR() {

    const ss = SpreadsheetApp.openById('1Ux_aSHfhQhmKIONJTOvqXzdpof80zgPzPfamQt_DqDg');
    const userLogin = ss.getSheetByName('user')


    const [header, ...data] = userLogin.getDataRange().getDisplayValues().map(i => {
        let nombre = i[3];
        let mail = i[1];
        let funcion = i[4];
        let rol = i[6];
        let equipo = i[7];
        let estado = i[8];
        let session = i[9];
        let timeSession = i[10];
        let e_click = i[11];
        let access_day = i[12];
        let timers = i[13];
        let checkin = i[14];
        let character_type = i[15];
        let dark_mode = i[16];

        return [nombre, mail, funcion, rol, equipo, estado, session, timeSession, e_click, access_day, timers, checkin, character_type, dark_mode]
    })
    console.log({ headers: header, data: data })

}