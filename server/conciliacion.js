function arregloConciliacion(user) {
  const ss = SpreadsheetApp.openById('1NkTqXq-rlYZ3wj7Gdek_BgfjoavRlAO2UrCj4nDtjqQ');
  const hoja = ss.getSheetByName("LJK");

  const rango = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  const usuario = user.completeName



  const registroUso = ss.getSheetByName("registro_uso");
  const registroUsoFiltrada = registroUso.getDataRange().getDisplayValues();
  const ultimaActualizacion = registroUsoFiltrada
    .filter(i => i[0] === "contable")
    .map(i => i[2]);

  const ultimaFecha = ultimaActualizacion[ultimaActualizacion.length - 1];


  console.log(ultimaFecha)
  registroUso.appendRow([usuario, "consultar conciliacion", new Date()]);

  return {
    datos: rango,
    ultimaFecha: ultimaFecha
  };
}