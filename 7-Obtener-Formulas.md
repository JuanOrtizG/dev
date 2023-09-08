# Obtener fórmulas de un rango de celdas
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");    //Obtenemos el objeto rango.
    range.load("formulas");                 //Cargamos la formula de la hoja excel al objeto rango.
                                            //No carga datos, ni direccion, solo carga las formulas.
    await context.sync();

    console.log(JSON.stringify(range.formulas, null, 4)); //Convertimos a JSON los datos.
});
```