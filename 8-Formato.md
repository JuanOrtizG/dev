# Establecimiento del formato de intervalo mediante la API de JavaScript de Excel
Color de la Fuente
Color de Relleno
Formato de Número

# Establecer el color de fuente y el color de relleno
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E2");    //Obtenemos el objeto rango.
    range.format.fill.color = "#4472C4";    //Color de relleno con objeto rango.
    range.format.font.color = "white";      //Color de Fuente con objeto rango.

    await context.sync();
});
```
# Establecer el formato de número
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ]; // Indicaremos con ello y la propiedad numberFormat el numero de digitos decimales.

    let range = sheet.getRange("D3:E5");
    range.numberFormat = formats;   //Ajustamos el formato de números

    await context.sync();
});
```


