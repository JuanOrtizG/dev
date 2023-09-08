# Establecimiento y obtención de valores de rango, texto o fórmulas mediante la API de JavaScript de Excel 

## Establecer fórmulas para un rango de celdas
El ejemplo de código siguiente establece las fórmulas de las celdas del rango E2:E6 y, después, establece el ancho de las columnas que mejor se ajusta a los datos.
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ]; // 4 filas y una columna de datos

    let range = sheet.getRange("E3:E6"); //obtenemos el objeto rango.
    range.formulas = data;               // Imprimimos los datos mediante el objeto rango.
    range.format.autofitColumns();       // Ajustamos las columnas según los datos.

    await context.sync();
});
```

## 