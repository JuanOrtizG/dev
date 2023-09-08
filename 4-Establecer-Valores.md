# Establecimiento y obtención de valores de rango, texto o fórmulas mediante la API de JavaScript de Excel

# 1. Establecer el valor para una única celda
En el ejemplo de código siguiente se establece el valor de la celda C3 en "5" y, después, se establece el ancho de las columnas que mejor se ajusta a los datos.
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("C3");   //Obtengo el objeto rango c3.
    range.values = [[ 5 ]];             // Cargamos la lista de sublista en 5.
    range.format.autofitColumns();      //AutoAjustamos las columnas.

    await context.sync();               //Sincronizamos con el servidor.
});
```

# 2. Establecer valores para un rango de celdas
El ejemplo de código siguiente establece los valores de las celdas del rango B5:D5 y, después, establece el ancho de las columnas que mejor se ajusta a los datos.
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let data = [
        ["Potato Chips", 10, 1.80], //[celda1, celda1, celda3]
    ];

    let range = sheet.getRange("B5:D5");// Obtenemos el objeto Rango
    range.values = data;                // Imprimimos mediante el objeto rango en la hoja.
    range.format.autofitColumns();      // Ajustamos las columnas al ancho de los datos.

    await context.sync();
});
```
