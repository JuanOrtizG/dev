## Obtener valores de un rango de celdas
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");    //Obtenemos el Objeto Rango.
    range.load("values");                   //Cargamos los valores de la hoja en el Objeto Rango.
    await context.sync();

    console.log(JSON.stringify(range.values, null, 4)); // Convertimos los datos en formato JSON.
});
```

## Obtener texto de un rango de celdas
La text propiedad de un rango especifica los valores de visualización de las celdas del rango. Incluso si algunas celdas de un rango contienen fórmulas.
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("text"); //Nos presenta las formulas y numeros en formato texto.
    await context.sync();

    console.log(JSON.stringify(range.text, null, 4));
});
```

1. Obtener valores con rango por direccion
2. Obtener valores con  rango por nombre
3. Obtener valores con  rango usado
4. Obtener valores con  rango completo

## 1. Obtener valores con rango por direccion

```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    
    let range = sheet.getRange("B2:C5"); //Obtengo el objeto rango
    range.load("values");   // Almaceno los valores de excel en el objeto rango.
    await context.sync();
    
    console.log(`The values of the range B2:C5 is "${range.values}"`); // Recupero los valores de los datos de Excel.
});
```
## 2. Obtener valores con  rango por nombre
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("MyRange"); //Obtengo el objeto rango, MyRange es el nombre rango.
    range.load("values");// Almaceno los valores de excel en el objeto rango.
    await context.sync(); //Sincronizamos.

    console.log(`The values of the range "MyRange" is "${range.values}"`); //Recupero los valores de Excel.
});
```
## 3. Obtener valores con  rango usado
El rango usado es el rango más pequeño que abarque todas las celdas de la hoja de cálculo que tengan asignado un valor o un formato. Si toda la hoja de cálculo está en blanco, el getUsedRange() método devuelve un rango que consta solo de la celda superior izquierda.

```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getUsedRange();
    range.load("values");
    await context.sync();
    
    console.log(`The values of the used range in the worksheet is "${range.values}"`);
});
```

## 4. Obtener valores con  rango completo
El ejemplo de código siguiente obtiene todo el intervalo de hojas de cálculo de la hoja de cálculo denominada Sample, carga su values propiedad y escribe un mensaje en la consola.
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange();
    range.load("values");
    await context.sync();
    
    console.log(`The values of the entire worksheet range is "${range.values}"`);
});
```