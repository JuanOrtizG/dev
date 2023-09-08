Es importante saber que la API de JavaScript para Excel no tiene un objeto o clase de "Celda". En su lugar, se definen todas las celdas de Excel como objetos Range. Los ejemplos que veremos son para
1. Obtener rango por direccion
2. Obtener rango por nombre
3. Obtener rango por rango usado
4. Obtener rango completo

## 1. Obtener rango por Dirección

```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    
    let range = sheet.getRange("B2:C5"); //Obtengo el objeto rango
    range.load("address");   // Almaceno las direcciones de excel en el objeto rango.
    await context.sync();
    
    console.log(`The address of the range B2:C5 is "${range.address}"`); // Recupero la direccion de los datos de Excel.
});
```
## 1. Obtener rango por Nombre
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("MyRange"); //Obtengo el objeto rango, MyRange es el nombre rango.
    range.load("address");// Almaceno los valores de excel en el objeto rango.
    await context.sync(); //Sincronizamos.

    console.log(`The address of the range "MyRange" is "${range.address}"`); //Recupero los valores de Excel.
});
```
## Obtener rango por rango usado
El rango usado es el rango más pequeño que abarque todas las celdas de la hoja de cálculo que tengan asignado un valor o un formato. Si toda la hoja de cálculo está en blanco, el getUsedRange() método devuelve un rango que consta solo de la celda superior izquierda.

```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getUsedRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the used range in the worksheet is "${range.address}"`);
});
```

## Obtener rango completo
El ejemplo de código siguiente obtiene todo el intervalo de hojas de cálculo de la hoja de cálculo denominada Sample, carga su address propiedad y escribe un mensaje en la consola.
```javascript
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the entire worksheet range is "${range.address}"`);
});
```