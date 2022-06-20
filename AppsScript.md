# [Apps Script](https://www.youtube.com/playlist?list=PLG1qdjD__qH4dyXq4sM03Rf0RFhB_4tbm)
Powered by: Aula en la nube

## Indice
- [Crear y abrir documentos en Google DOCS](#crear-y-abrir-documentos-en-google-docs)
- [Condiciones](#condiciones)
- [Bucle](#bucle)
- [Bucle FOR IN y análisis de atributos](#bucle-for-in-y-análisis-de-atributos)
- [Crear funciones con parámetros y menú personalizado](#crear-funciones-con-parámetros-y-menú-personalizado)
- [Crear formulario HTML en Google Docs](#crear-formulario-html-en-google-docs)   
- [Crear tablas personalizadas con formulario HTML](#crear-tablas-personalizadas-con-formulario-html)
- [Combinar correspondencia en Google Docs](#combinar-correspondencia-en-google-docs)
- [Combinar correspondencia múltiple](#combinar-correspondencia-múltiple)
- [Añadir ventana de verificación sobre combinar correspondencia](#añadir-ventana-de-verificación-sobre-combinar-correspondencia)


## Installar Google Apps Script
1. Google Apps Script and open it then start scripting. Also get access from Goolge Drive

## Crear y abrir documentos en Google DOCS   

*[Indice](#indice)*
1. Creacion de un Documento en Blanco en Google Docs
```js
function CrearDcumento() {
  DocumentApp.create('Data Analyst')
//Listo Creamos un Docs al Ejecutar la funcion. 
}
```
2. Para guardar el documento y manipularlo lo debemos guardar en una variable. El objetivo es poder acceder a metodos, a diferetes niveles de profundidad y que el codigo sea legible realemten podemos colocar una cadena de metodos sin afectar el funcionamiento del codigo.

```js
function CrearDcumento() {
 var documento = DocumentApp.create('Data Analyst')
 var documento_body= documento.getBody()
 documento_body.appendParagraph('Escribiendo en Docs Creado desde Apps Script')
//La legibilidad no es la mejor: 
documento.getBody().appendParagraph('Text')
}
```
3. Escribir sobre un Documento Existente. Los Docs normalmente tienen un Id el cual podemos extraer desde la URL del documento de Google.
```js 
function OpenDocumento() {
  var documento = DocumentApp.openById('Id')
  documento.getBody().appendParagraph('Text')
}
```
## Condiciones   

*[Indice](#indice)*

```js
function ModificarFormato() {

Logger.log('Imprime Texto en pantalla');

//Buscamos el documento a partir de su Id
//Accedemos al cuerpo del documento, luego al parrafo
//El parrafo un elemento tipo lista y lo accedemos a partir de getText()
  var documento = DocumentApp.openById('1S9s6tjqJapx7BL2fnxzsWRwjNmS3UoGgrCINLqYjRMQ');
  var parrafo = documento.getBody().getParagraphs();


//If: Si la condicion se cumple usamos setText() para modificar el parrafo[0]
if (parrafo[0].getText() == 'Hola Data Analyst Jr. ') parrafo[0].setText('Soy Data Scientist');

//If/Else: Alineamos a la Izquierda si se encuentra a la derecha y viceversa
if (parrafo[0].isLeftToRight) 
  parrafo[0].setLeftToRight(false);
else
  parrafo[0].setLeftToRight(true);


//If/Else If/Else: Obtiene el numero de espacios despues del parrafo señalado, a partir de ese numero
//modifica la cantidad de espacios si se cumple alguna de la condiciones indicadas.
if (parrafo[0].getSpacingAfter()==null)
  parrafo[0].setSpacingAfter(10);
else if (parrafo[0].getSpacingAfter()==10)
  parrafo[0].setSpacingAfter(20);
else if (parrafo[0].getSpacingAfter()==20)
  parrafo[0].setSpacingAfter(30);  
else
  parrafo[0].setSpacingAfter(null);

}
```
## Bucle   

*[Indice](#indice)*
A continuaciones aplicaremos loops en documentos, pero para eso debemos abrir el Documento sobre el que se escribira en la ejecucion del bucle.
1. Inicializamos `contador`, definimos una condicion a partir de contador e incrementamos `contador++`.
2. En cada ejecucion usamos `appendParagraph()` para escribir sobre el documento. 

- **While Loop**
```js
function CrearLoopParagraphs(){
  var documento = DocumentApp.openById('1S9s6tjqJapx7BL2fnxzsWRwjNmS3UoGgrCINLqYjRMQ');
  var contador = 0;
  while(contador<100){
    documento.getBody().appendParagraph(contador+' Data Analysis is the Future')
    contador++;
  }
}
```
- **For Loop**
```js
function CrearLoopParagraphs(){
  var documento = DocumentApp.openById('1S9s6tjqJapx7BL2fnxzsWRwjNmS3UoGgrCINLqYjRMQ');

  for(var contador = 0;contador < 100;contador++)
    {
      documento.getBody().appendParagraph(contador+' Data Analysis is the Future');
    }
}
```
- **While Loop/Condicional**
```js
function ChangeDocThruLoop(){
  var documento = DocumentApp.openById('1S9s6tjqJapx7BL2fnxzsWRwjNmS3UoGgrCINLqYjRMQ');
  var parrafo = documento.getBody().getParagraphs();
  var contador = 0;
  while(contador < parrafo.length){
  var thisparagraph = parrafo[contador].getText();
//Se modifica el parrafo si la condicion se cumple.
    if (thisparagraph =='2 Data Analysis is the Future'){
      parrafo[contador].setText('Analysis is the Future');
    }
    contador++;
  }
}
```

## Bucle FOR IN y análisis de atributos   
*[Indice](#indice)*

Este codigo ademas de usar `For In` similar a `For Each` utilizar metodos que hacen que un vector se convierta en otro vector y ademas se modifique a partir de sus atributos. 
```js
function AnalizarParrafos(){
  var doc = DocumentApp.openById('1S9s6tjqJapx7BL2fnxzsWRwjNmS3UoGgrCINLqYjRMQ');

//Acedemos al texto a partir de getParagraphs()
  var parrafos = doc.getBody().getParagraphs();
//Inicializamos el tamaño del texto
  var size = 10;

//Analizar cada parrafo por separado.
  for(var parrafo in parrafos){
//Crear un vector con los attributos del parrafo
    var paragraphattrs = parrafos[parrafo].getAttributes();
//Recorrer el vector y validar los condicionales
    for(var paragraphattr in paragraphattrs){
//Si este en BOLD entonces retiramos ese atributo
      if (paragraphattr == 'BOLD' && paragraphattrs[paragraphattr] == true){
        parrafos[parrafo].setAttributes({'BOLD':false});
      }
      if (paragraphattr == 'FOREGROUD_COLOR' && paragraphattrs[paragraphattr] != null){
        parrafos[parrafo].setAttributes({'FOREGROUD_COLOR':'#00ffff','FOND_SIZE':size++});
      }
    }
  }
}
```

## Crear funciones con parámetros y menú personalizado   
*[Indice](#indice)*   

La funcion `onOpen` recibe este nombre por comvension para la creacion de un menu en el panel de herramientas de google Docs
```js
function onOpen(){
//createMenu('name_nav_menu') Es el nombre que tendra la nueva opcion del panel
  DocumentApp.getUi().createMenu('Options')
//addItem('name_option','name_function') No solo el nombre de la opcion sino que funcion desencadena.
  .addItem('Direct','CreateParagraph')
  .addItem('Promt','PromptWindow')
  .addToUi();
}
```
![image](https://user-images.githubusercontent.com/60556632/174460372-a7fd2c75-876e-4346-84b2-66ba13526444.png)   

La funcion `CreateParagraph()` simplemente accede al documento activo, reemplaza todo lo que pueda haber en `getBody()` por un texto y luego del texto un bucle `For` escribe varias veces sobre el documento 
```js
function CreateParagraph(){
//Debido a que el codigo se encuentra enlazado al Documento solo es suficiente
//el metodo getActiveDocument()
  var documento =  DocumentApp.getActiveDocument();
  documento.getBody().setText('Aqui Inicial el Bucle: ')
  for(var contador = 0; contador < 5;contador++){
    documento.getBody().appendParagraph('Texto Creado a partir de un Loop en Apps Script \n')
  }
}
```
![image](https://user-images.githubusercontent.com/60556632/174460385-8c74495d-2bb2-4dab-9761-d48286afc94f.png)    

La funcion `PromptWindow()` hace exactamente lo mismo sin embargo le pregunta el texto al usuario y le da opciones para aceptar o cancelar la ejecucion de la funcion. 

1. Acedemos al metodo de `interfaz de usuario` del docuemento en general y luego le indicamos que queremos crear una ventana emergente `prompt`.
2. Almacenamos la respuesta del usuario en `button -> Button.CANCEL|Button.OK`. Y en texto respectivamente.
3. Usamos la infomacion de `button` para establecer la condicion y usamos `text` para indicar que se debe escribir en el documento.
```js
function PromptWindow(){
  var ui =  DocumentApp.getUi();
  var ventana = ui.prompt('Nombre de Ventana Emergente','Informacion hacerca de la Ventana',ui.ButtonSet.OK_CANCEL);
  var button = ventana.getSelectedButton();
  var text = ventana.getResponseText();

  if(button == ui.Button.OK){
    var documento =  DocumentApp.getActiveDocument();
    documento.getBody().setText('Aqui Inicial el Bucle: ')
    for(var contador = 0; contador < 5;contador++){
      documento.getBody().appendParagraph(text)
    }
  }

  else if (button == ui.Button.CANCEL){
    ui.alert('Cancelaste la Operacion, intenta nuevamente')
  }
}
```
## Crear formulario HTML en Google Docs
*[Indice](#indice)* 

1. Creamos un formulario usando HTML y luego usamos JavaScript para llamar la funcion de Apps Script y pasarle lo argumentos obtenidos a partir del formulario.
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <label>Cuantos Parrafos?</label>
    <input type="number" id="cantidad"><br>

    <label>Introduce el contenido de Parrafos</label>
    <input type="text" id="frase"><br>

    <label>Introduce el Color</label>
    <input type="color" id="colores"><br>

    <label>Fecha</label>
    <input type="date" id="fecha"><br>

    <button type="button" onclick="GetInfoForm()">Escribir Documento</button>

  <script>
    function GetInfoForm(){
      var texto = document.getElementById('frase').value;
      var number = document.getElementById('cantidad').value;
      var colores = document.getElementById('colores').value;
      var fecha = document.getElementById('fecha').value;

      google.script.run.OpenFromJS(texto,number,colores,fecha)
      google.script.host.close(); 
    }
  </script>
  </body>
</html>
```   
![image](https://user-images.githubusercontent.com/60556632/174464563-c9faac74-fc00-44fb-874f-9b2fb5bf32ff.png)

2. Creamos un menu facil de acceder desde Google Docs
```js
function onOpen(){
//createMenu('name_nav_menu') Es el nombre que tendra la nueva opcion del panel
  DocumentApp.getUi().createMenu('Options')
//addItem('name_option','name_function') No solo el nombre de la opcion sino que funcion desencadena.
  .addItem('Form HTML','OpenForm')
  .addToUi();
}
```
3. La primera funcion a Ejecutar es aquella que abre el archivo HTML

```js
//La funcion `OpenForm()` abre al formulario y es la funcion desencadenada por el menu de herramientas
function OpenForm(){
//Me permite crear un HTML a partir de un archivo
  var html = HtmlService.createHtmlOutputFromFile('index.html')
  .setWidth(1000)
  .setHeight(700)

//Usado para proteger a los usuarios de HTML malicioso o JS 
//y que se ejecuta en un Sandbox que impone restricciones al codigo.
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  DocumentApp.getUi().showModalDialog(html,'Name');
}
```
4. Luego de abrir el formulario, el usuario ingresara la informacion y esta sera recolectada a travez de JavaScript para pasar los argumentos a la siguiente funcion. 

```js
function OpenFromJS(texto,numero,color,fecha){
  var documento = DocumentApp.getActiveDocument();
  documento.getBody().clear();
  documento.getBody().setAttributes({'FOREGROUND_COLOR':color});
  var title = documento.getBody().insertParagraph(0,'Darn Good Title')
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  for(var contador = 0;contador<numero;contador++){
    var paraph = documento.getBody().appendParagraph(texto);
  }

  documento.addFooter();
  var footer = documento.getFooter().appendParagraph('Modified on: '+fecha);
  footer.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}
```

## Crear tablas personalizadas con formulario HTML   
*[Indice](#indice)* 

En esta seccion se explica como crear una tabla en Google Docs. El usuario indica filas, columnas y que colores debe tener el encabezado y el cuerpo de la tabla. El proceso de ejecucion involucra una interfaz de usuario creada a partir de **HTML**, **CSS** y **JavaScript** la cual se comunica con **Apps Script** y ejecuta la logica de la creacion de la tabla en este lenguage.

1. Creamos un menu `Options` en Google Docs con la opcion `Insert Table`.   

![image](https://user-images.githubusercontent.com/60556632/174496903-8fb27843-cf90-43ba-afa6-5eefdf974e27.png)

```js
function onOpen(){
  DocumentApp.getUi().createMenu('Options')
  .addItem('Form HTML','OpenForm')
  .addItem('Insert Table','OpenCreateTable')
  .addToUi();
}
```
2. La opcion `Insert Table` activa la funcion `OpenCreateTable()` que se encarga de abrir el archivo **HTML**.

```js
function OpenCreateTable(){
  var html = HtmlService.createHtmlOutputFromFile('table.html')
  .setWidth(1000)
  .setHeight(700)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  DocumentApp.getUi().showModalDialog(html,'Este Codigo nos Permite insertar tablas en el Documento');   
}
```
  - Usamos Bootstrap y le damos estilo al HTML. 

```html
<!DOCTYPE html>
<html>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-0evHe/X+R7YkIZDRvuzKMRqM+OrBnVFBL6DOitfPri4tjfHxaWutUpFmBp4vmVor" crossorigin="anonymous">
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/js/bootstrap.bundle.min.js" integrity="sha384-pprn3073KE6tl6bjs2QrFaJGz5/SUsLqktiwsUTF55Jfv3qYSDhgCecCxMW52nD2" crossorigin="anonymous"></script>

<style>
.contenedor{
  display:flex;
  text-align:right;
  justify-content:space-around;
  flex-direction:row;
  flex-wrap: wrap;
  fond-size:large;
  border-radius: 5px;
  padding:50px;
  margin:40px;
  width: 80%;
  background-color:rgb(243,247,255);
  box-shadow:0px 2px 2px #aaaaaa;
}
.caja{
  line-height:50px;
}
</style>
```
- Creamos el formulario en **HTML** 

![image](https://user-images.githubusercontent.com/60556632/174497210-6ec01137-a359-4bd0-8584-5dc4d9d8de79.png)

```html
<head>
  <base target="_top">
</head>
<body>
  <div class="contenedor">
    <div class="caja">
      <label>Introduce la cantidad de Filas</label>
      <input type="range" id="filas" min="1" max="20" step="1" value="5">
      <output class="filas-output"></output><br>

      <label>Introduce la cantidad de Columnas</label>
      <input type="range" id="columnas" min="1" max="20" step="1" value="5">
      <output class="columnas-output"></output><br>

      <label>Introduce el Color Enabezado</label>
      <input type="color" id="colorE"><br>

      <label>Introduce el Color Par</label>
      <input type="color" id="colorP"><br>

      <label>Introduce el Color Impar</label>
      <input type="color" id="colorI"><br>

      <center><button type="button" class="btn btn-outline-primary" onclick="CrearTable()">Crear Tabla</button></center>

    </div>
  </div>
```
  - Capturamos los datos del Usuario usando **JS**
```js
<script>
    function CrearTable(){
      var rows = document.getElementById('filas').value;
      var columns = document.getElementById('columnas').value;
      var colorE = document.getElementById('colorE').value;
      var colorP = document.getElementById('colorP').value;
      var colorI = document.getElementById('colorI').value;
      google.script.run.ParamsTable(rows,columns,colorE,colorP,colorI);
      google.script.host.close(); 
    }
```
- Esta parte del codigo solo se encarga de mostrar el usuario el valor correspondiente a los `<input>` tipo `rage`.

```js
   //Indica el Id del objeto que se esta consultando
    const numRows = document.querySelector('#filas');
    //Indica la clase del Objeto donde se va a mostrar el # de rows
    const outRows = document.querySelector('.filas-output');
    //Se coloca para mostrar el valor inicial addEventListener() no lo muestra a menos que cambie.
    outRows.textContent =  numRows.value;
    //Obtiene y Actualiza constatemente outRows que es igual a numRows.value
    numRows.addEventListener('input',function(){outRows.textContent =  numRows.value;})

    const numCols = document.querySelector('#columnas');
    const outCols = document.querySelector('.columnas-output');
    outCols.textContent =  numCols.value;
    numCols.addEventListener('input',function(){outCols.textContent =  numCols.value;})


  </script>
  </body>
</html>
```
3. Creamos una funcion en Apps Script que se encarga de crear la tabla segun los parametros indicados por el usuario.

```js
function ParamsTable(rows,columns,colorE,colorP,colorI){

  var documento =  DocumentApp.getActiveDocument().getBody();

  var styleDocP = {};
  styleDocP[DocumentApp.Attribute.BACKGROUND_COLOR] = colorP;

  var styleDocE = {};
  styleDocE[DocumentApp.Attribute.BACKGROUND_COLOR] = colorE;
  styleDocE[DocumentApp.Attribute.BOLD] = true

  var styleDocI = {};
  styleDocI[DocumentApp.Attribute.BACKGROUND_COLOR] = colorI;

  var createTable =  documento.appendTable();
  var fila;
  var celda;
  for(var i=0;i<rows;i++){
    fila = createTable.appendTableRow();
    for(var j=0;j<columns;j++){
      celda = fila.appendTableCell('Celda '+j)
      if (i==0) celda.setAttributes(styleDocE);
      else if (i%2 == 0) celda.setAttributes(styleDocP);
      else celda.setAttributes(styleDocI);
    }
  }
}
```
![image](https://user-images.githubusercontent.com/60556632/174497243-41a114b6-0c1f-4431-9765-3ea5678f3ae3.png)

## Combinar correspondencia en Google Docs
*[Indice](#indice)*

A continuacion se muestra como podemos crear documentos personalizables a partir de los parametros indicados por el usuario los cuales fueron previamente asignados a una variable la cual se reemplazara en el documento.

1. Creamos una platilla y colocamos las variables `<<variable>>`.   

![image](https://user-images.githubusercontent.com/60556632/174512603-3dc17c6d-f89d-43b2-b0d1-d9a3507c07de.png)

2. Creamos el menu de `Google Docs` para que el usuario intereacture con la funcion.    
![image](https://user-images.githubusercontent.com/60556632/174512658-cd38de2f-9027-4837-959a-2e87a1c24b2d.png)

3. Abrimos el HTML usando `HtmlService.createHtmlOutputFromFile` como se ha hecho anteriormente.   
![image](https://user-images.githubusercontent.com/60556632/174512767-95d07ee8-8f9e-4614-9211-a4390ee5c428.png)

4. La funcion principal `CreateCertificate()` crea un nuevo **Documento** y luego reemplaza las variables `<<>>` en el texto, asi podemos utilizar la plantilla un numero ilimitado de veces y crear tantos documentos como queramos. 
```js
function CreateCertificate(curso,plataforma,funcion){

  var idCurrentDoc =  DocumentApp.getActiveDocument().getId();
  var currentDoc = DriveApp.getFileById(idCurrentDoc)

//Creamos un nuevo documento, si no lo hacemos perderemos la plantilla
  var newDoc= currentDoc.makeCopy('Curso: '+curso)
  var documento = DocumentApp.openById(newDoc.getId());

//Obtenemos la fecha en la que se creo el Certificado
  var date = new Date();
  var mes = date.getMonth();
  var day = date.getDate();
  var year = date.getFullYear();

//La funcion switch pemite reemplazar el mes numerico a tipo texto
  switch(mes){
    case 0: mes = 'Enero'; break;
    case 1: mes = 'Febrero'; break;
    case 2: mes = 'Marzo'; break;
    case 3: mes = 'Abril'; break;
    case 4: mes = 'Mayo'; break;
    case 5: mes = 'Junio'; break;
    case 6: mes = 'Julio'; break;
    case 7: mes = 'Agosto'; break;
    case 8: mes = 'Septiembre'; break;
    case 9: mes = 'Octubre'; break;
    case 10: mes = 'Noviembre'; break;
    case 11: mes = 'Diciembre'; break;
  }

//Reemplazamos las variables <<>> en el texto por aquellas ingresadas por el usuario.
  var certyDate = 'Certificado Emitido el dia '+day+' de '+mes+' de '+year;
  documento.getBody().replaceText('<<curso>>',curso);
  documento.getBody().replaceText('<<plataforma>>',plataforma);
  documento.getBody().replaceText('<<funcion>>',funcion);
  documento.getBody().replaceText('<<fecha>>',certyDate);
}
```
![image](https://user-images.githubusercontent.com/60556632/174513478-a463c626-6017-428f-adc6-f71bb63d6828.png)

## Combinar correspondencia múltiple    
*[Indice](#indice)*   
A diferencia del codigo anterior(HTML Form) ahora utilizamos **Google Sheets** para crear varios documentos.
1. Como no usamos **HTML** no necesitamos funciones adicionales para abrir **HTML** que luego le pasa los datos a **JavaScript** y luego se los pasa a **Apps Script**. Nada de lo anterior, solo creamos del codigo **Apps Script** desde **Spread Sheet**, creamos una funcion que se conecta con el documento platilla que recibe los datos desde filas y columnas y que se ejecuta en un bucle hasta que no encuentra mas filas.

2. Creamos la tabla que contiene la informacion que se colocara en el Documento ademas de un boton que al hacer click desencadenara la funcion.    
![image](https://user-images.githubusercontent.com/60556632/174523324-c1970d0d-3fe8-47d0-887b-a0295009dd02.png)    

3. `row` indica desde que fila iniciamos, `nameCel` es el nombre de la celda donde iniciamos, `currentSheet` se conecta con el Spread Sheet y finalmente `currentCel` se ubica en el rango señalado `nameCel` usando la funcion `getRange(nameCel)` 
```js
//Nos conectamos al documento donde se encuentra la plantilla utilizando su Id
  var currentDoc = DriveApp.getFileById('1U0WAP6dFPRYfZOW8jSPSJ8Vo2UthIdNaQV0zWs3oFGY')
  var row = 2;
  var nameCel = 'A'+row;
  var currentSheet = SpreadsheetApp.getActive();
  var currentCel =  currentSheet.getRange(nameCel);
```
4. El bucle solo se ejecuta solo si la celda donde estamos ubicados no esta vacia. 
```js
 while(!currentCel.isBlank()){...}
```
5. Luego de confirmar que la celda no esta vacia, no s desplazamos por la columnas de la fila y asignamos el valor de cada celda en el documento platilla.
```js
documento.getBody().replaceText('<<curso>>',currentSheet.getRange('A'+row).getValue());
documento.getBody().replaceText('<<plataforma>>',currentSheet.getRange('B'+row).getValue());
documento.getBody().replaceText('<<funcion>>',currentSheet.getRange('C'+row).getValue());
```
6. El codigo completo es muy similar al de la seccion anterior. 

```js
function InsertData() {
  var currentDoc = DriveApp.getFileById('1U0WAP6dFPRYfZOW8jSPSJ8Vo2UthIdNaQV0zWs3oFGY')
  var row = 2;
  var nameCel = 'A'+row;
  var currentSheet = SpreadsheetApp.getActive();
  var currentCel =  currentSheet.getRange(nameCel);

//Ejecuta el bucle si la celda no esta en blanco
  while(!currentCel.isBlank()){
  //Creamos un nuevo documento, si no lo hacemos perderemos la plantilla
    var newDoc= currentDoc.makeCopy('Curso: '+currentSheet.getRange('A'+row).getValue())
    var documento = DocumentApp.openById(newDoc.getId());

  //Obtenemos la fecha en la que se creo el Certificado
    var date = new Date();
    var mes = date.getMonth();
    var day = date.getDate();
    var year = date.getFullYear();

  //La funcion switch pemite reemplazar el mes numerico a tipo texto
    switch(mes){
      case 0: mes = 'Enero'; break;
      case 1: mes = 'Febrero'; break;
      case 2: mes = 'Marzo'; break;
      case 3: mes = 'Abril'; break;
      case 4: mes = 'Mayo'; break;
      case 5: mes = 'Junio'; break;
      case 6: mes = 'Julio'; break;
      case 7: mes = 'Agosto'; break;
      case 8: mes = 'Septiembre'; break;
      case 9: mes = 'Octubre'; break;
      case 10: mes = 'Noviembre'; break;
      case 11: mes = 'Diciembre'; break;
    }
  //Reemplazamos las variables <<>> en el texto por aquellas ingresadas por el usuario.
    var certyDate = 'Certificado Emitido el dia '+day+' de '+mes+' de '+year;
    documento.getBody().replaceText('<<curso>>',currentSheet.getRange('A'+row).getValue());
    documento.getBody().replaceText('<<plataforma>>',currentSheet.getRange('B'+row).getValue());
    documento.getBody().replaceText('<<funcion>>',currentSheet.getRange('C'+row).getValue());
    documento.getBody().replaceText('<<fecha>>',certyDate);


    //Shifting Down Cells
    row++;
    nameCel = 'A'+row;
    currentCel = currentSheet.getRange(nameCel);
  }
}
```
## Añadir ventana de verificación sobre combinar correspondencia    
*[Indice](#indice)* 

