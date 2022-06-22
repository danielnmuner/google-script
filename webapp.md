### [Apps Script Fundamentos](https://www.youtube.com/playlist?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV)    

By: Mozart Alberto García de Haro

- [Creando una WebApp de 0 a 100](#creando-una-webApp-de-0-a-100)
- [Lista desplegable](#lista-desplegable)
- [Tablas html dinámicas](#tablas-html-dinámicas)
- [Calendar a Google Sheets](#calendar-a-google-sheets)


### Creando una WebApp de 0 a 100
1. Antes de empezar a desarrollar el codigo debemos verificar que el **HTML** este abriendo desde Apps Script y que este comunicado con los archivos **CSS** y **JS**.
- Codigo para abrir **HTML** y funcion `include()` para conectarnos con los demas archivos.
```js
//build-in function o funcion reservada
function doGet() {
  //Guardamos el HTML en template
  var template = HtmlService.createTemplateFromFile('form');
  //Evaluamos que el codigo HTML este bien
  var output = template.evaluate();
  return output;
}

//Creamos la funcion que nos permite conectar con los archivos 'css' y 'js'
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```
- **HTML** y uso de scriptleds como `<?!= include('css'); ?>` para conectar con los demas archivos.
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!--Utilizamos un scriptlet para conectar el CSS -->
    <?!= include('css'); ?>
  </head>
  <body>
  <p>Hola</p>
  <!--Utilizamos un scriptlet para conectar el JS -->  
  <?!= include('js'); ?> 
  </body>
</html>
```
- Modificamos un atributo basico de una etiqueta como `<p>` en css.
```html
<style>
  p{
    color:red;
  }
</style>
```
- Utilizamos la siguiente funcion para que se ejecute `console.log('Codigo del Lado del Cliente')` en el navegador al cargar la pagina.
```html
<script>
  window.addEventListener('load',function(){
    console.log('Codigo del Lado del Cliente')
  })
</script>
```
2. Implementar luego hacemos las implementaciones de prueba que necesitemos y finalmente `Gestionar Implementaciones` alli obtendremos un enlace y lo podemos cargar en google Sites. 

## Codigo Fuente
### Apps Script
```js
//Guardamos todo el libro en la variable SS
const SS = SpreadsheetApp.openById('16Qtw0raJ6a8_SV1eSQrIcn5F5gFVEcVgsWHPzym6GpQ');
//Guardamos la hoja Registros en una variable
const sheetRegistros = SS.getSheetByName('Registros');
//Guardamos la hoja BD en una variable
const sheetBd = SS.getSheetByName('BD');

//build-in function o funcion reservada
function doGet() {

  //Obtener datos de la hoja BD
  var datos = sheetBd.getDataRange().getValues();
  //Saltamos la fila de encabezado 
  datos.shift();

  //El arreglo almacena la opciones que se cargen en la pagina
  var talleres = [];
  //Cada fila es un arreglo de columnas y evaluamos la columna de lugares disponibles
  datos.forEach(row => {
    if(row[1] > 0){
      talleres.push(row[0]);
    }
  });

  //Guardamos el HTML en template
  var template = HtmlService.createTemplateFromFile('form');
  //Crear la propiedad talleres dentro del objeto template
  template.talleres= talleres;
  //Evaluamos que el codigo HTML este bien
  var output = template.evaluate();
  return output;
}

//Creamos la funcion que nos permite conectar con los archivos 'css' y 'js'
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
  
}

function addRecord(form){
  console.log(form.taller)
  //validar cupo

  var cupoDisponible = verCupoTaller(form.talleres);
  sheetRegistros.appendRow([new Date,form.talleres]);
  return 'Registro Recibido';
}

function verCupoTaller(chosen){
  var talleres = sheetBd.getDataRange().getValues();
  talleres.shift();
  talleres.map(taller => {
    if(taller[0] == chosen){
      if(taller[1]>0){
        return true;
      }
      else{
        throw 'Este taller ya no tiene cupo, recarga la pagina';
      }
    }
  });
}
```

### HTML

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!--Utilizamos un scriptlet para conectar el CSS -->
    <?!= include('css'); ?>
  </head>
  <body>
  <div id='output'>
    <!--Evento para validar respuesta -->
    <form id='talleres' onSubmit="validarRespuesta(this)">
      <h3>Seleccionar taller de tu preferencia</h3>

      <? talleres.forEach(taller => {?>
      <input type="radio" name="talleres" value="<?=taller?>"/> <?=taller?><br/>
      <?}); ?>

      <input type="submit" class="action" id="btnSubmit" class="action" value="Enviar"/>
    </form>
  </div>
  <!--Utilizamos un scriptlet para conectar el JS -->  
  <?!= include('js');?> 
  </body>
</html>
```

### CSS

```html
<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
<style>
  body{
    padding:20px;
  }

</style>
```

### JS

```html
<script>
  window.addEventListener('load',function(){
    console.log('Codigo del Lado del Cliente')
  })
  function validarRespuesta(form) {
    event.preventDefault();
    //Se ejecuta al enviar el formulario
    google.script.run
    //Se ejecuta al recibir un Ok del lado del servidor
    .withSuccessHandler( showSuccessMessage )
    //Cuando la respuesta del servidor es de rechazo
    .withFailureHandler(onFailure)
    //Cuando se presione el boton
    .addRecord( form );
  }

  function showSuccessMessage(answer,form){
    console.log(answer);
    var div = document.getElementById('output');
    div.innerHTML = '<br/><h4><font color="green">Tu respuesta ha sido Enviada</font></h4><br/><br/>'
  }
  function onFailure(error){
    console.log(error);
    var div = document.getElementById('output');
    div.innerHTML = '<br/><h4><font color="red">No se ha logrado envir tu Solicitud</font></h4><br/>'
  }

</script>
```
### Lista desplegable

- Codigo **Apps Script**
```js 
var ss = SpreadsheetApp.openById('16Qtw0raJ6a8_SV1eSQrIcn5F5gFVEcVgsWHPzym6GpQ')
var sheetLista = ss.getSheetByName('Lista')

function doGet() {


//ForEach para obtener valores de forma dinamica
  var data = sheetLista.getDataRange().getValues();
  data.shift();
  var lista = []
  data.forEach(list => {
    if(list != ""){
      lista.push(list[0]);
    }
  });

//Coneccion y validacion del HTML
  var template = HtmlService.createTemplateFromFile('lista');
  template.lista =  lista;
  var output = template.evaluate();

  return output;
}

//Creamos la funcion que nos permite conectar con los archivos 'css' y 'js'
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```
- Codigo **HTML**
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!--Utilizamos un scriptlet para conectar el CSS -->
    <?!= include('css'); ?>
  </head>
  <body>
  <form id="empleados">
    <div class="form-group">
      <label for="lista">Seleccion la opcion que prefieras</label>
      <!-- El atributo Id se usa en el frontend y el Name se usa en backend -->
      <select id="lista" name="lista">
        <?lista.forEach(list => {?>
           <option value="<?=list?>"><?=list?></option><br/>
        <?});?>   
      </select>
    </div>
    <br/>
    <input type="submit" class="blue" name="btnSummit" id="submit" value="Enviar"/>
  </form>
  <!--Utilizamos un scriptlet para conectar el JS -->  
  <?!= include('js');?> 
  </body>
</html>
```
### Tablas html dinámicas

- Apps Script
```js
var ss = SpreadsheetApp.openById('16Qtw0raJ6a8_SV1eSQrIcn5F5gFVEcVgsWHPzym6GpQ');
var sheetBd = ss.getSheetByName('BD');
//getDisplayValues(); obtiene los valores tal cual como se ven(con formato)
var data =  sheetBd.getDataRange().getDisplayValues();


function doGet() {
  console.log(data)
//Coneccion y validacion del HTML
  var template = HtmlService.createTemplateFromFile('page');
  template.dato = data;
  var output = template.evaluate();
  return output;
}

//Creamos la funcion que nos permite conectar con los archivos 'css' y 'js'
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```
- HTML
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!--Utilizamos un scriptlet para conectar el CSS -->
    <?!= include('css'); ?>
  </head>
  <body>
  <div id="output">
    <table>
      <? for(var i in data){?>
        <tr>
          <? for(var j in data[i]){
              if(i==0){?>
              <th><?=data[i][j]?></th>
              <?}
              else{?>
              <td><?= data[i][j]?></td>
              <?}?>
              <?}?>
           </tr>
       <?}?>
    </table>
  </div>
  <!--Utilizamos un scriptlet para conectar el JS -->  
  <?!= include('js');?> 
  </body>
</html>
```
- CSS
```html
<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
<style>
  body{
    padding:20px;
  }
  #output{
    width: 720px;
  }
  table{
    table-layout: fixed;
    width:100%;
  }
  th{
    background-color:#ebebeb;
    font-weight: bold;
    padding:8px;
  }
  td{
    padding:8px;
  }
</style>
```
### Calendar a Google Sheets

