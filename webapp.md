### [Apps Script Fundamentos](https://www.youtube.com/playlist?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV)    

By: Mozart Alberto García de Haro

- [Creando una WebApp de 0 a 100](#creando-una-webApp-de-0-a-100)
- [Lista desplegable](#lista-desplegable)
- [Tablas html dinámicas](#tablas-html-dinámicas)
- [Calendar a Google Sheets](#calendar-a-google-sheets)
- [Formulario de registro con Bootstrap](#formulario-de-registro-con-bootstrap)
- [Bootstrap, íconos, CSS y Apps Script](#bootstrap,-íconos-,-css-y-apps-script)

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
- Apps Script
```js
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Calendar');
  menu.addItem('Importar Eventos','ImportEvents')
  .addToUi();
}

function ImportEvents(){
//Pasos a seguir:
//1. Seleccionar la hoja de calculo
var ss = SpreadsheetApp.openById('16Qtw0raJ6a8_SV1eSQrIcn5F5gFVEcVgsWHPzym6GpQ');
var sheetEvents = ss.getSheetByName('Eventos');
//2. Acceder al calandario deseado

var calendars = CalendarApp.getCalendarsByName('danielnicolasmuner@gmail.com')
var calendar = calendars[0];
console.log('Se encontraron %s calendario(s) con este nombre: ', calendars.length, calendar.getName());
//3. Calendar App

var fechaInicio = new Date();
//Cuantos dias a partir de la fecha de inicio
var dias = 30;
//Obtenemos la fechaFinal a partir de la fechaInicio y mas los dias pero debemos manejar el tiempo en milisegundos
var fechaFinal = new Date(fechaInicio.getTime() + (dias*24*60*60*1000));
console.log(fechaInicio,fechaFinal)

var eventos = calendar.getEvents(fechaInicio,fechaFinal);
//4. Almacenar la informacion en un arreglo
var record = [];
eventos.forEach(event => {
  record.push([event.getTitle(),event.getDescription() || '', event.getStartTime() || '', event.getEndTime() || '']);
});

console.log(record)
//5. Importar la informacion a la hoja4
//getRange(startrow,startcol,#rowmaximo,#colsmax)
sheetEvents.getRange(2,1,record.length,record[0].length).setValues(record);
}

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
### Formulario de registro con Bootstrap

- Apps Script

```js
//La funcion doGet() se ejecuta automaticamente cada vez que se carga la pagina
function doGet() {
//Coneccion y validacion del HTML
  var template = HtmlService.createTemplateFromFile('registro');
  template.pubUrl = 'https://script.google.com/macros/s/AKfycbwCGzDVip5oYcEwiLcXnh4g2YnUNoiBfB98b0kZXYYH/dev';
  var output = template.evaluate();
  return output;
}

//Creamos la funcion que nos permite conectar con los archivos 'css' y 'js'
//La funcion include es llamada desde los archivos CSS y JS
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
} 
//Se ejecuta al realizar el http request
function doPost(e){
//El post es el que envia la informacion a SpreadSheet
//El documento
  var ss = SpreadsheetApp.openById('16Qtw0raJ6a8_SV1eSQrIcn5F5gFVEcVgsWHPzym6GpQ')
  var sheetLogs =  ss.getSheetByName('logs')

//Obtenermos todas la variables del formulario
  var id = new Date();
  var email = e.parameter.email;
  var gradoAcademico = e.parameter.gradoAcademico;
  var pais = e.parameter.pais;
  var fechaNacimiento = e.parameter.fechaNacimiento;
  var nombreCompleto = e.parameter.nombreCompleto;
  var password = e.parameter.password;
  var acuerdoPrivacidad = (e.parameter.acuerdoPrivacidad == 'on')? 'Aceptado' : 'Rechazado';

//Creamos un vector con los datos y los unimos la hoja indicada
  sheetLogs.appendRow([id,nombreCompleto,email,password,fechaNacimiento,pais,gradoAcademico,acuerdoPrivacidad]);

//Luego de haber ejecutado el registro le indicamos al usuario que el registro se completo.
  return HtmlService.createTemplateFromFile('registroterminado').evaluate();
}
```
- Codigo HTML
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!--Utilizamos un scriptlet para conectar el CSS -->
    <?!= include('css'); ?>
  </head>
  <body>
  <!--container le da estilo a todo el body-->
  <div class="container">
    <h2>Registro</h2>
  <!-- method="post" permite hacer un http request-->
    <form novalidate id="formularioRegistro" name="formularioRegistro" class="needs-validation" method="post" action="<?=pubUrl?>">
      <div class="mb-3">
        <label for="nombreCompleto" class="form-label">Nombre Completo*</label>
        <input type="text" name="nombreCompleto" id="nombreCompleto" class="form-control" placeholder="Nombre Completo" required/>
      </div>
      <div class="row">
        <div class="col">
          <div class="mb-3">
            <label for="email" class="form-label">Correo Electronico*</label>
            <input type="email" class="form-control" name="email" id="email" placeholder="example@domain.com" required/>
            <div class="invalid-feedback">
              Ingresa una direccion de correo Electronico valida
            </div>    
          </div>
        </div>

        <div class="col">  
          <div class="mb-3">
            <label for="password" class="form-label">Contraseña*</label>
            <input type="password" class="form-control" name="password" id="password" placeholder="*********" required/>
          </div>
        </div>
      </div> 

      <div class="row">
        <div class="col">
          <div class="mb-3">
            <label for="fechaNacimiento" class="form-label">Fecha Nacimiento</label>
            <input type="date" id="fechaNacimiento" class="form-control" name="fechaNacimiento" placeholder="dd/mm/aaaa"/>
          </div>
        </div>

        <div class="col">
          <div class="mb-3">
            <label for="pais" class="form-label">Pais</label>
            <select class="form-select" id="pais" name="pais">
              <option value="Argentina">Argentina</option>
              <option value="Colombia">Colombia</option>
              <option value="Ecuador">Ecuador</option>
              <option value="Peru">Peru</option>
              <option value="Mexico">Mexico</option>
              <option value="Brazil">Brazil</option>
            </select>
          </div>
        </div>

        <div class="col">
          <div class="mb-3">
            <label for="gradoAcademico" class="form-label">Grado Academico</label>
            <select class="form-select" id="gradoAcademico" name="gradoAcademico">
              <option value="Primaria">Primaria</option>
              <option value="Secundaria">Secundaria</option>
              <option value="Tecnico">Tecnico</option>
              <option value="Profesional">Profesional</option>
              <option value="Maestria">Maestria</option>
              <option value="Phd">Phd</option>
            </select>
          </div>
        </div>  
      </div>   
      <div class="mb-3">
        <div class="form-check">
          <input type="checkbox" id="acuerdoPrivacidad" name="acuerdoPrivacidad" class="check-form-input" required/>
          <label for="acuerdoPrivacidad" class="form-check-label">Estoy de acuerdo con la clausula de Privacidad</label>
          <div class="invalid-feedback">
            You must agree before submitting.
          </div>
        </div> 
      </div>    
      <div class="col-12">
        <button class="btn btn-primary" type="submit">Submit form</button>
      </div>  
    </form>
    <br/>
  </div>
  <!--Utilizamos un scriptlet para conectar el JS -->  
  <?!=include('js');?> 
  </body>
</html>
```
- Codigo CSS Botstrap
```html
<!-- <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css"> -->

<!-- Bootstrap Include via CDN CSS-->
<!-- CSS only -->
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-0evHe/X+R7YkIZDRvuzKMRqM+OrBnVFBL6DOitfPri4tjfHxaWutUpFmBp4vmVor" crossorigin="anonymous">

<style>

</style>
```
- Codigo JS Botstrap
```html
<!--Bootstrap - Include via CDN JS --->
<!-- JavaScript Bundle with Popper -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/js/bootstrap.bundle.min.js" integrity="sha384-pprn3073KE6tl6bjs2QrFaJGz5/SUsLqktiwsUTF55Jfv3qYSDhgCecCxMW52nD2" crossorigin="anonymous"></script>
<script>
(() => {
  'use strict'

  // Fetch all the forms we want to apply custom Bootstrap validation styles to
  const forms = document.querySelectorAll('.needs-validation')

  // Loop over them and prevent submission
  Array.from(forms).forEach(form => {
    form.addEventListener('submit', event => {
      if (!form.checkValidity()) {
        event.preventDefault()
        event.stopPropagation()
      }

      form.classList.add('was-validated')
    }, false)
  })
})()
</script>
```
- Codigo HTML confirmacion
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  <?!= include('css');?>
  </head>
  <body>
  <div class="container">
    <h2>Registro</h2>
    <br/>
    <div id='output2'>
      <p>Se ha enviado un correo a tu cuenta</p>
      <p>Sigue las instrucciones enviadas</p>
    </div>
  </div>
  <?!= include('js');?>  
  </body>
</html>
```

### Bootstrap, íconos, CSS y Apps Script
[Design](https://www.youtube.com/watch?v=mGK1LE8Q4AQ&list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV&index=21)
