### [Apps Script Fundamentos](https://www.youtube.com/playlist?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV)    

By: Mozart Alberto García de Haro

- [Creando una WebApp de 0 a 100](#creando-una-webApp-de-0-a-100)


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
2. Implementar $\to$ selecionar tipo $\to$ aplicacion web $\to$ implementar. Finalemente abrimos `Implementaciones de Prueba`.
