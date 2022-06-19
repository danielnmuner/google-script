# [Apps Script](https://www.youtube.com/playlist?list=PLG1qdjD__qH4dyXq4sM03Rf0RFhB_4tbm)
Powered by: Aula en la nube

## Indice
- [Crear y abrir documentos en Google DOCS](#crear-y-abrir-documentos-en-google-docs)
- [Condiciones](#condiciones)
- [Bucle](#bucle)
- [Bucle FOR IN y análisis de atributos](#bucle-for-in-y-análisis-de-atributos)
- [Crear funciones con parámetros y menú personalizado](#crear-funciones-con-parámetros-y-menú-personalizado)
- [Crear formulario HTML en Google Docs](#crear-formulario-html-en-google-docs)   
- 
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
