# [Apps Script](https://www.youtube.com/playlist?list=PLG1qdjD__qH4dyXq4sM03Rf0RFhB_4tbm)
Powered by: Aula en la nube

## Indice
- [Crear y abrir documentos en Google DOCS](#crear-y-abrir-documentos-en-google-docs)
- [Condiciones](#condiciones)
- [Bucle](#bucle)

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


