/**
 * Amigo Secreto
 * Asignación de personaje para jugar al Amigo Secreto
 * Elaborado por: jorge.e.forero@gmail.com
 */

/*
Instalación

Los correos que envía la solución incorporan 2 imágenes que deben estar almacenadas en Google Drive.  Estas imagenes las encontrará en el Drive compartido o en el repositorio de Github. Ubique las imagenes en su Drive y obtenga el id de cada imagen. 

Para obtener el id, complete los siguientes pasos:
1- En Drive, sobre la imagen de botón derecho y escoja la opción 'Get link'
2- El acceso general debe estar seleccionado con 'Anyone with the link'
3- Haga clic en el boton 'Copy link'
4- La url que se genera es algo como: 
https://drive.google.com/file/d/0BwA54cY-n_xFd1JNT25zdWpXdDA/view?usp=sharing&resourcekey=0-K_eGX39agewliqM-gPo21g
5- El id de la imagen es la cadena que esta entre d/ y /view es decir que para este ejemplo el id es 0BwA54cY-n_xFd1JNT25zdWpXdDA
6- Si este es el id para la imagen de mujeres, reemplace el id obtenido así

// Id de la imagen para mujer almacenada en Drive
const IMGWOMID = '0BwA54cY-n_xFd1JNT25zdWpXdDA';   // <= id de la imagen obtenida 

Estos IDs deben ser actualizados en las constantes globales del archivo Code en el editor de G.A.S. 
*/

/**
 * Constantes Globales
 */

// Id de la imagen para hombre almacenada en Drive
const IMGMANID = 'XXXXX';   //<= Actualice el Id
// Id de la imagen para mujer almacenada en Drive
const IMGWOMID = 'XXXXX';   //<= Actualice el Id

// Nombre de la hoja de datos
const WSDATANAME = 'PARTICIPANTES';
// Subject para el correo
const SUBJECT = 'YA tienes tu amigo secreto';

/**
 * Interface
 */

/**
 * EXEC_Asignacion
 * Llamado a la función de envío de Correo
 */
function EXEC_asignacion() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert( 'Se asignarán los personajes. Está seguro?', ui.ButtonSet.YES_NO );
  if ( response == ui.Button.YES ) {
    Asignacion();
    Browser.msgBox( 'Los personajes fueron asignados' );
  };
}

/**
 * EXEC_envio
 * Llamado a la función de envío de Correo
 */
function EXEC_envio() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert( 'Se enviarán los correos. Está seguro?', ui.ButtonSet.YES_NO );
  if ( response == ui.Button.YES ) { 
    let envios = sendEmails();
    Browser.msgBox( 'Se enviaron ' + envios + ' correos.' );
  }
}
