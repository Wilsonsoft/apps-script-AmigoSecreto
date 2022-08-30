/**
 * sendEmails
 * Envia los correos con la información de los amigos secretos de los participantes. Los correos son enviados
 * si la columna de envio esta vacia
 */
function sendEmails() {

  let msg = '';
  // Numero de envíos
  let envios = 0;
  // Marcador de correo enviado
  let sent = [];
  // Obtiene los datos para el envío
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( WSDATANAME );
  if ( sheet !== null ) {
    // Datos de los participantes
    let box = sheet.getDataRange().getValues();
    // Imagen para mujer
    let womImgBlob = getImageBlob( IMGWOMID );
    // Imagen para hombre
    let manImgBlob = getImageBlob( IMGMANID );
    // Template HTML del cuerpo del correo
    let tpl_email = HtmlService.createHtmlOutputFromFile( 'tpl_email.html' ).getContent();
    for ( let indx = 1; indx<box.length-1; indx++ ) {
      // Obtiene los datos de la persona en proceso como un objeto
      let persona = getObjectFromArray( box[ indx ], box[ 0 ] );
      let mensaje = tpl_email;
      // Si no se le ha realizado envio a la persona y tiene correo, 
      if ( persona.enviado == '' ) {
        // Determina la imagen a incluir en el correo de acuerdo al genero registrado
        let imgSecretBlob = ( persona.genero.toUpperCase() == 'MUJER' ) ? womImgBlob : manImgBlob;
        let attach = { imgsecret: imgSecretBlob };
        // Reemplazables con los datos de la persona que participa
        mensaje = mensaje.replace( '##NOMBRE##', persona.nombre )
                        .replace( '##PERSONAJE##', persona.personaje )
                        .replace( '##AMIGO##', persona.amigo_secreto );
        // Envio del correo
        GmailApp.sendEmail( persona.email,
                            SUBJECT,
                            '',
                            {
                              htmlBody: mensaje,
                              inlineImages: attach
                            });
        envios++;
        sent.push( [ 'SI' ] );
      } else {
        sent.push( [ persona.enviado ] );
      }
    };//for

  } else {
    msg = `No se encontró la hoja de datos ${WSDATANAME}`;
  };
  // Actualiza la columna de control de envíos
  if ( envios > 0 ) {
    sheet.getRange( 2, 7, sent.length, 1 ).setValues( sent );
  };
  msg += `\nEnviados ${envios}`;
  return msg;

};

/**
 * getImageBlob
 * Dado el Id de la imagen en Drive, retorna el blob de la imagen
 */
function getImageBlob( FileId ) {
  return DriveApp.getFileById( FileId ).getBlob();
};

/**
 * getObjectFromArray
 * Dato un arreglo de datos y uno de encabezados, retorna un objeto que representa los datos datos
 */
function getObjectFromArray( Row, Header ) {
  var obj = {};
  for ( var indx=0; indx<Row.length; indx++ ) {
    obj[ Header[ indx ].toLowerCase().replace( /\s/g, '_' ) ] = Row[ indx ];
  };//for
  return obj;
}; 
