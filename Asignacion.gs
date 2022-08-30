/**
 * Asignación
 * Asigna el amigo secreto basado en la lista de participantes
 * 
 * @param {void} - void
 * @return {array} - Arreglo con los datos generados 
 */
function Asignacion() {

  // Obtiene el id de la hoja de cálculo activa
  var Libro = SpreadsheetApp.getActiveSpreadsheet();
  var HojaAmigoSecreto = Libro.getSheetByName( WSDATANAME );
  Libro.toast( 'Asignando Personajes', 'Working...' );
  
  // Carga los datos de la hoja en un arreglo de objetos
  var RangoDatos = HojaAmigoSecreto.getDataRange().getValues(); 
  var numCandidatos = RangoDatos.length;
 
  // Recorrido de arreglo para asignar al amigo secreto de cada persona 
  // Verificación de que haya más de 1 persona para jugar
  if ( numCandidatos <= 2 ) {
    Browser.msgBox( "El juego no se puede jugar con " + ( numCandidatos-1 ).toString() + " Jugador <br>" );
    return;
  };

  // Verificación de que no haya nombres repetidos de personajes
  for ( var indice=1; indice < numCandidatos-1; indice++ )
    for( var indice2=indice+1; indice2< numCandidatos; indice2++ )
      if ( RangoDatos[indice][4] == RangoDatos[indice2][4] ){
        Browser.msgBox( "Nombre de personaje repetido: "+ RangoDatos[indice][1].toString()+ " y " + RangoDatos[indice2][1].toString() );
        return;
      };
 
  // Se crean los vectores Mujeres Y Hombres donde se guardan los nombres de los personajes asignados
  var Hombres=[];
  var Mujeres=[];
  // Recorrido para guardar los personajes en el arreglo asociado a el genero
  for ( var indice=1; indice < numCandidatos; indice++ ) {
    // Adiciona el personaje al arreglo según el género
    if ( RangoDatos[indice][3] == "Hombre" ) Hombres.push(RangoDatos[indice][4]);
     else Mujeres.push( RangoDatos[indice][4] );
  };//for
  
  // Acción de verificación para asignar parejas
  for ( var indice=1; indice < numCandidatos; indice++ ) {
  // Si la primera persona encontrada es hombre
  if ( RangoDatos[indice][3] == "Hombre" ) {
    // Verifica si hay mujeres, y si hay, le asigna una al azar
    if ( Mujeres.length != 0 ) RangoDatos[indice][5] = Mujeres.splice((Math.random()*10) % Mujeres.length, 1);        
    // Sino hay mujeres, verifica si hay más hombres y si hay le asigna uno al azar
    else {
     if ( Hombres.length != 0 ) RangoDatos[indice][5] = Hombres.splice((Math.random()*10) % Hombres.length, 1);
      // sino hay mas personas lo deja como No asignado
      else RangoDatos[indice][5] = "NO ASIGNADO";
    }
  }
  // Si la primera persona resulta ser una mujer
  else {
    // verifica si hay hombres, y si hay, le asigna uno al azar
    if ( Hombres.length != 0 ) RangoDatos[indice][5] = Hombres.splice((Math.random()*10) % Hombres.length, 1);        
    // sino hay hombres, verifica si hay más mujeres y si hay le asigna una al azar
    else {
     if ( Mujeres.length != 0 ) RangoDatos[indice][5] = Mujeres.splice((Math.random()*10) % Mujeres.length, 1);
       // sino hay más personas la deja como No asignado
      else RangoDatos[indice][5] = "NO ASIGNADO";
    };
   };
  };//for
  //

  var cont = 0;
  for (var indice=1; indice < numCandidatos; indice++) {
    if ( RangoDatos[indice][4] == RangoDatos[indice][5] ) {
      RangoDatos[indice][6] = cont++;
      return Asignacion();
    };
  };
 
  // Refleja las asignaciones en la hoja de cálculo
  HojaAmigoSecreto.getRange( "A1:G" + HojaAmigoSecreto.getLastRow() ).setValues( RangoDatos );
  return RangoDatos;
};
