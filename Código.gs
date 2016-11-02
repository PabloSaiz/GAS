// Esta es la función encargada de mandar el correo correspondiente segun el idioma seleccionado
function enviarEmail(){
  // Guardamos en la variable sheet la pagina sobre la que estamos trabajando
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  // Guardamos la ultima fila introducida para poder coger los datos de esta utlima linea
  var row=sheet.getLastRow();
  // Cogemos el valor del idioma mediante lo siguiente: E + la fila ultima introducida.
  var idioma=sheet.getRange("E"+row).getValue(); 
  // Lo mismo para el nombre y para el nivel de conocimientos
  var nombre=sheet.getRange("B"+row).getValue();
  var nivel=sheet.getRange("D"+row).getValue();
  // Del mismo modo cogemos el email al que le vamos a enviar el correo.
  var email=sheet.getRange("C"+row).getValue();
  
  // Ahora según el idioma seleccionado enviaremos uno u otro correo, primero evaluamos si es Ingles, en caso de serlo prepara el mensaje en Ingles.
  if (idioma=="Ingles"){
    var subject=nombre+" welcome to the course of GAS";
    var body="Now these underwritten and will receive soon tutorials in the language you've selected.\n"+
      "The knowledge level indicated on the form has been "+nivel+"\n\n"+
      "Lesson 1\n"+
      "https://www.youtube.com/watch?v=bDqajllxvGQ";
    // Si el idioma es Español mandara el mensaje en español
  }else if (idioma=="Español"){
    var subject=nombre+" bienvenido al curso de GAS";
    var body="Ahora estas subscrito y recibiras proximamente tutoriales en el idioma que has selecionado.\n"+
      "El nivel de conocimientos indicados en el formulario ha sido "+nivel+"\n\n"+
      "Leccion 1\n"+
      "https://www.youtube.com/watch?v=Wai_P69BCpc";
    // Si no es ni ingles ni español solo queda que sea italiano asi que preparamos el mensaje en italiano
  }else{
    var subject=nombre+" benvenuti al GAS corso";
    var body="Ora, questi sottoscritto e riceverà presto tutorial nella lingua che hai selezionato.\n"+
      "Il livello di conoscenza indicato sul modulo è stato "+nivel+"\n\n"+
      "lezione 1\n"+
      "https://www.youtube.com/watch?v=8p3vtjvpW1A";
  }
  // Llamamos a la funcion que mandara el correo dandole el email al que lo tiene que enviar junto con el asunto y el contenido del email:
      GmailApp.sendEmail(email, subject, body);
}
