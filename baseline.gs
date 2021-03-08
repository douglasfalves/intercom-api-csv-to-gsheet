var spreadsheet = SpreadsheetApp.getActive();
var resultSheet = spreadsheet.getSheetByName("CSV");
var apiUrl = "https://api.intercom.io/export/messages/data";
//var basicAuth = "lembre-se de trazer o token do Intercom nesta variável";
var dataCached = [];

function cacheData(data) {
  //Logger.log(data);
  this.dataCached = data;
}

function main() {
  initSpreadsheet();
  deleteIds();
  deleteTriggers();
  params = getParamsRange();
  makePost(params); // finaliza criando um gatilho checando a URL a partir do ID do job no Intercom
}

function afterGet(dUrl){ //fluxo padrão pós URL obtida
  deleteTriggers();  
  initSpreadsheet();
  cacheData(retrieveIntercomZipOnly(dUrl));
  fillDataOnly();
  waitSetDates();
}

function onOpen() {
 var menuItems = [
   {name: 'Começar', functionName: 'main'}
   // fazer pedido de CSV
   ,{name: 'Formatar os dados', functionName: 'limpaCSV'}
   ,{name: '(Sem URL) Download', functionName: 'getUrlButton'} // Pegar a URL pelo botão caso o getUrl padrão tenha falhado
   ,{name: '(Com URL) Download', functionName: 'retrieveIntercomZip'} // Faz o download quando já se tem a URL
 ];
 this.spreadsheet.addMenu('Importar CSV', menuItems);
}

function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "Atenção:"); 
}

function getParamsRange() {
  var paramsSheet = spreadsheet.getSheetByName("params");
  paramsSheet1 = paramsSheet.getRange(3, 1, 1, 1).getDisplayValues();
  paramsSheet2 = paramsSheet.getRange(3, 2, 1, 1).getDisplayValues();
  paramsSheetBody = '{ "created_at_after":' + paramsSheet1 + ',"created_at_before":'+ paramsSheet2 + '}';
  console.log(paramsSheetBody);
  return paramsSheetBody;
} 

function makePost(params) {
  displayToastAlert('Começando os trabalhos...');

  // params = getParamsRange(); //está aqui para teste unitário
  var options =
   {
     "method" : "post",
     "headers": {
       "Accept": "application/json",
       "Content-Type": "application/json",
       "Authorization" : this.basicAuth,
     },
     "payload" : JSON.stringify(JSON.parse(params))
   };

  var response = UrlFetchApp.fetch(this.apiUrl, options);
  var json = JSON.parse(response.getContentText());
  console.log(json);
  var job = json.job_identifier;

  wait(); // aqui gera uma segunda execução a partir do gatilho
  escreveJobId(job);
  console.log(job);

  return job;
}

function wait(){
  ScriptApp.newTrigger('getUrl')
    .timeBased()
    .everyMinutes(1)
    .inTimezone('America/Sao_Paulo')
    .create();
}

function waitSetDates(){
  ScriptApp.newTrigger('setFirstLastRow')
    .timeBased()
    .after(60000)
    .inTimezone('America/Sao_Paulo')
    .create();
}

function waitLimpaCSV(){
  ScriptApp.newTrigger('limpaCSV')
    .timeBased()
    .after(60000)
    .inTimezone('America/Sao_Paulo')
    .create();
}


function escreveJobId(job){
  var paramsSheet = spreadsheet.getSheetByName("params");
  var rowJob = paramsSheet.getRange(4, 2, 1, 1);
  rowJob.setValue(job);
  console.log('Job ID escrito');
  displayToastAlert('JobID: ' + job);
}

function escreveURLparams(dUrl){
  var paramsSheet = spreadsheet.getSheetByName("params");
  var rowJob = paramsSheet.getRange(5, 2, 1, 1);
  rowJob.setValue(dUrl);
  console.log('URL escrita na tela principal');
  displayToastAlert('Escrevendo a URL aqui para saber se está tudo certo: ' + dUrl);
}

function deleteIds(){
  var paramsSheet = spreadsheet.getSheetByName("params");
  var rowJob1 = paramsSheet.getRange(4, 2, 1, 1);
  rowJob1.setValue('');
  var rowJob2 = paramsSheet.getRange(5, 2, 1, 1);
  rowJob2.setValue('');
  console.log('ID e URL apagados');
  displayToastAlert('JobID e URL reiniciados');
}

function deleteTriggers(){ // Apaga todos os gatilhos criados
  console.log('Existe(m) ' + ScriptApp.getProjectTriggers().length + ' gatilho(s) gatilho(s) programado(s).');
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
    Logger.log('Trigger '+ i +' apagado');
  }
}

function getUrl(){

  var paramsSheet = spreadsheet.getSheetByName("params");
  var jbid = paramsSheet.getRange(4, 2, 1, 1).getDisplayValue(); // pega o JobId da planilha

  displayToastAlert('Peguei o JobID: ' + jbid);

  var options =
  {
    "method" : "get",
    "headers": {
      "Accept": "application/octet-stream",
      "Accept": "application/json",
      "Content-Type": "application/json",
      "Authorization" : this.basicAuth,
    },
  };

  var response = UrlFetchApp.fetch(this.apiUrl + '/' + jbid, options);

  var json = JSON.parse(response.getContentText());
  console.log(json.download_url);
  dUrl = json.download_url;

  if(dUrl != ''){ // se tiver a url de download, limpa 
    displayToastAlert('Checando URL: ' + dUrl);
    escreveURLparams(dUrl);
    afterGet(dUrl);      
  }
  else{ // se não tiver, vai gerar mais um gatilho até dUrl for preenchido
    wait();
  }

  console.log(dUrl);
  return dUrl;
}

function initSpreadsheet() {
  this.resultSheet.getRange('A:W').clear();
  
  //resultSheet.deleteRows(2,resultSheet.getLastRow());
}

function limpaCSV(){
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName("CSV");
  var range = sheet.getRange('P:P');
  var data = range.getValues();
  
  Logger.log(range);

  for (var i = 0 ; i<data.length ; i++) {
    data[i][0] = data[i][0].toString().replace(' -0300', '');
    //range[i][0] = range[i][0].toString().substring(0,19);
    //Logger.log(range);
  }
  range.setValues(data);

  apagaColunas();
  ocultaColunas();

  displayToastAlert('Tudo limpo!');
  msg = ui.alert("Finalizado!");
}

function retrieveIntercomZipOnly(dUrl){
  displayToastAlert('Fazendo Download...');

  var options =
  {
    "method" : "get",
    "headers": {
      "Accept": "application/octet-stream",
      "Authorization" : this.basicAuth,
    },
  };

  var csv = Utilities.parseCsv(UrlFetchApp.fetch(dUrl, options));
  //var csv = Utilities.parseCsv(UrlFetchApp.fetch(urlTeste, options)); // teste unitário

  displayToastAlert('Consegui pegar o CSV');

  this.dataCached = csv
  return csv;
}

function fillDataOnly(){
  var sheet = spreadsheet.getSheetByName("CSV");
  csv = this.dataCached;

  sheet.getRange(1, 1, csv.length, csv[0].length).setValues(csv);

  

  // MailApp.sendEmail(userMail,
  //                   'A importação na planilha do Intercom acabou!',
  //                   'Dá uma olhada lá:' + idSheets);

  // deleteIds(); // limpa os ids na params para um trabalho limpo
}




// EXTRA | Para o uso nos botões de correção


function getUrlButton(){ // CHECAR

  var paramsSheet = spreadsheet.getSheetByName("params");
  var jbid = paramsSheet.getRange(4, 2, 1, 1).getDisplayValue(); // pega o JobId da planilha

  displayToastAlert('Tentando novamente com o JobID: ' + jbid);

  var options =
  {
    "method" : "get",
    "headers": {
      "Accept": "application/octet-stream",
      "Accept": "application/json",
      "Content-Type": "application/json",
      "Authorization" : this.basicAuth,
    },
  };

  var response = UrlFetchApp.fetch(this.apiUrl + '/' + jbid, options);
  
  var json = JSON.parse(response.getContentText());
  console.log(json.download_url);
  dUrl = json.download_url;

  if(dUrl != ''){ // se tiver a url de download, faz o mundo 
    displayToastAlert('Peguei a URL: ' + dUrl);
    initSpreadsheet();
    deleteTriggers();
    escreveURLparams(dUrl);
    retrieveIntercomZip(dUrl);
    console.log(dUrl);
  }
  else{ // se não tiver, vai gerar mais um gatilho até dUrl for preenchido
    displayToastAlert('Tente novamente em 1 minuto');
    return null;
  }

  return dUrl;
}

function pegaUrlPlanilha(){
  var paramsSheet = spreadsheet.getSheetByName("params");
  var sUrl = paramsSheet.getRange(5, 2, 1, 1).getDisplayValue(); // pega a URL da planilha
  return sUrl;
}

function retrieveIntercomZip(dUrl){ // CHECAR
  
  if(dUrl == null){
    dUrl = pegaUrlPlanilha();
  }
  
  displayToastAlert('Fazendo Download...');

  var options =
  {
    "method" : "get",
    "headers": {
      "Accept": "application/octet-stream",
      "Authorization" : this.basicAuth,
    },
  };

  var csv = Utilities.parseCsv(UrlFetchApp.fetch(dUrl, options)); // Acredito que para cachear, deveria parar aqui e retornar a var csv
  //var csv = Utilities.parseCsv(UrlFetchApp.fetch(urlTeste, options)); // teste unitário

  displayToastAlert('Consegui pegar o CSV');

  initSpreadsheet();

  var sheet = spreadsheet.getSheetByName("CSV");
 sheet.getRange(1, 1, csv.length, csv[0].length).setValues(csv);


  // MailApp.sendEmail(userMail,
  //                   'A importação na planilha do Intercom acabou!',
  //                   'Dá uma olhada lá:' + idSheets);

  // deleteIds(); // limpa os ids na params
  deleteTriggers(); // limpa gatilhos

  return sheet.getName();
}

function setFirstLastRow(){

var painel = SpreadsheetApp.getActive().getSheetByName("Painel");
var lista = SpreadsheetApp.getActive().getSheetByName("listas").getRange('A:A').getDisplayValues();
  
  var first = '';
  var last = '';

  for(var i=0;i<lista.length; i++){
    if(i==0){first = lista[i];}
    if(lista[i] != '' ){last = lista[i];}
  }

  console.log(first);
  console.log(last);

  painel.getRange('A4').setValue(first);
  painel.getRange('A6').setValue(last);

}




