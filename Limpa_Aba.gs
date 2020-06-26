/*************************************************************************************/
function limpa_dados() 
{
  //Get the SpreadSheet Active
  //Obtém a planilha ativa
  let planilha = SpreadsheetApp.getActive();
  
  //Get the Sheet Active
  //Obtém a aba ativa
  let aba = planilha.getActiveSheet();
  
  //Get values from a range
  //Obtém valores de um intervalo específicado por parametro
  let dados = aba.getRange('A2:H');
  
  //Clean shett on defined range 
  //Limpa aba no intervalo definido anteriormente
  dados.clear();
} 
