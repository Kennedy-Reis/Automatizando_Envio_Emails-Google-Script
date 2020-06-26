//Create custom menu
//Cria menu personalizado
function onOpen() 
{
  
  var ui = SpreadsheetApp.getUi()                                             //Retorna a interface da planilha //Return spreadsheet interface
                                .createMenu('Menu_Personalizado')             //Cria um menu  //Create a menu
                                .addItem('Cria Certificados', 'menuItem1')    //Adiciona um item (funcão menuItem1) ao menu // Adding item (menuItem1) to menu
                                .addItem('Limpa Dados', 'menuItem2')           //Adiciona um item (funcão menuItem2) ao menu // Adding item (menuItem2) to menu                               
                                .addToUi();                                   //Adiciona menu na interface // Add menu to interface
}

function menuItem1() 
{
  //Execute function 
  //Executa a função
  criaCertificados();
  //Create alert when menu button is pressed 
  //Cria alerta quando for clicado
  SpreadsheetApp.getUi().alert('Deu bom!!!');  
}

function menuItem2() 
{
  //Execute function
  //Executa a função
  limpa_dados();
  //Create alert when button is pressed 
  //Cria alerta quando for clicado
  SpreadsheetApp.getUi().alert('Sucesso!!');  
}
