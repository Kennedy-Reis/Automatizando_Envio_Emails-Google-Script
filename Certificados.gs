function criaCertificados() 
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
  
  //Get value from a specific cell for the loop
  //Obtém o valor de uma célula específica
  let quantidadeDocs = aba.getRange('G2').getCell(1,1).getValue();
  
  //Get hours:minutes:seconds
  //Cria data atual (exemplo do formato padrão (Thu Jun 25 2020 16:14:32 GMT-0000 (UTC))
  let data = new Date();
  
  //Get today's date value 
  //Obtém o valor do dia atual
  let dia = data.getDate();
  
  //Get actual month value (Values start with January being "0", February "1", and follows..)
  //Obtém o valor do mês atual (Por padrão, "0 é o mês de Janeiro", "1 é o mês de Fevereiro",...)
  
  //Adding 1 
  //Soma-se + 1 para que "1 seja o mês de Janeiro", "2 seja o mês de Fevereiro",...
  let mes = data.getMonth() + 1;
  
  //Get actual year value
  //Obtém o valor do ano atual
  let ano = data.getFullYear();
  
  //Get actual hour value 
  //Obtém o valor da hora atual
  let horas = data.getHours();
  
  //Get actual minute value
  //Obtém o valor do minuto atual
  let minutos = data.getMinutes();
  
  //Get actual seconds value
  //Obtém o valor do segundo atual
  let segundos = data.getSeconds();
  
  /*----------------------------------Search template file------------------------------------------------------------*/
  
  //Command that searchs all folders on Drive with giving name as a parameter "Certificados_PDF"
  //Pesquisa no Drive todoas as pastas que tenham o nome passado por parametro
  let folders = DriveApp.getFoldersByName('TutorialGS');
  
  //Needs an iterator to navigate on list of folders
  //Necessita de um iterador para navegar na lista de pastas
  
  //In this case, we're not using iterators (loop) once we have only one folder with giving name.
  //Como possui apenas uma pasta com esse nome então não usaremos laços de iterações (loop)
  let folder = folders.next();
  
  //Creating once a new folder to upload certificates being that will be created 
  //Cria apenas vez uma nova pasta para guardar os certificados criados
  
  //The folder is created by a giving name as parameter "Certificados" + day + month + year
  //A pasta é criada com nome passado por parametro
  var newFolder = folder.createFolder("Certificados " + dia + "." + mes + "." + ano);
  
  //Search archive with giving name ("Template_Certificado") on created folder
  //Pesquisa dentro da pasta encontrada o arquivo com o nome passado por parametro
  let docs = folder.getFilesByName('Template_Certificado');
  
  //Necessita de um iterador para navegar na lista de arquivos
  //Needs an iterator to navigate on list of archives
  
  //As we don't have more than one archive with giving name, we'll not use iterators (loop) 
  //Como possui apenas um arquivo com esse nome então não usaremos laços de iterações (loop)
  let doc = docs.next();
  
  
  /*----------------------------------Loop ---------------------------------------------------------------------------*/
 
  //Loop needed to navigate on each line of active sheet of spreadsheet
  //Realiza um loop para percorrer cada linha da planilha-aba ativa
  for(var i = 1; i <= quantidadeDocs; i++)
  {
    //Get value from column "CRIOU DOC"
    //Obtém o valor da célula abaixo da coluna "CRIOU DOC"
    
    //If cell is empty (== ""), certifcate is not created
    //Caso o valor esteja vazio então não é criado o certificado
    if(dados.getCell(i,8).getValue() == '')
    {     
      //Get e-mail value
      //Obtém o valor do email
      let email = dados.getCell(i,1).getValue();
      
      //Get name value
      //Obtém o nome completo
      let nome = dados.getCell(i,2).getValue();
      
      //Get course name value
      //Obtém o nome do curso
      let nomecurso = dados.getCell(i,3).getValue();
      
      //Get starting course date value
      //Obtém a data de inicio do curso
      let dataInicio = dados.getCell(i,4).getValue();
      
      //Get ending course date value
      //Obtém a data de fim do curso
      let dataFim = dados.getCell(i,5).getValue();
      
      //Get course workload value
      //Obtém a carga horária do curso
      let cargaHoraria = dados.getCell(i,6).getValue();  
           
      //Make a copy of the template file
      //Realiza uma cópia do arquivo template
      let idDocumento = doc.makeCopy().getId();
      
      //Rename the copied file
      //Renomeia o arquivo cópia
      DriveApp.getFileById(idDocumento).setName(nome + " " + horas+":"+minutos+":"+segundos);
      
      //Get the document body as a variable
      //Obtém todas as variáveis do arquivo cópia
      let dadosDocumento = DocumentApp.openById(idDocumento).getBody();
      
      //Insert the entries into the document
      //Substitui ##Nome## pelo valor da variavel nome
      dadosDocumento.replaceText('##Nome##', nome);
      //Replace ##NomeCurso## by variable nomecurso
      //Substitui ##NomeCurso## pelo valor da variavel nomecurso
      dadosDocumento.replaceText('##NomeCurso##', nomecurso);
      //Replace ##DataInicio## by variable dataInicio
      //Substitui ##DataInicio## pelo valor da variavel dataInicio
      dadosDocumento.replaceText('##DataInicio##', dataInicio);
      //Replace ##DataFim## by variable dataFim
      //Substitui ##DataFim## pelo valor da variavel dataFim
      dadosDocumento.replaceText('##DataFim##', dataFim);
      //Replace ##Carga## by variable cargaHoraria
      //Substitui ##Carga## pelo valor da variavel cargaHoraria
      dadosDocumento.replaceText('##Carga##', cargaHoraria);
     
      //Saving docs
      //Salva e Fecha arquivo cópia
      DocumentApp.openById(idDocumento).saveAndClose();
      
      //Create a copy in pdf
      //Realiza uma cópia no formato PDF do arquivo cópia (documento) criado
      let thePDF = DriveApp.getFileById(idDocumento).getBlob().getAs('application/pdf').setName(nome + ".pdf" );
      newFolder.createFile(thePDF);
      
      //Delete the temporary docs
      //Apaga arquivo cópia (documento)
      DriveApp.getFileById(idDocumento).setTrashed(true);
      
      //HTML entries with giving name as parameter
      //Chamada do arquivo HTML com o nome passado por parametro
      let htmlOutput = HtmlService.createTemplateFromFile('Corpo_Email');
      
      //Get first name
      //Obtém o primeiro nome
      nome = nome.split(" ")[0];
      
      //Prepare variable for input as name on HTML
      //Prepara variável para input (entrada) do nome no HTML
      let guest_name = {"nome": nome};
      //Prepare variable for input as course name on HTML
      //Prepara variável para input (entrada) do nome do curso no HTML
      let guest_curso = {"nomecurso": nomecurso};
      //Input nome no HTML
      htmlOutput.nome = guest_name.nome;   
      //Input nome do curso no HTML
      htmlOutput.nomecurso = guest_curso.nomecurso;
      
      //Send personalized email
      //Envia email personalizado
      MailApp.sendEmail                                              
      ({
        //Destinatary 
        //Email do destinatário
        to: email.toString() , 
        
        //Assunto do email
        subject: "Envio de Certificado - Template GS", 
        
        //Corpo do email com o HTML
        htmlBody: htmlOutput.evaluate().getContent(),
        
        //Attach pdf certificate
        //Certificado em formato PDF como anexo
        attachments : thePDF 
        
       });
      //Set confirmation
      //Coloca o valor passado por parametro na célula abaixo da coluna "CRIOU DOC"
      dados.getCell(i,8).setValue(dia + "." + mes + "." + ano + " às " + horas+":"+minutos+":"+segundos);
    }
  }
  //Insert copy of sheet
  //Cria uma cópia da aba anterior
  planilha.insertSheet("Criada "+dia + "." + mes + "." + ano + " às " + horas+":"+minutos+":"+segundos,{template: aba});
}