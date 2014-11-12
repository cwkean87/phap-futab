
var submissionSSKey = '1cDKsgfwAXfFG-yL1oCRKn9zeTs37Frt_3pkPbvz4DOc';// replace with your spreadsheet ID


//====================================/////////////////////////////////============================================================================//
//===================================// DO NOT MODIFY ANY TEXT BELOW //============================================================================//
//====================================////////////////////////////////=============================================================================//

var Panelstyle = {'background':'#E6E6FA','padding':'40px','borderStyle':'solid','borderWidth':'5PX','borderColor':'red', 'borderRadius':'15px'};
var cssTextBox = {'borderRadius':'10px', 'paddingLeft':'8px'};
var refSs = SpreadsheetApp.openById(submissionSSKey);
var items = refSs.getRangeByName('DropdownList').getValues();
var $formTitle = refSs.getRangeByName('FormTitle').getValue();
var $title = refSs.getRangeByName('Title').getValue();
var $name = refSs.getRangeByName('Name').getValue();
var $email = refSs.getRangeByName('Email').getValue();
var $file = refSs.getRangeByName('File').getValue();
var $instruction = refSs.getRangeByName('Instruction').getValue();
var $postUpload = refSs.getRangeByName('PostUpload').getValue();
var $folderName = refSs.getRangeByName('FolderName').getValue();
var $peopleEmail = refSs.getRangeByName('PeopleEmail').getValues();
var refSs2 = SpreadsheetApp.openById('1cmA8CbRqfuoH9OrTwDChn7reGsuKBuimB3yU3TUoitk');
var $Worksheet = refSs2.getRangeByName('Worksheet').getValues();
var $DueDate = refSs2.getRangeByName('DueDate').getValues();

function doGet() {
  var app = UiApp.createApplication().setTitle($formTitle).setStyleAttribute('padding','50PX');
  var panel = app.createFormPanel().setStyleAttributes(Panelstyle).setPixelSize(400, 200);
  var title = app.createLabel($formTitle).setStyleAttribute('color','#4B0082').setStyleAttribute('fontSize','25PX').setStyleAttribute('fontWeight', 'bold');
  var grid = app.createGrid(6,2).setId('grid');
  var list1 = app.createListBox().setName('list1').setStyleAttributes(cssTextBox).setSize('160', '30');
   for(var i=0; i<items.length; ++i){list1.addItem(items[i])}
  var Textbox1 = app.createTextBox().setWidth('150px').setName('TB1').setStyleAttributes(cssTextBox).setSize('160', '30');
  var email = app.createTextBox().setWidth('150px').setName('mail').setStyleAttributes(cssTextBox).setSize('160', '30');
  var upLoad = app.createFileUpload().setName('uploadedFile');
  var submitButton = app.createSubmitButton('<B>Upload</B>').setStyleAttribute('borderRadius', '10px'); 
  var warning = app.createLabel($instruction).setStyleAttribute('background','#bbbbbb').setStyleAttribute('fontSize','18px');
  //file upload
  var cliHandler2 = app.createClientHandler()
  .validateLength(Textbox1, 1, 40).validateNotMatches(list1,'Select a Case Topic').validateEmail(email).validateNotMatches(upLoad, 'FileUpload')
  .forTargets(submitButton).setEnabled(true)
  .forTargets(warning).setHTML('Ready to upload').setStyleAttribute('background','#99FF99').setStyleAttribute('fontSize','14px');
  //Grid layout of items on form
  grid.setWidget(0, 1, title)
      .setText(1, 0, $title).setStyleAttribute('fontWeight', 'bold')
      .setWidget(1, 1, list1.addClickHandler(cliHandler2))
      .setText(2, 0, $name).setStyleAttribute('fontWeight', 'bold')
      .setWidget(2, 1, Textbox1.addClickHandler(cliHandler2))
      .setText(3, 0, $email).setStyleAttribute('fontWeight', 'bold')
      .setWidget(3, 1, email)
      .setText(4, 0, $file).setStyleAttribute('fontWeight', 'bold')
      .setWidget(4, 1, upLoad.addChangeHandler(cliHandler2))
      .setWidget(5, 0, submitButton)
      .setWidget(5, 1, warning);

  var cliHandler = app.createClientHandler().forTargets(warning).setText($postUpload).setStyleAttribute('background','yellow');
  submitButton.addClickHandler(cliHandler).setEnabled(false);  
  panel.add(grid);
  app.add(panel)
     .add(app.createLabel()           
       .setWidth("500")
       .setHeight("22")
       .setStyleAttribute("backgroundColor", "red")
       .setStyleAttribute("color", "white")
       .setStyleAttribute("position", "fixed")
       .setStyleAttribute("top", "0px")
       .setStyleAttribute("left", "0px"));
  return app;
}


function doPost(e) {

     var app = UiApp.getActiveApplication();
     var ListVal = e.parameter.list1;
     var today = new Date();
     var dateSubmit = Utilities.formatDate(today, 'GMT+8', 'dd/MM/yyyy');
try{
   for (var i=0; i<$Worksheet.length; i++){
      Logger.log($Worksheet[i]);
      if ($Worksheet[i] == ListVal ){
        var cutODate = $DueDate[i].toString();//date format from the spreadsheet must be in Plain Text.
        var cutDate = dateToString(cutODate);
        break;//stop the loop if found match of ListVal
      }
   }
    Logger.log(today+' is submit date');
    Logger.log(cutDate+' is cutoff date');
    if (today > cutDate){
       app.add(app.createLabel('Sorry you are either late for submission or the worksheet has not been uploaded by preceptor yet. Please contact your preceptor.'));
       return app;
     } 
  else {    
  var ListValS = ListVal.toString();
  var textVal = e.parameter.TB1;
  var Email = e.parameter.mail;
  var fileBlob = e.parameter.uploadedFile;
  var blob = fileBlob.setContentTypeFromExtension();
  var img = DocsList.createFile(blob);
  var imgName = img.getName();
  var attachment = img.getBlob();
    var rename = ListVal + ' - ' +textVal;
  var imgRename = img.rename(rename);
  
    try{
      var folder = DocsList.getFolder(ListValS);} 
    catch(e){DocsList.getFolder('For Preceptors').createFolder(ListValS);var folder = DocsList.getFolder(ListValS)}
  img.addToFolder(folder);
  img.removeFromFolder(DocsList.getRootFolder());
  var sheet = SpreadsheetApp.openById(submissionSSKey).getSheetByName('Submission');
  var lastRow = sheet.getLastRow();
  var today = new Date();
  var imgUrl =  img.getUrl();
  var folderUrl = folder.getUrl();
  var shortUrls = shortenUrl(imgUrl,rename);
  var shortUrls2 = shortenUrl(folderUrl,ListValS);
  var dateSubmit = Utilities.formatDate(today, 'GMT+8', 'dd/MM/yyyy');
  var targetRange = sheet.getRange(lastRow+1, 1, 1, 5).setValues([[dateSubmit,ListVal,textVal,Email,shortUrls]]);


    for (var i=1; i<items.length; i++){
      Logger.log(items[i]);
      if (items[i] == ListVal){
          var userEmail = $peopleEmail[i];

  var templateSheet = refSs.getSheetByName("Online Form");
  var subject = templateSheet.getRange("I11").getValue();
  var uploaderEmail = e.parameter.mail;
  var bodyHTML1 = templateSheet.getRange("I6:I8").getValue();
  var emailText = fillInTemplateFromEntry_(bodyHTML1, textVal, ListVal, Email, shortUrls,shortUrls2);
  var emailObject = {htmlBody: emailText, name:"Pre-reg Case Study Drive", attachments: [attachment]};
      var bodyForPrereg = "";
    bodyForPrereg += "<p>Dear"+textVal+",</p>";
    bodyForPrereg += "<p>Your case topic have been uploaded succesfully.</p>";
    bodyForPrereg += "<p>Thank you for supporting the system.</p>";
  MailApp.sendEmail(userEmail, subject, '', emailObject);
  MailApp.sendEmail(uploaderEmail, subject, "", {htmlBody: bodyForPrereg, name:"Case Topic Upload Notification"} )
    }
   }

  app.add(app.createLabel('File uploaded. A notification email will be sent to you shortly. Thank you.'));
  return app
  }
  }
catch(e){
  Logger.log(e)
}
   }
