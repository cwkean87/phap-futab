//this is the "plug-in" script for the gas-futab script

/*
function to find certain keywork in a text and replace with another value (from var).
*/

// this is the fucntion name,then(para1, para2)
function fillInTemplateFromEntry_(template, data1, data2, data3, data4, data5) {
    
  //text of email = parameter1 
  var email = template;//template is para1
  
  // Search for the keywork Pharmacist Initial in the email text template
  var templateVars1 = template.match("NAME");
  var templateVars2 = template.match("LIST");
  var templateVars3 = template.match("EMAIL");
  var templateVars4 = template.match("URL");
  var templateVars5 = template.match("CUTOFFDATE");

  // Replace text from the template with the actual values from parameter2.
  // If no value is available, replace with the empty string.
  
    var variableData1 = data1;//data is para2
    var variableData2 = data2;
    var variableData3 = data3;
    var variableData4 = data4;
    var variableData5 = data5;
    
  //now, para1=para1.replace(keyword with para2, or empty string
    email = email.replace(templateVars1, variableData1 || "");
    email = email.replace(templateVars2, variableData2 || "");
    email = email.replace(templateVars3, variableData3 || "");
    email = email.replace(templateVars4, variableData4 || "");
    email = email.replace(templateVars5, variableData5 || "");

  // then, show the new text after replaced
  return email;
}

//URL Shortener Function
function shortenUrl(url, title, keyword){
var url = url;
  var title = title;
  //generate random string
var chars = "abcdefghijklmnopqrstuvwxyz0123456789abcdefghiklmnopqrstuvwxyz";
	var text_length = 5;
	var custom = '';
	for (var i=0; i<text_length; i++) {
		var rnum = Math.floor(Math.random() * chars.length);
		custom += chars.substring(rnum,rnum+1);
	}

var shortener = "http://pharmfair.com/yourls-api.php?signature=c3f62f7271&action=shorturl&url="+url+"&keyword="+keyword+"&title="+title+"&format=simple";
var response = UrlFetchApp.fetch(shortener);
  
return response;
}

function toCamelCase(s) {
    // remove all characters that should not be in a variable name
    // as well underscores an numbers from the beginning of the string
    s = s.replace(/([^a-zA-Z0-9_\- ])|^[_0-9]+/g, "").trim().toLowerCase();
    // uppercase letters preceeded by a hyphen or a space
    s = s.replace(/([ -]+)([a-zA-Z0-9])/g, function(a,b,c) {
        return c.toLowerCase();
    });
    // uppercase letters following numbers
    s = s.replace(/([0-9]+)([a-zA-Z])/g, function(a,b,c) {
        return b + c.toLowerCase();
    });
    return s;
}


function dateToString(dateString) {
  var dateArray = dateString.split(" ");
  var month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var monthNum = [1,2,3,4,5,6,7,8,9,10,11,12];
  var m = dateArray[1];//month cant be used coz it return Jun, which is a string, think a way to convert it to integer
  var year = dateArray[3];
  for (var i=0; i<=month.length;i++){
    Logger.log(month[i]);
    if ( m == month[i]){
    var newM = monthNum[i];//convert month in string to interger
    }
  }
  var day = dateArray[2];
  var date = new Date(year, newM-1, day);

  return date;
}

function sendEmail(email, url){
  var templateSh = SpreadsheetApp.openById('1PvN2qDhxugopyUjaYrrzFCmlMxtxvJYsS8GSvfFRVek');
  var subject = templateSh.getRangeByName('subject').getValue();
  var paragraph1 = templateSh.getRangeByName('paragraph1').getValue();
  var paragraph2 = templateSh.getRangeByName('paragraph2').getValue();
  var paragraph3 = templateSh.getRangeByName('paragraph3').getValue();
  var closing = templateSh.getRangeByName('closing').getValue();  
  var body = "";
  var htmlBody = "";
  htmlBody += "<p>"+paragraph1+"</p>";
  htmlBody += "<p>"+paragraph2+"</p>";
  htmlBody += "<p>"+paragraph3+"</p>";
  htmlBody += "<p>Click this link to access your personal record: <a href='"+url+"'>My PHAP</a></p>";
  htmlBody += "<p>"+closing+"</p>";
  
  
  MailApp.sendEmail(email, subject, body, {name:"PHAP Admin", htmlBody:htmlBody});
  Logger.log("Email sent successfully");

}
