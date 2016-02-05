<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%@ Page Language="C#" %>
<%@ Register tagprefix="SharePoint" namespace="Microsoft.SharePoint.WebControls" assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<SharePoint:CssRegistration Name="default" runat="server"/>
		<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.SPServices/2014.02/jquery.SPServices-2014.02.min.js"></script>

<script>

	var fileNames=[];
function CreateItemFromTemplate()
{
    
    var fso = new ActiveXObject("Scripting.FileSystemObject"); 
	var f = fso.GetFolder($('#directoryPath').val());
	fileNames=[];
	var list = "";
	
	// Traverse through the FileCollection using the FOR loop
	for(var objEnum = new Enumerator(f.Files); !objEnum.atEnd(); objEnum.moveNext()) {
	   strFileName = objEnum.item();
	   if(/\.oft$/gi.test(strFileName)){
	   	list += (strFileName + "<BR>");
	   	fileNames.push(strFileName);
	   }
	}
	
      $('#bodyContent').html(list + "<br/><a href='javascript:loadOFTs();'>Click Here to Convert above OFTs to Snippets</a>");

	return;
	
	}
	
var theApp;

function loadNextOFT(){
    try
    {
    if (fileNames.length == 0){
    
    	$('#bodyContent').append("<hr/>Conversion Complete!  Go to snippets and update image references.<br/>");

    	return;
    }
    
    var i = fileNames.length - 1;	    
	    
	    var theMailItem = theApp.CreateItemFromTemplate(fileNames[i]);
	      //theMailItem.Body = (msg);
	      //Show the mail before sending for review purpose
	      //You can directly use the theMailItem.send() function
	      //if you do not want to show the message.
	      //theMailItem.display();
	       //debugger;
	      //$('#bodyContent').html(theMailItem.HTMLBody);
	      var atts = getAttachments(theMailItem);
	      
	      var filename = fileNames[i].Name.replace(/\.oft/gi, "");
	      //debugger;
	      InsertNewTemplate(atts, filename, theMailItem.subject, theMailItem.HTMLBody, (function(fn){ return function(){ $('#bodyContent').append('<br/>Created snippet: ' + fn); fileNames.pop(); loadNextOFT(); };})(filename));
	      theMailItem.close(1);
      }
    catch(err)
    {
    alert(err.toString() + "\nThe following may have cause this error: \n"+
     "1. The Outlook express 2007 is not installed on the machine.\n"+
     "2. The msoutl.olb is not availabe at the location "+
     "C:\\Program Files\\Microsoft Office\\OFFICE11\\msoutl.old on client's machine "+
     "due to bad installation of the office 2007."+
     "Re-Install office2007 with default settings.\n"+
     "3. The Initialize and Scripts ActiveX controls not marked as safe is not set to enable.")
    document.write(err.toString());
    }

}	
	
function getAttachments(theMailItem){
	var totalCount = theMailItem.attachments.Count;
	var files = [];
	for(var i = 1; i <= totalCount; i++){
		var att = theMailItem.attachments.Item(i);
		if(!att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B")){
			//if not hidden, we save
			//var bytes = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102");
			var attFileName  = $('#directoryPath').val() + guid();//att.GetTemporaryFilePath();
			att.SaveAsFile(attFileName);
			files.push([attFileName, att.DisplayName]);
		}
	}	
	return files;
}

function asc(c) //(string)->integer
{   // Objet : Renvoie le code ASCII du 1er caractère de la chaine "c"
    var i = c.charCodeAt(0);
    if (i < 256)
        return i;       // caractères ordinaires
    // (plage 128-159 excepté les 5 caractères 129,141,143,144,157, qui fonctionnent normalement)
    switch (i) {
        case 8364:      // "€" 
            return 128
        case 8218:
            return 130
        case 402:
            return 131
        case 8222:
            return 132
        case 8230:
            return 133
        case 8224:
            return 134
        case 8225:
            return 135
        case 710:
            return 136
        case 8240:
            return 137
        case 352:
            return 138
        case 8249:
            return 139
        case 338:
            return 140
        case 381:
            return 142
        case 8216:
            return 145
        case 8217:
            return 146
        case 8220:
            return 147
        case 8221:
            return 148
        case 8226:
            return 149
        case 8211:
            return 150
        case 8212:
            return 151
        case 732:
            return 152
        case 8482:
            return 153
        case 353:
            return 154
        case 8250:
            return 155
        case 339:
            return 156
        case 382:
            return 158
        case 376:
            return 159
        default:
            return -1       // provoquera une erreur, le cas ne devant pas se présenter
    }
}
	
function guid() {
  function s4() {
    return Math.floor((1 + Math.random()) * 0x10000)
      .toString(16)
      .substring(1);
  }
  return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
    s4() + '-' + s4() + s4() + s4();
}
	
function loadOFTs(){
	
    try
    {
    	theApp = new ActiveXObject("Outlook.Application");
    	$('#bodyContent').html("<hr/>Running Conversion...<br/>");
    	loadNextOFT();
      }
    catch(err)
    {
    alert(err.toString() + "\nThe following may have cause this error: \n"+
     "1. The Outlook express 2007 is not installed on the machine.\n"+
     "2. The msoutl.olb is not availabe at the location "+
     "C:\\Program Files\\Microsoft Office\\OFFICE11\\msoutl.old on client's machine "+
     "due to bad installation of the office 2007."+
     "Re-Install office2007 with default settings.\n"+
     "3. The Initialize and Scripts ActiveX controls not marked as safe is not set to enable.")
    document.write(err.toString());
    }
}

function InsertNewTemplate(Attachments, Title, Subject, HTML, callback){
	 $().SPServices({
                operation: "UpdateListItems",
                async: true,
                batchCmd: "New",
                listName: "Snippets",
                valuepairs: [["Title", Title], ["Subject", Subject], ["Content", htmlEscape(HTML)], ['Category', $('#category').val()], ['Mode', $('#mode').val()]],
                completefunc: 
                (function(innerAttachments){
                	return function (xData, Status) {
	                	
	                	if(innerAttachments.length > 0){
		                	for(var x=0; x<innerAttachments.length;x++){
		                		var file = innerAttachments[x][0];
								  $().SPServices({
								    operation: "AddAttachment",
								    listName: "Snippets",
								    listItemID: $(xData.responseText).find("z\\:row").attr("ows_ID"),
								    fileName: innerAttachments[x][1],
								    attachment: getFileBufferBase64(file)
								  });
								
							}
	                	}
	                	
	                	//debugger;
	                    callback();
	                };
	            })(Attachments)
                
            });
}



var getFileBufferBase64 = function(file) {

	var bytes = [];
	
	fso = new ActiveXObject("Scripting.FileSystemObject");
	var fsoFile = fso.OpenTextFile(file)
	while(!fsoFile.AtEndOfStream){
		bytes.push(asc(fsoFile.Read(1)));
	}
	return base64ArrayBuffer(bytes);

};


function base64ArrayBuffer(arrayBuffer) {
  var base64    = ''
  var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

  var bytes         = new Uint8Array(arrayBuffer)
  var byteLength    = bytes.byteLength
  var byteRemainder = byteLength % 3
  var mainLength    = byteLength - byteRemainder

  var a, b, c, d
  var chunk

  // Main loop deals with bytes in chunks of 3
  for (var i = 0; i < mainLength; i = i + 3) {
    // Combine the three bytes into a single integer
    chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2]

    // Use bitmasks to extract 6-bit segments from the triplet
    a = (chunk & 16515072) >> 18 // 16515072 = (2^6 - 1) << 18
    b = (chunk & 258048)   >> 12 // 258048   = (2^6 - 1) << 12
    c = (chunk & 4032)     >>  6 // 4032     = (2^6 - 1) << 6
    d = chunk & 63               // 63       = 2^6 - 1

    // Convert the raw binary segments to the appropriate ASCII encoding
    base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d]
  }

  // Deal with the remaining bytes and padding
  if (byteRemainder == 1) {
    chunk = bytes[mainLength]

    a = (chunk & 252) >> 2 // 252 = (2^6 - 1) << 2

    // Set the 4 least significant bits to zero
    b = (chunk & 3)   << 4 // 3   = 2^2 - 1

    base64 += encodings[a] + encodings[b] + '=='
  } else if (byteRemainder == 2) {
    chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1]

    a = (chunk & 64512) >> 10 // 64512 = (2^6 - 1) << 10
    b = (chunk & 1008)  >>  4 // 1008  = (2^6 - 1) << 4

    // Set the 2 least significant bits to zero
    c = (chunk & 15)    <<  2 // 15    = 2^4 - 1

    base64 += encodings[a] + encodings[b] + encodings[c] + '='
  }
  
  return base64
}

//This function makes the magic
function htmlEscape(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

</script>

</head>

<body>

<form id="form1" runat="server">
Enter Directory to pull OFT files from:
<input type="text" id="directoryPath" value="C:\temp\"></input><br/>
Enter Category to use for OFTs:
<input type="text" id="category" value="General Content"></input><br/>
Enter Mode to use for OFTs:
<select id="mode">
<option value="Snippet">Snippet</option>
<option value="Template">Template</option>

</select><br/>


<a href="Javascript:CreateItemFromTemplate();">Click Here to Find OFT Files</a>

<div id="bodyContent"></div>



</form>

</body>

</html>
