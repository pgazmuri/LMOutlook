﻿<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%@ Page Language="C#" inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register tagprefix="SharePoint" namespace="Microsoft.SharePoint.WebControls" assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<WebPartPages:AllowFraming runat="server" __WebPartId="{DB8A38A8-7746-432A-A083-D650E8D5AF19}"/>

<html>
    <head>
<meta name="ProgId" content="SharePoint.WebPartPage.Document" />
		<meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
		<link rel="stylesheet" href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css"/>
		<link rel="stylesheet" href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css"/>
        <link rel="stylesheet" type="text/css" href="OutlookApp.css" />
		<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
		<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
		<script>
		
		// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

/**
 * Spinner Component
 *
 * An animating activity indicator.
 *
 */

/**
 * @namespace fabric
 */
var fabric = fabric || {};

/**
 * @param {HTMLDOMElement} target - The element the Spinner will attach itself to.
 */

fabric.Spinner = function(target) {

    var _target = target;
    var eightSize = 0.18;
    var circleObjects = [];
    var animationSpeed = 80;
    var interval;
    var spinner;
    var numCircles;
    var offsetSize;
    var fadeIncrement = 0;

    /**
     * @function start - starts or restarts the animation sequence
     * @memberOf fabric.Spinner
     */
    function start() {
        interval = setInterval(function() {
            var i = circleObjects.length;
            while(i--) {
                _fade(circleObjects[i]);
            }
        }, animationSpeed);
    }

    /**
     * @function stop - stops the animation sequence
     * @memberOf fabric.Spinner
     */
    function stop() {
        clearInterval(interval);
    }

    //private methods

    function _init() {
        offsetSize = eightSize;
        numCircles = 8;
        _createCirclesAndArrange();
        _initializeOpacities();
        start();
    }

    function _initializeOpacities() {
        var i = 0;
        var j = 2;
        var opacity;
        fadeIncrement = (1 / (numCircles + 2));

        for (i; i < numCircles; i++) {
            var circleObject = circleObjects[i];
            opacity = (fadeIncrement * j++);
            _setOpacity(circleObject.element, opacity);
        }
    }

    function _fade(circleObject) {
        var opacity = Math.round((_getOpacity(circleObject.element) - fadeIncrement) * 100) * 0.01;

        if (opacity <= 0) {
            opacity = 0.8;
        }

        _setOpacity(circleObject.element, opacity);
    }

    function _getOpacity(element) {
        return parseFloat(window.getComputedStyle(element).getPropertyValue("opacity"));
    }

    function _setOpacity(element, opacity) {
        element.style.opacity = opacity;
    }

    function _createCircle() {
        var circle = document.createElement('div');
        var parentWidth = parseInt(window.getComputedStyle(spinner).getPropertyValue("width"), 10);
        circle.className = "ms-Spinner-circle";
        circle.style.width = circle.style.height = parentWidth * offsetSize + "px";
        return circle;
    }

    function _createCirclesAndArrange() {
        //for backwards compatibility
        if (_target.className !== "ms-Spinner") {
            spinner = document.createElement("div");
            spinner.className = "ms-Spinner";
            _target.appendChild(spinner);
        } else {
            spinner = _target;
        }

        var width = spinner.clientWidth;
        var height = spinner.clientHeight;
        var angle = 0;
        var offset = width * offsetSize;
        var step = (2 * Math.PI) / numCircles;
        var i = numCircles;
        var circleObject;
        var radius = (width- offset) * 0.5;

        while (i--) {
            var circle = _createCircle();
            var x = Math.round(width * 0.5 + radius * Math.cos(angle) - circle.clientWidth * 0.5) - offset * 0.5;
            var y = Math.round(height * 0.5 + radius * Math.sin(angle) - circle.clientHeight * 0.5) - offset * 0.5;
            spinner.appendChild(circle);
            circle.style.left = x + 'px';
            circle.style.top = y + 'px';
            angle += step;
            circleObject = {element:circle, j:i};
            circleObjects.push(circleObject);
        }
    }

    _init();

    return {
        start:start,
        stop:stop
    };
};

		
		</script>
		
		<script type="text/javascript" defer>
		var SystemEmail;
		var TestEmails;
		var isTestMode = false;
		var TestingEnabled = false;
		$(function(){var spin8 = fabric.Spinner(jQuery("#spinner-8point")[0]);});
		
		var isLoadedOK = false;
		var docLoadCalled = false;
		var errorOnLoad = null;
		
		// The initialize function is required for all apps.
		try{
		
		var originalOfficeAtInitialization = JSON.stringify(Office);
		Office.initialize = function (reason) {
			isLoadedOK = true;
			// Checks for the DOM to load using the jQuery ready function.
			$(function(){
				docLoadCalled = true;
				try{
				// After the DOM is loaded, app-specific code can run.
				// Add any initialization logic to this function.
				
				var claimId = null;

				try{
					claimId = Office.context.mailbox.item.getRegExMatches().ClaimID[0];
					}catch(ignore){}
				if(claimId == null){
					try{
					claimId = Office.context.mailbox.item.getRegExMatches().ClaimIDSubject[0];
					}catch(ignore){}
				}
				
				if(claimId == null){
					$('#noClaimPara').show();
					$('#claimPara').hide();
				}
				if(claimId != null){
					document.getElementById('ClaimNumber').value = claimId;
				}
				
				document.onkeydown = function(e){
					if(TestingEnabled){
						var evtobj = window.event? event : e;
						//console.log('Key Pressed: ' + evtobj.keyCode + ' - Alt: ' + evtobj.altKey);
						if (evtobj.keyCode == 80 && evtobj.altKey) {
							
							if(!isTestMode){
								//enable test mode
								
								//clear emailPicker
								$('#emailPicker').find('option').remove();
								
								//populate emailPicker
								$.each(TestEmails, function(key, value) {   
									     $('#emailPicker')
									          .append($('<option></option>')
									          .attr("value", value)
									          .text(value));
								});
								
								//show TestHarness
								$('#TestHarness').show();
							}else{
								//hide TestHarness
								$('#TestHarness').hide();

							}
							
							//update boolean var to track test mode
							isTestMode = !isTestMode;
						}
					}
				};
				
				}catch(E){
					errorOnLoad = E;
					document.getElementById('AppMain').style.display = "none";
					document.getElementById('Loading').style.display = "none";
					document.getElementById('Error').style.display = "block";
					document.getElementById('ErrorMessage').innerText = "The following error occured when loading the claimID: " + E.toString();
				}
				
				$.get('Environment.txt', null, function(data){
					eval(data);	
					document.getElementById('AppMain').style.display = "block";
					document.getElementById('Loading').style.display = "none";
				});
				});
				
				

		};
		}catch(err){
					document.getElementById('AppMain').style.display = "none";
					document.getElementById('Loading').style.display = "none";
					document.getElementById('Error').style.display = "block";
					document.getElementById('ErrorMessage').innerText = "The following error occured when bootstrapping Office: " + err;
		}
		
		    // This function handles the click event of the sendNow button.
    // It retrieves the current mail item, so that we can get its itemId property.
    // It also retrieves the mailbox, so that we can make an EWS request
    // to get more properties of the item. In our case, we are interested in the ChangeKey
    // property, becuase we need that to forward a mail item.
    function sendNow() {
		
		var claimID = document.getElementById("ClaimNumber").value;
		if(! /^[0-9]{9}((-[0-9]{2})?)$/.test(claimID)){
					document.getElementById('Error').style.display = "block";
					document.getElementById('ErrorMessage').innerText = "The Claim ID is invalid.";
					return;
		}else{
			document.getElementById('Error').style.display = "none";
		}
	
        var item = Office.context.mailbox.item;
        item_id = item.itemId;
        mailbox = Office.context.mailbox;

		document.getElementById('AppMain').style.display = "none";
		document.getElementById('Loading').style.display = "block";
		
        // The following string is a valid SOAP envelope and request for getting the properties
        // of a mail item. Note that we use the item_id value (which we obtained above) to specify the item
        // we are interested in.
        var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>' +
            '    <GetItem' +
            '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '      <ItemShape>' +
            '        <t:BaseShape>IdOnly</t:BaseShape>' +
            '      </ItemShape>' +
            '      <ItemIds>' +
            '        <t:ItemId Id="' + item_id + '"/>' +
            '      </ItemIds>' +
            '    </GetItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        // The makeEwsRequestAsync method accepts a string of SOAP and a callback function
        mailbox.makeEwsRequestAsync(soapToGetItemData, getEnvironmentAndContinueSending);
    }
    
    function getEnvironmentAndContinueSending(asyncResult){
    	if(isTestMode){
    		SystemEmail = $('#emailPicker').val();
    	}
    	soapToGetItemDataCallback(asyncResult, SystemEmail);
    	
	}

    // This function is the callback for the makeEwsRequestAsync method
    // In brief, it first checks for an error repsonse, but if all is OK
    // it then parses the XML repsonse to extract the ChangeKey attribute of the 
    // t:ItemId element.
    function soapToGetItemDataCallback(asyncResult, SystemEmail) {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null) {
            document.getElementById('AppMain').style.display = "none";
			document.getElementById('Error').style.display = "block";
			document.getElementById('ErrorMessage').innerText = "The following error was recieved: " + asyncResult.error.message;      
        }
        else {
            var response = asyncResult.value;
			var changeKey = "";
            if (window.DOMParser) {
                var parser = new DOMParser();
                xmlDoc = parser.parseFromString(response, "text/xml");
            }
            else // Older Versions of Internet Explorer
            {
                xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                xmlDoc.async = false;
                xmlDoc.loadXML(response);
				
            }
            
			//we have to try getting the element both ways....
			try{
				changeKey = xmlDoc.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");
			}catch(E){
				changeKey = xmlDoc.getElementsByTagName("ItemId")[0].getAttribute("ChangeKey");
			}

            // Now that we have a ChangeKey value, we can use EWS to forward the mail item.
            // The first thing we'll do is get an array of email addresses that the user
            // has typed into the To: text box.
            // We'll also get the comment that the user may have provided in the Comment: text box.
            //var toAddresses = document.getElementById("groupEmails").value;
            var addresses = [SystemEmail]//toAddresses.split(";"); //ClaimsCommunicationsNon-Prod@exdevlibertymutual.com
            var addressesSoap = "";

            // The following loop build an XML fragment that we will insert into the SOAP message
            for (var address = 0; address < addresses.length; address++) {
                addressesSoap += "<t:Mailbox><t:EmailAddress>" + addresses[address] + "</t:EmailAddress></t:Mailbox>";
            }
            var comment = 'Forwarded from PI Claims Navigator - Claim Number: ' + document.getElementById("ClaimNumber").value;
            
			var newSubject = Office.context.mailbox.item.subject;
			
			//if the claimID isn't already there, let's add it:
			if(Office.context.mailbox.item.subject.indexOf('Claim#:' + document.getElementById("ClaimNumber").value) == -1){			
				newSubject = 'Claim#:' + document.getElementById("ClaimNumber").value + " - " + Office.context.mailbox.item.subject;
			}
			
			//tag subject with "inbound"
			newSubject = newSubject + ' - inbound';
			

            // The following string is a valid SOAP envelope and request for forwarding
            // a mail item. Note that we use the item_id value (which we obtained in the click event handler)
            // to specify the item we are interested in,
            // along with its ChangeKey value that we have just determined near the top of this function.
            // We also provide the XML fragment that we built in the loop above to specify the recipient addresses,
            // and the comment that the user may have provided in the Comment: text box
            var soapToForwardItem = '<?xml version="1.0" encoding="utf-8"?>' +
                '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
                '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
                '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                '  <soap:Header>' +
                '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
                '  </soap:Header>' +
                '  <soap:Body>' +
                '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
                '      <m:Items>' +
                '        <t:ForwardItem>' +
                '          <t:Subject>' + newSubject + '</t:Subject>' +
                '          <t:ToRecipients>' + addressesSoap + '</t:ToRecipients>' +
                '          <t:ReferenceItemId Id="' + item_id + '" ChangeKey="' + changeKey + '" />' +/*
                '          <t:NewBodyContent BodyType="Text">' + comment + '</t:NewBodyContent>' +*/
                '        </t:ForwardItem>' +
                '      </m:Items>' +
                '    </m:CreateItem>' +
                '  </soap:Body>' +
                '</soap:Envelope>';

            // As before, the makeEwsRequestAsync method accepts a string of SOAP and a callback function.
            // The only difference this time is that the body of the SOAP message requests that the item
            // be forwarded (rather than retrieved as in the previous method call)
            mailbox.makeEwsRequestAsync(soapToForwardItem, soapToForwardItemCallback);
        }
    }

    // This function is the callback for the above makeEwsRequestAsync method
    // In brief, it first checks for an error repsonse, but if all is OK
    // it then parses the XML repsonse to extract the m:ResponseCode value.
    function soapToForwardItemCallback(asyncResult) {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null) {
			document.getElementById('AppMain').style.display = "none";
			document.getElementById('Error').style.display = "block";
			document.getElementById('ErrorMessage').innerText = "The following error was recieved: " + asyncResult.error.message;
        }
        else {
            var response = asyncResult.value;
			var result = "Error parsing result from SOAP XML";
            if (window.DOMParser) {
                parser = new DOMParser();
                xmlDoc = parser.parseFromString(response, "text/xml");
				
            }
            else // Older Versions of Internet Explorer
            {
                xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                xmlDoc.async = false;
                xmlDoc.loadXML(response);
				result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;
            }
			
			//we have to try getting the element both ways....
			try{
				result = xmlDoc.getElementsByTagName("ResponseCode")[0].textContent;
			}catch(E){
				result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;
			}

            // Get the required response, and if it's NoError then all has succeeded, so tell the user.
            // Otherwise, tell them what the problem was. (E.G. Recipient email addresses might have been
            // entered incorrectly --- try it and see for yourself what happens!!)
            if (result == "NoError") {
                document.getElementById('Success').style.display = "block";
				document.getElementById('AppMain').style.display = "none";
				document.getElementById('Loading').style.display = "none";
            }
            else {
				document.getElementById('AppMain').style.display = "block";
				document.getElementById('Loading').style.display = "none";
				document.getElementById('Error').style.display = "block";
				document.getElementById('ErrorMessage').innerText = "The following error was recieved: " + result;
            }
        }

	}
	
	
	function debugHelp(){
		document.write('isloaded: ' + isLoadedOK + "     docLoadCalled:" + docLoadCalled + "  Office:" + originalOfficeAtInitialization + " Initialize:" + Office.initialize + " <br/><br/>Context:" + JSON.stringify(Office.context));
		Office.context = "something to stop refresh";
	}
	
	function fixTheMacBrokenness(){
		window.location.href = "read_desktop.aspx?reload";
		
	}
	
	if(window.location.href.indexOf('?reload') == -1){
		setTimeout(function(){
			if(Office.context == undefined || Office.context.mailbox == undefined){
				fixTheMacBrokenness();
			}
		}, 2000);
	}
	
	</script>
		
    </head>
    <body class="ms-font-m-plus">
		<div style="float:left;">
		<img src="Logo.png"/><span style="position:relative;top:-25px; left:15px;" class="ms-font-xxl ms-fontColor-themePrimary">PI Claims Navigator</span>
		</div>
		<div style="float:right;">
			<img width=64 src="NavigatorIcon.jpg"/>
		</div>
		<div style="clear:both;"></div>
		<div id="AppMain" style="display: none;">
			<p id="claimPara">We've detected a claim number in this email.  Would you like to file this in Navigator?</p>
			<p id="noClaimPara" style="display:none;">Please enter the associated claim number below.</p>
			<div class="ms-TextField" style="float:left;">
				<label class="ms-font-l" style="display: inline-block;">Claim ID:&nbsp;</label>
				<input type="text" class="ms-TextField-field" id="ClaimNumber" style="width: 120px;font-size: larger;"/><br>
				<span class="ms-TextField-description">Please correct the claim ID if it is not correct.</span>
			</div>
			<div style="float:left;margin-left:15px;">
			
			<a href="javascript:sendNow();" class="ms-Button ms-Button--primary" style="text-decoration:none;"><span class="ms-Button-label" style="text-decoration:none;">File in Navigator</span></a><p></p>
		</div>
		<div style="clear:both;"></div>
		</div>
		<div id="Loading" style="margin:auto;">
			<div style="width:64px; margin:auto;">
				<div id="spinner-8point"></div><p>Working...</p>
			</div>
		</div>
		<div id="TestHarness" style="display:none;">
			System Email Address: <select id="emailPicker"></select>
		</div>
		<div id="Error" style="display:none;">
			<i class="ms-Icon ms-Icon--alert"></i>&nbsp;&nbsp;<span id="ErrorMessage"></span>
		</div>
		<div id="Success" style="display:none;">
			<i class="ms-Icon ms-Icon--check"></i>&nbsp;&nbsp;<span id="SuccessMessage">Email Forwarded Successfully!</span>
		</div>
    </body>
</html>