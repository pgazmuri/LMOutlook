<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%@ Page Language="C#" inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register tagprefix="SharePoint" namespace="Microsoft.SharePoint.WebControls" assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<WebPartPages:AllowFraming runat="server" __WebPartId="{DB8A38A8-7746-432A-A083-D650E8D5AF19}" />
<html>
    <head>
        <meta name="ProgId" content="SharePoint.WebPartPage.Document" />
		<meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
		<link rel="stylesheet" href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css"/>
		<link rel="stylesheet" href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css"/>
        <link rel="stylesheet" type="text/css" href="OutlookApp.css" />
		<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.SPServices/2014.02/jquery.SPServices-2014.02.min.js"></script>
		<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
		
		
		<script type="text/javascript">
		var isLoadedOK = false;
		var docLoadCalled = false;
		var errorOnLoad = null;
		// The initialize function is required for all apps.
		Office.initialize = function (reason) {
			isLoadedOK = true;
			// Checks for the DOM to load using the jQuery ready function.
			$(document).ready(function(){
				docLoadCalled = true;
				
				//load snippets
				//SnippetData = [];
				//layoutTemplates();
				getSnippetData(layoutTemplates);
				});
		};
		
		var SnippetData = [];//[{name:"Header", cat:"General Content", desc:"Standard Email Header", date:"1/5/2015", content:"<h1>This is the header</h1>And some other content<br/><br/>"}];
		
		function getSnippetData(callback){
			SnippetData = [];
			$().SPServices({
				operation: "GetListItems",
				async: true,
				listName: "Snippets",
				CAMLViewFields: "<ViewFields><FieldRef Name='Title' /><FieldRef Name='Description' /><FieldRef Name='Category' /><FieldRef Name='Content' /><FieldRef Name='Modified' /></ViewFields>",
				completefunc: function (xData, Status) {
				  $(xData.responseText).find("z\\:row").each(function() {
					var liHtml = "<li>" + $(this).attr("ows_Title") + "</li>";
					SnippetData.push({name:$(this).attr("ows_Title"), cat:$(this).attr("ows_Category"), desc:$(this).attr("ows_Description"), date:$(this).attr("ows_Modified"), content:$(this).attr("ows_Content")});
					
				  });
				  callback();
				}
			  });
		}
		
				
		function swapImagesForAttachmentUrls($item, callback){
			$imgs = $item.find('img');
			var imgCount = $imgs.length;
			var imgID = 0;
			$imgs.each(function(){
				var mySrc = $(this).attr('src');
				if(mySrc.indexOf('/') == 0){
					mySrc = window.location.protocol + "//" + window.location.host + mySrc;
				}
				var pos = mySrc.lastIndexOf('/');
				var fileName = mySrc.substr(pos + 1, mySrc.length - pos);
				Office.context.mailbox.item.addFileAttachmentAsync(mySrc, "img_" + imgID, {asyncContext: fileName}, function(ctx){
					$(this).attr('src', "cid:" + ctx.asyncContext);
					imgCount--;
				});
				imgID++;
			});
			
			var timer = setInterval(function(){
				if(imgCount == 0){
					clearInterval(timer);
					callback();
				}
			}, 500);
			
		}
		
		function layoutTemplates(){
			
			
			for(var i =0; i < SnippetData.length;i++){
				var item = SnippetData[i];
				var html = $('#itemTemplate')[0].outerHTML;
				html = html.replace('[name]', item.name).replace('[cat]', item.cat).replace('[desc]', item.desc).replace('[date]', item.date);
				var $item = $(html + "<br/>");
				$item.attr('id', 'snippet_' + i.toString());
				$item.attr('content', item.content);
				$item.click(function(){
					
						
						try{
							var $content = $($(this).attr('content') + '<br/>');
							swapImagesForAttachmentUrls($content, function(){
								var CurrentSnippetHTML = $content[0].outerHTML;
								
								Office.context.mailbox.item.body.setSelectedDataAsync(
									CurrentSnippetHTML, 
									{coercionType: Office.CoercionType.Html}, 
									function callback(resultinner){
														log('setAsync: ' + resultinner.status);
									}
								);
							});
							
							
						}catch(E){
							log('Error: ' + E.toString());
						}
					
					
				});
				$('body').append($item);
			}
		}
		
				
		function log(str){
			
				$('body').append('<br/>' + str);
		}
		
	</script>
	<style>
		#itemTemplate{display:none;}
	</style>
		
    </head>
    <body class="ms-font-m-plus">
	
		<div id="itemTemplate" class="ms-ListItem is-unread" style="cursor:pointer;">
		  <span class="ms-ListItem-primaryText">[name]</span>  
		  <span class="ms-ListItem-secondaryText">[date]</span>
		  <span class="ms-ListItem-tertiaryText">[desc]</span>
		  <div class="ms-ListItem-actions">
			<div class="ms-ListItem-action"></div>
			<div class="ms-ListItem-action"></div>
			<div class="ms-ListItem-action"></div>
			<div class="ms-ListItem-action"></div>
		  </div>
		</div>
    </body>
</html>