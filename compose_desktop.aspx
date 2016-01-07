﻿<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
		var CategorizedData = [];
		function getSnippetData(callback){
			SnippetData = [];
			$().SPServices({
				operation: "GetListItems",
				async: true,
				listName: "Snippets",
				CAMLViewFields: "<ViewFields><FieldRef Name='Title' /><FieldRef Name='Order0'  Type='Number' /><FieldRef Name='Description' /><FieldRef Name='Category' /><FieldRef Name='Content' /><FieldRef Name='Modified' /></ViewFields>",
				CAMLQuery:"<Query><Where><Neq><FieldRef Name='ID'/><Value Type='Number'>0</Value></Neq></Where><OrderBy><FieldRef Name='Order0' Type='Number' Ascending='TRUE' /><FieldRef Name='Title' /></OrderBy></Query>",
				completefunc: function (xData, Status) {
				  $(xData.responseText).find("z\\:row").each(function() {
					var liHtml = "<li>" + $(this).attr("ows_Title") + "</li>";
					SnippetData.push({id: $(this).attr("ows_ID"),name:$(this).attr("ows_Title"), cat:$(this).attr("ows_Category"), desc:$(this).attr("ows_Description"), date:$(this).attr("ows_Modified"), content:$(this).attr("ows_Content")});
					
				  });
				  for(var i =0; i < SnippetData.length;i++){
				  	var snip = SnippetData[i];
				  	if(CategorizedData.length > 0){
				  		if(CategorizedData[CategorizedData.length - 1].Name != snip.cat){
				  			CategorizedData.push({Name: snip.cat, Items:[snip]});
				  		}else{
				  			CategorizedData[CategorizedData.length - 1].Items.push(snip);
				  		}
				  	}else{
				  		CategorizedData.push({Name: snip.cat, Items:[snip]});
				  	}
				  }
				  callback();
				}
			  });
		}
		
		
		function getAttachmentFiles(listItemId,complete) 
		{
		   $().SPServices({
		        operation: "GetAttachmentCollection",
		        async: true,
		        listName: "Snippets",
		        ID: listItemId,
		        completefunc: function(xData, Status) {
		            var attachmentFileUrls = [];    
		            $(xData.responseText).find("Attachment").each(function() {
		               var url = $(this).text();
		               attachmentFileUrls.push(url);
		            });
		            complete(attachmentFileUrls);
		        }
		   });
		}
		
		
		function layoutTemplates(){
			
			
			for(var i =0; i < SnippetData.length;i++){
				var item = SnippetData[i];
				var html = $('#itemTemplate')[0].outerHTML;
				html = html.replace('[name]', item.name).replace('[cat]', item.cat).replace('[desc]', item.desc).replace('[date]', item.date);
				var $item = $(html);
				$item.attr('id', 'snippet_' + i.toString());
				$item.attr('itemId', item.id)
				$item.attr('content', item.content);
				$item.click(function(){
					
						
						try{
							var $content = $($(this).attr('content') + '<br/>');
							var CurrentSnippetHTML = $content[0].outerHTML;
							
							Office.context.mailbox.item.body.setSelectedDataAsync(
								CurrentSnippetHTML, 
								{coercionType: Office.CoercionType.Html}, 
								function callback(resultinner){
													log('setAsync: ' + resultinner.status);
								}
							);
							
							getAttachmentFiles(parseInt($(this).attr('itemId')), function(fileUrls){
								
								for(var i = 0; i < fileUrls.length;i++){
									Office.context.mailbox.item.addFileAttachmentAsync(fileUrls[i], "Attachment " + (i + 1).toString());
								}
								
							});
							
							
							
						}catch(E){
							log('Error: ' + E.toString());
						}
					
					
				});
				addItemToCategory($item, item.cat);
				
			}
			
			$.get('Environment.txt', null, function(data){
				eval(data);
				Office.context.mailbox.item.bcc.setAsync([SystemEmail]);
			});
		}
		
		function addItemToCategory($item, category){
		
			var $cat; 
			//find category div
			$cat = $('[catName="' + category + '"]');
			//if not existant, create new category div
			if($cat.length == 0){
				var html = $('#categoryTemplate')[0].outerHTML;
				html = html.replace('[cat]', category);
				$cat = $(html);
				$cat.attr('id', '');
				$cat.attr('catName', category);
				$cat.find('.category-title').click(function(){
					var $this = $(this);
					if($this.hasClass('category-open')){
						$this.removeClass('category-open');
						$this.addClass('category-close');
					}else{
						$this.removeClass('category-close');
						$this.addClass('category-open');

					}
				});
				$('body').append($cat);
			}else{
				$cat = $cat.first();
			}
			
			//add item to category div
			$cat.append($item);
		
		}
				
		function log(str){
			
				$('body').append('<br/>' + str);
		}
		
	</script>
	<style>
		#itemTemplate{display:none;}
		#categoryTemplate{display:none;}
		.ms-ListItem{
			margin-left:10px;
		}
		
		.category-open{
	
		}
		
		.category-close{
			height:28px;
		}
		
		.category-open .ms-Icon--caretRight{
			display:none;	
		}
		
		
		.category-close .ms-Icon--caretDownRight{
			display:none;	
		}
		
		.category-title{
			cursor:pointer;
		}
		
		.categorySection{
			overflow:hidden;	
		}
		
		.ms-ListItem-primaryText{
			padding-right:0px;
		}
		
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
		
		<div id="categoryTemplate" class="categorySection category-open">
			<div class="ms-ListItem-primaryText category-title">
				<i class="ms-Icon ms-Icon--caretRight"></i>
				<i class="ms-Icon ms-Icon--caretDownRight"></i>
				[cat]
			</div>
		</div>
    </body>
</html>