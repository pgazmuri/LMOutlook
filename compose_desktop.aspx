<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
		<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.SPServices/2014.02/jquery.SPServices-2014.02.min.js"></script>
		<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
				
		<script type="text/javascript">
		
		function genericErrorHandler(asyncResult){
			var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 log(error.name + ": " + error.message);
            }
		}
		
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
				try{
					getSnippetData(layoutTemplates);
				}catch(E){
					log("Error: " + E.toString());
				}
				
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
				CAMLViewFields: "<ViewFields><FieldRef Name='Title' /><FieldRef Name='Mode' /><FieldRef Name='Subject' /><FieldRef Name='Order0'  Type='Number' /><FieldRef Name='Description' /><FieldRef Name='Category' /><FieldRef Name='Content' /><FieldRef Name='Modified' /></ViewFields>",
				CAMLQuery:"<Query><Where><Neq><FieldRef Name='ID'/><Value Type='Number'>0</Value></Neq></Where><OrderBy><FieldRef Name='Order0' Type='Number' Ascending='TRUE' /><FieldRef Name='Title' /></OrderBy></Query>",
				completefunc: function (xData, Status) {
				try{
					  $(xData.responseText).find("z\\:row").each(function() {
						var liHtml = "<li>" + $(this).attr("ows_Title") + "</li>";
						SnippetData.push({id: $(this).attr("ows_ID"),mode: $(this).attr("ows_Mode"),subject:$(this).attr("ows_Subject"),name:$(this).attr("ows_Title"), cat:$(this).attr("ows_Category"), desc:$(this).attr("ows_Description"), date:$(this).attr("ows_Modified"), content:$(this).attr("ows_Content")});
						
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
				 }catch(E){
				 	log("Error: " + E.toString());
				 }
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
		        	try{
			            var attachmentFileUrls = [];    
			            $(xData.responseText).find("Attachment").each(function() {
			               var url = $(this).text();
			               attachmentFileUrls.push(url);
			            });
			            complete(attachmentFileUrls);
		            }catch(E){
				 		log("Error: " + E.toString());
				 	}

		        }
		   });
		}
		
		
		function layoutTemplates(){
			
			try{
				for(var i =0; i < SnippetData.length;i++){
					var item = SnippetData[i];
					var html = $('#itemTemplate')[0].outerHTML;
					html = html.replace('[name]', item.name).replace('[cat]', item.cat).replace('[desc]', item.desc).replace('[date]', item.date).replace('[mode]', item.mode);
					var $item = $(html);
					var $anchor = $item.find('a');
					$item.attr('id', 'snippet_' + i.toString());
					$anchor.attr('itemId', item.id)
					$anchor.attr('content', item.content);
					$anchor.attr('subject', item.subject);

					$anchor.click(function(){
						try{
						
													
							var $content = $($(this).attr('content') + '<br/>');
							var subject = $(this).attr('subject');
							var CurrentSnippetHTML = $content[0].outerHTML;
							
							if($(this).text() == "Insert"){
							Office.context.mailbox.item.body.setSelectedDataAsync(
								CurrentSnippetHTML, 
								{coercionType: Office.CoercionType.Html}, 
								genericErrorHandler);
							}else{
							Office.context.mailbox.item.body.setAsync(
								CurrentSnippetHTML, 
								{coercionType: Office.CoercionType.Html}, 
								genericErrorHandler);
								

							}
							
							if(subject != null && subject.length > 0){
								Office.context.mailbox.item.subject.getAsync(function(result){
									//set the email subject if it's just the claim ID
									if(result.value == 'Claim#:' + globalClaimNumber){
										Office.context.mailbox.item.subject.setAsync('Claim#:' + globalClaimNumber + ' - ' + subject, {}, genericErrorHandler);
									}
								});
							}
							
							getAttachmentFiles(parseInt($(this).attr('itemId')), function(fileUrls){
								
								for(var i = 0; i < fileUrls.length;i++){
									Office.context.mailbox.item.addFileAttachmentAsync(fileUrls[i], "Attachment " + (i + 1).toString(), {}, genericErrorHandler);
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
					Office.context.mailbox.item.bcc.setAsync([SystemEmail], {}, genericErrorHandler);
				});
				
				$('#selMode').change(function(){
					$('.mode-Snippet').hide();
					$('.mode-Template').hide();
					$('.mode-' + $(this).val()).show();
					
					//hide all but relevant categories
					$('.categorySection').each(function(){
					
						$this = $(this);
						$this.hide();
						
						//show categories with visible items
						if($this.find('.mode-' + $('#selMode').val()).length > 0){
							$this.show();
						}
					
					});
					
					
					
				});
				
				$('#selMode').change();
				
				
			}catch(E){
				log('Error: ' + E.toString());
			}

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
					var $category = $(this).parent();
					if($category.hasClass('category-open')){
						$category.removeClass('category-open');
						$category.addClass('category-close');
					}else{
						$category.removeClass('category-close');
						$category.addClass('category-open');

					}
				});
				$('#MainUI').append($cat);
			}else{
				$cat = $cat.first();
			}
			
			//add item to category div
			$cat.append($item);
		
		}
				
		function log(str){
			
				$('#ErrorMessage').append(str);
				$('#Error').show();
				$('#MainUI').hide();
		}
		
		var globalClaimNumber;
		
		function goClaim(){
		
		
			globalClaimNumber = $('#ClaimNumber').val();
			if(! /^[0-9]{9}((-[0-9]{2})?)$/.test(globalClaimNumber)){
						$('#ClaimIDMessage').css('color', 'red');
						return;
			}else{
				$('#StartUI').hide();
				$('#MainUI').show();
				$('#ClaimIDDisplay').text(globalClaimNumber);
				Office.context.mailbox.item.subject.setAsync('Claim#:' + globalClaimNumber, {}, genericErrorHandler);
			}
		}
		
		function resetApp(){
				$('#StartUI').show();
				$('#MainUI').hide();
		}
		
	</script>
	<style>
		#itemTemplate{display:none;}
		#categoryTemplate{display:none;}
		.ms-ListItem{
			margin-left:10px;
		}
		
		
		.ms-ListItem-primaryText{
			padding-right:0px;
			font-size:16px;
		}
		
		.ms-ListItem-secondaryText{
			font-size:12px;
		}

		.ms-ListItem-tertiaryText{
			font-size:12px;
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
			font-size:21px;
		}
		
		.categorySection{
			overflow:hidden;	
		}
		
		
		.ms-ListItem.is-unread{
			border-bottom:1px solid #EEEEEE;
			margin-bottom:2px;
			padding-left:10px;
			padding-top:5px;
		}
		
		
		.smallerFont{
		font-size:.7em;
}

.ms-Button-label{
	padding:5px;
}

.ms-Button{
	height:22px;
	min-width:20px;
	padding:5px;
	padding-top:0px;
	margin-bottom:5px;
}

.mode-Snippet .button-Replace{
	display:none;
}

.mode-Template .button-Insert{
	display:none;
}

	</style>
		
    </head>
    <body class="ms-font-m-plus">
	
		<div id="itemTemplate" class="ms-ListItem is-unread mode-[mode]" style="cursor:pointer;">
		  <span class="ms-ListItem-primaryText">[name]</span>  
		  <span class="ms-ListItem-secondaryText">[date]</span>
		  <span class="ms-ListItem-tertiaryText">[desc]</span>
		  <a href="#" class="ms-Button ms-Button--primary button-Insert" style="text-decoration:none;"><span class="ms-Button-label" style="text-decoration:none;">Insert</span></a>
		  <a href="#" class="ms-Button ms-Button--primary button-Replace" style="text-decoration:none;"><span class="ms-Button-label" style="text-decoration:none;">Replace</span></a>

		</div>
		
		<div id="categoryTemplate" class="categorySection category-open smallerFont">
			<div class="ms-ListItem-primaryText category-title">
				<i class="ms-Icon ms-Icon--caretRight"></i>
				<i class="ms-Icon ms-Icon--caretDownRight"></i>
				[cat]
			</div>
		</div>
		
		<div id="StartUI">
			<div class="ms-TextField" style="float:left;">
				<label class="ms-font-l" style="display: inline-block;">Claim ID:&nbsp;</label>
				<input type="text" class="ms-TextField-field" id="ClaimNumber" style="width: 120px;font-size: larger;"/><br/>
				<span id="ClaimIDMessage" class="ms-TextField-description">Please ensure the claim ID is correct before continuing.</span>
			</div>
			
			<a href="javascript:goClaim();" class="ms-Button ms-Button--primary" style="text-decoration:none;"><span class="ms-Button-label" style="text-decoration:none;">Continue</span></a><p></p>
		</div>

		<div id="MainUI" style="display:none;">
			Claim ID: <span id="ClaimIDDisplay"></span>&nbsp;&nbsp;<a href="javascript:resetApp();" class="ms-Button ms-Button--primary" style="text-decoration:none;"><i class="ms-Icon ms-Icon ms-Icon--reactivate" style="color:white;"></i><span class="ms-Button-label" style="text-decoration:none;">Restart</span></a>
			
			<div>
				Filter By:&nbsp;&nbsp;<select id="selMode">
					<option value="Snippet">Snippets</option>
					<option value="Template">Templates</option>
				</select>
			</div>
			
			<p></p>
		</div>
		
		<div id="Error" style="display:none;">
			<i class="ms-Icon ms-Icon--alert"></i>&nbsp;&nbsp;<span id="ErrorMessage"></span>
		</div>

		
    </body>
</html>