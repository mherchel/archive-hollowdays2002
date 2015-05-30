<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Guide - Web Wiz Journal
'**                                                              
'**  Copyright 2001-2002 Bruce Corkhill All Rights Reserved.                                
'**
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**
'**  You may not pass the whole or any part of this application off as your own work.
'**   
'**  All links to Web Wiz Guide and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**
'**  You should have received a copy of the License along with this program; 
'**  if not, write to:- Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom.
'**    
'**
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.webwizguide.info/forum
'**
'**  Support questions are NOT answered by e-mail ever!
'**
'**  For correspondence or non support questions contact: -
'**  info@webwizguide.com
'**
'**  or at: -
'**
'**  Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom
'**
'****************************************************************************************
%>

<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {
	
	var errorMsg = "";
	
	//Check for an Author
	if (document.frmJournal.author.value==""){
		errorMsg += "\n\tAuthor \t- Enter your name";
	}
	
	//Check for an Journal Title
	if (document.frmJournal.title.value==""){
		errorMsg += "\n\tTitle \t\t- Enter a title for the Journal Item";
	}
	
	//Check for journal Item
	if (document.frmJournal.journalItem.value==""){
		errorMsg += "\n\tJournal Item \t- Enter a Journal Item to post";
	}
	
	//If there is a problem with the form then display an error
	if (errorMsg != ""){
		msg = "____________________________________________________________________\n\n";
		msg += "Your Journal Item has not been submitted because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "____________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";
		
		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	
	return true;
}


//Function to format text in the text box
function FormatText(command, option){
	
  	frames.message.document.execCommand(command, true, option);
  	frames.message.focus();
}

//Function to add image
function AddImage(){	
	imagePath = prompt('Enter the web address of the image', 'http://');				
	
	if ((imagePath != null) && (imagePath != "")){					
		frames.message.document.execCommand('InsertImage', false, imagePath);
  		frames.message.focus();
	}
	frames.message.focus();			
}


//Function to clear form
function ResetForm(){

	if (window.confirm('Are you sure you want to clear the e-mail you have entered?')){
	 	frames.message.document.body.innerHTML = ''; 
	 	return true;
	 } 
	 return false;		
}

//Function to open pop up window
function openWin(theURL,winName,features) {
  	window.open(theURL,winName,features);
}
// -->
</script><body bgcolor="#000000" text="#CCCCCC">
<div align="center">
 <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
  The WYSIWYG HTML editor on this page is for Windows IE5+ users only.<br>
  <br>
  If you are are not using Windows IE5+, or you are having problems using the form below on your web browser then <br>
  <% If strMode = "edit" then 
 	%>
  <a href="edit_journal_item_form.asp?JournalID=<% = Request.QueryString("JournalID") %>">
  <% 
 Else 
 	%>
  </a><a href="add_journal_form.asp">
  <%
 End If 
 %>
  Click here to use the Standard HTML Editor</a><br>
  <br>
  <br>
  Use the formatting buttons to format your e-mail, if you type HTML source code into the text box <strong>it will be <br>
  shown as HTML source code</strong> in the journal item.<br>
  <br>
  You can copy and paste HTML form another web page in IE, but copy the actual page <b>not</b> the source code!</font></p>
 <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">If you are showing the Short Journal items on a page outside of the journal directory then make sure you use full paths,<br>
  including domain name for links and images.</font><br>
 </p>
</div>
<form method=post name="frmJournal" action="add_journal_item.asp" onSubmit="return CheckForm();" onReset="return ResetForm();">
 <table width="660" border="0" cellspacing="0" cellpadding="1" bgcolor="#000000" height="230" align="center">
    <tr> 
      <td height="66" width="967"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF" height="201">
          <tr> 
            <td height="199"> 
       <table width="100%" height="143" border="0" align="center" bgcolor="#000000">
                <tr> 
                  <td colspan="2" height="30" class="text" align="left">*Indicates 
                    required fields</td>
                </tr>
        <tr bgcolor="#FFFFFF" > 
         <td width="14%" height="12" align="right" bgcolor="#000000">Author*:</td>
         <td width="86%" height="12" bgcolor="#000000"> 
          <input type="text" name="author" size="30" maxlength="50" value="<% = strAuthor %>">
         </td>
        </tr>
        <tr bgcolor="#FFFFFF" > 
         <td width="14%" height="12" align="right" bgcolor="#000000">Author Email:</td>
         <td width="86%" height="12" bgcolor="#000000"> 
          <input type="text" name="email" size="30" maxlength="50" value="<% = strAuthorEmail %>">
         </td>
        </tr>
                <tr bgcolor="#FFFFFF" > 
                  <td width="15%" height="12" align="right" bgcolor="#000000">Journal Title*:</td>
                  <td width="86%" height="12" bgcolor="#000000"> <input type="text" name="title" size="30" maxlength="50" value="<% = strJournalTitle %>"> 
                  </td>
                </tr>
                <tr> 
                  <td valign="bottom" align="right" height="22" width="15%"><span class="text">Text 
                    Format:</span></td>
                  <td height="22" width="86%" valign="bottom"> <table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td> <select name="selectText" onChange="FormatText('FontName', selectText.options[selectText.selectedIndex].value);document.frmJournal.selectText.options[0].selected = true;" >
                            <option value="0" selected>-- Font Type --</option>
                            <option value="Arial, Helvetica, sans-serif">Arial</option>
                            <option value="Times New Roman, Times, serif">Times</option>
                            <option value="Courier New, Courier, mono">Courier 
                            New</option>
                            <option value="Verdana, Arial, Helvetica, sans-serif">Verdana</option>
                          </select> <select name="selectFontSize" onChange="FormatText('FontSize', selectFontSize.options[selectFontSize.selectedIndex].value);document.frmJournal.selectFontSize.options[0].selected = true;" >
                            <option value="0" selected>-- Font Size --</option>
                            <option value="1">1</option>
                            <option value="2">2</option>
                            <option value="3">3</option>
                            <option value="4">4</option>
                            <option value="5">5</option>
                            <option value="6">6</option>
                            <option value="7">7</option>
                          </select> <select name="selectFontColour" onChange="FormatText('ForeColor', selectFontColour.options[selectFontColour.selectedIndex].value);document.frmJournal.selectFontColour.options[0].selected = true;" >
                            <option value="0" selected>-- Font Colour --</option>
                            <option value="black">Black</option>
                            <option value="white">White</option>
                            <option value="blue">Blue</option>
                            <option value="red">Red</option>
                            <option value="green">Green</option>
                            <option value="yellow">Yellow</option>
                            <option value="orange">Orange</option>
                            <option value="brown">Brown</option>
                            <option value="magenta">Magenta</option>
                            <option value="cyan">Cyan</option>
                            <option value="limegreen">Lime Green</option>
                          </select> </td>
                      </tr>
                      <tr> 
                        <td><img src="journal_images/post_button_cut.gif" width="25" height="24" align="absmiddle" onClick="FormatText('cut')" style="cursor: hand;" alt="Cut"> 
                          <img src="journal_images/post_button_copy.gif" width="25" height="24" align="absmiddle" onClick="FormatText('copy')" style="cursor: hand;" alt="Copy"> 
                          <img src="journal_images/post_button_paste.gif" width="25" height="24" align="absmiddle" onClick="FormatText('paste')" style="cursor: hand;" alt="Paste"> 
                          <img src="journal_images/post_button_bold.gif" width="25" height="24" align="absmiddle" alt="Bold" onClick="FormatText('bold', '')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_italic.gif" width="25" height="24"  align="absmiddle" alt="Italic" onClick="FormatText('italic', '')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_underline.gif" width="25" height="24" align="absmiddle" alt="Underline" onClick="FormatText('underline', '')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_left_just.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyLeft', '')" style="cursor: hand;" alt="Left Justify"> 
                          <img src="journal_images/post_button_centre.gif" width="25" height="24" align="absmiddle" border="0" alt="Centre Justify" onClick="FormatText('JustifyCenter', '')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_right_just.gif" width="25" height="24" align="absmiddle" onClick="FormatText('JustifyRight', '')" style="cursor: hand;" alt="Right Justify"> 
                          <img src="journal_images/post_button_list.gif" width="25" height="24" align="absmiddle" border="0" alt="Unordered List" onClick="FormatText('InsertUnorderedList', '')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_outdent.gif" width="25" height="24" align="absmiddle" onClick="FormatText('Outdent', '')" style="cursor: hand;" alt="Outdent"> 
                          <img src="journal_images/post_button_indent.gif" width="25" height="24" align="absmiddle" border="0" alt="Indent" onClick="FormatText('indent', '')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_hyperlink.gif" width="25" height="24" align="absmiddle" border="0" alt="Add Hyperlink" onClick="FormatText('createLink')" style="cursor: hand;"> 
                          <img src="journal_images/post_button_image.gif" width="25" height="24" align="absmiddle" border="0" alt="Add Image" onClick="AddImage()" style="cursor: hand;"> 
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr > 
                  <td valign="top" align="right" height="61" width="15%" ><span class="text">Journal 
                    Item*:</span></td>
                  <td height="61" width="86%" valign="top"> 
                    <%
'This bit creates a random number to add to the end of the iframe link as IE will cashe the page
'Randomise the system timer
Randomize Timer
%>
                    <script language="javascript">
		    
		    	//Create an iframe and turn on the design mode for it
		    	document.write ('<iframe src="adv_message_textbox.asp?NoCache=<% = CInt(RND * 2000) %><% If strMode = "edit" Then Response.Write("&mode=edit&JournalItem=" & CLng(Request.QueryString("JournalID"))) %>" id="message" width="510" height="200"></iframe>')
                    	frames.message.document.designMode = "On";
                  
                    </script> 
                    <!-- Display a message for IE users with JavaScript turned off -->
                    <noscript>
                    <br>
                    <br>
                    JavaScript must be enabled on your web browser for you to 
                    you the WYSIWYG e-mail editor!</noscript> </td>
                </tr>
                <tr> 
                  <td valign="top" align="right" height="2" >&nbsp;</td>
                  <td height="2" align="left"><input name="comments" type="checkbox" id="comments" value="True"<% If blnComments = True Then Response.Write(" checked") %>>
                    Allow users to post comments on this journal item</td>
                </tr>
                <td valign="top" align="right" height="2" width="15%" > <input type="hidden" name="mode" value="<% = strMode %>"> 
                  <input type="hidden" name="JournalID" value="<% = lngJournalID %>"> 
                  <input name="shortJournal" type="hidden" id="shortJournal" value=""> 
                  <input type="hidden" name="journalItem" value=""> <input type="hidden" name="browser" value="IE"> 
                </td>
                <td height="2" width="86%" align="left"> 
                  <%
                            If strMode="edit" Then
                            %>
                  <input type="submit" name="Submit" value="Update Journal Item" OnClick="document.frmJournal.journalItem.value = frames.message.document.body.innerHTML;"> 
                  <%
                            Else
                            %>
                  <input type="submit" name="Submit" value="Add Journal Item" OnClick="document.frmJournal.journalItem.value = frames.message.document.body.innerHTML;"> 
                  <%
                            End If
                            %>
                  <input type="reset" name="Reset" value="Reset Form" OnClick="document.frmJournal.journalItem.value = frames.message.document.body.innerHTML;"> 
                </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <br>
  
 <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Please Note</font></b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">, 
  that images need to have the full URL to them and must be available on your site as they will <br>
  not be sutomactically uploaded by the this application.</font></div>
</form>