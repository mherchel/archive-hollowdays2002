<% Option Explicit %>
<!--#include file="common.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Guide - Web Wiz Site News
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




'Set the response buffer to true
Response.Buffer = True 


'Dimension variables
Dim rsNewsConfig 		'Recorset holding all the username in the database				
Dim strMode			'holds the mode of the page, set to true if changes are to be made to the database
Dim strUsername			'Holds the admin username
Dim strPassword			'Holds the admin password
Dim intPreviewNewsItems		'Holds the number of preview items to display
    
      
'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If    


'Read in the users colours for the site news administrator
strUsername = Request.Form("username")
strPassword = Request.Form("password")
strTitleImage = Request.Form("titleImage")
intRecordsPerPage = CInt(Request.Form("RecPerPage"))
strBgColour = Request.Form("bg")
strTextColour = Request.Form("text")
strTextType = Request.Form("FontType")
intHeadingTextSize = CInt(Request.Form("HeadingFontSize"))
intTextSize = CInt(Request.Form("FontSize"))
intSmallTextSize = CInt(Request.Form("SmallFontSize")) 
strTableColour = Request.Form("table")
strTableBorderColour = Request.Form("tableBorder")
strTableTitleColour = Request.Form("tableTitle")
strLinkColour = Request.Form("links")
strVisitedLinkColour = Request.Form("vLinks")
strActiveLinkColour = Request.Form("aLinks")
intMsgCharNo = CInt(Request.Form("CharNo"))
blnCookieSet = CBool(Request.Form("Cookies"))
blnIPBlocking = CBool(Request.Form("IP"))
intPreviewNewsItems = CInt(Request.Form("PrePerPage"))
strMode = Request.Form("mode")


'Intialise the ADO recordset object
Set rsNewsConfig  = Server.CreateObject("ADODB.Recordset")
	

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblConfiguration.* From tblConfiguration;"

'Set the cursor type property of the record set to Dynamic so we can navigate through the record set
rsNewsConfig.CursorType = 2

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsNewsConfig.LockType = 3
	
'Query the database
rsNewsConfig.Open strSQL, strCon

'If the user is changing tthe colours then update the database
If strMode = "change" Then

	
	'Update the recordset	
	rsNewsConfig.Fields("username") = strUsername
	rsNewsConfig.Fields("password") = strPassword
	rsNewsConfig.Fields("No_records_per_page") = intRecordsPerPage	
	rsNewsConfig.Fields("bg_colour") = strBgColour
	rsNewsConfig.Fields("text_colour") = strTextColour
	rsNewsConfig.Fields("text_type") = strTextType
	rsNewsConfig.Fields("heading_text_size") = intHeadingTextSize
	rsNewsConfig.Fields("text_size") = intTextSize
	rsNewsConfig.Fields("small_text_size") = intSmallTextSize
	rsNewsConfig.Fields("table_colour") = strTableColour
	rsNewsConfig.Fields("table_border_colour") = strTableBorderColour
	rsNewsConfig.Fields("table_title_colour") = strTableTitleColour
	rsNewsConfig.Fields("links_colour") = strLinkColour
	rsNewsConfig.Fields("visited_links_colour") = strVisitedLinkColour
	rsNewsConfig.Fields("active_links_colour") = strActiveLinkColour
	rsNewsConfig.Fields("No_of_preview_items") = intPreviewNewsItems
	rsNewsConfig.Fields("Title_image") = strTitleImage
	rsNewsConfig.Fields("Message_char_no") = intMsgCharNo
	rsNewsConfig.Fields("Cookie") = blnCookieSet
	rsNewsConfig.Fields("IP_blocking") = blnIPBlocking
			
				
	'Update the database with the new user's colours
	rsNewsConfig.Update
		
	'Re-run the query to read in the updated recordset from the database
	rsNewsConfig.Requery	
End If

'Read in the site news colours from the database
If NOT rsNewsConfig.EOF Then
	
	'Read in the colour info from the database
	strUsername = rsNewsConfig.Fields("username")
	strPassword = rsNewsConfig.Fields("password")
	intRecordsPerPage = rsNewsConfig.Fields("No_records_per_page")	
	strBgColour = rsNewsConfig.Fields("bg_colour")
	strTextColour = rsNewsConfig.Fields("text_colour")
	strTextType = rsNewsConfig.Fields("text_type")
	intHeadingTextSize = CInt(rsNewsConfig.Fields("heading_text_size"))
	intTextSize = CInt(rsNewsConfig.Fields("text_size"))
	intSmallTextSize = CInt(rsNewsConfig.Fields("small_text_size"))
	strTableColour = rsNewsConfig.Fields("table_colour")
	strTableBorderColour = rsNewsConfig.Fields("table_border_colour")
	strTableTitleColour = rsNewsConfig.Fields("table_title_colour")
	strLinkColour = rsNewsConfig.Fields("links_colour")
	strVisitedLinkColour = rsNewsConfig.Fields("visited_links_colour")
	strActiveLinkColour = rsNewsConfig.Fields("active_links_colour")
	intPreviewNewsItems = Cint(rsNewsConfig.Fields("No_of_preview_items"))
	strTitleImage = rsNewsConfig.Fields("Title_image")
	intMsgCharNo = CInt(rsNewsConfig.Fields("Message_char_no"))
	blnCookieSet = CBool(rsNewsConfig.Fields("Cookie"))
	blnIPBlocking = CBool(rsNewsConfig.Fields("IP_blocking"))
End If


'Reset Server Objects
Set adoCon = Nothing
Set strCon = Nothing
Set rsNewsConfig = Nothing


%>  
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Configure Site News</title>

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
		
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm() {

	var errorMsg = "";
	
	//Check for a username
	if (document.frmColours.username.value==""){
		errorMsg += "\n\tUsername \t- Enter a Username to use the Admin pages with";	
	}
	
	//Check for a password
	if (document.frmColours.password.value==""){
		errorMsg += "\n\tPassword \t- Enter a Password to use the Admin pages with";
	}
	
	//Check for a background colour
	if (document.frmColours.bg.value==""){
		errorMsg += "\n\tBackground \t- Enter a Background Colour";
	}
	
	//Check for a text colour
	if (document.frmColours.text.value==""){
		errorMsg += "\n\tText \t\t- Enter a Text Colour";
	}
	
	//Check for a Table Background Colour
	if (document.frmColours.table.value==""){
		errorMsg += "\n\tTable Background \t- Enter a Table Background Colour";
	}
	
	//Check for a Table Title Background Colour
	if (document.frmColours.tableTitle.value==""){
		errorMsg += "\n\tTbale Title \t- Enter a Table Title Background Colour";
	}
	
	//Check for a Table Border Colour
	if (document.frmColours.tableBorder.value==""){
		errorMsg += "\n\tTable Border \t- Enter a Table Border Colour";
	}
	
	//Check for a Links
	if (document.frmColours.links.value==""){
		errorMsg += "\n\tLink Colour \t- Enter a Link Colour";
	}
	
	//Check for a Visited Links Colour
	if (document.frmColours.vLinks.value==""){
		errorMsg += "\n\tVisited Link \t- Enter a Visited Links Colour";
	}
	
	//Check for a Active Links Colour
	if (document.frmColours.aLinks.value==""){
		errorMsg += "\n\tMouse Over Link \t- Enter a Mouse Over Link Colour";
	}	
	
	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "________________________________________________________________\n\n";
		msg += "The form has not been submitted because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";
		
		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	
	return true;
}
// -->
</script>
 
<style type="text/css">
<!--
.heading {font-family: <% = strTextType %>; font-size: <% = intHeadingTextSize %>px; color: <% = strTextColour %> font-weight: bold;}
.text {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strTextColour %>}
.smText {font-family: <% = strTextType %>; font-size: <% = intSmallTextSize %>px; color: <% = strTextColour %>}
a {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strLinkColour %>}
a:hover {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strActiveLinkColour %>}
a:visited {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strVisitedLinkColour %>}
a:visited:hover {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strActiveLinkColour %>}
-->
</style>
     	
</head>
<body bgcolor="#FFFFFF" text="#000000">
<h1 align="center"><font face="Arial, Helvetica, sans-serif">Configure Site News</font></h1>
<div align="center"><a href="admin_menu.asp" target="_self"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Return to the Site News Menu</font></a><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
 <br>
 </font> 
 <table width="645" border="0" cellspacing="0" cellpadding="1">
  <tr> 
   <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">It is highly recommended that you change the Username and Password, but don't forget what you change them to as you will <b>NOT</b> be able to Administer the Site News without them.</font></td>
  </tr>
 </table>
  <br>
  <br>
  
 <table width="97%" border="0" cellspacing="1" cellpadding="8" align="center" height="157">
  <tr>
      
   <td height="101" align="center" bgcolor="<% = strBgColour %>"><span class="text">Background</span> Colour<br>
    <br>
    <table width="98%" border="0" cellspacing="1" cellpadding="4" bgcolor="<% = strTableBorderColour %>" align="center">
     <tr> 
      <td bgcolor="<% = strTableColour %>"> <span class="text">Homepage Include File Table Colours<br>
       <br>
       <a href="configure_site_news.asp" target="_self">Links</a></span><span class="smText"> - <em>Small Text</em></span><span class="text"><br>
       Normal Text</span></td>
     </tr>
    </table>
        
    <br>
    <table width="98%" border="0" cellspacing="1" cellpadding="4" align="center">
     <tr> 
      <td><span class="heading">Heading</span><br> <span class="text">Normal Text</span><br> <span class="smText">Small Text</span><br> <a href="configure_site_news.asp" target="_self">Links</a></td>
     </tr>
    </table>
    <br>
    <table width="85%" border="0" cellspacing="0" cellpadding="1" bgcolor="<% = strTableBorderColour %>">
     <tr> 
      <td align="center" height="27"> <table width="100%" border="0" cellspacing="0" cellpadding="0" height="30">
        <tr> 
         <td bgcolor="<% = strTableTitleColour %>" align="center" class="text"><font color="<% = strTextColour %>">Comments Form Colour</font></td>
        </tr>
       </table></td>
     </tr>
    </table> </td>
    </tr>
  </table>
</div>
<form method="post" name="frmColours" action="configure_site_news.asp" onSubmit="return CheckForm();">
 <table width="550" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000">
  <tr>
   <td><table border="0" align="center" cellpadding="4" cellspacing="1">
     <tr align="left" bgcolor="#CCCCCC"> 
      <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">*Indicates required fields</font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td width="318" align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Admin Username:*</font></td>
      <td width="235" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="username" maxlength="20" value="<% = strUsername %>">
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Admin Password:*</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="password" maxlength="20" value="<% = strPassword %>">
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Site News Title Image Location</font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
       <font size="1">This is the image shown at the top of each page of the Site News eg. Your web site logo</font></font></td>
      <td valign="top"><input type="text" name="titleImage" maxlength="65" value="<% = strTitleImage %>" size="35"></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Background Colour*</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="bg" maxlength="10" value="<% = strBgColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Text Colour*</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="text" maxlength="10" value="<% = strTextColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Font Style</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <select name="FontType">
        <option value="Arial, Helvetica, sans-serif" <% If strTextType = "Arial, Helvetica, sans-serif" Then Response.Write("selected") %>>Arial, Helvetica, sans-serif</option>
        <option value="Times New Roman, Times, serif" <% If strTextType = "Times New Roman, Times, serif" Then Response.Write("selected") %>>Times New Roman, Times, serif</option>
        <option value="Courier New, Courier, mono" <% If strTextType = "Courier New, Courier, mono" Then Response.Write("selected") %>>Courier New, Courier, mono</option>
        <option value="Georgia, Times New Roman, Times, serif" <% If strTextType = "Georgia, Times New Roman, Times, serif" Then Response.Write("selected") %>>Georgia, Times New Roman, Times, serif</option>
        <option value="Verdana, Arial, Helvetica, sans-serif" <% If strTextType = "Verdana, Arial, Helvetica, sans-serif" Then Response.Write("selected") %>>Verdana, Arial, Helvetica, sans-serif</option>
        <option value="Geneva, Arial, Helvetica, san-serif" <% If strTextType = "Geneva, Arial, Helvetica, san-serif" Then Response.Write("selected") %>>Geneva, Arial, Helvetica, san-serif</option>
       </select>
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Heading Font Size</font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <select name="HeadingFontSize">
        <option value="10" <% If intHeadingTextSize = 10 Then Response.Write("selected") %>>10</option>
        <option value="11" <% If intHeadingTextSize = 11 Then Response.Write("selected") %>>11</option>
        <option value="12" <% If intHeadingTextSize = 12 Then Response.Write("selected") %>>12</option>
        <option value="13" <% If intHeadingTextSize = 13 Then Response.Write("selected") %>>13</option>
        <option value="14" <% If intHeadingTextSize = 14 Then Response.Write("selected") %>>14</option>
        <option value="15" <% If intHeadingTextSize = 15 Then Response.Write("selected") %>>15</option>
        <option value="16" <% If intHeadingTextSize = 16 Then Response.Write("selected") %>>16</option>
        <option value="17" <% If intHeadingTextSize = 17 Then Response.Write("selected") %>>17</option>
        <option value="18" <% If intHeadingTextSize = 18 Then Response.Write("selected") %>>18</option>
        <option value="19" <% If intHeadingTextSize = 19 Then Response.Write("selected") %>>19</option>
        <option value="20" <% If intHeadingTextSize = 20 Then Response.Write("selected") %>>20</option>
        <option value="21" <% If intHeadingTextSize = 21 Then Response.Write("selected") %>>21</option>
       </select>
       pixels</font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Normal Font Size</font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <select name="FontSize">
        <option value="10" <% If intTextSize = 10 Then Response.Write("selected") %>>10</option>
        <option value="11" <% If intTextSize = 11 Then Response.Write("selected") %>>11</option>
        <option value="12" <% If intTextSize = 12 Then Response.Write("selected") %>>12</option>
        <option value="13" <% If intTextSize = 13 Then Response.Write("selected") %>>13</option>
        <option value="14" <% If intTextSize = 14 Then Response.Write("selected") %>>14</option>
       </select>
       pixels</font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Small Font Size</font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <select name="SmallFontSize">
        <option value="9" <% If intSmallTextSize = 9 Then Response.Write("selected") %>>9</option>
        <option value="10" <% If intSmallTextSize = 10 Then Response.Write("selected") %>>10</option>
        <option value="11" <% If intSmallTextSize = 11 Then Response.Write("selected") %>>11</option>
        <option value="12" <% If intSmallTextSize = 12 Then Response.Write("selected") %>>12</option>
        <option value="13" <% If intSmallTextSize = 13 Then Response.Write("selected") %>>13</option>
        <option value="14" <% If intSmallTextSize = 14 Then Response.Write("selected") %>>14</option>
       </select>
       pixels</font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Table Background Colour* </font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="table" maxlength="10" value="<% = strTableColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Table Comments Background Colour* </font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="tableTitle" maxlength="10" value="<% = strTableTitleColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Table Border Colour*</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="tableBorder" maxlength="10" value="<% = strTableBorderColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Links*</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="links" maxlength="10" value="<% = strLinkColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Visited Links*</font></td>
      <td valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="vLinks" maxlength="10" value="<% = strVisitedLinkColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Mouse Over Link Colour*</font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <input type="text" name="aLinks" maxlength="10" value="<% = strActiveLinkColour %>" size="10" >
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Number of News Items Per Page:</font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <select name="RecPerPage">
        <option value="1" <% If intRecordsPerPage = 1 Then Response.Write("selected") %>>1</option>
        <option value="2" <% If intRecordsPerPage = 2 Then Response.Write("selected") %>>2</option>
        <option value="3" <% If intRecordsPerPage = 3 Then Response.Write("selected") %>>3</option>
        <option value="4" <% If intRecordsPerPage = 4 Then Response.Write("selected") %>>4</option>
        <option value="5" <% If intRecordsPerPage = 5 Then Response.Write("selected") %>>5</option>
        <option value="6" <% If intRecordsPerPage = 6 Then Response.Write("selected") %>>6</option>
        <option value="7" <% If intRecordsPerPage = 7 Then Response.Write("selected") %>>7</option>
        <option value="8" <% If intRecordsPerPage = 8 Then Response.Write("selected") %>>8</option>
        <option value="9" <% If intRecordsPerPage = 9 Then Response.Write("selected") %>>9</option>
        <option value="10" <% If intRecordsPerPage = 10 Then Response.Write("selected") %>>10</option>
        <option value="12" <% If intRecordsPerPage = 12 Then Response.Write("selected") %>>12</option>
        <option value="14" <% If intRecordsPerPage = 14 Then Response.Write("selected") %>>14</option>
        <option value="18" <% If intRecordsPerPage = 18 Then Response.Write("selected") %>>18</option>
        <option value="20" <% If intRecordsPerPage = 20 Then Response.Write("selected") %>>20</option>
       </select>
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Preview News Items:</font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
       <font size="1">This is the number of News Items shown on the homepage integration file</font></font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
       <select name="PrePerPage" id="PrePerPage">
        <option value="1" <% If intPreviewNewsItems = 1 Then Response.Write("selected") %>>1</option>
        <option value="2" <% If intPreviewNewsItems = 2 Then Response.Write("selected") %>>2</option>
        <option value="3" <% If intPreviewNewsItems = 3 Then Response.Write("selected") %>>3</option>
        <option value="4" <% If intPreviewNewsItems = 4 Then Response.Write("selected") %>>4</option>
        <option value="5" <% If intPreviewNewsItems = 5 Then Response.Write("selected") %>>5</option>
        <option value="6" <% If intPreviewNewsItems = 6 Then Response.Write("selected") %>>6</option>
        <option value="7" <% If intPreviewNewsItems = 7 Then Response.Write("selected") %>>7</option>
        <option value="8" <% If intPreviewNewsItems = 8 Then Response.Write("selected") %>>8</option>
        <option value="9" <% If intPreviewNewsItems = 9 Then Response.Write("selected") %>>9</option>
        <option value="10" <% If intPreviewNewsItems = 10 Then Response.Write("selected") %>>10</option>
        <option value="12" <% If intPreviewNewsItems = 12 Then Response.Write("selected") %>>12</option>
        <option value="14" <% If intPreviewNewsItems = 14 Then Response.Write("selected") %>>14</option>
        <option value="18" <% If intPreviewNewsItems = 18 Then Response.Write("selected") %>>18</option>
        <option value="20" <% If intPreviewNewsItems = 20 Then Response.Write("selected") %>>20</option>
       </select>
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF">
      <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Maximum Number of Characters per Comment</font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
       <font size="1">This is the maximum allowed amount of comments allowed in comments posted in the News</font></font></td>
      <td valign="top"><select name="CharNo">
        <option <% If intMsgCharNo = 150 Then Response.Write("selected") %>>150</option>
        <option <% If intMsgCharNo = 175 Then Response.Write("selected") %>>175</option>
        <option <% If intMsgCharNo = 200 Then Response.Write("selected") %>>200</option>
        <option <% If intMsgCharNo = 250 Then Response.Write("selected") %>>250</option>
        <option <% If intMsgCharNo = 500 Then Response.Write("selected") %>>500</option>
        <option <% If intMsgCharNo = 750 Then Response.Write("selected") %>>750</option>
        <option <% If intMsgCharNo = 1000 Then Response.Write("selected") %>>1000</option>
        <option <% If intMsgCharNo = 2000 Then Response.Write("selected") %>>2000</option>
        <option <% If intMsgCharNo = 5000 Then Response.Write("selected") %>>5000</option>
        <option <% If intMsgCharNo = 10000 Then Response.Write("selected") %>>10,000</option>
       </select></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Anti-Spam Cookies</font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
       <font size="1">This will mean a cookie is set on the uesers machine to stop them spamming News Items with multiple comments.</font></font></td>
      <td valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">On 
       <input type="radio" name="Cookies" value="True" <% If blnCookieSet = True Then Response.Write "checked" %>>
       &nbsp;&nbsp;&nbsp;Off 
       <input type="radio" name="Cookies" value="False" <% If blnCookieSet = False Then Response.Write "checked" %>>
       </font></td>
     </tr>
     <tr  bgcolor="#FFFFFF"> 
      <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Anti-Spam IP Blocking<br>
       <font size="1">This will mean the IP addrees of the user is checked to stop them spamming News Items with multiple comments.</font><br>
       </font></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">On 
       <input type="radio" name="IP" value="True" <% If blnIPBlocking = True Then Response.Write "checked" %>>
       &nbsp;&nbsp;&nbsp;Off 
       <input type="radio" name="IP" value="False" <% If blnIPBlocking = False Then Response.Write "checked" %>>
       </font></td>
     </tr>
     <tr bgcolor="#FFFFFF" align="center"> 
      <td valign="top" colspan="2" > <p> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="hidden" name="mode" value="change">
        <input type="submit" name="Submit" value="Update Site News Configuration">
        <input type="reset" name="Reset" value="Clear Form">
        </font></p></td>
     </tr>
    </table></td>
  </tr>
 </table>
</form>
<br>
</body>
</html>
