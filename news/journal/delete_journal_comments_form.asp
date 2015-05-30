<% Option Explicit %>
<!--#include file="common.asp" -->
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




Dim rsJournal	'Database recordset holding the Journal items
Dim rsComments	'Database recordset holding the comments for this Journal item
Dim lngJournalID	'Holds the Journal item ID number

'Read in the Journal Item ID number to ge the comments for
lngJournalID = CLng(Request.QueryString("JournalID"))

'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html>
<head>
<title>Delete Journal Item Comments</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->

<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Intialise variables
	var errorMsg = "";
	var errorMsgLong = "";

	//Check for a name
	if (document.frmJournalComments.name.value == ""){
		errorMsg += "\n\tName \t\t- Enter your Name";
	}
	
	//Check for a country
	if (document.frmJournalComments.country.value == ""){
		errorMsg += "\n\tCountry \t\t- Select the country you are in";
	}
	
	//Check for comments
	if (document.frmJournalComments.comments.value == ""){
		errorMsg += "\n\tComments \t- Enter a comment to add to the Guestbook";
	}
	
	//Check for HTML tags before submitting the form	
	for (var count = 0; count <= 7; ++count){
		if ((document.frmJournalComments.elements[count].value.indexOf("<", 0) >= 0) && (document.frmJournalComments.elements[count].value.indexOf(">", 0) >= 0)){
			errorMsgLong += "\n- HTML tags are not permitted, remove all HTML tags.";
		}			
	}
	
	//If there is aproblem with the form then display an error
	if ((errorMsg != "") || (errorMsgLong != "")){
		msg = "___________________________________________________________________\n\n";
		msg += "Your Comments have not been added because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "___________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";
		
		errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
		return false;
	}
	
	return true;
}

// Function to add the code for bold italic and underline, to the message
function AddMessageCode(code,promptText, InsertText) {

	if (code != "") {
		insertCode = prompt(promptText + "\n[" + code + "]xxx[/" + code + "]", InsertText);
			if ((insertCode != null) && (insertCode != "")){
				document.frmJournalComments.comments.value += "[" + code + "]" + insertCode + "[/"+ code + "] ";
			}
	}		
	document.frmJournalComments.comments.focus();
}


//Function to add the code to the message for the smileys
function AddSmileyIcon(iconCode) {	
		document.frmJournalComments.comments.value += iconCode + " ";
		document.frmJournalComments.comments.focus();
}
// -->
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"> <b><font size="6">Delete Journal Item Comments</font></b><br>
  <a href="admin_menu.asp" target="_self">Return to the Site Journal Administrator Menu</a><br>
  <a href="select_journal_item.asp" target="_self">Select Comments for another Journal Item to Delete</a><br>
  <br>
  <table width="563" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="563" height="2" align="center">To delete any of the comments 
        for this Journal Item place a tick in the check box at the top left corner 
        of the comment(s) you wish to delete and click on the Delete Comments 
        at the bottom of the page.</td>
    </tr>
  </table>
  <br>
    <%

'Create recorset object
Set rsJournal = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
strSQL = "SELECT tblJournal.* FROM tblJournal "
strSQL = strSQL & "WHERE tblJournal.Journal_ID = " & lngJournalID
strSQL = strSQL & " ORDER BY Date_stamp DESC;"
	
'Query the database
rsJournal.Open strSQL, adoCon


'If there are no records then exit for loop
If NOT rsJournal.EOF Then
	
	%>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr> 
          <td bgcolor="#CCCCCC"><b><% = rsJournal("Journal_title") %></b></td>
        </tr>
        <tr> 
          
     <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
        <td><% = rsJournal("Journal_item") %></td>
       </tr>
       <tr> 
        <td align="right"><font size="2"><i>Posted by <a href="mailto:<% = rsJournal("Author_email") %>"><% = rsJournal("Author") %></a> on <% = FormatDateTime(rsJournal("Date_stamp"), vbLongDate) %> at <% = FormatDateTime(rsJournal("Date_stamp"), vbShortTime) %></i></font></td>
       </tr>
      </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
	
End If

%>
<form name="frmDelete" method="post" action="delete_journal_comments.asp">
  <br>
<%
'Create recorset object
Set rsComments = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
strSQL = "SELECT tblComments.* FROM tblComments "
strSQL = strSQL & "WHERE tblComments.Journal_ID = " & lngJournalID
strSQL = strSQL & " ORDER BY Date_stamp DESC;"
	
'Query the database
rsComments.Open strSQL, adoCon

'Loop round to display all the comments for the Journal item
Do While NOT rsComments.EOF

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr> 
            <td bgcolor="#CCCCCC"> 
              <input type="checkbox" name="chkCommentsNo" value="<% = rsComments("Comment_ID") %>">
              Comments by <a href="mailto:<% = rsComments("EMail") %>"><% = rsComments("Name") %></a> from <% = rsComments("Country") %> on <% = FormatDateTime(rsComments("Date_stamp"), VbLongDate) %> at <% = FormatDateTime(rsComments("Date_stamp"), VbShortTime) %>
            </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF">
            <% = Replace(rsComments("Comments"), "<img src=""Journal_images", "<img src=""../Journal_images") %>
            </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  
    
  <br>
  <%
	'Move to the next record in the recordset
	rsComments.MoveNext

Loop

'Reset server objects
rsJournal.Close
Set rsJournal = Nothing
Set strCon = Nothing
Set adoCon = Nothing
%>
  <div align="center">
    <input type="hidden" name="JournalID" value="<% = lngJournalID %>">
    <input type="submit" name="Submit" value="Delete Comments">
  </div>
</form>
</body>
</html>