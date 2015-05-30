<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="send_mail_functions_inc.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Guide - Web Wiz Mailing List
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


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True 

'Declare variables
Dim rsNewListMember	'Database recordset variable
Dim rsRemoveListMember	'Database recordset variable
Dim strEmailAddress	'Holds the users e-mail address
Dim strEmailBody	'Holds the body of the e-mail
Dim strAppendToEmail	'Holds the link to get removed from the mailing list
Dim strSubject		'Holds the subject of the e-mail
Dim strMessage		'Holds the error message if the user is not entered into the database
Dim strUserCode		'Holds a unique code for the new list member
Dim strMode		'Holds whether the user is subscribing or un-subscribing
Dim blnError		'Set to true if the user is not enetered into the mailing list
Dim blnSendWelcome	'Set to ture when the welcome e-mail is sent

'Initialise variables
blnError = False

'Read in the form details
strEmailAddress = LCase(Request.Form("email"))
strMode = Request.Form("mode")

'Clean up the email address address getting rid of unwated characters
strEmailAddress = characterStrip(strEmailAddress)


'Check to see if the user has entered an e-mail address and that it is a valid address
If Len(strEmailAddress) < 5 OR NOT Instr(1, strEmailAddress, " ") = 0 OR InStr(1, strEmailAddress, "@", 1) < 2 OR InStrRev(strEmailAddress, ".") < InStr(1, strEmailAddress, "@", 1) Then
	
	'Set an error message if the users has not enetered a valid e-mail address
	strMessage = strMessage & "You must enter a valid e-mail address"
	
	'Set the error boolean to true
	blnError = True
End If	

'Slect the mode of the page whether we are subscribing or un-subscribing
Select Case strMode 
	
	'If the user wants to subscribe then add them to the list
	Case "add"

		'Create recordset object
		Set rsNewListMember = Server.CreateObject("ADODB.Recordset")
			
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT tblMailingList.* FROM tblMailingList;"
		
		'Set the cursor type property of the record set to Dynamic so we can navigate through the record set
		rsNewListMember.CursorType = 2
			
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsNewListMember.LockType = 3
			
		'Query the database
		rsNewListMember.Open strSQL, adoCon
		
		'Randomise the system timer
		Randomize Timer
				
		'Calculate a code for the user
		strUserCode = Left(strEmailAddress,2) & (9876989856 * CInt((RND * 32000) + 100))
		
		'Loop through all the records in the recordset to check that the user id and the e-mail address are not already in the database
		Do While NOT rsNewListMember.EOF	
			
			'If there is no user code or it is already in the database make a new one and serch the recordset from the begining again
			If strUserCode = rsNewListMember("ID_Code") Then
			
				'Randomise the system timer
				Randomize Timer
				
				'Calculate a code for the user
				strUserCode = Left(strEmailAddress,2) & (9876989856 * CInt((RND * 32000) + 100))
			
				'Move to the first record to make sure the new user code is not in the database
				rsNewListMember.MoveFirst
			End If	
			
			'If the e-mail address is already in the database then create an error message and exit loop
			If strEmailAddress = rsNewListMember("Email") Then
				
				'Create an error message
				strMessage = strMessage & "Your e-mail address is already in the mailing list"
			
				'Set the error boolean to true
				blnError = True
				
				'Exit the for loop
				Exit Do
			End If
			
			'Move to the next record in the recordset
			rsNewListMember.MoveNext	
		Loop	
		
		'If there is no error message then add the new user to the database
		If blnError = False Then
			
			'Add new record to a new recorset
			rsNewListMember.AddNew
			
			rsNewListMember.Fields("Email") = strEmailAddress
			rsNewListMember.Fields("ID_Code") = strUserCode
			
			rsNewListMember.Update
			
			'Set the suject of the e-mail thanking teh user for joining and the -mail body
			strSubject = "Thank you for joining " & strWebsiteName & "'s mailing list"
			
			'If there is a welcome message then send it to the new mailing list member
			If strWelcomeMessage <> "" Then
				'set the message body of th e-mail
				strEmailBody = strWelcomeMessage
				
				'Create email object
				Call createMailObject(strMailComponent)
				
				'Write a remove from mailing list message to add to the end of the e-mail in HTML Format
				strAppendToEmail = mailBody(strWelcomeFormat, strUserCode, blnLCode)
				
				'Send the email
				Call SendMail(strEmailAddress, strMailComponent, strWelcomeFormat)
			
				'Drop email component
				Call dropMailObject(strMailComponent)
			End If
			
			'Set the message to thank the user for joining
			strMessage = strSubject & ".<br>"
		End If
		
		'Reset recordset variable
		rsNewListMember.Close
		Set rsNewListMember = Nothing

	
	'If the mode is to delete then un-subscribe the user
	Case "delete"
	
		'Create recorset object
		Set rsRemoveListMember = Server.CreateObject("ADODB.Recordset")
			
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT tblMailingList.* FROM tblMailingList WHERE tblMailingList.Email = '" & strEmailAddress & "';"
		
		'Set the cursor type property of the record set to Dynamic so we can navigate through the record set
		rsRemoveListMember.CursorType = 2
			
		'Set the Lock Type for the records so that the record set is only locked when it is deleteed
		rsRemoveListMember.LockType = 3
			
		'Query the database
		rsRemoveListMember.Open strSQL, adoCon
		
		
		'If there is no error message then add the new user to the database
		If NOT rsRemoveListMember.EOF Then
			
			'Delete the record
			rsRemoveListMember.Delete
			
			'Create a mesage for the user
			strMessage = "You have been removed from " & strWebsiteName & "'s mailing list.<br>Sorry for any inconvenience. "
		Else
			'As there is no record returned we need an error message
			strMessage = "The e-mail address enetered did not match any in the database.<br>You have not been removed from " & strWebsiteName & "'s mailing list."			
		End If
		
		'Reset recordset variable
		rsRemoveListMember.Close
		Set rsRemoveListMember = Nothing
End Select

'Reset server objects
Set strCon = Nothing
Set adoCon = Nothing
%>
<html>
<head>
<title>Mailing List</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!-- The Web Wiz Guide - Web Wiz Mailing List is written and produced by Bruce Corkhill ©2001-2002
     	 If you want your own ASP Mailing List then goto http://www.webwizguide.info -->

<style type="text/css">
<!--
.text {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strTextColour %>}
.bold {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strTextColour %>; font-weight: bold;}
.heading {font-family: <% = strTextType %>; font-size: 24px; color: <% = strTextColour %>; font-weight:bold;}
a {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strLinkColour %>}
a:hover {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strHoverLinkColour %>}
a:visited {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strVisitedLinkColour %>}
a:visited:hover {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strHoverLinkColour %>}
-->
</style>
</head>
<body bgcolor="<% = strBgColour %>">
<div align="center">
  <span class="heading"><% = strWebsiteName %>'s Mailing List</span><br><br>
  <span class="text"><%
'If there is a message for the user then display it
If IsEmpty(strMessage) = False Then
	
	'Write the message for the user
	Response.Write(strMessage)
Else
  %>
  <br>
  An unspecified error has occurred.<br>Please try again.<br>
<%
End If
%></span>
  <br>
  <br>
  <a href="JavaScript:onClick=window.close()">Close Window</a> <br>
  <br>
  <br>
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:11px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:11px"">Web Wiz Mailing List</a> version 3.02</span>")
	Response.Write("<br><span class=""text"" style=""font-size:11px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
 %>
</div>
</body>
</html>
