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



'Set the response buffer to true as we are redirecting
Response.Buffer = True


'Dimension variables
Dim rsDeleteComments 	'Database Recordset Variable
Dim laryCommentsID	'Array to hold the ID number for each comment to be deleted


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If

	
'Create a ADO recordset odject for the database
Set rsDeleteComments = Server.CreateObject("ADODB.Recordset")
	 
'Set the Lock Type for the records so that the record set is only locked when it is deleted
rsDeleteComments.LockType = 3

'Run through till all checked comments are deleted
For each laryCommentsID in Request.Form("chkCommentsNo")
	
	'Initalise the strSQL variable with the SQL string
	strSQL = "SELECT tblComments.* FROM tblComments"
	strSQL = strSQL & " WHERE Comment_ID =" & CLng(laryCommentsID)	
	
	'Query the database
	rsDeleteComments.Open strSQL, strCon
	
	'Delete the record is one is returned by the recordset
	If NOT rsDeleteComments.EOF Then rsDeleteComments.Delete
	
	'Close the recordset
	rsDeleteComments.Close
Next
	 
'Reset server variable
Set adoCon = Nothing
Set strCon = Nothing
Set rsDeleteComments = Nothing

'Return to the Select News Item page
Response.Redirect "delete_news_comments_form.asp?NewsID=" & Request.Form("NewsID")
%>
