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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True

'Dimension variables
Dim rsDeleteComments		'Database Recordset holding the comments to be deleted
Dim rsDeleteJournalItem		'Database recordset to delete the Journal Item
Dim lngJournalID			'Holds the Journal item ID number


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If


'Read the Journal ID number
lngJournalID = CLng(Request.Form("JournalID"))


'First we need to delete any comments associated with the Journal Item so we don't get an error
'Create recorset object
Set rsDeleteComments = Server.CreateObject("ADODB.Recordset")

'Initalise the SQL string with a query to read in all the comments from the database
strSQL = "SELECT tblComments.* FROM tblComments WHERE tblComments.Journal_ID = " & lngJournalID & ";"

'Set the Lock Type for the records so that the record set is only locked when it is deleted
rsDeleteComments.LockType = 3

'Open the recordset
rsDeleteComments.Open strSQL, strCon
			
'Loop through all the comments for the Journal item
Do while NOT rsDeleteComments.EOF 
	
	'Delete the Comments
	rsDeleteComments.Delete
	
	'Move to the next record in the recordset
	rsDeleteComments.MoveNext
Loop

'Requery the database to make sure that the coomets have been deleted
'This will make the script wait until Database has updated itself as sometimes Access can be a little slow at updating
rsDeleteComments.Requery


'Now we can delete the Journal Item	
'Create recorset object
Set rsDeleteJournalItem = Server.CreateObject("ADODB.Recordset")

'Initalise the SQL string with a query to read in all the comments from the database
strSQL = "SELECT tblJournal.* FROM tblJournal WHERE tblJournal.Journal_ID = " & lngJournalID & ";"

'Set the Lock Type for the records so that the record set is only locked when it is deleted
rsDeleteJournalItem.LockType = 3

'Open the recordset
rsDeleteJournalItem.Open strSQL, strCon
			
'Delete the Journal Item from the database
If NOT rsDeleteJournalItem.EOF Then rsDeleteJournalItem.Delete

'Requery the database to make sure that the Journal Item has been deleted
'This will make the script wait until Database has updated itself as sometimes Access can be a little slow at updating
rsDeleteJournalItem.Requery
	
		 
'Reset Sever Objects 
Set rsDeleteComments = Nothing
Set rsDeleteJournalItem = Nothing
Set adoCon = Nothing
Set strCon = Nothing


'Return to the comments page
Response.Redirect "select_Journal_item.asp"
%>
