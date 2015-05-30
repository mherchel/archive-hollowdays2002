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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True

'Dimension variables
Dim rsDeleteComments		'Database Recordset holding the comments to be deleted
Dim rsDeleteNewsItem		'Database recordset to delete the News Item
Dim lngNewsID			'Holds the News item ID number


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If


'Read the News ID number
lngNewsID = CLng(Request.Form("NewsID"))


'First we need to delete any comments associated with the News Item so we don't get an error
'Create recorset object
Set rsDeleteComments = Server.CreateObject("ADODB.Recordset")

'Initalise the SQL string with a query to read in all the comments from the database
strSQL = "SELECT tblComments.* FROM tblComments WHERE tblComments.News_ID = " & lngNewsID & ";"

'Set the Lock Type for the records so that the record set is only locked when it is deleted
rsDeleteComments.LockType = 3

'Open the recordset
rsDeleteComments.Open strSQL, strCon
			
'Loop through all the comments for the news item
Do while NOT rsDeleteComments.EOF 
	
	'Delete the Comments
	rsDeleteComments.Delete
	
	'Move to the next record in the recordset
	rsDeleteComments.MoveNext
Loop

'Requery the database to make sure that the coomets have been deleted
'This will make the script wait until Database has updated itself as sometimes Access can be a little slow at updating
rsDeleteComments.Requery


'Now we can delete the News Item	
'Create recorset object
Set rsDeleteNewsItem = Server.CreateObject("ADODB.Recordset")

'Initalise the SQL string with a query to read in all the comments from the database
strSQL = "SELECT tblNews.* FROM tblNews WHERE tblNews.News_ID = " & lngNewsID & ";"

'Set the Lock Type for the records so that the record set is only locked when it is deleted
rsDeleteNewsItem.LockType = 3

'Open the recordset
rsDeleteNewsItem.Open strSQL, strCon
			
'Delete the News Item from the database
If NOT rsDeleteNewsItem.EOF Then rsDeleteNewsItem.Delete

'Requery the database to make sure that the News Item has been deleted
'This will make the script wait until Database has updated itself as sometimes Access can be a little slow at updating
rsDeleteNewsItem.Requery
	
		 
'Reset Sever Objects 
rsDeleteComments.Close
Set rsDeleteComments = Nothing
rsDeleteNewsItem.Close
Set rsDeleteNewsItem = Nothing
Set adoCon = Nothing
Set strCon = Nothing


'Return to the comments page
Response.Redirect "select_news_item.asp"
%>
