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
Dim rsAddNewsItem		'Database recordset to add the new News Item
Dim strInputNewsTitle 		'Holds the News Title
Dim strInputShortNewsItem	'Holds the short news item
Dim strInputNewsItem 		'Holds the News Item
Dim blnComments			'set to true if users can leave comments on the news item
Dim lngNewsID			'Holds the News item ID number
Dim strMode			'Holds whether the News Item is new or to be updated


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If

'Read in the mode of the page and the News ID number
strMode = Request.Form("mode")
If strMode = "edit" Then lngNewsID = CLng(Request.Form("NewsID"))


'Read in user details from the form
strInputNewsTitle = Request.Form("title")
strInputShortNewsItem = Request.Form("shortNews")
strInputNewsItem = Request.Form("newsItem")
blnComments = CBool(Request.Form("comments"))

'Strip out Norton Internet Security add blocking code that messes up news posts
strInputShortNewsItem = Replace(strInputShortNewsItem, "<SCRIPT> window.open=NS_ActualOpen; </SCRIPT>", "", 1, -1, 1) 
strInputNewsItem = Replace(strInputNewsItem, "<SCRIPT> window.open=NS_ActualOpen; </SCRIPT>", "", 1, -1, 1) 

'If this is not the WYSIWYG editir then format the text
If Request.Form("browser") <> "IE" AND Request.Form("lineBreak") = "true" Then
	
	'Replace the vb new line code for the HTML new break code
	strInputShortNewsItem = Replace(strInputShortNewsItem, vbCrLf, "<br>")
	strInputNewsItem = Replace(strInputNewsItem, vbCrLf, "<br>")
End If
	
'Create recorset object
Set rsAddNewsItem = Server.CreateObject("ADODB.Recordset")

'If the mode is edit then initialise the SQL query to get the Nerws Item to be updated
If strMode = "edit" Then
	strSQL = "SELECT tblNews.* FROM tblNews WHERE tblNews.News_ID = " & lngNewsID & ";"	
Else
	'Initalise the SQL string with a query to read in all the new items from the database
	strSQL = "SELECT tblNews.* FROM tblNews;"
End If

'Set the cursor type property of the record set to Dynamic so we can navigate through the record set
rsAddNewsItem.CursorType = 2

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsAddNewsItem.LockType = 3

'Open the recordset
rsAddNewsItem.Open strSQL, strCon
	
'Add a new record to the recordset if it's a new News Item
If NOT strMode = "edit" Then rsAddNewsItem.AddNew

rsAddNewsItem.Fields("News_title") = strInputNewsTitle 
rsAddNewsItem.Fields("Short_news") = strInputShortNewsItem
rsAddNewsItem.Fields("News_item") = strInputNewsItem
rsAddNewsItem.Fields("Comments") = blnComments
			
'Update the database with the new recordset
rsAddNewsItem.Update

'Requery the database to make sure that the News Item has been deleted
'This will make the script wait until Database has updated itself as sometimes Access can be a little slow at updating
rsAddNewsItem.Requery
	
		 
'Reset Sever Objects 
rsAddNewsItem.Close
Set rsAddNewsItem = Nothing
Set adoCon = Nothing
Set strCon = Nothing


'If this is an update then go back to the select news item page
If strMode = "edit" then Response.Redirect "select_news_item.asp"
%>
<html>
<head>
<title>Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"><b><font size="6">Add News Item</font></b> </div>
<div align="center"><a href="admin_menu.asp" target="_self"> Return to the Site News Administrator Menu</a><br>
</div>
<br>
<br>
<table width="581" border="0" cellspacing="0" cellpadding="1" align="center">
  <tr> 
  <td align="center">You new News Item has been entered into the Database.<br>
   <br>
   <a href="add_news_form.asp<% If Request.Form("browser") = "IE" Then Response.Write("?browser=IE") %>" target="_top">Add another use news item</a></td>
  </tr>
</table>
</body>
</html>