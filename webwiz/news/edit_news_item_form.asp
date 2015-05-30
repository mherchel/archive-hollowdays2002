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



Dim strMode		'Holds whether the page is to add a new item or amend a news item
Dim lngNewsID		'Holds the ID number of the News Item
Dim rsNews		'Database recordset holding the news items
Dim strNewsTitle	'Holds the title of the news item
Dim strShortNewsItem	'holds the short news item
Dim strNewsItem		'Holds the news item
Dim blnComments		'set to true if the users can leave comments


'Initialise variables
strMode = "edit"


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If


'Create recorset object
Set rsNews = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblNews.* "
strSQL = strSQL & "FROM tblNews "
strSQL = strSQL & "WHERE tblNews.News_ID = " & CLng(Request.QueryString("NewsID")) & ";"
	
'Query the database
rsNews.Open strSQL, adoCon


'If there are records in the recordset then read them in
If NOT rsNews.EOF then
	
	'Read in the values from the recordset
	strNewsTitle = rsNews("News_title")
	strShortNewsItem = rsNews("Short_news")
	strNewsItem = rsNews("News_item")
	lngNewsID = CLng(rsNews("News_ID"))
	blnComments = CBool(rsNews("Comments"))
End If

'Replace HTML new lines with VB new lines in the news item
strNewsItem = Replace(strNewsItem, "<br>", vbCrLf)


'Reset server objects
rsNews.Close
Set rsNews = Nothing
Set strCon = Nothing
Set adoCon = Nothing
%>
<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Edit or Delete News Item</title>

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"><b><font size="6">Edit or Delete News Item</font></b> </div>
<div align="center"><a href="admin_menu.asp" target="_self"> Return to the Site 
  News Administrator Menu</a><br>
  <a href="select_news_item.asp" target="_self">Select another News Item to Edit 
  or Delete</a><br>
  <br>
  <table width="563" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="563" height="66" align="center">To amend the News Item change 
        the details in the form below, or if you wish to delete the News Item 
        click on the button at the bottom of the page.<br>
        <br>
        HTML can be added to the News Item for formatting etc. <br>
        If you are not familiar with HTML you can use the buttons to create the 
        HTML for you that will format your News Item.</td>
    </tr>
  </table>
  
  <br>
</div>
 <div align="center"> 
<% 
'If the browser type selected is IE then have the WYSIWYG editor
If Request.QueryString("browser") = "IE" Then %>
	<!--#include file="advanced_message_form_inc.asp" -->
<% Else %>
	<!--#include file="message_form_inc.asp" -->
<% End If %>
  <form name="frmDelete" method="post" action="delete_news_item.asp" onSubmit="return confirm('Are you sure you want to delete this news item?');">
    <input type="hidden" name="NewsID" value="<% = lngNewsID %>">
    <input type="submit" name="Delete" value="Delete News Item">
  </form>
 <br>
</div>
</body>
</html>