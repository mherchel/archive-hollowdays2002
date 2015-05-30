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

Dim lngNewsID		'Holds the ID number of the News Item
Dim rsNews		'Database recordset holding the news items
Dim strAuthor		'Holds the username of the author
Dim strNewsItem		'Holds the news item



'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If

'If thiss is editing a news item then get it from the database
If Request.QueryString("mode") = "edit" Then
	'Create recorset object
	Set rsNews = Server.CreateObject("ADODB.Recordset")
		
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT tblNews.News_item, tblNews.Short_news "
	strSQL = strSQL & "FROM tblNews "
	strSQL = strSQL & "WHERE tblNews.News_ID = " & CLng(Request.QueryString("NewsItem")) & ";"
		
	'Query the database
	rsNews.Open strSQL, adoCon
	
	'Get teh short news item
	If NOT rsNews.EOF AND Request.QueryString("item") = "short" Then strNewsItem = rsNews("Short_news")
	
	'Get the long news item
	If NOT rsNews.EOF AND Request.QueryString("item") = "long" Then strNewsItem = rsNews("News_item")
	
	'Close recordset
	rsNews.Close
	Set rsNews = Nothing
End If

'Reset Sever Objects 
Set adoCon = Nothing
Set strCon = Nothing
%>
<html>
<head>
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

</head>
<body bgcolor="#FFFFFF" text="#000000" class="text" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<% = strNewsItem %>
</body>
</html>