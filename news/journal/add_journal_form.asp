<% Option Explicit %>
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

Dim strMode		'Holds whether the page is to add a new item or amend a journal item
Dim lngJournalID		'Holds the ID number of the Journal Item
Dim rsJournal		'Database recordset holding the journal items
Dim strAuthor		'Holds the username of the author
Dim strAuthorEmail	'Holds the authors e-mail
Dim strJournalTitle	'Holds the title of the journal item
Dim strShortJournalItem	'holds the short journal item
Dim strJournalItem	'Holds the journal item
Dim blnComments		'set to true if the users can leave comments

'Initialise variables
strMode = "New"


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
<title>Add Journal Item</title>
<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->
<style type="text/css">
<!--
a:link {
	color: #FFFFFF;
}
a:visited {
	color: #FFFFFF;
}
a:hover {
	color: #FFFFFF;
	text-decoration: none;
}
a:active {
	color: #FFFFFF;
}
-->
</style>
</head>
<body bgcolor="#000000" text="#CCCCCC">
<div align="center"></div>
<div align="center"><a href="admin_menu.asp" target="_self"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Return 
  to the Site Journal Administrator Menu</font></a><br>
</div>
 <div align="center"> 
<% 
'If the browser type selected is IE then have the WYSIWYG editor
If Request.QueryString("browser") = "IE" Then %>
	<!--#include file="advanced_message_form_inc.asp" -->
<% Else %>
	<!--#include file="message_form_inc.asp" -->
<% End If %>
 <br>
  <a href="/" target="_parent">home</a>| <a href="/news/" target="_parent">news</a> 
  | <a href="/lyrics/" target="_parent">lyrics</a> | <a href="/shows/" target="_parent">shows</a> 
  | <a href="/press/">press</a> | <a href="/bios" target="_parent">bios</a> 
  | <a href="/mp3s/" target="_parent">mp3s</a> | <a href="/pictures/" target="_parent">pictures</a> 
  | <a href="/contact/" target="_parent">contact</a> 
  | <a href="/list/" target="_parent">mailing list</a> 
  | <a href="/links/" target="_parent">links</a> | <a href="/">&copy;2002 
  Hollow Days</a><font size="2"><br>
  </font><a href="http://www.herchel.com" target="_blank">Website designed by 
  Herchel Services</a> </div>
</body>
</html>