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



'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html>
<head>
<title>Hollow Days News Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
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
<p align="center"><img name="index" src="../index.jpg" width="756" height="200" border="0" usemap="#m_index" alt=""> 
  <map name="m_index">
    <area shape="rect" coords="690,178,747,197" href="/links/" target="_parent" alt="" >
    <area shape="rect" coords="580,179,686,196" href="/list/" target="_parent" alt="" >
    <area shape="rect" coords="504,177,578,200" href="/contact/" target="_parent" alt="" >
    <area shape="rect" coords="424,177,501,200" href="/pictures/" target="_parent" alt="" >
    <area shape="rect" coords="361,177,419,200" href="/mp3s/" target="_parent" alt="" >
    <area shape="rect" coords="311,178,353,200" href="/bios/" target="_parent" alt="" >
    <area shape="rect" coords="249,177,304,200" href="/press/" target="_parent" alt="" >
    <area shape="rect" coords="180,179,247,197" href="/shows/" target="_parent" alt="" >
    <area shape="rect" coords="125,175,178,200" href="/lyrics/" target="_parent" alt="" >
    <area shape="rect" coords="70,176,117,200" href="/news/" target="_parent" alt="" >
    <area shape="rect" coords="0,179,62,196" href="/" target="_parent" alt="" >
    <area shape="rect" coords="8,9,336,152" href="/" target="_parent">
  </map>
</p>
<table width="531" border="0" cellspacing="0" cellpadding="0" align="center">
  <!--DWLayoutTable-->
  <tr> 
    <td width="531" height="16"></td>
  </tr>
  <tr> 
    <td height="48" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="add_journal_form.asp?browser=IE" target="_self">Add 
      New News Item</a> (Windows IE 5+ WYSIWYG HTML Editor) &lt;---Use this one!<br>
      Add New News Item to the web site</font></td>
  </tr>
  <tr> 
    <td height="5"></td>
  </tr>
  <tr>
    <td height="37" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="select_journal_item.asp" target="_self">Amend 
      or Delete News Items </a><br>
      Amend or Delete News Items from the web site</font></td>
  </tr>
  <tr>
    <td height="51">&nbsp;</td>
  </tr>
</table>
<div align="center"><br>
  <a href="/" target="_parent">home</a>| <a href="http://www.hollowdays.com/news/" target="_parent">news</a> 
  | <a href="/lyrics/" target="_parent">lyrics</a> | <a href="http://www.hollowdays.com/shows/" target="_parent">shows</a> 
  | <a href="/press/">press</a> | <a href="http://www.hollowdays.com/bios" target="_parent">bios</a> 
  | <a href="http://www.hollowdays.com/mp3s/" target="_parent">mp3s</a> | <a href="http://www.hollowdays.com/pictures/" target="_parent">pictures</a> 
  | <a href="http://www.hollowdays.com/contact/" target="_parent">contact</a> 
  | <a href="http://www.hollowdays.com/list/" target="_parent">mailing list</a> 
  | <a href="http://www.hollowdays.com/links/" target="_parent">links</a> | <a href="/">&copy;2002 
  Hollow Days</a><font size="2"><br>
  </font><a href="http://www.herchel.com" target="_blank">Website designed by 
  Herchel Services</a> </div>
</body>
</html>
