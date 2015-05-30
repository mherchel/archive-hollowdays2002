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



'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html>
<head>
<title>Site News Administrator Menu</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

</head>
<body bgcolor="#FFFFFF" text="#000000">
<h1 align="center"><font face="Arial, Helvetica, sans-serif">Site News Administrator Menu</font></h1>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>    
  <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">It is highly recommend for security reasons that you change the username 
   and password.</font></td>
  </tr>
</table>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="add_news_form.asp?browser=IE" target="_self">Add New News Item</a> (Windows IE 5+ WYSIWYG HTML Editor)<br>
   Add New Site News Item to the web site<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="add_news_form.asp" target="_self">Add New News Item</a> (Standard HTML Editor)<br>
   Add New Site News Item to the web site<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="select_news_item.asp" target="_self">Amend or Delete News Items and Related User Comments</a><br>
   Amend or Delete News Itmes from the web site, also you can delete any inappropriate user comments<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="configure_site_news.asp" target="_self">Configure Site News and Change Admin Username and Password</a><br>
   Configure the Site News Application to look and feel like the rest of your site by changing graphics, colours, fonts, etc. Also change the admin username and password<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="setup_email_notification.asp" target="_self">E-mail notification Setup</a><br>
   Set up the e-mail notification so you can be notified when someone posts a comments for a News Item.<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="remove_link_buttons.asp">Remove Powered By Web Wiz Guide links</a><br>
   Remove the Powered by Web Wiz Site News links.</font> <br> <br> </td>
 </tr>
</table>
<div align="center"><br>
 <table width="700" border="0" cellspacing="0" cellpadding="1" bgcolor="#000000">
  <tr> 
   <td width="986"> <table width="100%" border="0" cellspacing="0" cellpadding="4" bgcolor="#EFEFEF">
     <tr> 
      <td align="center" height="186" width="100%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">I have spent many 1000's of unpaid hours in development and support this and the other applications<br>
       available for free from Web Wiz Guide. </font> <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">If you like using this application then please help support the development and update of <br>
        this and future applications from Web Wiz Guide.</font><br>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
        <a href="http://www.webwizguide.info/donations/default.asp" target="_blank">Click here to make a donation to Web Wiz Guide for this Application</a><br>
        <br>
        The <b>Web Wiz Site News application remains free</b> and you may still use it as much as you like both <br>
        privately and commercially, <b>the donation is only a request</b> to help me cover some of the costs involved.<br>
        <br>
        <b>For more info contact: -</b><br>
        <a href="mailto:donations@webwizguide.com">donations@webwizguide.com</a><br>
        Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom. </font></p></td>
     </tr>
    </table></td>
  </tr>
 </table>
 <br>
 <br>
 <font size="3" face="Verdana, Arial, Helvetica, sans-serif">
 <% 
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	Response.Write("<span class=""text"" style=""font-size:10px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:10px"">Web Wiz Site News</a> version 3.06</span>")
	Response.Write("<br><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
 </font><br>
</div>
</body>
</html>
