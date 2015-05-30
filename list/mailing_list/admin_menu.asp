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


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html>
<head>
<title>Mailing List Admin Menu</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- The Web Wiz Guide - Web Wiz Mailing List is written and produced by Bruce Corkhill ©2001-2002
     	 If you want your own ASP Mailing List then goto http://www.webwizguide.info -->
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

<body bgcolor="#000000" text="#CCCCCC"><h2 align="center">
<div align="center"><img src="/list/top.jpg" width="700" height="209" border="0" usemap="#Map"> 
  <map name="Map">
    <area shape="rect" coords="-160,-25,328,134" href="/" target="_parent">
  </map>
  <br>
  <br>
</div>
<table width="704" border="0" cellspacing="0" cellpadding="0" align="center">
  <!--DWLayoutTable-->
  <tr> 
    <td width="704" height="51" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="send_mail.asp?Format=advHTML" target="_self">Send 
      HTML E-mail to Mailing List Members </a><br>
      Send an e-mail in HTML format to all the Members of the Mailing List using 
      the WYSIWYG HTML e-mail editor.<br>
      </font></td>
  </tr>
  <tr> 
    <td height="67" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="send_mail.asp?Format=HTML" target="_self">Send 
      HTML E-mail to Mailing List Members </a><br>
      Send an e-mail in HTML format to all the Members of the Mailing List using 
      the standard HTML e-mail editor.<br>
      <br>
      </font></td>
  </tr>
  <tr> 
    <td height="67" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
      <a href="send_mail.asp?Format=Plain" target="_self">Send Plain Text E-mail 
      to Mailing List Members </a><br>
      Send an e-mail in plain text format to all the Members of the Mailing List.<br>
      <br>
      </font></td>
  </tr>
  <tr> 
    <td height="50" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="delete_list_members_form.asp" target="_self">View 
      or Remove Mailing List Members</a><br>
      View mailing list members or remove those who no longer want to be part 
      of your mailing list.</font></td>
  </tr>
  <tr>
    <td height="17"></td>
  </tr>
</table>
  
<div align="center"><br>
  <font size="3" face="Verdana, Arial, Helvetica, sans-serif"> 
  <% 
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	Response.Write("<span class=""text"" style=""font-size:10px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:10px"">Web Wiz Mailing List</a> version 3.02</span>")
	Response.Write("<br><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
  </font><br></div>
  </div>
</body>
</html>
