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

</head>

<body bgcolor="#FFFFFF" text="#000000">
<h2 align="center"><font face="Arial, Helvetica, sans-serif">Mailing List Admin Menu</font></h2>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">It is highly recommend for security reasons that you change the username 
      and password.<br>
      This can be done from the <a href="configure_script.asp" target="_self">Configure Mailing List Script</a> page.</font></td>
  </tr>
</table>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="send_mail.asp?Format=Plain" target="_self">Send Plain Text E-mail to Mailing List Members </a><br>
   Send an e-mail in plain text format to all the Members of the Mailing List.<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="send_mail.asp?Format=HTML" target="_self">Send HTML E-mail to Mailing List Members </a><br>
   Send an e-mail in HTML format to all the Members of the Mailing List using the standard HTML e-mail editor.<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="send_mail.asp?Format=advHTML" target="_self">Send HTML E-mail to Mailing List Members </a>(Hotmail style e-mail editor for<strong> Windows IE 
   5+ users only</strong>)<br>
   Send an e-mail in HTML format to all the Members of the Mailing List using the WYSIWYG HTML e-mail editor.<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td height="47"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="delete_list_members_form.asp" target="_self">View or Remove Mailing List Members</a><br>
   View mailing list members or remove those who no longer want to be part of your mailing list.<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="configure_script.asp" target="_self">Configure Mailing List Application</a><br>
   This must be done before you send any e-mail's to your mailing list members or the script will fail to send the e-mail's.<br>
   <br>
   </font></td>
 </tr>
 <tr>
  <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="configure_mailing_list_colours.asp" target="_self">Change the Mailing List Colours and Text Type</a><br>
   Set the colours of the Mailing List and the text type that your site visitors will see.<br>
   <br>
   </font></td>
 </tr>
 <tr> 
  <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="remove_link_buttons.asp">Remove Powered By Web Wiz Guide links</a><br>
   Remove the Powered by Web Wiz Mailing List links.</font> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
   <br>
   </font></td>
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
        The <b>Web Wiz Mailing List application remains free</b> and you may still use it as much as you like both <br>
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
	Response.Write("<span class=""text"" style=""font-size:10px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:10px"">Web Wiz Mailing List</a> version 3.02</span>")
	Response.Write("<br><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
 </font><br>
</div>
</body>
</html>
