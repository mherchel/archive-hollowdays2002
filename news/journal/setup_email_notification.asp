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


'Set the response buffer to true
Response.Buffer = True 

'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If

'Dimension variables
Dim rsEmailNotfify 		'Recorset holding all the username in the database				
Dim strMode			'holds the mode of the page, set to true if changes are to be made to the database
Dim blnEmailNotify		'Set to true to turn e-mail notify on

'Initialise variables
blnEmailNotify = False
      
'Read in the details from the form
strMailComponent = Request.Form("component")
strSMTPServer = Request.Form("mailServer")
strWebSiteEmailAddress = Request.Form("email")
blnEmailNotify = CBool(Request.Form("userNotify"))
strMode = Request.Form("mode")


'Intialise the ADO recordset object
Set rsEmailNotfify  = Server.CreateObject("ADODB.Recordset")
	

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblConfiguration.* From tblConfiguration;"

'Set the cursor type property of the record set to Dynamic so we can navigate through the record set
rsEmailNotfify.CursorType = 2

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsEmailNotfify.LockType = 3
	
'Query the database
rsEmailNotfify.Open strSQL, strCon

'If the user is changing the email setup then update the database
If strMode = "change" Then

	
	'Update the recordset
	rsEmailNotfify.Fields("mail_component") = strMailComponent
	rsEmailNotfify.Fields("mail_server") = strSMTPServer
	rsEmailNotfify.Fields("email_address") = strWebSiteEmailAddress
	rsEmailNotfify.Fields("email_notify") = blnEmailNotify
			
				
	'Update the database with the new user's details
	rsEmailNotfify.Update
		
	'Re-run the query to read in the updated recordset from the database
	rsEmailNotfify.Requery	
End If

'Read in the deatils from the database
If NOT rsEmailNotfify.EOF Then
	
	'Read in the e-mail setup from the database
	strMailComponent = rsEmailNotfify("mail_component")
	strSMTPServer = rsEmailNotfify("mail_server")
	strWebSiteEmailAddress = rsEmailNotfify("email_address")
	blnEmailNotify = CBool(rsEmailNotfify("email_notify"))
End If	


'Reset Server Objects
Set adoCon = Nothing
Set strCon = Nothing
Set rsEmailNotfify = Nothing


%>  
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
<title>E-mail Notification Setup</title>

<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->
		
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Check for a mail server
	if (((document.frmEmailsetup.component.value=="AspEmail") || (document.frmEmailsetup.component.value=="Jmail")) && (document.frmEmailsetup.mailServer.value=="")){
		alert("Please enter an working incoming mail server \nWithout one the Jmail/AspEmail component will fail");
		document.frmEmailsetup.mailServer.focus();
		return false;
	}
	
	//Check for an e-mail address
	if (document.frmEmailsetup.email.value==""){
		alert("Please enter your E-mail Address");
		document.frmEmailsetup.email.focus();
		return false;
	}
	
	//Check that the e-mail address is valid
	if (document.frmEmailsetup.email.value.length>0&&(document.frmEmailsetup.email.value.indexOf("@",0)==-1||document.frmEmailsetup.email.value.indexOf(".",0)==-1)) { 
		alert("Please enter your valid E-mail address\nWithout a valid e-mail address the e-mail notification will not work"); 
		document.frmEmailsetup.email.focus();
		return false;
	}
	
	
	
	return true
}
// -->
</script>
     	
</head>
<body bgcolor="#FFFFFF" text="#000000">
<h1 align="center">E-mail Notification Setup</h1>
<div align="center"><a href="admin_menu.asp" target="_self">Return to the the 
  Administration Menu</a><br>
  <br>
  <table width="97%" border="1" cellspacing="0" cellpadding="4" bordercolor="#000000">
    <tr> 
      
   <td align="center" bgcolor="#CCCCCC"> <b><font size="5">Important - Please Read</font></b></td>
    </tr>
    <tr>
      
   <td bgcolor="#EFEFEF"> 
    <p>To be able to use the e-mail notification you need to have CDOSYS, the CDONTS e-mail component, the W3 JMail component, or Persists AspEmail component installed on the web server.</p>
    <p><b>Windows Win2k and XP Pro users</b> - CDOSYS comes installed on Win2k and XP Pro.<br>
     <br>
     <b>Windows NT4 and Win2k users</b> - IIS 4 and 5 on NT4 and Win2k instals the CDONTS e-mail component by default, but you need the SMTP server that comes with IIS installed on the web server as well (This is the e-mail component that most web hosts will use).<br>
     <br>
     <b>Windows 9x users</b> - I'm afraid Windows 98 does not support the CDOSYS or CDONTS e-mail components so if you enable this feature and try to test it on a Windows 9x system the Journal will crash!!<br>
     <br>
     The personal version of the JMail e-mail component is free and can run under Win98, NT4, and Win2k, Win XP, but you must install the component on the web server and requires that you enter the address of a working SMTP server.<br>
     <br>
     If you are not sure what mail component, if any, your web host uses then contact them to find out. <br>
     <br>
     If Web Wiz Journal crashes or you receive no e-mail's, you are either using the wrong component or your web host may not support sending mail from your web site.</p>
        </td>
    </tr>
  </table>
</div>
<form method="post" name="frmEmailsetup" action="setup_email_notification.asp" onSubmit="return CheckForm();">
  <table width="680" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#000000">
    <tr> 
      <td width="680"> 
        <table width="100%" border="0" align="center" class="normal" height="233" cellpadding="4" cellspacing="1">
          <tr align="left" bgcolor="#FFFFFF"> 
            <td colspan="2" class="arial_sm2" height="30"><font size="2">*Indicates required fields</font></td>
          </tr>
          <tr class="arial" bgcolor="#FFFFFF"> 
            <td align="left" width="59%" class="arial" height="12">E-mail Component to use:</td>
            <td height="12" width="41%" valign="top"> 
              <select name="component">
        <option value="CDOSYS"<% If strMailComponent = "CDOSYS" Then Response.Write(" selected") %>>CDOSYS (Win2k/XP Pro)</option>
        <option value="CDONTS"<% If strMailComponent = "CDONTS" Then Response.Write(" selected") %>>CDONTS (NT4/Win2k)</option>
        <option value="Jmail"<% If strMailComponent = "Jmail" Then Response.Write(" selected") %>>JMail</option>
        <option value="AspEmail"<% If strMailComponent = "AspEmail" Then Response.Write(" selected") %>>AspEmail</option>
        <option value="AspMail"<% If strMailComponent = "AspMail" Then Response.Write(" selected") %>>AspMail</option>
       </select>
            </td>
          </tr>
          <tr class="arial" bgcolor="#FFFFFF"> 
            <td align="left" width="59%" class="arial" height="12">Outgoing SMTP Mail Server (<b>NOT needed for CDONTS</b>):<br>
              <font size="2">You only need this if you are using an e-mail component other than CDONTS. It must be a working mail server or the script will 
              crash.</font></td>
            <td height="12" width="41%" valign="top"> 
              <input type="text" name="mailServer" maxlength="50" value="<% = strSMTPServer %>" size="30" >
              <br>
              (eg. mail.myweb.com)</td>
          </tr>
          <tr class="arial" bgcolor="#FFFFFF"> 
            <td align="left" width="59%" class="arial" height="23">Your E-mail Address* <br>
              <font size="2">Without a valid e-mail address receive e-mail notification of comments posted</font><br>
            </td>
            <td height="23" width="41%" valign="top"> 
              <input type="text" name="email" maxlength="50" value="<% = strWebSiteEmailAddress %>" size="30">
              &nbsp;</td>
          </tr>
          <tr class="arial" bgcolor="#FFFFFF"> 
            <td align="left" width="59%" class="arial" height="7">Admin E-mail Notify<font size="2"><br>
       Turn this function on if you wish to recieve e-mail notofication if one of your web visitors eneters a comments for a Journal item</font></td>
            <td height="7" width="41%" valign="top">On 
              <input type="radio" name="userNotify" value="True" <% If blnEmailNotify = True Then Response.Write "checked" %>>
              &nbsp;&nbsp;&nbsp;Off 
              <input type="radio" name="userNotify" value="False" <% If blnEmailNotify = False Then Response.Write "checked" %>>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF" align="center"> 
            <td valign="top" height="2" colspan="2" class="arial"> 
              <p> 
                <input type="hidden" name="mode" value="change">
                <input type="submit" name="Submit" value="Update Details">
                <input type="reset" name="Reset" value="Clear Form">
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<br>
</body>
</html>
