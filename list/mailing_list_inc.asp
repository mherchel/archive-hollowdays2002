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


Dim adoMailingListCon 		'Database Connection Variable
Dim strMailingListCon		'Holds the Database driver and the path and name of the database
Dim strMailingListAccessDB 	'Holds the Access Database Name
Dim rsMailingListConfiguration	'Holds the configuartion recordset
Dim strMailingListSQL		'Holds the SQL query for the database
Dim blnMailingListShow		'Set to true
Dim strMailingListTextColour	'Holds the text colour of the Mailing List
Dim strMailingListTextType	'Holds the font type of the Mailing List
Dim strMailingListTextSize	'Holds the font size of the Mailing List
Dim strMailingListPath		'Holds the path to the mailing list directory



'Path to the mailing list directory
strMailingListPath = "mailing_list/"

'Create database connection

'Initialise the strMailingListAccessDB variable with the name and path to the Access Database
strMailingListAccessDB = "mailing_list/mailing_list.mdb"

'Create a connection odject
Set adoMailingListCon = Server.CreateObject("ADODB.Connection")
			 
'------------- If you are having problems with the script then try using a diffrent driver or DSN by editing the lines below --------------
			 
'Database connection info and driver (if this driver does not work then comment it out and use one of the alternative drivers)
strMailingListCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath(strMailingListAccessDB)
'Alternative drivers
'strMailingListCon = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & Server.MapPath(strMailingListAccessDB) 'This one is for Access 97
'strMailingListCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(strMailingListAccessDB)  'This one is for Access 2000

'If you wish to use DSN then comment out the driver above and uncomment the line below (DSN is slower than the above drivers)
'strMailingListCon = "DSN=guestbook" 'Place the DSN name after the DSN=

'---------------------------------------------------------------------------------------------------------------------------------------------

'Set an active connection to the Connection object
adoMailingListCon.Open strMailingListCon


'Read in the mailing list configuration
'Intialise the ADO recordset object
Set rsMailingListConfiguration = Server.CreateObject("ADODB.Recordset")

'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strMailingListSQL = "SELECT tblConfiguration.* From tblConfiguration;"

'Query the database
rsMailingListConfiguration.Open strMailingListSQL, strMailingListCon

'If there is config deatils in the recordset then read them in
If NOT rsMailingListConfiguration.EOF Then

	'Read in the configuration details from the recordset
	strMailingListTextColour = rsMailingListConfiguration("text_colour")
	strMailingListTextType = rsMailingListConfiguration("text_type")
	strMailingListTextSize = CInt(rsMailingListConfiguration("text_size"))
	blnMailingListShow = CBool(rsMailingListConfiguration("Code"))
End If

'Reset server object
rsMailingListConfiguration.Close
Set rsMailingListConfiguration = Nothing

%>

<!-- The Web Wiz Guide - Web Wiz Mailing List is written and produced by Bruce Corkhill ©2001-2002
     	 If you want your own ASP Mailing List then goto http://www.webwizguide.info -->

<style type="text/css">
<!--
.text {font-family: <% = strMailingListTextType %>; font-size: <% = strMailingListTextSize %>px; color: <% = strMailingListTextColour %>}
-->
</style>
<form name="frmMailingList" method="post" action="<% = strMailingListPath %>mailing_list.asp" target="mailingList" onSubmit="window.open('', 'mailingList', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=400,height=200')">
  <table width="143" border="0" cellspacing="2" cellpadding="0" align="center">
    <tr> 
      <td class="text">Your E-mail Address:</td>
    </tr>
    <tr> 
      <td> 
        <input type="text" name="email" maxlength="35">
      </td>
    </tr>
    <tr> 
      <td align="left" class="text"> 
        <input type="radio" name="mode" value="add" id="add" checked><label for="add">Subscribe</label><br>
        <input type="radio" name="mode" value="delete" id="delete"><label for="delete">Un-Subscribe</label></td>
    </tr>
    <tr>
      <td align="center">
        <input type="submit" name="Submit" value="Submit">
      </td>
    </tr>
  </table>
</form>