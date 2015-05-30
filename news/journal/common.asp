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


Dim adoCon 			'Database Connection Variable
Dim rsConfiguration		'Holds the configuartion recordset
Dim strCon			'Holds the Database driver and the path and name of the database
Dim strSQL			'Holds the SQL query for the database
Dim intRecordsPerPage		'Holds the number of files shown on each page
Dim strBgColour			'Holds the background colour of the Journal
Dim strTextColour		'Holds the text colour of the Journal
Dim strTextType			'Holds the font type of the Journal
Dim intHeadingTextSize		'Holds the heading font size
Dim intTextSize			'Holds the font size of the Journal
Dim intSmallTextSize		'Holds the small font size
Dim strLinkColour		'Holds the link colour of the Journal
Dim strTableColour		'Holds the table colour
Dim strTableBorderColour	'Holds the table border colour
Dim strTableTitleColour		'Holds the table title colour
Dim strVisitedLinkColour	'Holds the visited link colour of the Journal
Dim strActiveLinkColour		'Holds the active link colour of the Journal
Dim blnLCode			'set to true
Dim blnEmail			'Boolean set to true if e-mail is on
Dim strCode			'Holds the page code
Dim strCodeField		'Holds the code type
Dim strWebSiteEmailAddress	'Holds the e-mail address for the web site the Site Journal is on
Dim strMailComponent		'Email coponent the site Journal app useses
Dim strSMTPServer		'SMTP server for sending the e-mails through
Dim strLoggedInUserCode		'Holds the user code of the user
Dim strTitleImage		'Holds the path and name for the title image for the site Journal
Dim intMsgCharNo		'Holds the number of characters allowed for the messages
Dim blnCookieSet		'Set to true if cookies are to be set to stop multiple posts
Dim blnIPBlocking		'Set to true if IP blooking is to be used to stop multiple posts



'Create database connection

'Create a connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")
			 
'------------- If you are having problems with the script then try using a diffrent driver or DSN by editing the lines below --------------
			 
'Database connection info and driver
strCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("journal.mdb")

'Database driver info for Brinkster
'strCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/USERNAME/db/site_journal.mdb") 'This one is for Brinkster users place your Brinster username where you see USERNAME

'Alternative OLE drivers faster than the basic one above
'strCon = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & Server.MapPath("site_journal.mdb") 'This one is if you convert the database to Access 97
'strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("site_journal.mdb")  'This one is for Access 2000/2002

'If you wish to use DSN then comment out the driver above and uncomment the line below (DSN is slower than the above drivers)
'strCon = "DSN = DSN_NAME" 'Place the DSN where you see DSN_NAME

'---------------------------------------------------------------------------------------------------------------------------------------------

'Set an active connection to the Connection object
adoCon.Open strCon

'Set up the page encoding
strCodeField = "C&#111;&#100;&#101;"
strCode = "&#110;&#111;&#108;&#105;&#110;&#107;&#115;&#050;&#048;&#048;&#050;"

'Read in the configuration for the script
'Intialise the ADO recordset object
Set rsConfiguration = Server.CreateObject("ADODB.Recordset")

'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT tblConfiguration.* From tblConfiguration;"

'Query the database
rsConfiguration.Open strSQL, strCon

'If there is config deatils in the recordset then read them in
If NOT rsConfiguration.EOF Then

	'Read in the configuration details from the recordset
	intRecordsPerPage = CInt(rsConfiguration("No_records_per_page"))
	strBgColour = rsConfiguration("bg_colour")
	strTextColour = rsConfiguration("text_colour")
	strTextType = rsConfiguration("text_type")
	intHeadingTextSize = CInt(rsConfiguration("heading_text_size"))
	intTextSize = CInt(rsConfiguration("text_size"))
	intSmallTextSize = CInt(rsConfiguration("small_text_size"))	
	strTableColour = rsConfiguration("table_colour")
	strTableBorderColour = rsConfiguration("table_border_colour")
	strTableTitleColour = rsConfiguration("table_title_colour")
	strLinkColour = rsConfiguration("links_colour")
	strVisitedLinkColour = rsConfiguration("visited_links_colour")
	strActiveLinkColour = rsConfiguration("active_links_colour")
	strWebSiteEmailAddress = rsConfiguration("email_address")
	blnLCode = CBool(rsConfiguration("Code"))
	blnEmail = CBool(rsConfiguration("email_notify"))
	strTitleImage = rsConfiguration("Title_image")
	intMsgCharNo = rsConfiguration("Message_char_no")
	blnCookieSet = CBool(rsConfiguration("Cookie"))
	blnIPBlocking = CBool(rsConfiguration("IP_blocking"))
	strMailComponent = rsConfiguration("mail_component")
	strSMTPServer = rsConfiguration("mail_server")
End If

'Reset server object
rsConfiguration.Close
Set rsConfiguration = Nothing
%>