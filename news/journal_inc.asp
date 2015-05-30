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


Dim adoJournalCon 			'Database Connection Variable
Dim strJournalCon			'Holds the Database driver and the path and name of the database
Dim strJournalSQL			'Holds the SQL query for the database
Dim rsJournalRecordsConfiguration	'Holds the configuartion recordset
Dim rsJournalRecords			'Database recordset holding the Journal items
Dim rsJournalCommentsCount		'Database recordset holding the number of comments for a Journal Item
Dim intJournalItems			'Loop counter for displaying the Journal items
Dim strJournalBgColour			'Holds the background colour of the Journal
Dim strJournalTextColour		'Holds the text colour of the Journal
Dim strJournalTextType			'Holds the font type of the Journal
Dim intJournalHeadingTextSize		'Holds the heading font size
Dim intJournalTextSize			'Holds the font size of the Journal
Dim intJournalSmallTextSize		'Holds the small font size
Dim strJournalLinkColour		'Holds the link colour of the Journal
Dim strJournalTableColour		'Holds the table colour
Dim strJournalTableBorderColour		'Holds the table border colour
Dim strJournalTableTitleColour		'Holds the table title colour
Dim strJournalVisitedLinkColour		'Holds the visited link colour of the Journal
Dim strJournalActiveLinkColour		'Holds the active link colour of the Journal
Dim blnJournalLCode			'set to true
Dim intJournalPreviewItems		'Holds the number of preview journal items to display



'Create database connection

'Create a connection odject
Set adoJournalCon = Server.CreateObject("ADODB.Connection")
			 
'------------- If you are having problems with the script then try using a diffrent driver or DSN by editing the lines below --------------
			 
'Database connection info and driver
strJournalCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("journal/journal.mdb")

'Database driver info for Brinkster
'strJournalCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/USERNAME/db/site_journal.mdb") 'This one is for Brinkster users place your Brinster username where you see USERNAME

'Alternative OLE drivers faster than the basic one above
'strJournalCon = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & Server.MapPath("site_journal.mdb") 'This one is if you convert the database to Access 97
'strJournalCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("site_journal.mdb")  'This one is for Access 2000/2002

'If you wish to use DSN then comment out the driver above and uncomment the line below (DSN is slower than the above drivers)
'strJournalCon = "DSN = DSN_NAME" 'Place the DSN where you see DSN_NAME

'---------------------------------------------------------------------------------------------------------------------------------------------

'Set an active connection to the Connection object
adoJournalCon.Open strJournalCon

'Read in the configuration for the script
'Intialise the ADO recordset object
Set rsJournalRecordsConfiguration = Server.CreateObject("ADODB.Recordset")

'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strJournalSQL = "SELECT tblConfiguration.* From tblConfiguration;"

'Query the database
rsJournalRecordsConfiguration.Open strJournalSQL, strJournalCon

'If there is config deatils in the recordset then read them in
If NOT rsJournalRecordsConfiguration.EOF Then

	'Read in the configuration details from the recordset
	strJournalBgColour = rsJournalRecordsConfiguration("bg_colour")
	strJournalTextColour = rsJournalRecordsConfiguration("text_colour")
	strJournalTextType = rsJournalRecordsConfiguration("text_type")
	intJournalHeadingTextSize = CInt(rsJournalRecordsConfiguration("heading_text_size"))
	intJournalTextSize = CInt(rsJournalRecordsConfiguration("text_size"))
	intJournalSmallTextSize = CInt(rsJournalRecordsConfiguration("small_text_size"))	
	strJournalTableColour = rsJournalRecordsConfiguration("table_colour")
	strJournalTableBorderColour = rsJournalRecordsConfiguration("table_border_colour")
	strJournalTableTitleColour = rsJournalRecordsConfiguration("table_title_colour")
	strJournalLinkColour = rsJournalRecordsConfiguration("links_colour")
	strJournalVisitedLinkColour = rsJournalRecordsConfiguration("visited_links_colour")
	strJournalActiveLinkColour = rsJournalRecordsConfiguration("active_links_colour")
	blnJournalLCode = CBool(rsJournalRecordsConfiguration("Code"))
	intJournalPreviewItems = rsJournalRecordsConfiguration("No_of_preview_items")
End If

'Reset server object
rsJournalRecordsConfiguration.Close
Set rsJournalRecordsConfiguration = Nothing

%>

<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->
<style type="text/css">
<!--
.heading {font-family: <% = strJournalTextType %>; font-size: <% = intJournalHeadingTextSize %>px; color: <% = strJournalTextColour %>; font-weight: bold;}
.text {font-family: <% = strJournalTextType %>; font-size: <% = intJournalTextSize %>px; color: <% = strJournalTextColour %>}
.smText {font-family: <% = strJournalTextType %>; font-size: <% = intJournalSmallTextSize %>px; color: <% = strJournalTextColour %>}
a:hover {font-family: <% = strJournalTextType %>; font-size: <% = intJournalTextSize %>px; color: <% = strJournalActiveLinkColour %>
	color: #FFFFFF;
	text-decoration: underline;
}
a:visited:hover {font-family: <% = strJournalTextType %>; font-size: <% = intJournalTextSize %>px; color: <% = strJournalActiveLinkColour %>}
-->
</style>
<style type="text/css">
<!--
a:link {
	color: #FFFFFF;
	text-decoration: none;
}
a:visited {
	color: #FFFFFF;
	text-decoration: none;
}
-->
</style>
<body bgcolor="#000000" text="#CCCCCC"> 
<div align="center"> 
  <table width="730" border="0" cellpadding="0" cellspacing="0" bgcolor="#000000">
    <!--DWLayoutTable-->
    <tr> 
      <td width="730" height="164" align="center" valign="top" class="text"> 
        <%


'Create recorset object
Set rsJournalRecords = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strJournalSQL variable with an SQL statement to query the database
strJournalSQL = "SELECT TOP " & intJournalPreviewItems & "  tblJournal.* FROM tblJournal ORDER BY Date_stamp DESC;"
	
'Query the database
rsJournalRecords.Open strJournalSQL, adoJournalCon

'Create recorset object
Set rsJournalCommentsCount = Server.CreateObject("ADODB.Recordset")

'If there are no Journal item to display then display a message seying so
If rsJournalRecords.EOF Then Response.Write("<span class=""text"">Sorry, There is no Site Journal Items to display</span>")

'Loop round to display each of the news items
For intJournalItems = 1 to intJournalPreviewItems

	'If there are no records then exit for loop
	If rsJournalRecords.EOF Then Exit For
		
	'Get the count of comments from the db
	strJournalSQL = "SELECT Count(tblComments.Journal_ID) AS CountOfJournalItems "
	strJournalSQL = strJournalSQL & "FROM tblComments "
	strJournalSQL = strJournalSQL & "WHERE tblComments.Journal_ID = " & CLng(rsJournalRecords("Journal_ID")) & ";"
				
	'Query the database
	rsJournalCommentsCount.Open strJournalSQL, adoJournalCon
	
	%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
          <!--DWLayoutTable-->
          <tr> 
            <td width="2" height="84">&nbsp;</td>
            <td width="728" valign="top"> <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#000000">
                <!--DWLayoutTable-->
                <tr bgcolor="#000000"> 
                  <td width="8" height="25"></td>
                  <td width="705" valign="top" class="heading"> 
                    <% = rsJournalRecords("Journal_title") %>
                  </td>
                </tr>
                <tr bgcolor="#000000">
                  <td height="2"></td>
                  <td></td>
                </tr>
                <tr bgcolor="#000000"> 
                  <td height="47" colspan="2" valign="top" bgcolor="#000000" class="text"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <!--DWLayoutTable-->
                      <tr> 
                        <td width="10" height="18"></td>
                        <td width="710" valign="top" bgcolor="#000000" class="text"> 
                          <% = rsJournalRecords("Journal_item") %>
                        </td>
                      </tr>
                      <tr> 
                        <td height="19"></td>
                        <td align="right" valign="top" bgcolor="#000000" class="smText">~ 
                          <%
                  
        'If there is an email address entered make it a mailto link
  	If rsJournalRecords("Author_email") <> "" Then Response.Write("<a href=""mailto:" & rsJournalRecords("Author_email") & """ style=""font-size: " & intJournalSmallTextSize & "px;"">" & rsJournalRecords("Author") & "</a>") Else Response.Write(rsJournalRecords("Author"))
    
                  %>
                        </td>
                      </tr>
                      <tr> 
                        <td height="3"></td>
                        <td></td>
                      </tr>
                    </table></td>
                </tr>
                <tr bgcolor="#000000"> 
                  <td height="5"></td>
                  <td></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <%
	'Close the count recordset
	rsJournalCommentsCount.Close
		
	'Move to the next record in the recordset
	rsJournalRecords.MoveNext
Next

%>
        <br> <a href="archive.htm" target="_self" class="mail">View Archive</a> 
        <%

'Reset server objects
rsJournalRecords.Close
Set rsJournalRecords = Nothing
Set rsJournalCommentsCount = Nothing
Set strJournalCon = Nothing
Set adoJournalCon = Nothing
%>
        <br> </td>
      </tr>
  </table>
</div>
