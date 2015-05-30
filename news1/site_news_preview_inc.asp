<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Site News
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

Dim adoNewsCon 			'Database Connection Variable
Dim rsNewsConfiguration		'Holds the configuartion recordset
Dim strAdoNewsConfig		'Holds the Database driver and the path and name of the database
Dim strNewsSQL			'Holds the SQL query for the database
Dim intPreviewNewsItems		'Number of files shown on each page
Dim blnNewsLCode		'News page code set to true
Dim strNewsBgColour		'Holds the background colour of the News Administrator
Dim strNewsTextColour		'Holds the text colour of the News Administrator
Dim strNewsTextType		'Holds the font type of the News Administrator
Dim intNewsTextSize		'Holds the font size of the News Administrator
Dim intNewsSmallTextSize	'Holds the size of small fonts
Dim strNewsLinkColour		'Holds the link colour of the News Administrator
Dim strNewsTableColour		'Holds the table colour
Dim strNewsTableBorderColour	'Holds the table border colour
Dim strNewsTableTitleColour	'Holds the table title colour
Dim strNewsVisitedLinkColour	'Holds the visited link colour of the News Administrator
Dim strNewsActiveLinkColour	'Holds the active link colour of the News Administrator





'Create database connection

'Create a connection odject
Set adoNewsCon = Server.CreateObject("ADODB.Connection")
			 
'------------- If you are having problems with the script then try using a diffrent driver or DSN by editing the lines below --------------
			 
'Database connection info and driver
strAdoNewsConfig = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("news/news.mdb")

'Database driver info for Brinkster
'strAdoNewsConfig = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/USERNAME/db/news.mdb") 'This one is for Brinkster users place your Brinster username where you see USERNAME

'Alternative drivers faster than the basic one above
'strAdoNewsConfig = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & Server.MapPath("news/news.mdb") 'This one is if you convert the database to Access 97
'strAdoNewsConfig = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("news/news.mdb")  'This one is for Access 2000/2002

'If you wish to use DSN then comment out the driver above and uncomment the line below (DSN is slower than the above drivers)
'strAdoNewsConfig = "DSN = DSN_NAME" 'Place the DSN where you see DSN_NAME

'---------------------------------------------------------------------------------------------------------------------------------------------

'Set an active connection to the Connection object
adoNewsCon.Open strAdoNewsConfig

'Read in the configuration for the script
'Intialise the ADO recordset object
Set rsNewsConfiguration = Server.CreateObject("ADODB.Recordset")

'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strNewsSQL = "SELECT tblConfiguration.* From tblConfiguration;"

'Query the database
rsNewsConfiguration.Open strNewsSQL, strAdoNewsConfig

'If there is config deatils in the recordset then read them in
If NOT rsNewsConfiguration.EOF Then

	'Read in the configuration details from the recordset
	strNewsTextColour = rsNewsConfiguration("text_colour")
	strNewsTextType = rsNewsConfiguration("text_type")
	intNewsTextSize = CInt(rsNewsConfiguration("text_size"))
	intNewsSmallTextSize = CInt(rsNewsConfiguration("small_text_size"))		
	strNewsTableColour = rsNewsConfiguration("table_colour")
	strNewsTableBorderColour = rsNewsConfiguration("table_border_colour")
	strNewsTableTitleColour = rsNewsConfiguration("table_title_colour")
	strNewsLinkColour = rsNewsConfiguration("links_colour")
	strNewsVisitedLinkColour = rsNewsConfiguration("visited_links_colour")
	strNewsActiveLinkColour = rsNewsConfiguration("active_links_colour")
	blnNewsLCode = CBool(rsNewsConfiguration("Code"))
	intPreviewNewsItems = rsNewsConfiguration("No_of_preview_items")
End If

'Reset server object
rsNewsConfiguration.Close
Set rsNewsConfiguration = Nothing

%>

<!-- Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<style type="text/css">
<!--
.text {font-family: <% = strNewsTextType %>; font-size: <% = intNewsTextSize %>px; color: <% = strNewsTextColour %>}
.smText {font-family: <% = strNewsTextType %>; font-size: <% = intNewsSmallTextSize %>px; color: <% = strNewsTextColour %>}
a {font-family: <% = strNewsTextType %>; font-size: <% = intNewsTextSize %>px; color: <% = strNewsLinkColour %>}
a:hover {font-family: <% = strNewsTextType %>; font-size: <% = intNewsTextSize %>px; color: <% = strNewsActiveLinkColour %>}
a:visited {font-family: <% = strNewsTextType %>; font-size: <% = intNewsTextSize %>px; color: <% = strNewsVisitedLinkColour %>}
a:visited:hover {font-family: <% = strNewsTextType %>; font-size: <% = intNewsTextSize %>px; color: <% = strNewsActiveLinkColour %>}
-->
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="1" align="center" bgcolor="<% = strNewsTableBorderColour %>">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="3" bgcolor="<% = strNewsTableColour %>">
        <tr>
          <td align="center" class="text">
            <%
Dim rsNews		'Database recordset holding the news items
Dim intNewsItems	'Loop counter for displaying the news items

'Create recorset object
Set rsNews = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strNewsSQL variable with an SQL statement to query the database
strNewsSQL = "SELECT TOP " & intPreviewNewsItems & " tblNews.* FROM tblNews ORDER BY News_Date DESC;"
	
'Query the database
rsNews.Open strNewsSQL, adoNewsCon

'If there are no news item to display then display a message seying so
If rsNews.EOF Then Response.Write("<span class=""text"">Sorry, There is no Site News Items to display</span>")

'Loop round to display each of the news items
For intNewsItems = 1 to intPreviewNewsItems

	'Iv there are no records then exit for loop
	If rsNews.EOF Then Exit For
	
	%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr> 
        <td align="left" class="text"><strong><a href="news/news_item.asp?NewsID=<% = rsNews("News_ID") %>" target="_self"><% = rsNews("News_title") %></a></strong> <span class="smText"><em>- <% = FormatDateTime(rsNews("News_Date"), vbLongDate) %></em></span><br> 
   	<% = rsNews("Short_news") %> (<a href="news/news_item.asp?NewsID=<% = rsNews("News_ID") %>" target="_self">full story</a>)</td>
       </tr>
      </table> 
      <br>
            <%
	'Move to the next record in the recordset
	rsNews.MoveNext
Next

'Reset server objects
rsNews.Close
Set rsNews = Nothing
Set strAdoNewsConfig = Nothing
Set adoNewsCon = Nothing
%>
           <a href="news/default.asp" target="_self">Site News and Archive</a> 
           <br>
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnNewsLCode = True Then
	Response.Write("<br>")
	Response.Write("<span class=""text"" style=""font-size:10px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:10px"">Web Wiz Site News</a> version 3.06</span>")
	Response.Write("<br><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
 %>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>