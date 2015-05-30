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



Dim strInputJournalItem 	'Holds the Users Message
Dim strJournalItemTitle	'Holds the title of the Journal item


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If


'Read in the message to be previewed from the cookie set
strInputJournalItem = Request.Cookies("JournalItem")
strJournalItemTitle = Request.Cookies("Title")

'If there is nothing in the Journal title then put something in the title string
If strJournalItemTitle = "" Then strJournalItemTitle = "&nbsp;"


'Format the Journal Item by replacing characters with HTML equivelents
strInputJournalItem = Replace(strInputJournalItem, ":)", "<img src=""Journal_images/smiley1.gif"" width=""17"" height=""17"">")
strInputJournalItem = Replace(strInputJournalItem, ";)", "<img src=""Journal_images/smiley2.gif"" width=""17"" height=""17"">")
strInputJournalItem = Replace(strInputJournalItem, ":o", "<img src=""Journal_images/smiley3.gif"" width=""17"" height=""17"">")
strInputJournalItem = Replace(strInputJournalItem, ":D", "<img src=""Journal_images/smiley4.gif"" width=""17"" height=""17"">")
strInputJournalItem = Replace(strInputJournalItem, ":errr:", "<img src=""Journal_images/smiley5.gif"" width=""17"" height=""17"">")
strInputJournalItem = Replace(strInputJournalItem, ":(", "<img src=""Journal_images/smiley6.gif"" width=""17"" height=""17"">")
strInputJournalItem = Replace(strInputJournalItem, "X(", "<img src=""Journal_images/smiley7.gif"" width=""17"" height=""17"">")

strInputJournalItem = Replace(strInputJournalItem, Chr(10), "<br>", 1, -1, 1)


'If there is nothing to preview then say so
If strInputJournalItem = "" OR IsNull(strInputJournalItem) Then
	strInputJournalItem = "<br><br><div align=""center"">There is nothing to preview</div><br><br>"
End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Journal Item Preview</title>

<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->

<style type="text/css">
<!--
body {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strTextColour %>}
h1 {font-family: <% = strTextType %>; font-size: 24px; color: <% = strTextColour %>}
td {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strTextColour %>}
-->
</style>

</head>
<body bgcolor="<% = strBgColour %>" text="<% = strTextColour %>" link="<% = strLinkColour %>" vlink="<% = strVisitedLinkColour %>" alink="<% = strActiveLinkColour %>">
<table width="98%" border="0" cellspacing="0" cellpadding="1" align="center" height="53">
  <tr> 
    <td align="center" height="17"><h1>Journal Item Preview</h1></td>
  </tr>
  <tr>
    <td align="center" height="39"><a href="JavaScript:onClick=window.close()">Close 
      Window</a></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="<% = strTableBorderColour %>">
  <tr>
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="<% = strTableBorderColour %>">
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="1" cellpadding="2">
              <tr> 
                <td bgcolor="<% = strTableTitleColour %>"><span style="font-size: <% = intTextSize + 1 %>;font-weight: bold;">
                  <% = strJournalItemTitle %>
                  </span></td>
              </tr>
              <tr> 
                <td bgcolor="<% = strTableColour %>"><% = strInputJournalItem %></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="1" align="center">
  <tr>
    <td align="center" height="49"><a href="JavaScript:onClick=window.close()">Close 
      Window</a></td>
  </tr>
</table>
</body>
</html>