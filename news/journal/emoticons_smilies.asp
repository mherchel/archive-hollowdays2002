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

Response.Buffer = True 

'Reset Server Objects
Set adoCon = Nothing
Set strCon = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
<title>Emoticon Smilies</title>

<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->
		
<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to add the code to the message for the smileys
function AddSmileyIcon(iconCode) {	
	window.opener.document.frmJournalComments.comments.value += iconCode + " ";
}

// -->
</script>

<style type="text/css">
<!--
.heading {font-family: <% = strTextType %>; font-size: <% = intHeadingTextSize %>px; color: <% = strTextColour %>; font-weight: bold;}
.text {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strTextColour %>}
.smText {font-family: <% = strTextType %>; font-size: <% = intSmallTextSize %>px; color: <% = strTextColour %>}
a {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strLinkColour %>}
a:hover {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strActiveLinkColour %>}
a:visited {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strVisitedLinkColour %>}
a:visited:hover {font-family: <% = strTextType %>; font-size: <% = intTextSize %>px; color: <% = strActiveLinkColour %>}
-->
</style>

</head>
<body bgcolor="<% = strBgColour %>" text="<% = strTextColour %>">
<table width="98%" border="0" cellspacing="0" cellpadding="1" align="center">
  <tr>
    <td align="center"><span class="heading">Emoticon Smilies</span></td>
  </tr>
</table>
<br>
  <table width="350" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="<% = strTableBorderColour %>" height="138">
  <tr> 
      <td height="174"> 
        
      <table border="0" align="center" cellpadding="4" cellspacing="1" width="350">
        <tr align="left" bgcolor="<% = strTableColour %>"> 
          <td colspan="2" class="text" align="center"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="4">
              <tr> 
                <td align="center">Click on the emoticon smiley you would like to add to your comments.</td>
              </tr>
            </table>
            <table width="340" border="0" cellspacing="0" cellpadding="4">
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:)]')"><img src="journal_images/smiley1.gif" width="17" height="17" border="0" alt="[:)]"></a></td>
                <td width="62" class="text">Smile</td>
                <td width="64" class="text">[:)]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:oops:]')"><img src="journal_images/smiley9.gif" width="17" height="17" border="0" alt="[:oops:]"></a></td>
                <td width="75" class="text">Embarassed</td>
                <td width="57" class="text">[:oops:]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[;)]')"><img src="journal_images/smiley2.gif" width="17" height="17" border="0" alt="[;)]"></a></td>
                <td width="62" class="text">Wink</td>
                <td width="64" class="text">[;)]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:star:]')"><img src="journal_images/smiley10.gif" width="17" height="17" border="0" alt="[:star:]"></a></td>
                <td width="75" class="text">Star</td>
                <td width="57" class="text">[:star:]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:o]')"><img src="journal_images/smiley3.gif" width="17" height="17" border="0" alt="[o)]"></a></td>
                <td width="62" class="text">Shocked</td>
                <td width="64" class="text">[:O]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[xx(]')"><img src="journal_images/smiley11.gif" width="17" height="17" border="0" alt="[xx(]"></a></td>
                <td width="75" class="text">Dead</td>
                <td width="57" class="text">[xx(]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:D]')"><img src="journal_images/smiley4.gif" width="17" height="17" border="0" alt="[:D]"></a></td>
                <td width="62" class="text">Big smile</td>
                <td width="64" class="text">[:D]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[|)]')"><img src="journal_images/smiley12.gif" width="17" height="17" border="0" alt="[|)]"></a></td>
                <td width="75" class="text">Sleepy</td>
                <td width="57" class="text">[|)]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:errr:]')"><img src="journal_images/smiley5.gif" width="17" height="17" border="0" alt="[:errr:]"></a></td>
                <td width="62" class="text">Confused</td>
                <td width="64" class="text">[:errr:]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:V:]')"><img src="journal_images/smiley13.gif" width="17" height="17" border="0" alt="[:V:]"></a></td>
                <td width="75" class="text">Disapprove</td>
                <td width="57" class="text">[:V:]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:(]')"><img src="journal_images/smiley6.gif" width="17" height="17" border="0" alt="[:(]"></a></td>
                <td width="62" class="text">Unhappy</td>
                <td width="64" class="text">[:(]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:^:]')"><img src="journal_images/smiley14.gif" width="17" height="17" border="0" alt="[:^:]"></a></td>
                <td width="75" class="text">Approve</td>
                <td width="57" class="text">[:^:]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:x]')"><img src="journal_images/smiley7.gif" width="17" height="17" border="0" alt="[:x]"></a></td>
                <td width="62" class="text">Angry</td>
                <td width="64" class="text">[:x]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[}:)]')"><img src="journal_images/smiley15.gif" width="17" height="17" border="0" alt="[}:)]"></a></td>
                <td width="75" class="text">Evil smile</td>
                <td width="57" class="text">[}:)]</td>
              </tr>
              <tr> 
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[:o)]')"><img src="journal_images/smiley8.gif" width="17" height="17" border="0" alt="[:o)]"></a></td>
                <td width="62" class="text">Clown</td>
                <td width="64" class="text">[:o)]</td>
                <td width="17" class="text"><a href="JavaScript:AddSmileyIcon('[8D]')"><img src="journal_images/smiley16.gif" width="17" height="17" border="0" alt="[8D]"></a></td>
                <td width="75" class="text">Cool</td>
                <td width="57" class="text">[8D]</td>
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
    <td align="center" height="34"><a href="JavaScript:onClick=window.close()">Close 
      Window</a></td>
  </tr>
</table>
<div align="center">
 <%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:10px"">Web Wiz Journal</a> version 1.0</span>")
	Response.Write("<br><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
 %>
</div>
</body>
</html>
