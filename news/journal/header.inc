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
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center"><% 
'If there is a Site News image then dsiplay it
If NOT strTitleImage = "" Then
    Response.Write("<img src=""" & strTitleImage) & """></td>"
End If
%>   </tr>
</table>