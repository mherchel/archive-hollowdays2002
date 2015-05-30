<% Option Explicit %>
<!--#include file="common.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Guide - Web Wiz Site News
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

'Declare variables
Dim rsNews	'Database recordset holding the news items
Dim rsComments	'Database recordset holding the comments for this news item
Dim lngNewsID	'Holds the News item ID number

'Read in the News Item ID number to ge the comments for
lngNewsID = CLng(Request.QueryString("NewsID"))

'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html>
<head>
<title>Delete News Item Comments</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"> <b><font size="6">Delete News Item Comments</font></b><br>
  <a href="admin_menu.asp" target="_self">Return to the Site News Administrator 
  Menu</a><br>
  <a href="select_news_item.asp" target="_self">Select Comments for another News 
  Item to Delete</a><br>
  <br>
  <table width="563" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="563" height="2" align="center">To delete any of the comments 
        for this News Item place a tick in the check box at the top left corner 
        of the comment(s) you wish to delete and click on the Delete Comments 
        at the bottom of the page.</td>
    </tr>
  </table>
  <br>
    <%

'Create recorset object
Set rsNews = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
strSQL = "SELECT tblNews.* FROM tblNews "
strSQL = strSQL & "WHERE tblNews.News_ID = " & lngNewsID & ";"
	
'Query the database
rsNews.Open strSQL, adoCon


'If there are no records then exit for loop
If NOT rsNews.EOF Then
	
	%>
 <table width="90%" border="0" cellpadding="1" cellspacing="0" bgcolor="#000000">
  <tr>
   <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
     <tr> 
      <td width="645" bgcolor="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="4"><% = rsNews("News_title") %></font></b> - <em><% = FormatDateTime(rsNews("News_Date"), vbLongDate) %></em><br> 
       <% = rsNews("News_item") %>
      </td>
     </tr>
    </table></td>
  </tr>
 </table>
 
</div>
<br>
<%
	
End If

%>
<form name="frmDelete" method="post" action="delete_news_comments.asp" onSubmit="return confirm('Are you sure you want to delete these comments?');">
 <strong><font size="4" face="Arial, Helvetica, sans-serif">Comments</font></strong><br>
 <br>
<%
'Create recorset object
Set rsComments = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
strSQL = "SELECT tblComments.* FROM tblComments "
strSQL = strSQL & "WHERE tblComments.News_ID = " & lngNewsID
strSQL = strSQL & " ORDER BY Comments_Date DESC;"
	
'Query the database
rsComments.Open strSQL, adoCon

'Loop round to display all the comments for the news item
Do While NOT rsComments.EOF

%>
 <table width="90%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#000000">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="2">
     <tr> 
            
      <td bgcolor="#FFFFFF"> 
       <input type="checkbox" name="chkCommentsNo" value="<% = rsComments("Comment_ID") %>">Comments by <a href="mailto:<% = rsComments("EMail") %>"><% = rsComments("Name") %></a> from <% = rsComments("Country") %> on <% = FormatDateTime(rsComments("Comments_Date"), VbLongDate) %> at <% = FormatDateTime(rsComments("Comments_Date"), VbShortTime) %> - IP: <% = rsComments("IP") %>
      </td>
        </tr>
        <tr>
      <td bgcolor="#FFFFFF"><% = rsComments("Comments") %></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <br>
  <%
	'Move to the next record in the recordset
	rsComments.MoveNext

Loop

'Reset server objects
rsNews.Close
Set strCon = Nothing
Set adoCon = Nothing
%>
  <div align="center">
    <input type="hidden" name="NewsID" value="<% = lngNewsID %>">
    <input type="submit" name="Submit" value="Delete Comments">
  </div>
</form>
</body>
</html>