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



Dim rsNews			'Database recordset holding the news items
Dim intRecordPositionPageNum	'Holds the number of the page the user is on
Dim intRecordLoopCounter	'Loop counter to loop through each record in the recordset
Dim intTotalNumNewsEntries	'Holds the number of News Items there are in the database
Dim intTotalNumNewsPages	'Holds the number of pages the News Items cover
Dim intLinkPageNum		'Holds the number of the other pages of news itmes to link to


'If this is the first time the page is displayed then set the record position is set to page 1
If Request.QueryString("PagePosition") = "" Then
	intRecordPositionPageNum = 1

'Else the page has been displayed before so the news item record postion is set to the Record Position number
Else
	intRecordPositionPageNum = CInt(Request.QueryString("PagePosition"))
End If	
%>
<html>
<head>
<title>Site News</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

<!-- #include file="header.inc" -->
 <%
'Create recorset object
Set rsNews = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
strSQL = "SELECT tblNews.* FROM tblNews ORDER BY News_Date DESC;"

'Set the cursor type property of the record set to dynamic so we can naviagate through the record set
rsNews.CursorType = 3
	
'Query the database
rsNews.Open strSQL, adoCon

'Set the number of records to display on each page by the constant set in the common.asp file
rsNews.PageSize = intRecordsPerPage
	
'Get the record poistion to display from
If NOT rsNews.EOF Then rsNews.AbsolutePage = intRecordPositionPageNum


'If there are no rcords in the database display an error message
If rsNews.EOF Then
	'Tell the user there are no records to show
	Response.Write "<span class=""text""><br>There are no News Items to read"
	Response.Write "<br>Please check back later</span>"
	Response.End
	


'Display the News Items
Else	
	
	'Count the number of News Items database
	intTotalNumNewsEntries = rsNews.RecordCount	
	
	'Count the number of pages of News Items there are in the database calculated by the PageSize attribute set above
	intTotalNumNewsPages = rsNews.PageCount


	'Display the HTML number number the total number of pages and total number of records
	%>
 <br>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
 <tr> 
  <td align="center" class="text">There are <% = intTotalNumNewsEntries %> News Items in <% = intTotalNumNewsPages %> pages and your are on page number <% = intRecordPositionPageNum %></td>
 </tr>
</table>
<br><%

	'For....Next Loop to display the News Items in the database
	For intRecordLoopCounter = 1 to intRecordsPerPage

		'If there are no records then exit for loop
		If rsNews.EOF Then Exit For	
	%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="0">
 <tr> 
  <td class="text"><strong><a href="news_item.asp?NewsID=<% = rsNews("News_ID") %>" target="_self"><% = rsNews("News_title") %></a></strong> <span class="smText"><em>- <% = FormatDateTime(rsNews("News_Date"), vbLongDate) %></em></span><br> 
   <% = rsNews("Short_news") %> (<a href="news_item.asp?NewsID=<% = rsNews("News_ID") %>" target="_self">full story</a>)
  </td>
 </tr>
</table>
<br><%
		'Move to the next record in the recordset
		rsNews.MoveNext
	Next
End If

'Display an HTML table with links to the other News Items
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
 <tr> 
  <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
     <td width="50%" align="center" class="text"><%
'If there are more pages to display then add a title to the other pages
If intRecordPositionPageNum > 1 or NOT rsNews.EOF Then
	Response.Write vbCrLf & "		Page:&nbsp;"
End If

'If the News Items page number is higher than page 1 then display a back link    	
If intRecordPositionPageNum > 1 Then 
	Response.Write vbCrLf & "		 <a href=""default.asp?PagePosition=" &  intRecordPositionPageNum - 1  & """ target=""_self"">&lt;&lt;&nbsp;Prev</a>&nbsp;"   	     	
End If     	


'If there are more pages to display then display links to all the pages
If intRecordPositionPageNum > 1 or NOT rsNews.EOF Then 
	
	'Display a link for each page in the News Items     	
	For intLinkPageNum = 1 to intTotalNumNewsPages		
		
		'If the page to be linked to is the page displayed then don't make it a hyper-link
		If intLinkPageNum = intRecordPositionPageNum Then
			Response.Write vbCrLf & "		     " & intLinkPageNum
		Else
		
			Response.Write vbCrLf & "		     <a href=""default.asp?PagePosition=" &  intLinkPageNum  & """ target=""_self"">" & intLinkPageNum & "</a>&nbsp;"			
		End If
	Next
End If


'If it is Not the End of the News Items entries then display a next link for the next News Items page      	
If NOT rsNews.EOF then   	
	Response.Write vbCrLf & "		<a href=""default.asp?PagePosition=" &  intRecordPositionPageNum + 1  & """ target=""_self"">Next&nbsp;&gt;&gt;</a>"	   	
End If      	


'Finsh HTML the table 
%> </td>
    </tr>
   </table>
   </td>
 </tr>
</table>
<% 

'Reset server objects
rsNews.Close
Set rsNews = Nothing
Set strCon = Nothing
Set adoCon = Nothing
%>
<br>
<div align="center">
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:11px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:11px"">Web Wiz Site News</a> version 3.06</span>")
	Response.Write("<br><span class=""text"" style=""font-size:11px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
 %>
</div>
<!--#include file="footer.inc" -->