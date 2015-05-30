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
Dim rsComments			'Database recordset holding the count of comments for each news item
Dim intRecordPositionPageNum	'Holds the number of the page the user is on
Dim intRecordLoopCounter	'Loop counter to loop through each record in the recordset
Dim intTotalNumNewsEntries	'Holds the number of News Items there are in the database
Dim intTotalNumNewsPages	'Holds the number of pages the News Items cover
Dim intLinkPageNum		'Holds the number of the other pages of news itmes to link to


'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If


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
<title>Amend or Delete News Item </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"><b><font size="6" face="Arial, Helvetica, sans-serif">Amend or Delete News Item</font></b> <br>
 <a href="admin_menu.asp" target="_self"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Return to the Site News Administrator Menu</font></a><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
 <br>
 </font> 
 <table width="612" border="0" cellspacing="0" cellpadding="0">
  <tr> 
   <td width="612" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Use the links at the bottom of each news item to Edit/Delete and Review News Items and any user comments that may have been left for a news item.</font></td>
  </tr>
 </table>
 <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
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


'Create recorset object
Set rsComments = Server.CreateObject("ADODB.Recordset")

'If there are no rcords in the database display an error message
If rsNews.EOF Then
	'Tell the user there are no records to show
	Response.Write "<br>There are no News Items to read"
	Response.Write "<br>Please check back later"
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
 </font></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
 <tr> 
  <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> There are <% = intTotalNumNewsEntries %> News Items in <% = intTotalNumNewsPages %> pages and your are on page number <% = intRecordPositionPageNum %></font></td>
 </tr>
</table>
      <br>
      <%

	'For....Next Loop to display the News Items in the database
	For intRecordLoopCounter = 1 to intRecordsPerPage

		'If there are no records then exit for loop
		If rsNews.EOF Then Exit For
		
		'Read in if there any comments for this news item
		'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
		strSQL = "SELECT TOP 1 tblComments.Comment_ID FROM tblComments WHERE tblComments.News_ID = " & rsNews("News_ID") & ";"
			
		'Query the database
		rsComments.Open strSQL, adoCon
	
	%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#000000">
 <tr>
  <td><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
     <td bgcolor="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="4"><% = rsNews("News_title") %></font></b> <em>- <% = FormatDateTime(rsNews("News_Date"), vbLongDate) %></em><br> 
      <% = rsNews("Short_news") %>
      <br><a href="edit_news_item_form.asp?NewsID=<% = rsNews("News_ID") %>&browser=IE" target="_self">Edit with IE 5 WYSIWYG HTML editor</a> | <a href="edit_news_item_form.asp?NewsID=<% = rsNews("News_ID") %>" target="_self">Edit with Standard HTML editor</a>
      <%
   		
   		'If there are comments realated to this news item display a link to edit or delete them
   		If NOT rsComments.EOF Then %> | <a href="delete_news_comments_form.asp?NewsID=<% = rsNews("News_ID") %>" target="_self">Review/Delete Comments</a></td>
     <%
   
		End If 
		
		'Close the comments recordset
		rsComments.Close
		
		%>
    </tr>
   </table></td>
 </tr>
</table>
<br>
<%
		'Move to the next record in the recordset
		rsNews.MoveNext
	Next
End If

'Display an HTML table with links to the other News Items
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="50%" align="center"> 
                  <%
'If there are more pages to display then add a title to the other pages
If intRecordPositionPageNum > 1 or NOT rsNews.EOF Then
	Response.Write vbCrLf & "		Page:&nbsp;&nbsp;"
End If

'If the News Items page number is higher than page 1 then display a back link    	
If intRecordPositionPageNum > 1 Then 
	Response.Write vbCrLf & "		 <a href=""select_news_item.asp?PagePosition=" &  intRecordPositionPageNum - 1  & """ target=""_self"">&lt;&lt;&nbsp;Prev</a>&nbsp;"   	     	
End If     	


'If there are more pages to display then display links to all the pages
If intRecordPositionPageNum > 1 or NOT rsNews.EOF Then 
	
	'Display a link for each page in the News Items     	
	For intLinkPageNum = 1 to intTotalNumNewsPages		
		
		'If the page to be linked to is the page displayed then don't make it a hyper-link
		If intLinkPageNum = intRecordPositionPageNum Then
			Response.Write vbCrLf & "		     " & intLinkPageNum
		Else
		
			Response.Write vbCrLf & "		     &nbsp;<a href=""select_news_item.asp?PagePosition=" &  intLinkPageNum  & """ target=""_self"">" & intLinkPageNum & "</a>&nbsp;"			
		End If
	Next
End If


'If it is Not the End of the News Items entries then display a next link for the next News Items page      	
If NOT rsNews.EOF then   	
	Response.Write vbCrLf & "		&nbsp;<a href=""select_news_item.asp?PagePosition=" &  intRecordPositionPageNum + 1  & """ target=""_self"">Next&nbsp;&gt;&gt;</a>"	   	
End If      	


'Finsh HTML the table 
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
<% 

'Reset server objects
rsNews.Close
Set rsNews = Nothing
Set rsComments = Nothing
Set strCon = Nothing
Set adoCon = Nothing
%>
</body>
</html>