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



Dim rsJournal			'Database recordset holding the Journal items
Dim rsCommentsCount		'Database recordset holding the count of comments for each Journal item
Dim intRecordPositionPageNum	'Holds the number of the page the user is on
Dim intRecordLoopCounter	'Loop counter to loop through each record in the recordset
Dim intTotalNumJournalEntries	'Holds the number of Journal Items there are in the database
Dim intTotalNumJournalPages	'Holds the number of pages the Journal Items cover
Dim intLinkPageNum		'Holds the number of the other pages of Journal itmes to link to


'If this is the first time the page is displayed then set the record position is set to page 1
If Request.QueryString("PagePosition") = "" Then
	intRecordPositionPageNum = 1

'Else the page has been displayed before so the Journal item record postion is set to the Record Position number
Else
	intRecordPositionPageNum = CInt(Request.QueryString("PagePosition"))
End If
%>
<html>
<head>
<title>Journal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->
     
<!-- #include file="header.inc" -->
<div align="center"><span class="heading" style="font-size: <% = intHeadingTextSize + 2 %>px;">Journal</span><br>
  <a href="../default.asp" target="_self">Return to the homepage</a><br>
  <%
'Create recorset object
Set rsJournal = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblJournal.* FROM tblJournal ORDER BY Date_stamp DESC;"

'Set the cursor type property of the record set to dynamic so we can naviagate through the record set
rsJournal.CursorType = 3
	
'Query the database
rsJournal.Open strSQL, adoCon

'Set the number of records to display on each page by the constant set in the common.asp file
rsJournal.PageSize = intRecordsPerPage
	
'Get the record poistion to display from
If NOT rsJournal.EOF Then rsJournal.AbsolutePage = intRecordPositionPageNum


'Create recorset object
Set rsCommentsCount = Server.CreateObject("ADODB.Recordset")
		

'If there are no rcords in the database display an error message
If rsJournal.EOF Then
	'Tell the user there are no records to show
	Response.Write "<span class=""text""><br>There are no Journal Items to read"
	Response.Write "<br>Please check back later</span>"
	Response.End
	


'Display the Journal Items
Else	
	
	'Count the number of Journal Items database
	intTotalNumJournalEntries = rsJournal.RecordCount	
	
	'Count the number of pages of Journal Items there are in the database calculated by the PageSize attribute set above
	intTotalNumJournalPages = rsJournal.PageCount


	'Display the HTML number number the total number of pages and total number of records
	%>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="center" class="text"> <br>
      There are <% = intTotalNumJournalEntries %> Journal Items in <% = intTotalNumJournalPages %> pages and your are on page number <% = intRecordPositionPageNum %>
    </td>
  </tr>
</table>
<br>
<%

	'For....Next Loop to display the Journal Items in the database
	For intRecordLoopCounter = 1 to intRecordsPerPage

		'If there are no records then exit for loop
		If rsJournal.EOF Then Exit For
		
		'Get the count of comments from the db
		strSQL = "SELECT Count(tblComments.Journal_ID) AS CountOfJournalItems "
		strSQL = strSQL & "FROM tblComments "
		strSQL = strSQL & "WHERE tblComments.Journal_ID = " & CLng(rsJournal("Journal_ID")) & ";"
				
		'Query the database
		rsCommentsCount.Open strSQL, adoCon
	
	%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="<% = strTableBorderColour %>">
  <tr> 
    <td> <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr> 
          <td bgcolor="<% = strTableTitleColour %>" class="heading"><% = rsJournal("Journal_title") %></td>
        </tr>
        <tr> 
          <td bgcolor="<% = strTableColour %>" class="text"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="text"><% = rsJournal("Journal_item") %>
                </td>
              </tr>
              <tr> 
                <td align="right" class="smText" height="20">Posted by <%
                  
                'If there is an email address entered make it a mailto link
  		If rsJournal("Author_email") <> "" Then Response.Write("<a href=""mailto:" & rsJournal("Author_email") & """ style=""font-size: " & intSmallTextSize & "px;"">" & rsJournal("Author") & "</a>") Else Response.Write(rsJournal("Author"))
    
                  %> on <% = FormatDateTime(rsJournal("Date_stamp"), vbLongDate) %> at <% = FormatDateTime(rsJournal("Date_stamp"), vbShortTime) %><%
                  
		'If commets are allowed for this itm show a links to the comments page
		If CBool(rsJournal("Comments")) = True Then  
                  
                  %>
                  <a href="journal_comments.asp?JournalID=<% = rsJournal("Journal_ID") %>&PagePosition=<% = intRecordPositionPageNum %>" target="_self" style="font-size: <% = intSmallTextSize %>px">Comments</a>
                  <% 
                                                                          
	            	If NOT rsCommentsCount.EOF Then 
	            		Response.Write "(" & rsCommentsCount("CountOfJournalItems") & ")"
	            	Else
	            		Response.Write "(0)"
	            	End If
		End If
                     %>
                  </td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<br>
<%
		'Close the count recordset
		rsCommentsCount.Close
		
		'Move to the next record in the recordset
		rsJournal.MoveNext
	Next
End If

'Display an HTML table with links to the other Journal Items
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="50%" align="center" class="text"> 
            <%
'If there are more pages to display then add a title to the other pages
If intRecordPositionPageNum > 1 or NOT rsJournal.EOF Then
	Response.Write vbCrLf & "		Page:&nbsp;"
End If

'If the Journal Items page number is higher than page 1 then display a back link    	
If intRecordPositionPageNum > 1 Then 
	Response.Write vbCrLf & "		 <a href=""default.asp?PagePosition=" &  intRecordPositionPageNum - 1  & """ target=""_self"">&lt;&lt;&nbsp;Prev</a>&nbsp;"   	     	
End If     	


'If there are more pages to display then display links to all the pages
If intRecordPositionPageNum > 1 or NOT rsJournal.EOF Then 
	
	'Display a link for each page in the Journal Items     	
	For intLinkPageNum = 1 to intTotalNumJournalPages		
		
		'If the page to be linked to is the page displayed then don't make it a hyper-link
		If intLinkPageNum = intRecordPositionPageNum Then
			Response.Write vbCrLf & "		     " & intLinkPageNum
		Else
		
			Response.Write vbCrLf & "		     <a href=""default.asp?PagePosition=" &  intLinkPageNum  & """ target=""_self"">" & intLinkPageNum & "</a>&nbsp;"			
		End If
	Next
End If


'If it is Not the End of the Journal Items entries then display a next link for the next Journal Items page      	
If NOT rsJournal.EOF then   	
	Response.Write vbCrLf & "		<a href=""default.asp?PagePosition=" &  intRecordPositionPageNum + 1  & """ target=""_self"">Next&nbsp;&gt;&gt;</a>"	   	
End If      	


'Finsh HTML the table 
%>
          </td>
        </tr>
      </table>
      </td>
  </tr>
</table>
<br>
<br>
<% 

'Reset server objects
rsJournal.Close
Set rsJournal = Nothing
Set rsCommentsCount = Nothing
Set strCon = Nothing
Set adoCon = Nothing
%>
<div align="center"> 
  <%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:11px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:11px"">Web Wiz Journal</a> version 1.0</span>")
	Response.Write("<br><span class=""text"" style=""font-size:11px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
 %>
</div>
<!--#include file="footer.inc" -->