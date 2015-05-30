<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="send_mail_function.inc" -->
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



'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'***********************************************
'Function to strip non alphanumeric characters for links and email addresses
Private Function characterStrip(strTextInput)

	'Dimension variable
	Dim intLoopCounter 	'Holds the loop counter
	
	'Loop through the ASCII characters
	For intLoopCounter = 0 to 37
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters
	For intLoopCounter = 39 to 44
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters numeric characters to lower-case characters
	For intLoopCounter = 65 to 94
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the extended ASCII characters
	For intLoopCounter = 123 to 125
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the extended ASCII characters
	For intLoopCounter = 127 to 255
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Strip individul ASCII characters left out from above left over
	strTextInput = Replace(strTextInput, CHR(59), "", 1, -1, 0)
	strTextInput = Replace(strTextInput, CHR(60), "", 1, -1, 0)
	strTextInput = Replace(strTextInput, CHR(62), "", 1, -1, 0)
	strTextInput = Replace(strTextInput, CHR(96), "", 1, -1, 0)
	
	
	'Return the string
	characterStrip = strTextInput
	
End Function
'*******************************************************



'Dimension variables
Dim rsSmut 				'Database Recordset holding the smut table
Dim rsAddNewsComments			'Database recordset to add new comments
Dim strInputName 			'Holds the Users name
Dim strInputCountry 			'Holds the users country
Dim strInputEmailAddress		'Holds the Users e-mail address
Dim strInputComments 			'Holds the Users comments
Dim saryCommentWord 			'Array to hold each word in the comments enetred by the user
Dim intCheckWordLengthLoopCounter	'Loop counter
Dim intWordLength			'Holds the length of the word to be checked
Dim blnWordLenthOK			'Boolean set to False if any words in the description are above 30 characters
Dim intLongestWordLength 		'Holds the number of characters in the longest word entered in the description
Dim lngNewsID				'Holds the News item ID number
Dim strEmailSubject			'Holds the subject of the e-mail notification
Dim strEmailBody			'Holds the body of the e-mail
Dim blnEmailSent			'Set to tru if the e-mail is sent
Dim blnAlreadyPostsed			'Set to true if the person has already posted comments in for this news item


'Read in the ID number of the news item we are looking at the comments of
If isNull(Request.QueryString("NewsID")) = True Or isNumeric(Request.QueryString("NewsID")) = False Then
	Response.Write "news_comments.asp"
Else
	lngNewsID = CLng(Request.QueryString("NewsID"))
	
End If


'Read in user deatils from the comments form
strInputName = Trim(Mid(Request.Form("name"), 1, 30))
strInputCountry = Trim(Mid(Request.Form("country"), 1, 40))
strInputEmailAddress = Trim(Mid(Request.Form("email"), 1, 50))
strInputComments = Trim(Request.Form("comments"))


'Strip HTML tags
strInputName = Replace(strInputName, "<", "&lt;", 1, -1, 1)
strInputName = Replace(strInputName, ">", "&gt;", 1, -1, 1)
strInputComments = Replace(strInputComments, "<", "&lt;", 1, -1, 1)
strInputComments = Replace(strInputComments, ">", "&gt;", 1, -1, 1)

'Strip malicious code from the homepage and email links
strInputEmailAddress = characterStrip(LCase(strInputEmailAddress))



'Split-up each word in the comments from the user to check that no word entered is over 50 characters
saryCommentWord = Split(Trim(strInputComments), " ")
	
'Initialse the word length variable
blnWordLenthOK = True
	
'Loop round to check that each word in the comments entered by the user is not above 50 characters
For intCheckWordLengthLoopCounter = 0 To UBound(saryCommentWord)
	
	'Initialise the intWordLength variable with the length of the word to be searched
	intWordLength = Len(saryCommentWord(intCheckWordLengthLoopCounter))
	
	'Get the number of characters in the longest word
	If intWordLength => intLongestWordLength Then 
		intLongestWordLength = intWordLength
	End If
		
	'If the word length to be searched is more than or equal to 50 then set the blnWordLegthOK to false
	If intWordLength => 50 Then 
		blnWordLenthOK = False				
	End If
Next



'Change my own codes for bold and italic HTML tags back to the normal satndrd HTML tags now that the check for unwated HTML tags is over
strInputComments = Replace(strInputComments, "[B]", "<b>", 1, -1, 1)
strInputComments = Replace(strInputComments, "[/B]", "</b>", 1, -1, 1)
strInputComments = Replace(strInputComments, "[I]", "<i>", 1, -1, 1)
strInputComments = Replace(strInputComments, "[/I]", "</i>", 1, -1, 1)
strInputComments = Replace(strInputComments, "[U]", "<u>", 1, -1, 1)
strInputComments = Replace(strInputComments, "[/U]", "</u>", 1, -1, 1)

'Change the emotion symbols for the path to the relative smiley icon
strInputComments = Replace(strInputComments, "[:)]", "<img src=""news_images/smiley1.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[;)]", "<img src=""news_images/smiley2.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:o]", "<img src=""news_images/smiley3.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:D]", "<img src=""news_images/smiley4.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:errr:]", "<img src=""news_images/smiley5.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:(]", "<img src=""news_images/smiley6.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:x]", "<img src=""news_images/smiley7.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:o)]", "<img src=""news_images/smiley8.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:oops:]", "<img src=""news_images/smiley9.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:star:]", "<img src=""news_images/smiley10.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[xx(]", "<img src=""news_images/smiley11.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[|)]", "<img src=""news_images/smiley12.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:V:]", "<img src=""news_images/smiley13.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[:^:]", "<img src=""news_images/smiley14.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[}:)]", "<img src=""news_images/smiley15.gif"" border=""0"">", 1, -1, 1)
strInputComments = Replace(strInputComments, "[8D]", "<img src=""news_images/smiley16.gif"" border=""0"">", 1, -1, 1)


'Replace the vb new line code for the HTML new break code
strInputComments = Replace(strInputComments, vbCrLf, "<br>")

'Get rid of repeated return key hits so there arn't two many new lines going half way down the page (<br> is the HTML tag for new line)
'Loop though the comments entered by the user till all cases of two new lines togather are replaced by one new line
Do While InStr(1, strInputComments, "<br><br>" ,vbTextCompare) > 0
	
	'Replace <br><br> with one case of <br>
	strInputComments = Replace(strInputComments , "<br><br>", "<br>")
Loop

'Create recordset object
Set rsSmut = Server.CreateObject("ADODB.Recordset")

'Replace swear words with other words with ***	
'Initalise the SQL string with a query to read in all the words from the smut table
strSQL = "SELECT tblSmut.* FROM tblSmut;"

'Open the recordset
rsSmut.Open strSQL, strCon

'Loop through all the words to check for
Do While NOT rsSmut.EOF
	
	'Replace the swear words with the words in the database the swear words
	strInputComments = Replace(strInputComments, rsSmut("Smut"), rsSmut("Word_replace"), 1, -1, 1)
	strInputName = Replace(strInputName, rsSmut("Smut"), rsSmut("Word_replace"), 1, -1, 1)
	strInputCountry = Replace(strInputCountry, rsSmut("Smut"), rsSmut("Word_replace"), 1, -1, 1)

	'Move to the next word in the recordset
	rsSmut.MoveNext
Loop

'Reset recordset
rsSmut.Close
Set rsSmut = Nothing



'Create recorset object
Set rsAddNewsComments = Server.CreateObject("ADODB.Recordset")
	
'Initalise the SQL string with a query to read in all the comments from the database
strSQL = "SELECT TOP 1 tblComments.*, tblNews.News_title FROM tblNews INNER JOIN tblComments ON tblNews.News_ID = tblComments.News_ID WHERE tblComments.News_ID = " & lngNewsID & " ORDER BY tblComments.Comment_ID DESC;"
	
'Set the cursor type property of the record set to Dynamic so we can navigate through the record set
rsAddNewsComments.CursorType = 2
	
'Set the Lock Type for the records so that the record set is only locked when it is updated
rsAddNewsComments.LockType = 3
	
'Open the recordset
rsAddNewsComments.Open strSQL, strCon


'If cookies anti spam settings are enabled check a cookie has not already been set
If blnCookieSet = True Then
	If CBool(Request.Cookies("WWGNews")("Comments" & lngNewsID)) = True Then blnAlreadyPostsed = True
End If

'If IP blooking ant-spam settings are enabled check the IP address of the last poster
If blnIPBlocking = True Then
	If NOT rsAddNewsComments.EOF Then 
		If rsAddNewsComments("IP") = Request.ServerVariables("REMOTE_ADDR") Then blnAlreadyPostsed = True 
	End If
End If


'Write to the database if there are no unwanted HTML tags or the word lengths in the commets entered by the user are OK
If blnWordLenthOK = True AND blnAlreadyPostsed = False Then
	
	'Add a new record to the recordset
	rsAddNewsComments.AddNew
	
	rsAddNewsComments.Fields("Name") =  strInputName
	rsAddNewsComments.Fields("Country") = strInputCountry
	rsAddNewsComments.Fields("EMail") = strInputEmailAddress
	rsAddNewsComments.Fields("Comments") = strInputComments
	rsAddNewsComments.Fields("News_ID") = lngNewsID
	rsAddNewsComments.Fields("IP") = Request.ServerVariables("REMOTE_ADDR")
				
	'Update the database with the new recordset
	rsAddNewsComments.Update
	
	'Requery the database to make sure that the new comments have been added 
	'This will make the script wait until Database has updated itself as sometimes Access can be a little slow at updating
	rsAddNewsComments.Requery
	
	'If cookies anti-spam settings are enabled set a cookie on the users machine
	If blnCookieSet = True Then
		Response.Cookies("WWGNews")("Comments" & lngNewsID) = True
		Response.Cookies("WWGNews").Expires = DateAdd("n", 30, Now())
	End If
		
	
	'If the news is configured to send an e-mail then send one
	If blnEmail = True Then
		
		'Turn the smiley image paths back into text :)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley1.gif"" border=""0"">", ":)", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley2.gif"" border=""0"">", ";)", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley3.gif"" border=""0"">", ":o", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley4.gif"" border=""0"">", ":D", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley5.gif"" border=""0"">", ":errr:", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley6.gif"" border=""0"">", ":(", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley7.gif"" border=""0"">", ":x", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley8.gif"" border=""0"">", ":o)", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley9.gif"" border=""0"">", "[:oops:]", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley10.gif"" border=""0"">", ":X:", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley11.gif"" border=""0"">", "xx(", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley12.gif"" border=""0"">", "|)", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley13.gif"" border=""0"">", ":V:", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley14.gif"" border=""0"">", ":^:", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley15.gif"" border=""0"">", "}:)", 1, -1, 1)
		strInputComments = Replace(strInputComments, "<img src=""news_images/smiley16.gif"" border=""0"">", "8D", 1, -1, 1)
	
		'Initilise the subject of the e-mail
		strEmailSubject = "Site News Comment Notification"
	
		'Initailise the e-mail body variable with the body of the e-mail
		strEmailBody = "Hi "
		strEmailBody = strEmailBody & "<br><br>This e-mail is automactically generated by the Site News on your web site."
		strEmailBody = strEmailBody & "<br>The following comment has been posted in the News Item, " & rsAddNewsComments.Fields("News_title") & ": -"
		strEmailBody = strEmailBody & "<br><br><b>Name: </b>" & strInputName
		strEmailBody = strEmailBody & "<br><b>E-Mail: </b>" & strInputEmailAddress
		strEmailBody = strEmailBody & "<br><b>Country: </b>" & strInputCountry
		strEmailBody = strEmailBody & "<br><b>Comments: -</b><br>" & strInputComments
		
	
		'Call the funtion to send the e-mail
		blnEmailSent = SendMail(strEmailBody, strWebSiteEmailAddress, strEmailSubject, strMailComponent)
	End If
	
		 
	'Reset Sever Objects
	rsAddNewsComments.Close
	Set rsAddNewsComments = Nothing 
	Set adoCon = Nothing
	Set strCon = Nothing
	

	'Return to the comments page
	Response.Redirect "news_item.asp?NewsID=" & lngNewsID

End If

'Reset Sever Objects
rsAddNewsComments.Close
Set rsAddNewsComments = Nothing 
Set adoCon = Nothing
Set strCon = Nothing
%>
<HTML> 
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
<TITLE>Sign the Guest Book</TITLE>

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<!--#include file="header.inc" -->
<div align="center"><span class="heading">News Item Comments</span><br>
<a href="news_item.asp?NewsID=<% = Request.QueryString("NewsID") %>" target="_self">Return to the the News Item</a></div>
<div align="center">
 <table width="100%" height="178" border="0" cellpadding="1" cellspacing="0">
  <tr>
   <td align="center"><% 
'If word length is to long display an error message
If blnAlreadyPostsed = True Then %>
   <span class="text">Our records show that you have already posted comments for this News Item</span><br><%

'If the user has already posted display an error message
Else %>
    <span class="text">Sorry, one or more of the words used in your Comments where to long</span><br>
    <br>
    <a href="javascript:history.back(1);">Edit my comments</a><br><%
End If %>
    <br><br><br>
   </td>
  </tr>
 </table>
 <br>
 <%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Powered by <a href=""http://www.webwizguide.info"" target=""_blank"" style=""font-size:10px"">Web Wiz Site News</a> version 3.06</span>")
	Response.Write("<br><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2002 Web Wiz Guide</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
 %>
 <br>
</div>
<!--#include file="footer.inc" -->