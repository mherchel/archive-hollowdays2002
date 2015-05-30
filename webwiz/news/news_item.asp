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




Dim rsNews	'Database recordset holding the news items
Dim rsComments	'Database recordset holding the comments for this news item
Dim lngNewsID	'Holds the News item ID number
Dim blnComments	'Set to true if comments are allowed for this item

'Read in the ID number of the news item we are looking at the comments of
If isNull(Request.QueryString("NewsID")) = True Or isNumeric(Request.QueryString("NewsID")) = False Then
	Response.Write "default.asp"
Else
	lngNewsID = CLng(Request.QueryString("NewsID"))
	
End If
%>
<html>
<head>
<title>News Item</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">

<!-- Web Wiz Guide - Web Wiz Site News is written by Bruce Corkhill ©2001-2002
    	 If you want your own Web Wiz Site News then goto http://www.webwizguide.info --> 

<!-- #include file="header.inc" -->
<br>
<div align="center">
 <a href="default.asp" target="_self">View All News Items</a> <br>
 <br><%

'Create recorset object
Set rsNews = Server.CreateObject("ADODB.Recordset")
	
'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
strSQL = "SELECT tblNews.* FROM tblNews WHERE tblNews.News_ID = " & lngNewsID & " ORDER BY News_Date DESC;"
	
'Query the database
rsNews.Open strSQL, adoCon


'If there are no records then exit for loop
If NOT rsNews.EOF Then
	
	'Read in if comments are allowed for this news item
	blnComments = CBool(rsNews("Comments"))
	
	%>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 <tr> 
  <td class="heading"><% = rsNews("News_title") %> <span class="text"><em>- <% = FormatDateTime(rsNews("News_Date"), vbLongDate) %> at <% = FormatDateTime(rsNews("News_Date"), vbShortTime) %><em></span>
   </td>
 </tr>
 <tr> 
  <td class="text"><% = rsNews("News_item") %></td>
 </tr>
</table>
<br>
<%
End If

'Clean up
rsNews.Close
Set rsNews = Nothing



'If comments are allowed then show any comments and a text form
If blnComments = True Then

%>
<br>
<br>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="1">
 <tr>
  <td align="center" class="heading">Comments</td>
 </tr>
</table><%

	'Create recorset object
	Set rsComments = Server.CreateObject("ADODB.Recordset")
		
	'Initalise the strSQL variable with an SQL statement to query the database by selecting all tables ordered by the decending date
	strSQL = "SELECT tblComments.* FROM tblComments WHERE tblComments.News_ID = " & lngNewsID & " ORDER BY Comments_Date DESC;"
		
	'Query the database
	rsComments.Open strSQL, adoCon
	
	'Loop round to display all the comments for the news item
	Do While NOT rsComments.EOF

%>
<table width="100%" border="0" cellspacing="1" cellpadding="2">
 <tr> 
  <td class="text"><strong>Comments by <%
  
  'If there is an email address entered make it a mailto link
  If rsComments("EMail") <> "" Then Response.Write("<a href=""mailto:" & rsComments("EMail") & """>" & rsComments("Name") & "</a>") Else Response.Write(rsComments("Name")) 
  
  %> from <% = rsComments("Country") %> on <% = FormatDateTime(rsComments("Comments_Date"), VbLongDate) %> at <% = FormatDateTime(rsComments("Comments_Date"), VbShortTime) %></strong><span class="smText"> - IP Logged</span></td>
 </tr>
 <tr> 
  <td class="text"><% = rsComments("Comments") %>
  </td>
 </tr>
</table>
<br>
<%
	'Move to the next record in the recordset
	rsComments.MoveNext

Loop

	'Reset server objects
	rsComments.Close
	Set rsComments = Nothing
	Set strCon = Nothing
	Set adoCon = Nothing	
%>
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Intialise variables
	var errorMsg = "";
	var errorMsgLong = "";

	//Check for a name
	if (document.frmNewsComments.name.value == ""){
		errorMsg += "\n\tName \t\t- Enter your Name";
	}
	
	//Check for a country
	if (document.frmNewsComments.country.value == "0"){
		errorMsg += "\n\tCountry \t\t- Select the country you are in";
	}
	
	//Check for comments
	if (document.frmNewsComments.comments.value == ""){
		errorMsg += "\n\tComments \t- Enter a comment to add to the News";
	}
	
	//Check the description length before submiting the form	
	if (document.frmNewsComments.comments.value.length > <% = intMsgCharNo %>){
		errorMsgLong += "\n- Your comments are " + document.frmNewsComments.comments.value.length + " chracters long, they need to be shortned to below <% = intMsgCharNo %> chracters.";
	}
	
	//Check the word length before submitting
	words = document.frmNewsComments.comments.value.split(' ');
	for (var loop = 0; loop <= words.length - 1; ++loop){
		if (words[loop].length >= 50){
		errorMsgLong += "\n- A word in your comments contains " + words[loop].length + " characters, this needs to be shortened to below 50 characters.";
		}	
	}	
		
	//Check for HTML tags before submitting the form	
	for (var count = 0; count <= 7; ++count){
		if ((document.frmNewsComments.elements[count].value.indexOf("<", 0) >= 0) && (document.frmNewsComments.elements[count].value.indexOf(">", 0) >= 0)){
			errorMsgLong += "\n- HTML tags are not permitted, remove all HTML tags.";
		}			
	}
	
	//If there is aproblem with the form then display an error
	if ((errorMsg != "") || (errorMsgLong != "")){
		msg = "___________________________________________________________________\n\n";
		msg += "Your Comments have not been added because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "___________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";
		
		errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
		return false;
	}
	
	return true;
}

//Function to count the number of characters in the description text box
function DescriptionCharCount() {
	document.frmNewsComments.countcharacters.value = document.frmNewsComments.comments.value.length;	
}

// Function to add the code for bold italic and underline, to the message
function AddMessageCode(code,promptText, InsertText) {

	if (code != "") {
		insertCode = prompt(promptText + "\n[" + code + "]xxx[/" + code + "]", InsertText);
			if ((insertCode != null) && (insertCode != "")){
				document.frmNewsComments.comments.value += "[" + code + "]" + insertCode + "[/"+ code + "] ";
			}
	}		
	document.frmNewsComments.comments.focus();
}


//Function to add the code to the message for the smileys
function AddSmileyIcon(iconCode) {	
		document.frmNewsComments.comments.value += iconCode + " ";
		document.frmNewsComments.comments.focus();
}

//Function to open pop up window
function openWin(theURL,winName,features) {
  	window.open(theURL,winName,features);
}
// -->
</script>

<br>
<form method=post name="frmNewsComments" action="add_comments.asp?NewsID=<% = lngNewsID %>" onSubmit="return CheckForm();">
 <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
   <td width="83%"> <table width="100%" border="0" cellspacing="0" cellpadding="1" align="center" bgcolor="<% = strTableBorderColour %>">
     <tr> 
      <td height="280"> <table width="100%" border="0" align="center" height="164" bgcolor="<% = strTableTitleColour %>" cellpadding="2" cellspacing="0">
        <tr align="left"> 
         <td colspan="2" class="text" height="30">*Indicates required fields</td>
        </tr>
        <tr> 
         <td align="right" width="19%" height="14" class="text">Name*: </td>
         <td height="14" width="81%"> <input type="text" name="name" size="30" maxlength="30" > <a href="http://www.webwizguide.info"></a> </td>
        </tr>
        <tr class="arial"> 
         <td align="right" width="19%" class="text" height="12">Country*:</td>
         <td height="12" width="81%"> <select name=country>
           <option value="0" selected>Pull down to select</option>
           <option>United Kingdom</option>
           <option>United States</option>
           <option>Afghanistan</option>
           <option>Albania</option>
           <option>Algeria</option>
           <option>American Samoa</option>
           <option>Andorra</option>
           <option>Angola</option>
           <option>Anguilla</option>
           <option>Antarctica</option>
           <option>Antigua And Barbuda</option>
           <option>Argentina</option>
           <option>Armenia</option>
           <option>Aruba</option>
           <option>Australia</option>
           <option>Austria</option>
           <option>Azerbaijan</option>
           <option>Bahamas</option>
           <option>Bahrain</option>
           <option>Bangladesh</option>
           <option>Barbados</option>
           <option>Belarus</option>
           <option>Belgium</option>
           <option>Belize</option>
           <option>Benin</option>
           <option>Bermuda</option>
           <option>Bhutan</option>
           <option>Bolivia</option>
           <option>Bosnia Hercegovina</option>
           <option>Botswana</option>
           <option>Bouvet Island</option>
           <option>Brazil</option>
           <option>Brunei Darussalam</option>
           <option>Bulgaria</option>
           <option>Burkina Faso</option>
           <option>Burundi</option>
           <option>Byelorussian SSR</option>
           <option>Cambodia</option>
           <option>Cameroon</option>
           <option>Canada</option>
           <option>Cape Verde</option>
           <option>Cayman Islands</option>
           <option>Central African Republic</option>
           <option>Chad</option>
           <option>Chile</option>
           <option>China</option>
           <option>Christmas Island</option>
           <option>Cocos (Keeling) Islands</option>
           <option>Colombia</option>
           <option>Comoros</option>
           <option>Congo</option>
           <option>Cook Islands</option>
           <option>Costa Rica</option>
           <option>Cote D'Ivoire</option>
           <option>Croatia</option>
           <option>Cuba</option>
           <option>Cyprus</option>
           <option>Czech Republic</option>
           <option>Czechoslovakia</option>
           <option>Denmark</option>
           <option>Djibouti</option>
           <option>Dominica</option>
           <option>Dominican Republic</option>
           <option>East Timor</option>
           <option>Ecuador</option>
           <option>Egypt</option>
           <option>El Salvador</option>
           <option>England</option>
           <option>Equatorial Guinea</option>
           <option>Eritrea</option>
           <option>Estonia</option>
           <option>Ethiopia</option>
           <option>Falkland Islands</option>
           <option>Faroe Islands</option>
           <option>Fiji</option>
           <option>Finland</option>
           <option>France</option>
           <option>Gabon</option>
           <option>Gambia</option>
           <option>Georgia</option>
           <option>Germany</option>
           <option>Ghana</option>
           <option>Gibraltar</option>
           <option>Great Britain</option>
           <option>Greece</option>
           <option>Greenland</option>
           <option>Grenada</option>
           <option>Guadeloupe</option>
           <option>Guam</option>
           <option>Guatemela</option>
           <option>Guernsey</option>
           <option>Guiana</option>
           <option>Guinea</option>
           <option>Guinea-Bissau</option>
           <option>Guyana</option>
           <option>Haiti</option>
           <option>Heard Islands</option>
           <option>Honduras</option>
           <option>Hong Kong</option>
           <option>Hungary</option>
           <option>Iceland</option>
           <option>India</option>
           <option>Indonesia</option>
           <option>Iran</option>
           <option>Iraq</option>
           <option>Ireland</option>
           <option>Isle Of Man</option>
           <option>Israel</option>
           <option>Italy</option>
           <option>Jamaica</option>
           <option>Japan</option>
           <option>Jersey</option>
           <option>Jordan</option>
           <option>Kazakhstan</option>
           <option>Kenya</option>
           <option>Kiribati</option>
           <option>Korea, South</option>
           <option>Korea, North</option>
           <option>Kuwait</option>
           <option>Kyrgyzstan</option>
           <option>Lao People's Dem. Rep.</option>
           <option>Latvia</option>
           <option>Lebanon</option>
           <option>Lesotho</option>
           <option>Liberia</option>
           <option>Libya</option>
           <option>Liechtenstein</option>
           <option>Lithuania</option>
           <option>Luxembourg</option>
           <option>Macau</option>
           <option>Macedonia</option>
           <option>Madagascar</option>
           <option>Malawi</option>
           <option>Malaysia</option>
           <option>Maldives</option>
           <option>Mali</option>
           <option>Malta</option>
           <option>Marshall Islands</option>
           <option>Martinique</option>
           <option>Mauritania</option>
           <option>Mauritius</option>
           <option>Mayotte</option>
           <option>Mexico</option>
           <option>Micronesia</option>
           <option>Moldova</option>
           <option>Monaco</option>
           <option>Mongolia</option>
           <option>Montserrat</option>
           <option>Morocco</option>
           <option>Mozambique</option>
           <option>Myanmar</option>
           <option>Namibia</option>
           <option>Nauru</option>
           <option>Nepal</option>
           <option>Netherlands</option>
           <option>Netherlands Antilles</option>
           <option>Neutral Zone</option>
           <option>New Caledonia</option>
           <option>New Zealand</option>
           <option>Nicaragua</option>
           <option>Niger</option>
           <option>Nigeria</option>
           <option>Niue</option>
           <option>Norfolk Island</option>
           <option>Mariana Islands</option>
           <option>Norway</option>
           <option>Oman</option>
           <option>Pakistan</option>
           <option>Palau</option>
           <option>Panama</option>
           <option>Papua New Guinea</option>
           <option>Paraguay</option>
           <option>Peru</option>
           <option>Philippines</option>
           <option>Pitcairn</option>
           <option>Poland</option>
           <option>Polynesia</option>
           <option>Portugal</option>
           <option>Puerto Rico</option>
           <option>Qatar</option>
           <option>Reunion</option>
           <option>Romania</option>
           <option>Russian Federation</option>
           <option>Rwanda</option>
           <option>Saint Helena</option>
           <option>Saint Kitts</option>
           <option>Saint Lucia</option>
           <option>Saint Pierre</option>
           <option>Saint Vincent</option>
           <option>Samoa</option>
           <option>San Marino</option>
           <option>Sao Tome and Principe</option>
           <option>Saudi Arabia</option>
           <option>Senegal</option>
           <option>Seychelles</option>
           <option>Sierra Leone</option>
           <option>Singapore</option>
           <option>Slovakia</option>
           <option>Slovenia</option>
           <option>Solomon Islands</option>
           <option>Somalia</option>
           <option>South Africa</option>
           <option>South Georgia</option>
           <option>Spain</option>
           <option>Sri Lanka</option>
           <option>Sudan</option>
           <option>Suriname</option>
           <option>Svalbard</option>
           <option>Swaziland</option>
           <option>Sweden</option>
           <option>Switzerland</option>
           <option>Syrian Arab Republic</option>
           <option>Taiwan</option>
           <option>Tajikista</option>
           <option>Tanzania</option>
           <option>Thailand</option>
           <option>Togo</option>
           <option>Tokelau</option>
           <option>Tonga</option>
           <option>Trinidad and Tobago</option>
           <option>Tunisia</option>
           <option>Turkey</option>
           <option>Turkmenistan</option>
           <option>Turks and Caicos Islands</option>
           <option>Tuvalu</option>
           <option>Uganda</option>
           <option>Ukraine</option>
           <option>United Arab Emirates</option>
           <option>United Kingdom</option>
           <option>United States</option>
           <option>Uruguay</option>
           <option>Uzbekistan</option>
           <option>Vanuatu</option>
           <option>Vatican City State</option>
           <option>Venezuela</option>
           <option>Vietnam</option>
           <option>Virgin Islands</option>
           <option>Western Sahara</option>
           <option>Yemen</option>
           <option>Yugoslavia</option>
           <option>Zaire</option>
           <option>Zambia</option>
           <option>Zimbabwe</option>
          </select> </td>
        </tr>
        <tr> 
         <td align="right" width="19%" class="text" height="12">E-mail:</td>
         <td height="12" width="81%"> <input type="text" name="email" size="30" maxlength="50"> </td>
        </tr>
        <tr><a href="http://www.webwizguide.info"></a> 
         <td valign="top" align="right" height="31" width="19%" class="text">&nbsp;</td>
         <td height="31" width="81%" valign="bottom"> <a href="JavaScript:AddMessageCode('B','Enter text you want formatted in Bold', '')"><img src="news_images/post_button_bold.gif" width="25" height="24" align="absmiddle" border="0" alt="Bold"></a> <a href="JavaScript:AddMessageCode('I','Enter text you want formatted in Italic', '')"><img src="news_images/post_button_italic.gif" width="25" height="24" align="absmiddle" border="0" alt="Italic"></a> <a href="JavaScript:AddMessageCode('U','Enter text you want Underlined', '')"><img src="news_images/post_button_underline.gif" width="25" height="24" align="absmiddle" border="0" alt="Underline"></a> 
          <a href="javascript:openWin('emoticons_smilies.asp','smilies','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=400,height=400')"><img src="news_images/post_button_smiley.gif" width="25" height="24" align="absmiddle" alt="Emoticon Smilies" border="0"></a></td>
        </tr>
        <tr> 
         <td valign="top" align="right" height="61" width="19%" class="text">Comments*:<br> <span style="font-size: 10px;">(max. 
          <% = intMsgCharNo %>
          characters)</span></td>
         <td height="61" width="81%" valign="top"> <textarea name="comments" cols="40" rows="6" onKeyDown="DescriptionCharCount();" onKeyUp="DescriptionCharCount();"></textarea> </td>
        </tr>
        <tr> 
         <td valign="top" align="right" height="2" width="19%" class="text">Character Count: </td>
         <td height="2" width="81%"> <input size="5" value="0" name="countcharacters" maxlength="5"> <input onClick="DescriptionCharCount();" type="button" value="Count" name="Count"> </td>
        </tr>
        <tr> 
         <td valign="top" align="right" height="2" width="19%" class="text">&nbsp; </td>
         <td height="2" width="81%"> <p> 
           <input type="submit" name="Submit" value="Submit Comments">
           <input type="reset" name="Reset" value="Clear Form">
          </p></td>
        </tr>
       </table></td>
     </tr>
    </table></td>
  </tr>
 </table>
</form>
<%
End If
%>
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