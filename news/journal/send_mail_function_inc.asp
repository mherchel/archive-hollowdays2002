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


'Function to send an e-mail
Function SendMail(strEmailBody, strEmailAddress, strSubject, strMailComponent)

	'Dimension variables
	Dim objCDOSYSMail	'Holds the CDOSYS mail object
	Dim objCDOMail		'Holds the CDONTS mail object
	Dim objJMail		'Holds the Jmail object
	Dim objAspEmail		'Holds the Persits AspEmail email object
	Dim objAspMail		'Holds the Server Objects AspMail email object

	'Select which email component to use
	Select Case strMailComponent
	
		'CDOSYS mail component
		Case "CDOSYS"
			
			'Dimension variables
			Dim objCDOSYSCon
			
			'Create the e-mail server object
			Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		    	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
        		
	        	'Set and update fields properties
	        	'Out going SMTP server
	        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
	        	'SMTP port
	        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  = 25
	        	'CDO Port
	        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	        	'Timeout
	        	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        		objCDOSYSCon.Fields.Update 
		
			'Update the CDOSYS Configuration
			Set objCDOSYSMail.Configuration = objCDOSYSCon
			
			
			'Who the e-mail is from
			objCDOSYSMail.From = "<" & strEmailAddress & ">"
						
			'Who the e-mail is sent to
			objCDOSYSMail.To = "<" & strEmailAddress & ">"
								
			'The subject of the e-mail
			objCDOSYSMail.Subject = strSubject
						
			'The main body of the e-amil (HTML format)
			objCDOSYSMail.HTMLBody = strEmailBody
			
			'The main body of the e-amil (Plain text format)
			'objCDOSYSMail.TextBody = strEmailBody
						
			'Send the e-mail
			If NOT strSMTPServer = "" Then objCDOSYSMail.Send				
							
			'Close the server mail object
			Set objCDOSYSMail = Nothing
		
		'CDONTS mail component
		Case "CDONTS"
		
			'Create the e-mail server object
			Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
		
			'Who the e-mail is from
			objCDOMail.From = "<" & strEmailAddress & ">"
						
			'Who the e-mail is sent to
			objCDOMail.To = "<" & strEmailAddress & ">"
								
			'The subject of the e-mail
			objCDOMail.Subject = strSubject
						
			'The main body of the e-amil
			objCDOMail.Body = strEmailBody
						
			'Set the e-mail body format (0=HTML 1=Text)
			objCDOMail.BodyFormat = 0
			
			'Set the mail format (0=MIME 1=Text)
			objCDOMail.MailFormat = 0
						
			'Importance of the e-mail (0=Low, 1=Normal, 2=High)
			objCDOMail.Importance = 1 
						
			'Send the e-mail
			objCDOMail.Send				
							
			'Close the server mail object
			Set objCDOMail = Nothing
		
		
		'JMail component
		Case "Jmail"
	
			'Create the e-mail server object
			Set objJMail = Server.CreateObject("JMail.SMTPMail")
		
			'Out going SMTP mail server address
			objJMail.ServerAddress = strSMTPServer
		
			'Who the e-mail is from
			objJMail.Sender = strEmailAddress
						
			'Who the e-mail is sent to
			objJMail.AddRecipient strEmailAddress
									
			'The subject of the e-mail
			objJMail.Subject = strSubject
						
			'The main body of the e-amil
			objJMail.HTMLBody = strEmailBody
						
			'Importance of the e-mail
			objJMail.Priority = 3 
						
			'Send the e-mail
			If NOT strSMTPServer = "" Then objJMail.Execute				
							
			'Close the server mail object
			Set objJMail = Nothing
			
		'AspEmail component
		Case "AspEmail"
	
			'Create the e-mail server object
			Set objAspEmail = Server.CreateObject("Persits.MailSender")
			
			'Out going SMTP mail server address
			objAspEmail.Host = strSMTPServer
				
			'Who the e-mail is from
			objAspEmail.From = strEmailAddress
				
			'Who the e-mail is sent to
			objAspEmail.AddAddress strEmailAddress
												
			'The subject of the e-mail
			objAspEmail.Subject = strSubject
			
			'Set the e-mail body to HTML
			objAspEmail.IsHTML = True 
			
			'The main body of the e-mail
			objAspEmail.Body = strEmailBody
									
			'Send the e-mail
			If NOT strSMTPServer = "" Then objAspEmail.Send			
						
			'Close the server mail object
			Set objAspEmail = Nothing
		
		'AspMail component
		Case "AspMail"
	
			'Create the e-mail server object
			Set objAspMail = Server.CreateObject("SMTPsvg.Mailer")
			
			'Out going SMTP mail server address
			objAspMail.RemoteHost = strSMTPServer
				
			'Who the e-mail is from
			objAspMail.FromAddress = strEmailAddress
				
			'Who the e-mail is sent to
			objAspMail.AddRecipient " ", strEmailAddress
												
			'The subject of the e-mail
			objAspMail.Subject = strSubject
			
			'Set the e-mail body to HTML
			objAspMail.ContentType = "text/html" 
			
			'The main body of the e-mail
			objAspMail.BodyText = strEmailBody
									
			'Send the e-mail
			If NOT strSMTPServer = "" Then objAspMail.SendMail 			
						
			'Close the server mail object
			Set objAspMail = Nothing
	End Select	
	
	'Set the returned value of the function to true
	SendMail = True
End Function
%>