
MAXIMIZE FIRST TO GET THE PROPER VIEW !!

'*******************************
'*                             *
'*     Database Access CTL     *
'*       Copyright ©2001       *
'*      Austin K. Hayward      *
'*                             *
'*******************************

Option Explicit


Dim strCompanyInformation As String
Dim strVersionNumber As String
Dim strMiscInformation As String
Dim strEmail As String


Private Sub ReadMe_Load()

    strCompanyInformation = "Copyright ©2001 Austin Hayward"

    strVersionNumber = "2.0.0"	       'Version 1.0.0 was a DLL, also available at PSC

    strEmail = "achayward@frontlinefocus.com"

    strMiscInformation = "This product is supplied as is.  Any changes or deletions due to " & _
			 "yourself seeing room for improvement should be emailed to Austin Hayward " & _
			 "so as to better improve the software for future releases." & vbcrlf & _
			 "" & _
			 "You may use this control freely in your projects, but please do not distribute it " & _
			 "without the expressed written consent of Austin Hayward." & vbcrlf & _
			 "" & _
			 "This product was designed exclusively with ADO 2.1 and will work with Access " & _
			 "and SQL Server databases quite freely." & _
			 "" & _
			 "Any comments or suggestions are always welcome.  Please email me at: " & strEmail

End Sub