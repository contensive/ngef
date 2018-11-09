VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BioClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
Option Explicit
'
Const RequestNameButton = "button"
Const RequestNameFormID = "formid"
Const RequestNameSourceFormID = "sourceformid"
'
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNamePhone = "phone"
Const RequestNameEmail = "email"
Const RequestNameDetails = "details"
Const RequestNameSpam = "spmdtct"
'
Const FormButtonSubmit = "Submit"
'
Const FormContact = 100
Const FormThankYou = 102
'
Const ContentNamePeople = "People"
'
Const DataSource = "Default"
'
Const SystemEmailNotification = "Guard Muster Bio Update Notification"
Const SystemEmailAutoResponder = "Guard Muster Bio Update Auto Responder"
'
Private ButtonValue As String
Private WorkingQueryString As String
Private ErrorString As String
'
Private FormID As Long
Private SourceFormID As Long
'
Private FirstName As String
Private LastName As String
Private Phone As String
Private eMail As String
Private Details As String
'
Private Main As Object
'
'----------------------------------------------
'   Init()
'----------------------------------------------
Public Sub Init(MainObject As Object)
    '
    Set Main = MainObject
    '
    ButtonValue = Main.GetStreamText(RequestNameButton)
    WorkingQueryString = Main.RefreshQueryString
    '
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    '
    FirstName = Main.GetStreamText(RequestNameFirstName)
    LastName = Main.GetStreamText(RequestNameLastName)
    eMail = Main.GetStreamText(RequestNameEmail)
    Phone = Main.GetStreamText(RequestNamePhone)
    Details = Main.GetStreamText(RequestNameDetails)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("BioClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim LocalArgument As String
    '
    If Not (Main Is Nothing) Then
        '
        Call ProcessForm
        '
        Select Case FormID
            Case FormThankYou
                Stream = Stream & Main.GetContentCopy("Guard Muster Bio Update Thank You") & "<br />&nbsp;<br />"
            Case Else
                Stream = Stream & GetFormContact()
        End Select
        GetContent = Stream
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("BioClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'----------------------------------------------
'   Handles the source form post
'----------------------------------------------
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    Select Case SourceFormID
        Case FormContact
            Call ProcessFormContact
    End Select
    '
    Exit Sub
ErrorTrap:
    Call HandleError("BioClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormContact() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    '
    Stream = Stream & Main.GetFormStart()
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    '
    If ErrorString <> "" Then
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=2 width=""100%""><span class=""ccError"">" & ErrorString & "</span></TD>"
        Stream = Stream & "</TR>"
    Else
        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            FirstName = Main.GetCS(CSPointer, "FirstName")
            LastName = Main.GetCS(CSPointer, "LastName")
            eMail = Main.GetCS(CSPointer, "Email")
            Phone = Main.GetCS(CSPointer, "Phone")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%"">" & Main.GetContentCopy("RSVP Form Instructions") & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>First Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 35) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Last Name:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameLastName, LastName, , 35) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>eMail:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameEmail, eMail, , 35) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Phone Number:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNamePhone, Phone, , 35) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Updates:</TD>"
    Stream = Stream & "<TD width=""100%"">" & Main.GetFormInputText(RequestNameDetails, Details, 10, 30) & "</TD>"
    Stream = Stream & "</TR>"
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=175 height=1></TD>"
    Stream = Stream & "<TD width=100%>" & Main.GetFormInputHidden(RequestNameSpam, "") & Main.GetFormInputHidden(RequestNameSourceFormID, FormContact) & Main.GetFormButton(FormButtonSubmit) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetFormContact = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("BioClass", "GetFormContact", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormContact()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim EmailString As String
    Dim LocalGroupID As Long
    '
    If ButtonValue = FormButtonSubmit Then
        '
        '   Check for required
        '
        If FirstName = "" Then
            ErrorString = ErrorString & "Please enter your first name.<br />"
            End If
        If LastName = "" Then
            ErrorString = ErrorString & "Please enter your last name.<br />"
            End If
        If Not isValidEmailAddress(eMail) Then
            ErrorString = ErrorString & "Please enter a valid email address.<br />"
            End If
        If Phone = "" Then
            ErrorString = ErrorString & "Please provide your phone number.<br />"
            End If
        If Details = "" Then
            ErrorString = ErrorString & "Please provide your biographical information you would like updated.<br />"
            End If
        '
        If ErrorString = "" Then
            '
            If Main.GetStreamText(RequestNameSpam) = "" Then
                EmailString = EmailString & "<br /><br /><b>First Name:</b> " & FirstName & "<br />"
                EmailString = EmailString & "<b>Last Name:</b> " & LastName & "<br />"
                EmailString = EmailString & "<b>Email:</b> " & eMail & "<br />"
                EmailString = EmailString & "<b>Phone:</b> " & Phone & "<br />"
                EmailString = EmailString & "<b>Details:</b> " & Details & "<br />"
                '
                Call Main.SendSystemEmail(SystemEmailNotification, EmailString)
                Call Main.SendSystemEmail(SystemEmailAutoResponder, , Main.MemberID)
            End If
            FormID = FormThankYou
            '
            End If
        '
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("BioClass", "ProcessFormContact", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function isValidEmailAddress(EmailAddress As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim ATPosition As Long
    Dim Dotposition As Long
    '
    isValidEmailAddress = False
    '
    '   check for an @ anywhere
    '
    ATPosition = InStr(1, EmailAddress, "@")
    If ATPosition <> 0 Then
        '
        '   check for a dot past the found @ or first character
        '
        Dotposition = InStr(ATPosition, EmailAddress, ".")
        '
        '   make sure there is at least 1 character past the @
        '
        If (Dotposition > ATPosition + 2) And (ATPosition <> 1) Then
            isValidEmailAddress = True
            End If
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("RegistrationClass", "isValidEmailAddress", Err.Number, Err.Source, Err.Description, True, False)
    End Function
