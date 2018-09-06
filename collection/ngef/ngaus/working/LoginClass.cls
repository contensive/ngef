VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameUsername = "username"
Const RequestNamePassword = "password"
Const RequestNameEmail = "eMail"
'
Const RequestNameFormID = "formid"
'
Const FormButtonLogin = "Log In"
Const FormButtonSendPassword = "Send Password"
'
Private FormID As Long
Private SourceFormID As Long
Private ErrorString As String
Private CompletedFlag As Boolean
'
Private Main As Object
'Private Main As MainClass
'
'----------------------------------------------
'   Init()
'----------------------------------------------
Public Sub Init(MainObject As Object)
'Public Sub Init(MainObject As ccWeb3.MainClass)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("LoginClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetSiteLoginForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("LoginClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
Private Function GetSiteLoginForm() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim Title As String
    Dim FirstName As String
    Dim LastName As String
    Dim NameString As String
    '
    If Main.IsAuthenticated Then
        CSPointer = Main.OpenContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            Title = Main.GetCSText(CSPointer, "Title")
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            If Title <> "" Then
                NameString = Title
                End If
            NameString = NameString & " " & FirstName & " " & LastName
            End If
        Stream = Stream & "<span class=""TopNav"">Welcome <b>" & NameString & "</b></span><br /><br />"
        Stream = Stream & "If you are not " & NameString & "<br />"
        Stream = Stream & "<a href=""" & Main.ServerPage & "?method=logout"">Click Here</a><br /><br />"
        Stream = Stream & "<a href=""" & Main.ServerPage & "?method=logout""><b>LOG OUT</b></a>"
    Else
        Stream = Stream & Main.GetFormStart()
        Stream = Stream & "Username:<br />"
        Stream = Stream & Main.GetFormInputText(RequestNameUsername, , , 15) & "<br />"
        Stream = Stream & "Password:<br />"
        Stream = Stream & Main.GetFormInputText(RequestNamePassword, , , 15, , True) & "&nbsp;<input border=1 src=""" & Main.ServerAppRootPath & "images/login.gif"" name=""" & FormButtonLogin & """ width=56 height=19 type=image><br />"
        Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" height=10 width=10><br />"
        Stream = Stream & "Forgot Your Password?<br />"
        Stream = Stream & "eMail:<br />"
        Stream = Stream & Main.GetFormInputText(RequestNameEmail, , , 15) & "&nbsp;<input border=1 src=""" & Main.ServerAppRootPath & "images/send.gif"" name=""" & FormButtonSendPassword & """ width=56 height=19 type=image><br />"
        Stream = Stream & Main.GetFormEnd()
        End If
    '
    GetSiteLoginForm = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("LoginClass", "GetSiteLoginForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function

