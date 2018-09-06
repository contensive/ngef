VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DonationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFormID = "formid"
Const RequestNameSourceFormID = "sourceformid"
Const RequestNameButtonValue = "button"
'
Const ContentNamePeople = "People"
Const ContentNameDonations = "Donations"
'
Const FormDonate = 100
Const FormThankYou = 300
'
Const FormButtonSubmit = "Contribute"
'
Const FormSpacer = "<img src=""/ccLib/images/spacer.gif"" width=150 height=1>"
Const ButtonSpacer = "&nbsp;&nbsp;&nbsp;"
'
Const SystemEmailAutoResponder = "Donation Auto Responder"
Const SystemEmailNotification = "Donation Notification"
'
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameAddress = "address"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNamePhone = "phone"
Const RequestNameEmail = "email"
Const RequestNameMemory = "memory"
Const RequestNameAmount = "amount"
Const RequestNameCardNumber = "cardnumber"
Const RequestNameCardExpiration = "cardexpiration"
'
Private formID As Long
Private sourceFormID As Long
Private ButtonValue As String
'
Private FirstName As String
Private LastName As String
Private Address As String
Private City As String
Private State As String
Private Zip As String
Private Phone As String
Private eMail As String
Private memory As String
Private amount As String
Private CardNumber As String
Private CardExpiration As String
'
Private ErrorString As String
'
Private Main As Object
'
Public Sub Init(MainObject As Variant)
    '
    Set Main = MainObject
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        Call loadForm
        Call ProcessForm
        '
        GetContent = GetForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("DonationClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub loadForm()
    On Error GoTo ErrorTrap
    '
    formID = Main.GetStreamInteger(RequestNameFormID)
    sourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    ButtonValue = Main.GetStreamText(RequestNameButtonValue)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    Select Case sourceFormID
        Case FormDonate
            Call LoadFormDonate
            Call ProcessFormDonate
    End Select
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetForm() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    Select Case formID
        Case FormThankYou
            stream = stream & Main.GetContentCopy("Donation Thank You")
            stream = stream & "<br />"
        Case Else
            Call LoadFormDonate
            formID = FormDonate
            stream = stream & GetFormDonate
    End Select
    '
    GetForm = stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetPageHead(PageName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=""100%"">"
    '
    If ErrorString <> "" Then
        stream = stream & "<span class=""ccError"">" & ErrorString & "</span>"
    Else
        stream = stream & Main.GetContentCopy(PageName)
        End If
    '
    stream = stream & "</TD>"
    stream = stream & "</TR>"
    '
    GetPageHead = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DonationClass", "GetPageHead", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow(FieldCaption As String, Innards As String, Optional AlignLeft As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim AlignmentString As String
    '
    If Not AlignLeft Then
        AlignmentString = " align=right"
    Else
        AlignmentString = " align=left"
        End If
    '
    stream = stream & "<TR>"
    stream = stream & "<TD " & AlignmentString & ">" & FieldCaption & "</TD>"
    stream = stream & "<TD width=""100%"">" & Innards & "</TD>"
    stream = stream & "</TR>"
    '
    GetFormRow = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DonationClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow2(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    stream = stream & "<TR>"
    stream = stream & "<TD colspan=2 width=""100%"">" & Innards & "</TD>"
    stream = stream & "</TR>"
    '
    GetFormRow2 = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DonationClass", "GetFormRow2", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormDonate() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim ButtonString As String
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, formID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonSubmit)
    '
    stream = stream & Main.GetFormStart
    stream = stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    stream = stream & GetPageHead("Donation Instructions")
    '
    stream = stream & GetFormRow(GetProperty("First Name"), GetField(RequestNameFirstName, FirstName))
    stream = stream & GetFormRow(GetProperty("Last Name"), GetField(RequestNameLastName, LastName))
    stream = stream & GetFormRow(GetProperty("Address"), GetField(RequestNameAddress, Address))
    stream = stream & GetFormRow(GetProperty("City"), GetField(RequestNameCity, City))
    stream = stream & GetFormRow(GetProperty("State"), GetField(RequestNameState, State))
    stream = stream & GetFormRow(GetProperty("Zip"), GetField(RequestNameZip, Zip))
    stream = stream & GetFormRow(GetProperty("Phone"), GetField(RequestNamePhone, Phone))
    stream = stream & GetFormRow(GetProperty("Email"), GetField(RequestNameEmail, eMail))
    stream = stream & GetFormRow(GetProperty("In Memory Of"), GetField(RequestNameMemory, memory))
    stream = stream & GetFormRow(GetProperty("Amount"), GetField(RequestNameAmount, amount))
    stream = stream & GetFormRow(GetProperty("Card Number"), GetField(RequestNameCardNumber, CardNumber))
    stream = stream & GetFormRow(GetProperty("Card Expiration"), GetField(RequestNameCardExpiration, CardExpiration) & " <b>MM/YY</b>")
    '
    stream = stream & GetFormRow(FormSpacer, ButtonString)
    stream = stream & "</TABLE>"
    stream = stream & Main.GetFormEnd
    '
    GetFormDonate = stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "GetFormDonate", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadFormDonate()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    If sourceFormID = FormDonate Then
        amount = Main.GetStreamText(RequestNameAmount)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Address = Main.GetStreamText(RequestNameAddress)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        eMail = Main.GetStreamText(RequestNameEmail)
        memory = Main.GetStreamText(RequestNameMemory)
        CardNumber = Main.GetStreamText(RequestNameCardNumber)
        CardExpiration = Main.GetStreamText(RequestNameCardExpiration)
    Else
        '
        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.MemberID)
        If Main.CSOK(CSPointer) Then
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Address = Main.GetCSText(CSPointer, "Address")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Phone = Main.GetCSText(CSPointer, "Phone")
            eMail = Main.GetCSText(CSPointer, "Email")
            End If
        Call Main.CloseCS(CSPointer)
        '
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "LoadFormRequest", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetField(RequestName As String, Value As String, Optional Height As String, Optional Width As String, Optional MaxLenghth As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    If Height = "" Then
        Height = 1
        End If
    If Width = "" Then
        Width = 25
        End If
    '
    stream = Main.GetFormInputText(RequestName, Value, Height, Width)
    stream = Replace(stream, "<INPUT ", "<INPUT maxlength=" & MaxLenghth & " ", 1, 99, 1)
    '
    GetField = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DonationClass", "GetField", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetProperty(PropertyName As String, Optional DefaultValue As String) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    '
    If DefaultValue = "" Then
        DefaultValue = PropertyName
        End If
    '
    stream = Main.GetSiteProperty(PropertyName, DefaultValue, True)
    '
    GetProperty = stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("DonationClass", "GetProperty", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormDonate()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim ProcPay As New ProcessClass3
    Dim Copy As String
    Dim DonationID As Long
    '
    If ButtonValue <> "" Then
        '
        If amount = "" Then
            ErrorString = ErrorString & "In order to continue, please specify a donation amount.<br />"
            End If
        If FirstName = "" Then
            ErrorString = ErrorString & "Please enter your first name in order to continue.<br />"
            End If
        If LastName = "" Then
            ErrorString = ErrorString & "Please enter your last name in order to continue.<br />"
            End If
        If Address = "" Then
            ErrorString = ErrorString & "Please enter your address in order to continue.<br />"
            End If
        If City = "" Then
            ErrorString = ErrorString & "Please enter your city in order to continue.<br />"
            End If
        If State = "" Then
            ErrorString = ErrorString & "Please enter your state in order to continue.<br />"
            End If
        If Zip = "" Then
            ErrorString = ErrorString & "Please enter your zip in order to continue.<br />"
            End If
        If Phone = "" Then
            ErrorString = ErrorString & "Please enter your phone number in order to continue.<br />"
            End If
        If eMail = "" Then
            ErrorString = ErrorString & "Please enter your email address in order to continue.<br />"
            End If
        If CardNumber = "" Then
            ErrorString = ErrorString & "Please enter a credit card number in order to continue.<br />"
            End If
        If CardExpiration = "" Then
            ErrorString = ErrorString & "Please enter the credit card's expiration in order to continue.<br />"
            End If
        '
        If ErrorString = "" Then
            CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.MemberID)
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "Address", Address)
                Call Main.SetCS(CSPointer, "City", City)
                Call Main.SetCS(CSPointer, "State", State)
                Call Main.SetCS(CSPointer, "Zip", Zip)
                Call Main.SetCS(CSPointer, "Phone", Phone)
                Call Main.SetCS(CSPointer, "Email", eMail)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
            Call ProcPay.Init(Main)
            If ProcPay.PaymentProcessed(CardNumber, CardExpiration, amount, , "Online donation - " & Now(), FirstName & " " & LastName) Then
                CSPointer = Main.InsertCSContent(ContentNameDonations)
                If Main.CSOK(CSPointer) Then
                    DonationID = Main.GetCSInteger(CSPointer, "ID")
                    Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName & " " & Now())
                    Call Main.SetCS(CSPointer, "memoryOF", memory)
                    Call Main.SetCS(CSPointer, "Amount", amount)
                    Call Main.SetCS(CSPointer, "MemberID", Main.MemberID)
                    Call Main.SetCS(CSPointer, "VisitID", Main.VisitID)
                    Call Main.SetCS(CSPointer, "ReferenceNumber", ProcPay.ReferenceNumber)
                    Call Main.SaveCSRecord(CSPointer)
                    End If
                Call Main.CloseCS(CSPointer)
            Else
                ErrorString = ErrorString & ProcPay.ResponseMessage
                End If
            '
            If ErrorString = "" Then
                '
                Copy = "Donation submitted " & Now() & "<br />"
                Copy = Copy & "Name: " & FirstName & " " & LastName & "<br />"
                Copy = Copy & "Address: " & Address & "<br />"
                Copy = Copy & "City: " & City & "<br />"
                Copy = Copy & "State: " & State & "<br />"
                Copy = Copy & "Zip: " & Zip & "<br />"
                Copy = Copy & "eMail: " & eMail & "<br />"
                Copy = Copy & "In memory of: " & memory & "<br />"
                Copy = Copy & "Phone: " & Phone & "<br />"
                Copy = Copy & "Amount: $" & amount & "<br /><br />"
                Copy = Copy & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?af=4&id=" & DonationID & "&cid=" & Main.GetContentID(ContentNameDonations) & """>Click here for the entire details of this transaction</a>"
                '
                Call Main.AddGroupMember("Donators")
                Call Main.SendSystemEmail(SystemEmailAutoResponder, , Main.MemberID)
                Call Main.SendSystemEmail(SystemEmailNotification, Copy)
                formID = FormThankYou
            Else
                formID = FormDonate
                End If
            '
        Else
            formID = FormDonate
            End If
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("DonationClass", "ProcessFormDonate", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

