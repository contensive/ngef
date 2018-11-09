VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "projectDonation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const rnprojectID = "pid"
Const rnDelProjID = "rpid"
Const rnButton = "button"

Const rnName = "dnnm"
Const rnAmount = "dnat"

Const rnFormID = "frmid"
Const rnSourceFormID = "sfrmid"

Const rnAcctNumber = "actnm"
Const rnAcctName = "actnme"
Const rnAcctMo = "actmo"
Const rnAcctYr = "actyr"
Const rnAcctCSV = "csv"

Const btnContinue = "Continue"
Const btnUpdate = "Update"
Const btnPrevious = "Previous"

Const rnPrefix = "prfx"
Const rnFName = "fnam"
Const rnMName = "mnam"
Const rnLName = "lnam"
Const rnSufix = "suf"
Const rnAddress = "addr"
Const rnAddress2 = "addr2"
Const rnCity = "cit"
Const rnState = "stat"
Const rnZip = "zip"
Const rnCountry = "cnt"
Const rnPhone = "phn"
Const rnEmail = "eml"
Const rnSubscribe = "sub"

Const frmReview = 100
Const frmPerInfo = 150
Const frmPay = 200
Const frmThanks = 300

Const required = " <span class=""ccError"">*</span> "

Private projID As Long
Private delProjID As Long
Private btnValue As String

Private formID As Long
Private sourceFormID As Long

Private Name As String
Private amount As Double

Private actNam As String
Private actNum As String
Private actMo As String
Private actYr As String
Private actCSV As String
Private ttlAmount As Double
Private transID As Integer

Private custPrefix As String
Private custFName As String
Private custMName As String
Private custLName As String
Private custSufix As String
Private custAddress As String
Private custAddress2 As String
Private custCity As String
Private custState As String
Private custZip As String
Private custCountry As String
Private custPhone As String
Private custEmail As String
Private custSubscribe As Boolean

Private errMessage As String

Private projectPage As String

Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
        
    Set Main = MainObject
    Set CSV = CsvObject

    Dim stream As String
    Dim cs As Long
    
    projectPage = Main.GetAggrOption("Projects Page", OptionString)
    
    Call loadForm
    Call updateDonation
    Call ProcessForm
    
    If Main.RenderedPageID <> 0 Then
        Call Main.SetSiteProperty("Location - Project Donation Form", "?bid=" & Main.RenderedPageID)
    End If
    
    Select Case formID
        Case frmThanks
            stream = getFormThankYou()
        Case frmPay
            stream = getFormPay()
        Case frmPerInfo
            Call loadFormInfo
            stream = getFormInfo()
        Case Else
            formID = frmReview
            stream = getFormReview()
    End Select
    
    
    stream = "<div id=""donationNavigation"">" & getNavigation() & "</div>" & stream
    stream = "<div id=""donationContainer"">" & Main.GetFormStart & Main.GetFormInputHidden(rnSourceFormID, formID) & stream & Main.GetFormEnd & "</div>"
    
    Execute = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
End Function

Private Sub updateDonation()
On Error GoTo ErrorTrap

    Dim cs As Long
    
    If projID <> 0 Then
        cs = Main.OpenCSContent("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ") AND (ProjectID=" & projID & ")")
        If Not Main.CSOK(cs) Then
            Call Main.CloseCS(cs)
            cs = Main.InsertCSContent("Donations")
            If Main.CSOK(cs) Then
                Call Main.SetCS(cs, "ProjectID", projID)
                Call Main.SetCS(cs, "MemberID", Main.MemberID)
                Call Main.SetCS(cs, "VisitID", Main.VisitID)
                Call Main.SetCS(cs, "Name", Main.MemberName)
                Call Main.SetCS(cs, "Amount", "25.00")
                Call Main.SaveCSRecord(cs)
            End If
        End If
        Call Main.CloseCS(cs)
    End If
    
    If delProjID <> 0 Then
        Call Main.DeleteContentRecords("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ") AND (ProjectID=" & delProjID & ")")
    End If

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "updateDonation", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Function getDonationRow(cs As Long) As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim locProjID As Long
    Dim QS As String
    
    locProjID = Main.GetCSInteger(cs, "projectID")
    
    QS = "?" & Main.RefreshQueryString
    QS = QS & "&" & rnDelProjID & "=" & locProjID
    
    stream = "<div id=""donationRowContainer"">"
    stream = stream & "<h3>You are making a gift to " & Main.GetRecordName("Projects", locProjID) & "</h3>"
'    stream = stream & "<div id=""donationRowLeft"">"
'    stream = stream & Main.GetFormInputText(rnName & "_" & locProjID, Main.GetCSText(CS, "Name"), , "20")
'    stream = stream & " Let them know who cares!"
'    stream = stream & "</div>"
    'stream = stream & "<div id=""donationRowRight"">"
    'stream = stream & "<div>"
    stream = stream & Replace(Main.GetFormInputText(rnAmount & "_" & locProjID, Main.GetCSNumber(cs, "Amount"), , "20", rnAmount & "_" & locProjID), "<input ", "<input onChange=""setAmount(this)"" ")
    stream = stream & " Amount ($25 minimum)"
    stream = stream & "</div>"
    stream = stream & "<div id=""donationRemoveContainer""><a href=""" & Main.ServerPage & QS & """>[Remove]</a></div>"
    stream = stream & "</div>"
    
    getDonationRow = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getDonationRow", Err.Number, Err.Source, Err.Description, True, False)
End Function
    
Private Sub ProcessForm()
On Error GoTo ErrorTrap

    Dim cs As Long
    Dim objPay As New ProcessClass6
    
    Select Case sourceFormID
        Case frmPerInfo
            Select Case btnValue
                Case btnContinue
                    Call loadFormInfo
                    Call processFormInfo
                Case btnPrevious
                    formID = frmReview
            End Select
        Case frmPay
            Select Case btnValue
                Case btnContinue
                    Call loadFormPay
                    Call processFormPay
                Case btnPrevious
                    formID = frmPerInfo
            End Select
        Case frmReview
            Select Case btnValue
                Case btnContinue
                    Call saveReview
                    formID = frmPerInfo
                Case btnUpdate
                    Call saveReview
                    formID = frmReview
            End Select
    End Select

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "processForm", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Function getFormReview() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim cs As Long

    cs = Main.OpenCSContent("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ")")
    If Main.CSOK(cs) Then
        Do While Main.CSOK(cs)
            stream = stream & getDonationRow(cs)
            Call Main.NextCSRecord(cs)
        Loop
        stream = stream & "<div id=""donationButtonContainer"">" & Main.GetFormButton(btnContinue) & " " & Main.GetFormButton(btnUpdate) & "</div>"
        
        stream = stream & "<div id=""returnContainer""><a href=""?bid=" & projectPage & """>Make Another Donation</a></div>"
    Else
        stream = stream & "You have not selected a Project to donate to - please <a href=""" & Main.GetSiteProperty("Location - Projects") & """>select a project</a>"
    End If
    Call Main.CloseCS(cs)
    
    getFormReview = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getFormReview", Err.Number, Err.Source, Err.Description, True, False)
End Function

Private Sub loadForm()
    On Error GoTo ErrorTrap

    projID = Main.GetStreamInteger(rnprojectID)
    delProjID = Main.GetStreamInteger(rnDelProjID)
    btnValue = Main.GetStreamText(rnButton)
    formID = Main.GetStreamInteger(rnFormID)
    sourceFormID = Main.GetStreamInteger(rnSourceFormID)

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "loadForm", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Sub loadFormInfo()
    On Error GoTo ErrorTrap

    Dim cs As Long

    If sourceFormID = frmPerInfo Then
        custFName = Main.GetStreamText(rnFName)
        custLName = Main.GetStreamText(rnLName)
        custAddress = Main.GetStreamText(rnAddress)
        custAddress2 = Main.GetStreamText(rnAddress2)
        custCity = Main.GetStreamText(rnCity)
        custState = Main.GetStreamText(rnState)
        custZip = Main.GetStreamText(rnZip)
        custCountry = Main.GetStreamText(rnCountry)
        custPhone = Main.GetStreamText(rnPhone)
        custEmail = Main.GetStreamText(rnEmail)
    Else
        cs = Main.OpenCSContent("People", "ID=" & Main.MemberID)
        If Main.CSOK(cs) Then
            custFName = Main.GetCSText(cs, "FirstName")
            custLName = Main.GetCSText(cs, "LastName")
            custAddress = Main.GetCSText(cs, "Address")
            custAddress2 = Main.GetCSText(cs, "Address2")
            custCity = Main.GetCSText(cs, "City")
            custState = Main.GetCSText(cs, "State")
            custZip = Main.GetCSText(cs, "Zip")
            custCountry = Main.GetCSText(cs, "Country")
            custPhone = Main.GetCSText(cs, "Phone")
            custEmail = Main.GetCSText(cs, "Email")
        End If
        Call Main.CloseCS(cs)
    End If

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "loadFormInfo", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Sub saveReview()
On Error GoTo ErrorTrap

    Dim cs As Long
    Dim locProjID As Long
    
    cs = Main.OpenCSContent("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ")")
    If Main.CSOK(cs) Then
        Do While Main.CSOK(cs)
            locProjID = Main.GetCSInteger(cs, "projectID")
            Call Main.SetCS(cs, "Name", Main.GetStreamText(rnName & "_" & locProjID))
            Call Main.SetCS(cs, "Amount", Main.GetStreamNumber(rnAmount & "_" & locProjID))
            Call Main.NextCSRecord(cs)
        Loop
    End If
    Call Main.CloseCS(cs)

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "saveReview", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Function getFormPay() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim cs As Long
    Dim locTotal As Double
    Dim donString As String

    cs = Main.OpenCSContent("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ")")
    If Main.CSOK(cs) Then
        Do While Main.CSOK(cs)
            locTotal = locTotal + Main.GetCSNumber(cs, "Amount")
            donString = donString & FormatCurrency(Main.GetCSNumber(cs, "Amount"), 2) & " - " & Main.GetRecordName("Projects", Main.GetCSInteger(cs, "ProjectID")) & "<br />"
            Call Main.NextCSRecord(cs)
        Loop
    End If
    Call Main.CloseCS(cs)
    
    errMessage = Main.GetUserError
    
    If errMessage <> "" Then
        stream = stream & errMessage
    Else
        stream = stream & "<div id=""instructionContainer"">" & Main.GetContentCopy2("Instructions - Project Donation Payment", , "Use the form to complete your donation.") & "</div>"
    End If
    
    stream = stream & "<div class=""donField"">" & donString & "</div>"
    stream = stream & "<div class=""donField"">Total Amount Due<br />" & FormatCurrency(locTotal, 2) & "</div>"
    stream = stream & "<div class=""donField"">Account Number" & required & "<br />" & Main.GetFormInputText(rnAcctNumber, "", , 20) & "</div>"
    stream = stream & "<div class=""donField"">Expiration" & required & "<br />" & getExpField() & "</div>"
    stream = stream & "<div class=""donField"">Security Code" & required & "<br />" & Main.GetFormInputText(rnAcctCSV, "", , 20) & "</div>"
    stream = stream & "<div id=""donationButtonContainer"">" & Main.GetFormButton(btnContinue) & " " & Main.GetFormButton(btnPrevious) & "</div>"
    stream = stream & Main.GetFormInputHidden(rnAmount, locTotal)
    
    getFormPay = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getFormPay", Err.Number, Err.Source, Err.Description, True, False)
End Function

Private Function getExpField() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim DateMonth As String
    Dim DateYear As String
    Dim YearStart As Long
    Dim YearEnd As Long
    Dim Counter As Long
    
    YearStart = Year(Now)
    YearEnd = Year(Now) + 10
    
    stream = stream & "<select size=""1"" name=""" & rnAcctMo & """>"
    stream = stream & "<option selected>MM</option>"
    For Counter = 1 To 12
        stream = stream & "<option value=""" & Counter & """>" & Counter & " - " & MonthName(Counter) & "</option>"
    Next
    stream = stream & "</select> "
    
    stream = stream & "<select size=""1"" name=""" & rnAcctYr & """>"
    stream = stream & "<option selected>YYYY</option>"
    For Counter = YearStart To YearEnd
        stream = stream & "<option value=""" & Right(Counter, 2) & """>" & Counter & "</option>"
    Next
    stream = stream & "</select>"
    
    getExpField = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getExpField", Err.Number, Err.Source, Err.Description, True, False)
    End Function
    
Private Function getNavigation() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String

    If (formID = frmReview) Then
        stream = stream & "<span class=""navHighlighted"">"
    Else
        stream = stream & "<span class=""navNormal"">"
    End If
    stream = stream & "Donation Information</span>"
    stream = stream & " | "
    If (formID = frmPerInfo) Then
        stream = stream & "<span class=""navHighlighted"">"
    Else
        stream = stream & "<span class=""navNormal"">"
    End If
    stream = stream & "Personal Information</span>"
    stream = stream & " | "
    If (formID = frmPay) Then
        stream = stream & "<span class=""navHighlighted"">"
    Else
        stream = stream & "<span class=""navNormal"">"
    End If
    stream = stream & "Billing Information</span>"
    stream = stream & " | "
    If (formID = frmThanks) Then
        stream = stream & "<span class=""navHighlighted"">"
    Else
        stream = stream & "<span class=""navNormal"">"
    End If
    stream = stream & "Print Donation Information</span>"
    
    getNavigation = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getNavigation", Err.Number, Err.Source, Err.Description, True, False)
    End Function
    
Private Function getFormInfo() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim formLayout As String
    Dim cs As Long
    Dim altStream As String

    formLayout = "Form Layout - Project Donation Contact Information"
    cs = Main.OpenCSContent("Copy Content", "name=" & KmaEncodeSQLText(formLayout))
    If Main.IsCSOK(cs) Then
        stream = Main.GetCS(cs, "Copy")
    End If
    Call Main.CloseCS(cs)
    
    If stream = "" Then
        stream = stream & "<div class=""donField"">First Name" & required & "<br />$First$</div>"
        stream = stream & "<div class=""donField"">Last Name" & required & "<br />$Last$</div>"
        stream = stream & "<div class=""donField"">Address" & required & "<br />$Address$</div>"
        stream = stream & "<div class=""donField"">$Address2$</div>"
        stream = stream & "<div class=""donField"">City" & required & "<br />$City$</div>"
        stream = stream & "<div class=""donField"">State" & required & "<br />$State$</div>"
        stream = stream & "<div class=""donField"">Zip" & required & "<br />$Zip$</div>"
        stream = stream & "<div class=""donField"">Country" & required & "<br />$Country$</div>"
        stream = stream & "<div class=""donField"">Phone" & required & "<br />$Phone$</div>"
        stream = stream & "<div class=""donField"">Email" & required & "<br />$Email$</div>"

        stream = stream & "<div class=""btnField"">$Button$</div>"

        altStream = Main.GetContentCopy2(formLayout, , stream)
        
    End If
    
    stream = Main.GetUserError & stream
    
    stream = Replace(stream, "$First$", Main.GetFormInputText(rnFName, custFName, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Last$", Main.GetFormInputText(rnLName, custLName, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Address$", Main.GetFormInputText(rnAddress, custAddress, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Address2$", Main.GetFormInputText(rnAddress2, custAddress2, , 45), , , vbTextCompare)
    stream = Replace(stream, "$City$", Main.GetFormInputText(rnCity, custCity, , 45), , , vbTextCompare)
    stream = Replace(stream, "$State$", Main.GetFormInputText(rnState, custState, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Zip$", Main.GetFormInputText(rnZip, custZip, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Country$", Main.GetFormInputText(rnCountry, custCountry, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Phone$", Main.GetFormInputText(rnPhone, custPhone, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Email$", Main.GetFormInputText(rnEmail, custEmail, , 45), , , vbTextCompare)
    stream = Replace(stream, "$Button$", Main.GetFormButton(btnContinue) & " " & Main.GetFormButton(btnPrevious), , , vbTextCompare)
    
    getFormInfo = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getFormInfo", Err.Number, Err.Source, Err.Description, True, False)
End Function

Private Sub processFormInfo()
    On Error GoTo ErrorTrap

    Dim cs As Long
    Dim errFlag As Boolean
    
    If (custFName = "") Or (custLName = "") Or (custAddress = "") Or (custCity = "") Or (custState = "") Or (custZip = "") Or (custCountry = "") Or (custPhone = "") Or (custEmail = "") Then
        Call Main.AddUserError("Please complete all required fields in order to continue")
        errFlag = True
    End If
    
    If Not errFlag Then
        cs = Main.OpenCSContent("People", "ID=" & Main.MemberID)
        If Main.CSOK(cs) Then
            Call Main.SetCS(cs, "FirstName", custFName)
            Call Main.SetCS(cs, "LastName", custLName)
            Call Main.SetCS(cs, "Address", custAddress)
            Call Main.SetCS(cs, "Address2", custAddress2)
            Call Main.SetCS(cs, "City", custCity)
            Call Main.SetCS(cs, "State", custState)
            Call Main.SetCS(cs, "Zip", custZip)
            Call Main.SetCS(cs, "Country", custCountry)
            Call Main.SetCS(cs, "Phone", custPhone)
            Call Main.SetCS(cs, "Email", custEmail)
        End If
        Call Main.CloseCS(cs)
        
        formID = frmPay
    Else
        formID = sourceFormID
    End If

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "processFormInfo", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Sub loadFormPay()
    On Error GoTo ErrorTrap

    Dim cs As Long

    If sourceFormID = frmPay Then
        actNum = Main.GetStreamText(rnAcctNumber)
        actMo = Main.GetStreamText(rnAcctMo)
        actYr = Main.GetStreamText(rnAcctYr)
        actCSV = Main.GetStreamText(rnAcctCSV)
        ttlAmount = Main.GetStreamNumber(rnAmount)
        
        cs = Main.OpenCSContent("People", "ID=" & Main.MemberID)
        If Main.CSOK(cs) Then
            custFName = Main.GetCSText(cs, "FirstName")
            custLName = Main.GetCSText(cs, "LastName")
            custAddress = Main.GetCSText(cs, "Address")
            custAddress2 = Main.GetCSText(cs, "Address2")
            custCity = Main.GetCSText(cs, "City")
            custState = Main.GetCSText(cs, "State")
            custZip = Main.GetCSText(cs, "Zip")
            custCountry = Main.GetCSText(cs, "Country")
            custPhone = Main.GetCSText(cs, "Phone")
            custEmail = Main.GetCSText(cs, "Email")
        End If
        Call Main.CloseCS(cs)
    End If

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "loadFormPay", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Sub processFormPay()
    On Error GoTo ErrorTrap

    Dim cs As Long
    Dim errFlag As Boolean
    Dim donString As String
    Dim resultString As String
    Dim Approved As Boolean
    Dim Response As String
    Dim Reference As String
    Dim OptionString As String
    Dim doc As New DOMDocument60
    
    If (actNum = "") Or (actMo = "MM") Or (actYr = "YYYY") Or (actCSV = "") Then
        Call Main.AddUserError("Please complete all required information in order to continue.")
        errFlag = True
    Else
        OptionString = ""
        OptionString = OptionString & "&CreditCardNumber=" & encodeNvaArgument(actNum)
        OptionString = OptionString & "&CreditCardExpiration=" & encodeNvaArgument(CStr(actMo & "/" & actYr))
        OptionString = OptionString & "&PaymentAmount=" & encodeNvaArgument(CStr(ttlAmount))
        OptionString = OptionString & "&SecurityCode=" & encodeNvaArgument(actCSV)
        
        OptionString = OptionString & "&custFirstName=" & encodeNvaArgument(custFName)
        OptionString = OptionString & "&custLastName=" & encodeNvaArgument(custLName)
        OptionString = OptionString & "&custAddress=" & encodeNvaArgument(custAddress)
        OptionString = OptionString & "&custCity=" & encodeNvaArgument(custCity)
        OptionString = OptionString & "&custState=" & encodeNvaArgument(custState)
        OptionString = OptionString & "&custZip=" & encodeNvaArgument(custZip)
        OptionString = OptionString & "&custCountry=" & encodeNvaArgument(custCountry)
        OptionString = OptionString & "&custPhone=" & encodeNvaArgument(custPhone)
        OptionString = OptionString & "&custEmail=" & encodeNvaArgument(custEmail)
        
        cs = Main.OpenCSContent("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ")")
        If Main.CSOK(cs) Then
            Do While Main.CSOK(cs)
                If donString <> "" Then
                    donString = donString & " - "
                End If
                donString = donString & FormatCurrency(Main.GetCSNumber(cs, "Amount"), 2) & " - " & Main.GetRecordName("Projects", Main.GetCSInteger(cs, "ProjectID"))
                Call Main.NextCSRecord(cs)
            Loop
        End If
        Call Main.CloseCS(cs)
        
        OptionString = OptionString & "&transDesc=" & encodeNvaArgument(donString)
        
        resultString = Main.ExecuteAddon(0, "{F71E8C9B-38A4-446E-8CAC-07548EE602BB}", OptionString, ContextAdmin, "", 0, "", "", -1)
        Call AppendLogFile("resultString: " & resultString)
        If resultString = "" Then
            Call Main.AddUserError("There was a problem processing your payment - please try again.")
            errFlag = True
        Else
            Call doc.loadXML(resultString)
            If doc.hasChildNodes Then
                Approved = kmaEncodeBoolean(doc.getElementsByTagName("status").Item(0).Text)
                Response = doc.getElementsByTagName("responseMessage").Item(0).Text
                Reference = doc.getElementsByTagName("referenceNumber").Item(0).Text
            End If
            If Not Approved Then
                Call Main.AddUserError("There was a problem processing your payment - " & Response)
                errFlag = True
            End If
        End If
        
        cs = Main.InsertCSContent("Payment Transactions")
        If Main.CSOK(cs) Then
            transID = Main.GetCSInteger(cs, "ID")
            Call Main.SetCS(cs, "Name", "Donation")
            Call Main.SetCS(cs, "VisitID", Main.VisitID)
            Call Main.SetCS(cs, "MemberID", Main.MemberID)
            Call Main.SetCS(cs, "Amount", ttlAmount)
            Call Main.SetCS(cs, "Account", Right(actNum, 4))
            Call Main.SetCS(cs, "Reference", Reference)
            Call Main.SetCS(cs, "Outcome", Response)
        End If
        Call Main.CloseCS(cs)
        
    End If
    
    If Not errFlag Then
        cs = Main.OpenCSContent("Donations", "(Completed=0) AND (VisitID=" & Main.VisitID & ")")
        Do While Main.CSOK(cs)
            Call Main.SetCS(cs, "Completed", 1)
            Call Main.SetCS(cs, "paymentTransactionID", transID)
            Call Main.NextCSRecord(cs)
        Loop
        Call Main.CloseCS(cs)
        formID = frmThanks
    Else
        formID = frmPay
    End If

    Exit Sub
ErrorTrap:
    Call HandleError("donationClass", "processFormInfo", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Private Function getFormThankYou() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim donString As String
    Dim innerString As String
    Dim cs As Long

    stream = Main.GetContentCopy2("Project Donation - Thank You", , "Thank you for your donation")
    
    stream = stream & "<div id=""donationOverview"">"
    
    cs = Main.OpenCSContent("Donations", "(Completed=1) AND (VisitID=" & Main.VisitID & ")")
    If Main.CSOK(cs) Then
        Do While Main.CSOK(cs)
            donString = donString & FormatCurrency(Main.GetCSNumber(cs, "Amount"), 2) & " - " & Main.GetRecordName("Projects", Main.GetCSInteger(cs, "ProjectID")) & "<br />"
            Call Main.NextCSRecord(cs)
        Loop
    End If
    Call Main.CloseCS(cs)
    
    innerString = "<h3>Donation Information</h3>"
    
    innerString = innerString & "<span class=""caption"">Total Billed:</span> " & FormatCurrency(ttlAmount, 2) & "<br />"
    innerString = innerString & "<span class=""caption"">Date:</span> " & Date & "<br />"
    innerString = innerString & "<span class=""caption"">Receipt #:</span> " & transID & "<br />"
    innerString = innerString & donString & "<br />"
    innerString = innerString & "<span class=""caption"">Name:</span> " & custFName & "<br />"
    innerString = innerString & "<span class=""caption"">Address:</span> " & custLName & "<br />"
    innerString = innerString & "<span class=""caption"">City, State, Zip:</span> " & custCity & " " & custState & " " & custZip & "<br />"
    innerString = innerString & "<span class=""caption"">Country:</span> " & custCountry & "<br />"
    innerString = innerString & "<span class=""caption"">Email:</span> " & custEmail & "<br />"
    innerString = innerString & "<span class=""caption"">Phone:</span> " & custPhone & "<br />"
    innerString = innerString & "<span class=""caption"">Card Number:</span> " & Right(actNum, 4) & "<br />"
    innerString = innerString & "<span class=""caption"">Card Expiration:</span> " & actMo & "/" & actYr & "<br />"
    
    stream = stream & innerString
    stream = stream & "</div>"
    
    Call Main.SendSystemEmail("Project Donation Auto Responder", innerString, Main.MemberID)
    Call Main.SendSystemEmail("Project Donation Notification", innerString)
    
    getFormThankYou = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("donationClass", "getFormThankYou", Err.Number, Err.Source, Err.Description, True, False)
    End Function

