Namespace Contensive.Addons.ngausDonationCampaigns
    '
    Public Class DonationClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                '
                CP.Doc.Var("initialHit") = True
                s = CP.Utils.ExecuteAddon("{46310ACA-584E-4117-B6FD-247929CE730A}")
                '
                s = CP.Html.div(s, , , "donationFormContainer")
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.ngausDonationCampaigns.donationClass.execute")
            End Try
        End Function
    End Class
    '
    Public Class DonationFormClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim lS As String = ""   '   layout stream
                Dim campaignID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("Campaign"))
                Dim campaignName As String = CP.Content.GetRecordName("Donation Campaigns", campaignID)
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim sS As String = ""   '   script stream
                Dim initialHit As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("initialHit"))
                '
                If initialHit Then
                    loadForm(CP)
                End If
                '
                s += CP.Content.GetCopy(campaignName & " Donation Instructions", "Please use the following form to make a donation.")
                '
                cs.Open("Layouts", "ccGuid=" & CP.Db.EncodeSQLText("{106E459D-8F41-4056-8CB4-8EC6591C39D6}"), , , "layout")
                If cs.OK() Then
                    lS = cs.GetText("layout")
                End If
                cs.Close()
                '
                s += CP.Html.Hidden("campaign", campaignID) & lS
                s = CP.Html.Form(s).Replace("<form ", "<form id=""dontaion-form"" ")
                '
                sS += "$(document).ready(function() {" & vbCrLf
                sS += " $('#formButtonSubmit').click(function(){" & vbCrLf
                sS += "     GetAjax('donationSubmitHandler','','dontaion-form','donationFormContainer','','');" & vbCrLf
                sS += "     setSpinner('donationFormContainer','Processing Form...');" & vbCrLf
                sS += " })" & vbCrLf
                '
                sS += vbCrLf
                sS += "  $('#donationFirstName').val('" & CP.Doc.Var("firstName") & "');" & vbCrLf
                sS += "  $('#donationLastName').val('" & CP.Doc.Var("lastName") & "');" & vbCrLf
                sS += "  $('#donationAddress').val('" & CP.Doc.Var("address") & "');" & vbCrLf
                sS += "  $('#donationAddress2').val('" & CP.Doc.Var("address2") & "');" & vbCrLf
                sS += "  $('#donationCity').val('" & CP.Doc.Var("city") & "');" & vbCrLf
                sS += "  $('#donationState').val('" & CP.Doc.Var("state") & "');" & vbCrLf
                sS += "  $('#donationZip').val('" & CP.Doc.Var("zip") & "');" & vbCrLf
                sS += "  $('#donationPhone').val('" & CP.Doc.Var("phone") & "');" & vbCrLf
                sS += "  $('#donationEmail').val('" & CP.Doc.Var("email") & "');" & vbCrLf
                '
                sS += "  $('#donationInHonor').val('" & CP.Doc.Var("inHonor") & "');" & vbCrLf
                '
                sS += "  $('#donationAmount').val('" & CP.Doc.Var("amount") & "');" & vbCrLf
                sS += "  $('#donationAccountName').val('" & CP.Doc.Var("accountName") & "');" & vbCrLf
                sS += "  $('#donationAccountNumber').val('" & CP.Doc.Var("accountNumber") & "');" & vbCrLf
                sS += "  $('#donationExpiration').val('" & CP.Doc.Var("expiration") & "');" & vbCrLf
                '
                sS += "});"
                '
                CP.Doc.AddHeadJavascript(sS)
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.ngausDonationCampaigns.donationFormClass.execute")
            End Try
        End Function
        '
        Private Function LoadForm(ByVal CP As BaseClasses.CPBaseClass)
            Try
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                '
                cs.Open("People", "ID=" & CP.User.Id)
                If cs.OK() Then
                    CP.Doc.Var("firstName") = cs.GetText("firstName")
                    CP.Doc.Var("lastName") = cs.GetText("lastName")
                    CP.Doc.Var("address") = cs.GetText("address")
                    CP.Doc.Var("address2") = cs.GetText("address2")
                    CP.Doc.Var("city") = cs.GetText("city")
                    CP.Doc.Var("state") = cs.GetText("state")
                    CP.Doc.Var("zip") = cs.GetText("zip")
                    CP.Doc.Var("phone") = cs.GetText("phone")
                    CP.Doc.Var("email") = cs.GetText("email")
                End If
                cs.Close()
                '
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.ngausDonationCampaigns.donationFormClass.execute")
            End Try
        End Function
    End Class
    '
    Public Class DonationSubmitClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim campaignID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("campaign"))
                Dim campaignName As String = CP.Content.GetRecordName("Donation Campaigns", campaignID)
                Dim errFlag As Boolean = False
                Dim iSS As String = ""  '  inner script stream
                Dim sS As String = ""  '   script stream
                Dim hS As String = ""   '   honor stream
                '
                Dim firstName As String = CP.Doc.Var("firstName")
                Dim lastName As String = CP.Doc.Var("lastName")
                Dim amount As Double = CP.Utils.EncodeNumber(CP.Doc.Var("amount"))
                Dim acctNumber As String = CP.Doc.Var("accountNumber")
                Dim expDate As String = CP.Doc.Var("expiration")
                '
                Dim address As String = CP.Doc.Var("address")
                Dim address2 As String = CP.Doc.Var("address2")
                Dim city As String = CP.Doc.Var("city")
                Dim state As String = CP.Doc.Var("state")
                Dim zip As String = CP.Doc.Var("zip")
                Dim phone As String = CP.Doc.Var("phone")
                Dim email As String = CP.Doc.Var("email")
                Dim inhonor As String = CP.Doc.Var("inHonor")
                '
                Dim rS As String = ""   '   response stream
                Dim errMsg As String = ""
                Dim processed As Boolean = False
                Dim resultDoc As New System.Xml.XmlDocument
                Dim ppResponse As String = ""
                Dim ppApproved As Boolean = False
                Dim ppReference As String = ""
                Dim ppCSCMatch As Boolean = False
                Dim ppAVSMatch As Boolean = False
                '
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim copy As String = ""
                Dim copyPlus As String = ""
                Dim inC As String = ""  '   inner copy
                Dim donationID As Integer = 0
                '
                Dim systemEmailID As Integer = 0
                '
                If firstName = "" Then
                    iSS += "  $('#donationFirstName').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If lastName = "" Then
                    iSS += "  $('#donationLastName').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If address = "" Then
                    iSS += "  $('#donationAddress').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If city = "" Then
                    iSS += "  $('#donationCity').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If state = "" Then
                    iSS += "  $('#donationState').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If zip = "" Then
                    iSS += "  $('#donationZip').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If email = "" Then
                    iSS += "  $('#donationEmail').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If amount = 0 Then
                    iSS += "  $('#donationAmount').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If CP.Doc.Var("accountName") = "" Then
                    iSS += "  $('#donationAccountName').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If acctNumber = "" Then
                    iSS += "  $('#donationAccountNumber').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                If expDate = "" Then
                    iSS += "  $('#donationExpiration').css('background', '#FF9494');" & vbCrLf
                    errFlag = True
                End If
                '
                '   process the payment
                '
                If Not errFlag Then
                    '
                    CP.Doc.Var("CreditCardNumber") = acctNumber
                    CP.Doc.Var("CreditCardExpiration") = expDate
                    CP.Doc.Var("PaymentAmount") = amount.ToString
                    CP.Doc.Var("Comment1") = "Donation - " & campaignName
                    CP.Doc.Var("Comment2") = FormatCurrency(amount, 2) & " from " & firstName & " " & lastName
                    '
                    rS = CP.Utils.ExecuteAddon("{F71E8C9B-38A4-446E-8CAC-07548EE602BB}")
                    If rS = "" Then
                        errMsg += "There was a problem processing your payment. Please try again."
                        errFlag = True
                    Else
                        Call resultDoc.LoadXml(rS)
                        If resultDoc.HasChildNodes Then
                            ppApproved = CP.Utils.EncodeBoolean(resultDoc.GetElementsByTagName("status").Item(0).InnerText)
                            ppResponse = resultDoc.GetElementsByTagName("responseMessage").Item(0).InnerText
                            ppReference = resultDoc.GetElementsByTagName("referenceNumber").Item(0).InnerText
                            ppCSCMatch = CP.Utils.EncodeBoolean(resultDoc.GetElementsByTagName("securityMatchPassed").Item(0).InnerText)
                            ppAVSMatch = CP.Utils.EncodeBoolean(resultDoc.GetElementsByTagName("avsMatchPassed").Item(0).InnerText)
                        End If
                        If (Not ppApproved) Then
                            errMsg += "There was a problem processing your payment. -- " & ppResponse
                            errFlag = True
                        Else
                            processed = True
                        End If
                    End If
                Else
                    errMsg = "Please complete all required fields in order to continue."
                End If
                '
                If errFlag Then
                    s = CP.Html.p(errMsg, , "ccError")
                    s += CP.Utils.ExecuteAddon("{46310ACA-584E-4117-B6FD-247929CE730A}")
                    '
                    sS += vbCrLf
                    sS += "$(document).ready(function() {" & vbCrLf
                    sS += iSS & vbCrLf
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                Else
                    cs.Insert("Campaign Donations")
                    If cs.OK() Then
                        cs.SetField("name", campaignName & " by " & firstName & " " & lastName & " - " & Date.Now())
                        '
                        cs.SetField("donationCampaignID", campaignID)
                        cs.SetField("amount", amount.ToString)
                        cs.SetField("inHonor", inhonor)
                        '
                        cs.SetField("accountName", CP.Doc.Var("accountName"))
                        cs.SetField("accountNumber", "XXXXXXXXXXXX" & Right(acctNumber, 4))
                        cs.SetField("accountExpiration", expDate)
                        '
                        cs.SetField("firstName", firstName)
                        cs.SetField("lastName", lastName)
                        cs.SetField("address", address)
                        cs.SetField("address2", address2)
                        cs.SetField("city", city)
                        cs.SetField("state", state)
                        cs.SetField("zip", zip)
                        cs.SetField("phone", phone)
                        cs.SetField("email", email)
                        '
                        cs.SetField("visitID", CP.Visit.Id)
                        cs.SetField("memberID", CP.User.Id)
                        '
                        donationID = cs.GetInteger("ID")
                    End If
                    cs.Close()
                    '
                    '   open the campaign to get the system email id
                    '
                    If cs.Open("Donation Campaigns", "id=" & campaignID, , , "systemEmailID") Then
                        systemEmailID = cs.GetInteger("systemEmailID")
                    End If
                    cs.Close()
                    '
                    If inhonor <> "" Then
                        hS = "In Honor / Memory of " & inhonor
                    End If
                    '
                    copy = CP.Html.h2("Your Payment Receipt")
                    copy += CP.Html.h3("Campaign: " & campaignName)
                    copy += CP.Html.p("Amount: " & amount.ToString & " " & hS)
                    '
                    copy += CP.Html.h3("Payment Information")
                    '
                    inC += firstName & " " & lastName & "<br />"
                    inC += address & "<br />"
                    If address2 <> "" Then
                        inC += address2 & "<br />"
                    End If
                    inC += city & ", " & state & " " & zip & "<br />"
                    If phone <> "" Then
                        inC += "Phone: " & phone
                    End If
                    If email <> "" Then
                        inC += "Email: " & email
                    End If
                    '
                    copy += CP.Html.p(inC)
                    '
                    copyPlus = copy
                    copyPlus += "<a href=""" & CP.Site.Domain & "/admin/" & CP.Site.PageDefault & "?cid=" & CP.Content.GetRecordID("Content", "Campaign Donations") & "&id=" & donationID & "&af=4"">Click here to view the donation record</a>"
                    '
                    '   send notification
                    '
                    CP.Email.sendSystem("Campaign Donation Notification", copyPlus)
                    '
                    '   get the auto responder from the donation campaign record
                    '
                    If systemEmailID <> 0 Then
                        CP.Email.sendSystem(CP.Content.GetRecordName("System Email", systemEmailID), copy, CP.User.Id)
                    End If
                    '
                    '   used to send one email for every donation
                    '
                    'CP.Email.sendSystem("Campaign Donation Thank You", copy)
                    '
                    s = CP.Html.div(CP.Utils.ExecuteAddon("{EFD4A584-8B5C-4675-810C-1637FC3B3EC3}"), , "thermometerContainer")
                    s += CP.Content.GetCopy(campaignName & " Donation Thank You", "Thank You for your donation!")
                End If
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.ngausDonationCampaigns.donationSubmitClass.execute")
            End Try
        End Function

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
    '
End Namespace