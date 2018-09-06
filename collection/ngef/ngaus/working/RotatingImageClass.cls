VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rotatingImageClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Private Main As Object
'
Public Sub Init(MainObject As Object)
    Set Main = MainObject
    Exit Sub
    '
ErrorTrap:
    Call HandleError("RotatingImageClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        '
        Dim CSImage As Long
        Dim ImageSelect As Long
        Dim ImageCount As Long
        Dim Stream As String
        Dim Caption As String
        Dim SQL As String
        Dim CurrentImageRecordID As Long
        Dim CSImage2 As Long
        Dim CSImage3 As Long
        '
        Stream = ""
        CurrentImageRecordID = Main.GetSiteProperty("Current Image Record ID", 0)
        SQL = "Select Top 2 ID, ImageFileName, Caption from ccImages Where ( ContentControlID=" & Main.GetContentID("Images") & " ) AND ( Active = 1 ) AND ( ID >= " & CurrentImageRecordID & " ) Order By ID Asc ;"
        '
        CSImage = Main.OpenCSSQL("Default", SQL)
        If Main.CSOK(CSImage) Then
            Caption = Main.GetCSText(CSImage, "Caption")
            Stream = Stream & "<img border=0 src=""" & Main.ServerFilePath & Main.GetCSText(CSImage, "ImageFileName") & """ border=""0"" width=""780"" height=""109"" title=""" & Caption & """ alt=""" & Caption & """>"
            Call Main.NextCSRecord(CSImage)
            If Main.CSOK(CSImage) Then
                Call Main.SetSiteProperty("Current Image Record ID", Main.GetCSInteger(CSImage, "ID"))
            Else
                Call Main.SetSiteProperty("Current Image Record ID", 0)
            End If
        End If
        Call Main.CloseCS(CSImage)
        '
        GetContent = Stream
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("RotatingImageClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

