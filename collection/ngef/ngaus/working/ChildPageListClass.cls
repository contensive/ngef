VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChildPageListClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    Call HandleError("ChildPageListClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetChilPageList()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("ChildPageListClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'----------------------------------------------
'   Returns the child page list
'----------------------------------------------
'
Private Function GetChilPageList() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ContentPageStructure As String
    '
    Dim ContentPageStructureArray() As String
    Dim RowPointer As Long
    Dim RowCount As Long
    Dim ContentPageStructureLastChar As String
    '
    Dim ColumnArray() As String
    Dim ColumnPointer As Long
    Dim ColumnCount As Long
    Dim ColumnDelimiter As String
    '
    Dim LinkSignifier As String
    Dim LinkCaption As String
    '
    ContentPageStructure = Main.ContentPageStructure
Main.TestPoint ("ContentPageStructure: " & ContentPageStructure)
    ContentPageStructureArray = Split(ContentPageStructure, vbCrLf)
    RowCount = UBound(ContentPageStructureArray) + 1
    For RowPointer = 0 To RowCount - 1
        ColumnDelimiter = Left(ContentPageStructureArray(RowPointer), 1)
Main.TestPoint ("ColumnDelimiter: '" & ColumnDelimiter & "'")
Main.TestPoint ("Row: " & ContentPageStructureArray(RowPointer))
        ColumnArray = Split(ContentPageStructureArray(RowPointer), ColumnDelimiter)
        ColumnCount = UBound(ColumnArray)
        '
        If ColumnCount > 1 Then
            'Stream = Stream & "ColumnCount: " & ColumnCount & "<br />"
            'Stream = Stream & "Array0: " & ColumnArray(1) & "<br />"
            'Stream = Stream & "Array1: " & ColumnArray(2) & "<br />"
            'Stream = Stream & "Array2: " & ColumnArray(3) & "<br />"
            'Stream = Stream & "Array3: " & ColumnArray(4) & "<br />"
            'Stream = Stream & "Array4: " & ColumnArray(5) & "<br />"
            'Stream = Stream & "Array5: " & ColumnArray(6) & "<br />"
            'Stream = Stream & "<br /><br />"
            '
            If ColumnArray(1) = "3" Then
                '
                '   Child Page - show if menu headline is present
                '
                LinkCaption = ColumnArray(4)
                '
                If LinkCaption <> "" Then
                    Stream = Stream & Main.GetRecordEditLink("Page Content", ColumnArray(2))
                    Stream = Stream & "<a href=""" & ColumnArray(6) & """>"
                    Stream = Stream & LinkCaption
                    Stream = Stream & "</a><br />"
                    End If
                End If
            Stream = Stream & Main.GetRecordAddLink("Page Content", "ParentID=" & ColumnArray(3))
            '
            End If
    Next
    '
    GetChilPageList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("ChildPageListClass", "GetChilPageList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
