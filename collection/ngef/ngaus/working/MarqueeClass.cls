VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarqueeClass"
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
    Call HandleError("MarqueeClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
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
        GetContent = GetMarquee()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("MarqueeClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetMarquee() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim Message As String
    '
    Message = Main.GetSiteProperty("Marquee Message", "", True)
    '
    If Message <> "" Then
        Message = Replace(Message, "'", "\'")
        Message = Replace(Message, vbCrLf, "")
        Stream = "<script language=""JavaScript1.2"">"
        Stream = Stream & vbCrLf & "/*"
        Stream = Stream & vbCrLf & " Cross browser Marquee script- © Dynamic Drive (www.dynamicdrive.com)"
        Stream = Stream & vbCrLf & " For full source code, 100's more DHTML scripts, and Terms Of Use, visit http://www.dynamicdrive.com"
        Stream = Stream & vbCrLf & "*/"
        Stream = Stream & vbCrLf & " var marqueewidth='344px'"
        Stream = Stream & vbCrLf & " var marqueeheight='28px'"
        Stream = Stream & vbCrLf & " //marquee speed (larger is faster 1-10)"
        Stream = Stream & vbCrLf & " var marqueespeed=2"
        Stream = Stream & vbCrLf & " var marqueebgcolor=''"
        Stream = Stream & vbCrLf & " //Pause marquee onMousever (0=no. 1=yes)"
        Stream = Stream & vbCrLf & " var pauseit=1"
        Stream = Stream & vbCrLf & " var marqueecontent='<nobr>" & Message & "</nobr>'"
        Stream = Stream & vbCrLf & " marqueespeed=(document.all)? marqueespeed : Math.max(1, marqueespeed-1) //slow speed down by 1 for NS"
        Stream = Stream & vbCrLf & " var copyspeed=marqueespeed;"
        Stream = Stream & vbCrLf & " var pausespeed=(pauseit==0)? copyspeed: 0;"
        Stream = Stream & vbCrLf & " var iedom=document.all||document.getElementById;"
        Stream = Stream & vbCrLf & " if (iedom) {"
        Stream = Stream & vbCrLf & "   document.write('<span id=""temp"" style=""visibility:hidden;position:absolute;top:-100px;left:-9000px"">'+marqueecontent+'</span>');"
        Stream = Stream & vbCrLf & " }"
        Stream = Stream & vbCrLf & " var actualwidth='';"
        Stream = Stream & vbCrLf & " var cross_marquee, ns_marquee;"
        Stream = Stream & vbCrLf & " function populate(){"
        Stream = Stream & vbCrLf & " if (iedom){"
        Stream = Stream & vbCrLf & " cross_marquee=document.getElementById? document.getElementById(""iemarquee"") : document.all.iemarquee;"
        Stream = Stream & vbCrLf & " cross_marquee.style.left=parseInt(marqueewidth)+8+'px'"
        Stream = Stream & vbCrLf & " cross_marquee.innerHTML=marqueecontent"
        Stream = Stream & vbCrLf & " actualwidth=document.all? temp.scrollWidth : document.getElementById(""temp"").scrollWidth"
        Stream = Stream & vbCrLf & " } else if (document.layers){"
        Stream = Stream & vbCrLf & " ns_marquee=document.ns_marquee.document.ns_marquee2;"
        Stream = Stream & vbCrLf & " ns_marquee.left=parseInt(marqueewidth)+8;"
        Stream = Stream & vbCrLf & " ns_marquee.document.write(marqueecontent);"
        Stream = Stream & vbCrLf & " ns_marquee.document.close();"
        Stream = Stream & vbCrLf & " actualwidth=ns_marquee.document.width;"
        Stream = Stream & vbCrLf & " }"
        Stream = Stream & vbCrLf & " lefttime=setInterval('scrollmarquee()',20);"
        Stream = Stream & vbCrLf & " }"
        Stream = Stream & vbCrLf & " window.onload=populate;"
        Stream = Stream & vbCrLf & " function scrollmarquee(){"
        Stream = Stream & vbCrLf & "   if (iedom){"
        Stream = Stream & vbCrLf & "     if (parseInt(cross_marquee.style.left)>(actualwidth*(-1)+8)) {"
        Stream = Stream & vbCrLf & "       cross_marquee.style.left=parseInt(cross_marquee.style.left)-copyspeed+'px' }"
        Stream = Stream & vbCrLf & "     else {"
        Stream = Stream & vbCrLf & "       cross_marquee.style.left=parseInt(marqueewidth)+8+'px' }"
        Stream = Stream & vbCrLf & "   } else if (document.layers){"
        Stream = Stream & vbCrLf & "     if (ns_marquee.left>(actualwidth*(-1)+8))"
        Stream = Stream & vbCrLf & "       ns_marquee.left-=copyspeed"
        Stream = Stream & vbCrLf & "     else"
        Stream = Stream & vbCrLf & "       ns_marquee.left=parseInt(marqueewidth)+8"
        Stream = Stream & vbCrLf & "   }"
        Stream = Stream & vbCrLf & " }"
        Stream = Stream & vbCrLf & " if (iedom||document.layers){"
        Stream = Stream & vbCrLf & "   with (document){"
        Stream = Stream & vbCrLf & "     document.write('<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" bgcolor=""'+marqueebgcolor+'""><td align=""left"">');"
        Stream = Stream & vbCrLf & "     if (iedom){"
        Stream = Stream & vbCrLf & "       write('<div style=""position:relative;width:'+marqueewidth+';height:'+marqueeheight+';overflow:hidden"">');"
        Stream = Stream & vbCrLf & "       write('<div style=""position:absolute;width:'+marqueewidth+';height:'+marqueeheight+';background-color:'+marqueebgcolor+'"" onMouseover=""copyspeed=pausespeed"" onMouseout=""copyspeed=marqueespeed"">');"
        Stream = Stream & vbCrLf & "       write('<div id=""iemarquee"" style=""position:absolute;left:0px;top:0px""></div>');"
        Stream = Stream & vbCrLf & "       write('</div></div>');"
        Stream = Stream & vbCrLf & "     } else if (document.layers){"
        Stream = Stream & vbCrLf & "       write('<ilayer width='+marqueewidth+' height='+marqueeheight+' name=""ns_marquee"" bgColor='+marqueebgcolor+'>');"
        Stream = Stream & vbCrLf & "       write('<layer name=""ns_marquee2"" left=0 top=0 onMouseover=""copyspeed=pausespeed"" onMouseout=""copyspeed=marqueespeed""></layer>');"
        Stream = Stream & vbCrLf & "       write('</ilayer>');"
        Stream = Stream & vbCrLf & "     }"
        Stream = Stream & vbCrLf & "   document.write('</td></table>');"
        Stream = Stream & vbCrLf & "   }"
        Stream = Stream & vbCrLf & " }"
        Stream = Stream & vbCrLf & " </script>"
    End If
    '
    GetMarquee = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("MarqueeClass", "GetMarquee", Err.Number, Err.Source, Err.Description, True, False)
    End Function
