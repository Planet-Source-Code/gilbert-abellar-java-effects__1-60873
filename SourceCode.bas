Attribute VB_Name = "SourceCode"
Option Explicit
Option Compare Text
Option Base 1
Option Private Module

Public Function DetectEntry(Text As String) As String
    If Not InStr(Text, ":") = 0 Then
        DetectEntry = Left$(Mid$(Text, InStrRev(Text, "/") + 1, Len(Text) - InStrRev(Text, "/")), InStrRev(Mid$(Text, InStrRev(Text, "/") + 1, Len(Text) - InStrRev(Text, "/")), ".") - 1)
        Mode = False
    Else
        DetectEntry = Left$(Text, InStrRev(Text, ".") - 1)
        Mode = True
    End If
End Function

Public Function ApplicationRunning() As Boolean
    ApplicationRunning = Not (App.PrevInstance = False)
End Function

Public Function FormMousePointer(FormName As Form, MousePointerNumber As Long) As Long
    FormMousePointer = MousePointerNumber
End Function

Public Function FormMouse(FormName As Form, Number As Long) As Long
    With FormName
        .Label1.MousePointer = FormMousePointer(FormName, Number)
        .Label2.MousePointer = FormMousePointer(FormName, Number)
        .Label3.MousePointer = FormMousePointer(FormName, Number)
        .Label4.MousePointer = FormMousePointer(FormName, Number)
        .lblDisplay.MousePointer = FormMousePointer(FormName, Number)
        .picProgress.MousePointer = FormMousePointer(FormName, Number)
        .cmbOptions.MousePointer = FormMousePointer(FormName, Number)
        .cmdClear.MousePointer = FormMousePointer(FormName, Number)
        .cmdCopy.MousePointer = FormMousePointer(FormName, Number)
        .cmdCreate.MousePointer = FormMousePointer(FormName, Number)
        .cmdTextGen.MousePointer = FormMousePointer(FormName, Number)
        .Frame1.MousePointer = FormMousePointer(FormName, Number)
        .Frame2.MousePointer = FormMousePointer(FormName, Number)
        .Frame3.MousePointer = FormMousePointer(FormName, Number)
        .txtInt.MousePointer = FormMousePointer(FormName, Number)
        .txtMes.MousePointer = FormMousePointer(FormName, Number)
        .txtOutput.MousePointer = FormMousePointer(FormName, Number)
        .VS.MousePointer = FormMousePointer(FormName, Number)
        For X = 0 To 3
            .txtMen(X).MousePointer = FormMousePointer(FormName, Number)
        Next X
        FormMouse = FormMousePointer(FormName, Number)
    End With
End Function

Public Function CreateDocument(ByVal Index As Integer, Optional ByVal Text As String = vbNullString, Optional ByVal Interval As String = 10, Optional ByVal Title As String = vbNullString) As String
    CreateDocument = "<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "iso-8859-1" & """" & "?>" & vbNewLine
    CreateDocument = CreateDocument & "<%@LANGUAGE=" & """" & "JAVASCRIPT" & """" & " CODEPAGE=" & """" & "1252" & """" & "%>" & vbNewLine
    CreateDocument = CreateDocument & "<html>" & vbNewLine
    CreateDocument = CreateDocument & "</head" & vbNewLine
    CreateDocument = CreateDocument & "<META NAME=" & """" & "Author" & """" & " GILBERT O. ABELLAR" & """" & ">" & vbNewLine
    CreateDocument = CreateDocument & "<meta http-equiv=" & """" & "pragma" & """" & " content=" & """" & "no-cache" & """" & ">" & vbNewLine
    CreateDocument = CreateDocument & "<title>" & Title & "&reg;</title>" & vbNewLine
    CreateDocument = CreateDocument & "<meta http-equiv=" & """" & "Content-Type" & """" & " content=" & """" & "text/html; charset=iso-8859-1" & """" & ">" & vbNewLine
    CreateDocument = CreateDocument & "</head>" & vbNewLine
    Print #1, CreateDocument
    Print #1, TextEffects(Index, Text, Interval)
    If Index >= 0 And Index <= 7 Then
        Print #1, "</head>"
        Print #1, vbNullString
        Print #1, "<body>"
        Print #1, vbNullString
        Print #1, "<a href=" & """" & "javascript:void(window.close())" & """" & ">Click here to close this window.</a>"
        Print #1, "</body>"
        Print #1, "</html>"
    ElseIf Index = 10 Or Index = 11 Or Index = 12 Or Index = 16 Or Index = 17 Or Index = 18 Then
        Print #1, vbNullString
        Print #1, "<a href=" & """" & "javascript:void(window.close())" & """" & ">Click here to close this window.</a>"
        Print #1, "</body>"
        Print #1, "</html>"
    Else
        Print #1, "</html>"
    End If
End Function

Public Function TextEffects(ByVal Index As Integer, Optional ByVal Text As String = vbNullString, Optional ByVal Interval As Integer = 10) As String
    Select Case Index
        Case 0
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var ip=" & """" & Text & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var op= " & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var cur=0;" & vbNewLine
            TextEffects = TextEffects & "function sstatus(){" & vbNewLine
            TextEffects = TextEffects & "op = ip.substring(cur,ip.length) + ip.substring(0,cur);" & vbNewLine
            TextEffects = TextEffects & "window.status = op;" & vbNewLine
            TextEffects = TextEffects & "cur++;" & vbNewLine
            TextEffects = TextEffects & "if(cur>ip.length)" & vbNewLine
            TextEffects = TextEffects & "cur=0;" & vbNewLine
            TextEffects = TextEffects & "setTimeout('sstatus();'" & "," & Interval & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "sstatus();" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 1
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var ip=" & """" & Text & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var op= " & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var cur=0;" & vbNewLine
            TextEffects = TextEffects & "function sstatus(){" & vbNewLine
            TextEffects = TextEffects & "op = ip.substring(cur,ip.length) + ip.substring(0,cur);" & vbNewLine
            TextEffects = TextEffects & "window.status = op;" & vbNewLine
            TextEffects = TextEffects & "cur--;" & vbNewLine
            TextEffects = TextEffects & "if(cur<0)" & vbNewLine
            TextEffects = TextEffects & "cur=ip.length;" & vbNewLine
            TextEffects = TextEffects & "setTimeout('sstatus();'" & "," & Interval & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "sstatus();" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 2
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var ip=" & """" & Text & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var op= " & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var cur=0;" & vbNewLine
            TextEffects = TextEffects & "function sstatus(){" & vbNewLine
            TextEffects = TextEffects & "op = ip.substring(0,cur);" & vbNewLine
            TextEffects = TextEffects & "window.status = op;" & vbNewLine
            TextEffects = TextEffects & "cur++;" & vbNewLine
            TextEffects = TextEffects & "if(cur>ip.length)" & vbNewLine
            TextEffects = TextEffects & "cur=0;" & vbNewLine
            TextEffects = TextEffects & "setTimeout('sstatus();'" & "," & Interval & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "sstatus();" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 3
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var ip=" & """" & Text & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var op= " & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var cur=0;" & vbNewLine
            TextEffects = TextEffects & "var space=10;" & vbNewLine
            TextEffects = TextEffects & "var sp=" & """" & "             " & """" & ";" & vbNewLine
            TextEffects = TextEffects & "function sstatus(){" & vbNewLine
            TextEffects = TextEffects & "op = ip.substring(0,cur) + sp.substring(0,space) + " & vbNewLine
            TextEffects = TextEffects & "ip.substring(cur,cur+1);" & vbNewLine
            TextEffects = TextEffects & "window.status = op;" & vbNewLine
            TextEffects = TextEffects & "space--;" & vbNewLine
            TextEffects = TextEffects & "if (space<-1){" & vbNewLine
            TextEffects = TextEffects & "space=10;" & vbNewLine
            TextEffects = TextEffects & "cur++;" & vbNewLine
            TextEffects = TextEffects & "if(cur>ip.length)" & vbNewLine
            TextEffects = TextEffects & "cur=0;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "setTimeout('sstatus();'" & "," & Interval & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "sstatus();" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 4
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var ip=" & """" & Text & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var op= " & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var cur=0;" & vbNewLine
            TextEffects = TextEffects & "var space=10;" & vbNewLine
            TextEffects = TextEffects & "var sp=" & """" & "                " & """" & ";" & vbNewLine
            TextEffects = TextEffects & "function sstatus(){" & vbNewLine
            TextEffects = TextEffects & "op =" & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "for(var v=0;v<ip.length;v++)" & vbNewLine
            TextEffects = TextEffects & "op = op + ip.substring(v,v+1) + sp.substring(0,cur);" & vbNewLine
            TextEffects = TextEffects & "window.status = op;" & vbNewLine
            TextEffects = TextEffects & "cur--;" & vbNewLine
            TextEffects = TextEffects & "if(cur<-10)" & vbNewLine
            TextEffects = TextEffects & "cur=20;" & vbNewLine
            TextEffects = TextEffects & "setTimeout('sstatus();'" & "," & Interval & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "sstatus();" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 5
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var osd = " & """" & "   " & """" & vbNewLine
            TextEffects = TextEffects & "osd +=" & """" & Text & """" & ";" & vbNewLine
            TextEffects = TextEffects & "osd +=" & """" & "          " & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var timer;" & vbNewLine
            TextEffects = TextEffects & "var msg = """";" & vbNewLine
            TextEffects = TextEffects & "function scrollMaster () {" & vbNewLine
            TextEffects = TextEffects & "msg = customDateSpring(new Date())" & vbNewLine
            TextEffects = TextEffects & "clearTimeout(timer)" & vbNewLine
            TextEffects = TextEffects & "msg += "" "" + showtime() + "" "" + osd" & vbNewLine
            TextEffects = TextEffects & "for (var i= 0; i < 100; i++){" & vbNewLine
            TextEffects = TextEffects & "msg = "" "" + msg;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "scrollMe()" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function scrollMe(){" & vbNewLine
            TextEffects = TextEffects & "window.status = msg;" & vbNewLine
            TextEffects = TextEffects & "msg = msg.substring(1, msg.length) + msg.substring(0,1);" & vbNewLine
            TextEffects = TextEffects & "timer = setTimeout(""scrollMe()"", " & Interval & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function showtime (){" & vbNewLine
            TextEffects = TextEffects & "var now = new Date();" & vbNewLine
            TextEffects = TextEffects & "var hours= now.getHours();" & vbNewLine
            TextEffects = TextEffects & "var minutes= now.getMinutes();" & vbNewLine
            TextEffects = TextEffects & "var seconds= now.getSeconds();" & vbNewLine
            TextEffects = TextEffects & "var months= now.getMonth();" & vbNewLine
            TextEffects = TextEffects & "var dates= now.getDate();" & vbNewLine
            TextEffects = TextEffects & "var years= now.getYear();" & vbNewLine
            TextEffects = TextEffects & "var timeValue = """"" & vbNewLine
            TextEffects = TextEffects & "timeValue += ((months >9) ? """" : "" "")" & vbNewLine
            TextEffects = TextEffects & "timeValue += ((dates >9) ? """" : "" "")" & vbNewLine
            TextEffects = TextEffects & "timeValue = ( months +1)" & vbNewLine
            TextEffects = TextEffects & "timeValue +=""/""+ dates" & vbNewLine
            TextEffects = TextEffects & "timeValue +=""/""+  years" & vbNewLine
            TextEffects = TextEffects & "var ap = ""A.M.""" & vbNewLine
            TextEffects = TextEffects & "if (hours == 12) {" & vbNewLine
            TextEffects = TextEffects & "ap=""P.M.""" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if (hours == 0) {" & vbNewLine
            TextEffects = TextEffects & "hours = 12" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if(hours >= 13){" & vbNewLine
            TextEffects = TextEffects & "hours -= 12;" & vbNewLine
            TextEffects = TextEffects & "ap = ""P.M.""" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "var timeValue2 = "" "" + hours" & vbNewLine
            TextEffects = TextEffects & "timeValue2 += ((minutes < 10) ? "":0"":"":"") + minutes + "" "" + ap" & vbNewLine
            TextEffects = TextEffects & "return timeValue2;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function MakeArray(n) {" & vbNewLine
            TextEffects = TextEffects & "this.length = n" & vbNewLine
            TextEffects = TextEffects & "return this" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "monthNames = new MakeArray(12)" & vbNewLine
            TextEffects = TextEffects & "monthNames[1] = ""Janurary""" & vbNewLine
            TextEffects = TextEffects & "monthNames[2] = ""February""" & vbNewLine
            TextEffects = TextEffects & "monthNames[3] = ""March""" & vbNewLine
            TextEffects = TextEffects & "monthNames[4] = ""April""" & vbNewLine
            TextEffects = TextEffects & "monthNames[5] = ""May""" & vbNewLine
            TextEffects = TextEffects & "monthNames[6] = ""June""" & vbNewLine
            TextEffects = TextEffects & "monthNames[7] = ""July""" & vbNewLine
            TextEffects = TextEffects & "monthNames[8] = ""August""" & vbNewLine
            TextEffects = TextEffects & "monthNames[9] = ""Sept.""" & vbNewLine
            TextEffects = TextEffects & "monthNames[10] = ""Oct.""" & vbNewLine
            TextEffects = TextEffects & "monthNames[11] = ""Nov.""" & vbNewLine
            TextEffects = TextEffects & "monthNames[12] = ""Dec.""" & vbNewLine
            TextEffects = TextEffects & "daysNames = new MakeArray(7)" & vbNewLine
            TextEffects = TextEffects & "daysNames[1] = ""Sunday""" & vbNewLine
            TextEffects = TextEffects & "daysNames[2] = ""Monday""" & vbNewLine
            TextEffects = TextEffects & "daysNames[3] = ""Tuesday""" & vbNewLine
            TextEffects = TextEffects & "daysNames[4] = ""Wednesday""" & vbNewLine
            TextEffects = TextEffects & "daysNames[5] = ""Thursday""" & vbNewLine
            TextEffects = TextEffects & "daysNames[6] = ""Friday""" & vbNewLine
            TextEffects = TextEffects & "daysNames[7] = ""Saturday""" & vbNewLine
            TextEffects = TextEffects & "function customDateSpring(oneDate) {" & vbNewLine
            TextEffects = TextEffects & "var theDay = daysNames[oneDate.getDay() +1]" & vbNewLine
            TextEffects = TextEffects & "var theDate =oneDate.getDate()" & vbNewLine
            TextEffects = TextEffects & "var theMonth = monthNames[oneDate.getMonth() +1]" & vbNewLine
            TextEffects = TextEffects & "var dayth=""th""" & vbNewLine
            TextEffects = TextEffects & "if ((theDate == 1) || (theDate == 21) || (theDate == 31)) {" & vbNewLine
            TextEffects = TextEffects & "dayth=""st"";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if ((theDate == 2) || (theDate ==22)) {" & vbNewLine
            TextEffects = TextEffects & "dayth=""nd"";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if ((theDate== 3) || (theDate  == 23)) {" & vbNewLine
            TextEffects = TextEffects & "dayth=""rd"";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "return theDay + "", "" + theMonth + "" "" + theDate + dayth + "",""" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "scrollMaster();" & vbNewLine
            TextEffects = TextEffects & "</SCRIPT>" & vbNewLine
        Case 6
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "dCol='000000';" & vbNewLine
            TextEffects = TextEffects & "fCol='000000';" & vbNewLine
            TextEffects = TextEffects & "sCol='000000';" & vbNewLine
            TextEffects = TextEffects & "mCol='000000';" & vbNewLine
            TextEffects = TextEffects & "hCol='000000';" & vbNewLine
            TextEffects = TextEffects & "ClockHeight=40;" & vbNewLine
            TextEffects = TextEffects & "ClockWidth=40;" & vbNewLine
            TextEffects = TextEffects & "ClockFromMouseY=0;" & vbNewLine
            TextEffects = TextEffects & "ClockFromMouseX=100;" & vbNewLine
            TextEffects = TextEffects & "d=new" & vbNewLine
            TextEffects = TextEffects & "Array(" & """" & "SUNDAY" & """" & "," & """" & "MONDAY" & """" & "," & """" & "TUESDAY" & """" & "," & """" & "WEDNESDAY" & """" & "," & """" & "THURSDAY" & """" & "," & """" & "FRIDAY" & """" & "," & """" & "SATURDAY" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "m=new" & vbNewLine
            TextEffects = TextEffects & "Array(" & """" & "JANUARY" & """" & "," & """" & "FEBRUARY" & """" & "," & """" & "MARCH" & """" & "," & """" & "APRIL" & """" & "," & """" & "MAY" & """" & "," & """" & "JUNE" & """" & "," & """" & "JULY" & """" & "," & """" & "AUGUST" & """" & "," & """" & "SEPTEMBER" & """" & "," & """" & "OCTOBER" & """" & "," & """" & "NOVEMBER" & """" & "," & """" & "DECEMBER" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "date=new Date();" & vbNewLine
            TextEffects = TextEffects & "day=date.getDate();" & vbNewLine
            TextEffects = TextEffects & "year=date.getYear();" & vbNewLine
            TextEffects = TextEffects & "if (year < 2000) year = year+1900;" & vbNewLine
            TextEffects = TextEffects & "TodaysDate=" & """" & " " & """" & "+d[date.getDay()]+" & """" & " " & """" & "+day+" & """" & " " & """" & "+m[date.getMonth()]+" & """" & " " & """" & "+year;" & vbNewLine
            TextEffects = TextEffects & "D=TodaysDate.split('');" & vbNewLine
            TextEffects = TextEffects & "H='...';" & vbNewLine
            TextEffects = TextEffects & "H=H.split('');" & vbNewLine
            TextEffects = TextEffects & "M='....';" & vbNewLine
            TextEffects = TextEffects & "M=M.split('');" & vbNewLine
            TextEffects = TextEffects & "S='.....';" & vbNewLine
            TextEffects = TextEffects & "S=S.split('');" & vbNewLine
            TextEffects = TextEffects & "Face='1 2 3 4 5 6 7 8 9 10 11 12';" & vbNewLine
            TextEffects = TextEffects & "font='Arial';" & vbNewLine
            TextEffects = TextEffects & "size=1;" & vbNewLine
            TextEffects = TextEffects & "speed=0.6;" & vbNewLine
            TextEffects = TextEffects & "ns=(document.layers);" & vbNewLine
            TextEffects = TextEffects & "ie=(document.all);" & vbNewLine
            TextEffects = TextEffects & "Face=Face.split(' ');" & vbNewLine
            TextEffects = TextEffects & "n=Face.length;" & vbNewLine
            TextEffects = TextEffects & "a=size*10;" & vbNewLine
            TextEffects = TextEffects & "ymouse=0;" & vbNewLine
            TextEffects = TextEffects & "xmouse=0;" & vbNewLine
            TextEffects = TextEffects & "scrll=0;" & vbNewLine
            TextEffects = TextEffects & "props=" & """" & "<font face=" & """" & "+font+" & """" & " size=" & """" & "+size+" & """" & " color=" & """" & "+fCol+" & """" & "><B>" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "props2=" & """" & "<font face=" & """" & "+font+" & """" & " size=" & """" & "+size+" & """" & " color=" & """" & "+dCol+" & """" & "><B>" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "Split=360/n;" & vbNewLine
            TextEffects = TextEffects & "Dsplit=360/D.length;" & vbNewLine
            TextEffects = TextEffects & "HandHeight = ClockHeight / 4.5" & vbNewLine
            TextEffects = TextEffects & "HandWidth = ClockWidth / 4.5" & vbNewLine
            TextEffects = TextEffects & "HandY=-7;" & vbNewLine
            TextEffects = TextEffects & "HandX=-2.5;" & vbNewLine
            TextEffects = TextEffects & "scrll=0;" & vbNewLine
            TextEffects = TextEffects & "step=0.06;" & vbNewLine
            TextEffects = TextEffects & "currStep=0;" & vbNewLine
            TextEffects = TextEffects & "y=new Array();x=new Array();Y=new Array();X=new Array();" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < n; i++){y[i]=0;x[i]=0;Y[i]=0;X[i]=0}" & vbNewLine
            TextEffects = TextEffects & "Dy=new Array();Dx=new Array();DY=new Array();DX=new Array();" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < D.length; i++){Dy[i]=0;Dx[i]=0;DY[i]=0;DX[i]=0;}" & vbNewLine
            TextEffects = TextEffects & "if (ns){" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < D.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<layer name=" & """" & "nsDate'+i+'" & """" & " top=0 left=0 height='+a+' width='+a+'><center>'+props2+D[i]+'</font></center></layer>');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < n; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<layer name=" & """" & "nsFace'+i+'" & """" & " top=0 left=0 height='+a+' width='+a+'><center>'+props+Face[i]+'</font></center></layer>');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < S.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<layer name=nsSeconds'+i+' top=0 left=0 width=15 height=15><font face=Arial size=3 color='+sCol+'><center><b>'+S[i]+'</b></center></font></layer>');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < M.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<layer name=nsMinutes'+i+' top=0 left=0 width=15 height=15><font face=Arial size=3 color='+mCol+'><center><b>'+M[i]+'</b></center></font></layer>');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < H.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<layer name=nsHours'+i+' top=0 left=0 width=15 height=15><font face=Arial size=3 color='+hCol+'><center><b>'+H[i]+'</b></center></font></layer>');" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if (ie){" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "Od" & """" & " style=" & """" & "position:absolute;top:0px;left:0px" & """" & "><div style=" & """" & "position:relative" & """" & ">');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < D.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "ieDate" & """" & " style=" & """" & "position:absolute;top:0px;left:0;height:'+a+';width:'+a+';text-align:center" & """" & ">'+props2+D[i]+'</B></font></div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('</div></div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "Of" & """" & " style=" & """" & "position:absolute;top:0px;left:0px" & """" & "><div style=" & """" & "position:relative" & """" & ">');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < n; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "ieFace" & """" & " style=" & """" & "position:absolute;top:0px;left:0;height:'+a+';width:'+a+';text-align:center" & """" & ">'+props+Face[i]+'</B></font></div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('</div></div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "Oh" & """" & " style=" & """" & "position:absolute;top:0px;left:0px" & """" & "><div style=" & """" & "position:relative" & """" & ">');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < H.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "ieHours" & """" & " style=" & """" & "position:absolute;width:16px;height:16px;font-family:Arial;font-size:16px;color:'+hCol+';text-align:center;font-weight:bold" & """" & ">'+H[i]+'</div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('</div></div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "Om" & """" & " style=" & """" & "position:absolute;top:0px;left:0px" & """" & "><div style=" & """" & "position:relative" & """" & ">');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < M.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "ieMinutes" & """" & " style=" & """" & "position:absolute;width:16px;height:16px;font-family:Arial;font-size:16px;color:'+mCol+';text-align:center;font-weight:bold" & """" & ">'+M[i]+'</div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('</div></div>')" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "Os" & """" & " style=" & """" & "position:absolute;top:0px;left:0px" & """" & "><div style=" & """" & "position:relative" & """" & ">');" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < S.length; i++)" & vbNewLine
            TextEffects = TextEffects & "document.write('<div id=" & """" & "ieSeconds" & """" & " style=" & """" & "position:absolute;width:16px;height:16px;font-family:Arial;font-size:16px;color:'+sCol+';text-align:center;font-weight:bold" & """" & ">'+S[i]+'</div>');" & vbNewLine
            TextEffects = TextEffects & "document.write('</div></div>')" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "(ns)?window.captureEvents(Event.MOUSEMOVE):0;" & vbNewLine
            TextEffects = TextEffects & "function Mouse(evnt){" & vbNewLine
            TextEffects = TextEffects & "ymouse = (ns)?evnt.pageY+ClockFromMouseY-(window.pageYOffset):event.y+ClockFromMouseY;" & vbNewLine
            TextEffects = TextEffects & "xmouse = (ns)?evnt.pageX+ClockFromMouseX:event.x+ClockFromMouseX;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "(ns)?window.onMouseMove=Mouse:document.onmousemove=Mouse;" & vbNewLine
            TextEffects = TextEffects & "function ClockAndAssign(){" & vbNewLine
            TextEffects = TextEffects & "time = new Date ();" & vbNewLine
            TextEffects = TextEffects & "secs = time.getSeconds();" & vbNewLine
            TextEffects = TextEffects & "sec = -1.57 + Math.PI * secs/30;" & vbNewLine
            TextEffects = TextEffects & "mins = time.getMinutes();" & vbNewLine
            TextEffects = TextEffects & "min = -1.57 + Math.PI * mins/30;" & vbNewLine
            TextEffects = TextEffects & "hr = time.getHours();" & vbNewLine
            TextEffects = TextEffects & "hrs = -1.575 + Math.PI * hr/6+Math.PI*parseInt(time.getMinutes())/360;" & vbNewLine
            TextEffects = TextEffects & "if (ie){" & vbNewLine
            TextEffects = TextEffects & "Od.style.top=window.document.body.scrollTop;" & vbNewLine
            TextEffects = TextEffects & "Of.style.top=window.document.body.scrollTop;" & vbNewLine
            TextEffects = TextEffects & "Oh.style.top=window.document.body.scrollTop;" & vbNewLine
            TextEffects = TextEffects & "Om.style.top=window.document.body.scrollTop;" & vbNewLine
            TextEffects = TextEffects & "Os.style.top=window.document.body.scrollTop;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < n; i++){" & vbNewLine
            TextEffects = TextEffects & "var F=(ns)?document.layers['nsFace'+i]:ieFace[i].style;" & vbNewLine
            TextEffects = TextEffects & "F.top=y[i] + ClockHeight*Math.sin(-1.0471 + i*Split*Math.PI/180)+scrll;" & vbNewLine
            TextEffects = TextEffects & "F.left=x[i] + ClockWidth*Math.cos(-1.0471 + i*Split*Math.PI/180);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < H.length; i++){" & vbNewLine
            TextEffects = TextEffects & "var HL=(ns)?document.layers['nsHours'+i]:ieHours[i].style;" & vbNewLine
            TextEffects = TextEffects & "HL.top=y[i]+HandY+(i*HandHeight)*Math.sin(hrs)+scrll;" & vbNewLine
            TextEffects = TextEffects & "HL.left=x[i]+HandX+(i*HandWidth)*Math.cos(hrs);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < M.length; i++){" & vbNewLine
            TextEffects = TextEffects & "var ML=(ns)?document.layers['nsMinutes'+i]:ieMinutes[i].style;" & vbNewLine
            TextEffects = TextEffects & "ML.top=y[i]+HandY+(i*HandHeight)*Math.sin(min)+scrll;" & vbNewLine
            TextEffects = TextEffects & "ML.left=x[i]+HandX+(i*HandWidth)*Math.cos(min);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < S.length; i++){" & vbNewLine
            TextEffects = TextEffects & "var SL=(ns)?document.layers['nsSeconds'+i]:ieSeconds[i].style;" & vbNewLine
            TextEffects = TextEffects & "SL.top=y[i]+HandY+(i*HandHeight)*Math.sin(sec)+scrll;" & vbNewLine
            TextEffects = TextEffects & "SL.left=x[i]+HandX+(i*HandWidth)*Math.cos(sec);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "for (i=0; i < D.length; i++){" & vbNewLine
            TextEffects = TextEffects & "var DL=(ns)?document.layers['nsDate'+i]:ieDate[i].style;" & vbNewLine
            TextEffects = TextEffects & "DL.top=Dy[i] + ClockHeight*1.5*Math.sin(currStep+i*Dsplit*Math.PI/180)+scrll;" & vbNewLine
            TextEffects = TextEffects & "DL.left=Dx[i] + ClockWidth*1.5*Math.cos(currStep+i*Dsplit*Math.PI/180);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "currStep-=step;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function Delay(){" & vbNewLine
            TextEffects = TextEffects & "scrll=(ns)?window.pageYOffset:0;" & vbNewLine
            TextEffects = TextEffects & "Dy[0]=Math.round(DY[0]+=((ymouse)-DY[0])*speed);" & vbNewLine
            TextEffects = TextEffects & "Dx[0]=Math.round(DX[0]+=((xmouse)-DX[0])*speed);" & vbNewLine
            TextEffects = TextEffects & "for (i=1; i < D.length; i++){" & vbNewLine
            TextEffects = TextEffects & "Dy[i]=Math.round(DY[i]+=(Dy[i-1]-DY[i])*speed);" & vbNewLine
            TextEffects = TextEffects & "Dx[i]=Math.round(DX[i]+=(Dx[i-1]-DX[i])*speed);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "y[0]=Math.round(Y[0]+=((ymouse)-Y[0])*speed);" & vbNewLine
            TextEffects = TextEffects & "x[0]=Math.round(X[0]+=((xmouse)-X[0])*speed);" & vbNewLine
            TextEffects = TextEffects & "for (i=1; i < n; i++){" & vbNewLine
            TextEffects = TextEffects & "y[i]=Math.round(Y[i]+=(y[i-1]-Y[i])*speed);" & vbNewLine
            TextEffects = TextEffects & "x[i]=Math.round(X[i]+=(x[i-1]-X[i])*speed);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "ClockAndAssign();" & vbNewLine
            TextEffects = TextEffects & "setTimeout('Delay()'," & frmJava.txtInt & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if (ns||ie)window.onload=Delay;" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 7
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var message=" & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "function clickIE() {if (document.all) {(message);return false;}}" & vbNewLine
            TextEffects = TextEffects & "function clickNS(e) {if " & vbNewLine
            TextEffects = TextEffects & "(document.layers||(document.getElementById&&!document.all)) {" & vbNewLine
            TextEffects = TextEffects & "if (e.which==2||e.which==3) {(message);return false;}}}" & vbNewLine
            TextEffects = TextEffects & "if (document.layers)" & vbNewLine
            TextEffects = TextEffects & "{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}" & vbNewLine
            TextEffects = TextEffects & "else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}" & vbNewLine
            TextEffects = TextEffects & "document.oncontextmenu=new Function(" & """" & "return false" & """" & ")" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 8
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "function shake() {" & vbNewLine
            TextEffects = TextEffects & "if (parent.moveBy) {" & vbNewLine
            TextEffects = TextEffects & "for (i = " & Interval & "; i > 0; i--) {" & vbNewLine
            TextEffects = TextEffects & "for (j = 2; j > 0; j--) {" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine & vbNewLine
            TextEffects = TextEffects & "<body>" & vbNewLine
            TextEffects = TextEffects & "<div align=" & """" & "center" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<input type=button onClick=" & """" & "shake()" & """" & " value=" & """" & "Shake Me!" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "</div>" & vbNewLine & vbNewLine
            TextEffects = TextEffects & "</body>" & vbNewLine
        Case 9
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "function shake() {" & vbNewLine
            TextEffects = TextEffects & "if (parent.moveBy) {" & vbNewLine
            TextEffects = TextEffects & "for (i = " & Interval & "; i > 0; i--) {" & vbNewLine
            TextEffects = TextEffects & "for (j = 2; j > 0; j--) {" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(i,0);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(0,-i);" & vbNewLine
            TextEffects = TextEffects & "parent.moveBy(-i,0);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine & vbNewLine
            TextEffects = TextEffects & "<body>" & vbNewLine
            TextEffects = TextEffects & "<div align=" & """" & "center" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<input type=button onClick=" & """" & "shake()" & """" & " value=" & """" & "Shake Screen" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "</div>" & vbNewLine & vbNewLine
            TextEffects = TextEffects & "</body>" & vbNewLine
        Case 10
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "function checkVersion4() {" & vbNewLine
            TextEffects = TextEffects & "var x = navigator.appVersion;" & vbNewLine
            TextEffects = TextEffects & "y = x.substring(0,4);" & vbNewLine
            TextEffects = TextEffects & "if (y>=4) setVariables();moveOB();" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function setVariables() {" & vbNewLine
            TextEffects = TextEffects & "if (navigator.appName == " & """" & "Netscape" & """" & ") {" & vbNewLine
            TextEffects = TextEffects & "h=" & """" & ".left=" & """" & ";v=" & """" & ".top=" & """" & ";dS=" & """" & "document." & """" & ";sD=" & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "else{" & vbNewLine
            TextEffects = TextEffects & "h=" & """" & ".pixelLeft=" & """" & ";v=" & """" & ".pixelTop=" & """" & ";dS=" & """" & """" & ";sD=" & """" & ".style" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "objectX = " & """" & "object11" & """" & vbNewLine
            TextEffects = TextEffects & "XX=-70;" & vbNewLine
            TextEffects = TextEffects & "YY=-70;" & vbNewLine
            TextEffects = TextEffects & "OB=11;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function setObject(a) {" & vbNewLine
            TextEffects = TextEffects & "objectX=" & """" & "object" & """" & "+a;" & vbNewLine
            TextEffects = TextEffects & "OB=a;" & vbNewLine
            TextEffects = TextEffects & "XX=eval(" & """" & "xpos" & """" & "+a);" & vbNewLine
            TextEffects = TextEffects & "YY=eval(" & """" & "ypos" & """" & "+a);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function getObject() {" & vbNewLine
            TextEffects = TextEffects & "if (isNav) document.captureEvents(Event.MOUSEMOVE);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function releaseObject() {" & vbNewLine
            TextEffects = TextEffects & "if (isNav) document.releaseEvents(Event.MOUSEMOVE);" & vbNewLine
            TextEffects = TextEffects & "check=" & """" & "no" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "objectX=" & """" & "object11" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "document.close();" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function moveOB() {" & vbNewLine
            TextEffects = TextEffects & "eval(dS + objectX + sD + h + Xpos);" & vbNewLine
            TextEffects = TextEffects & "eval(dS + objectX + sD + v + Ypos);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "var isNav = (navigator.appName.indexOf(" & """" & "Netscape" & """" & ") !=-1);" & vbNewLine
            TextEffects = TextEffects & "var isIE = (navigator.appName.indexOf(" & """" & "Microsoft" & """" & ") !=-1);" & vbNewLine
            TextEffects = TextEffects & "nsValue=(document.layers);" & vbNewLine
            TextEffects = TextEffects & "check=" & """" & "no" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "function MoveHandler(e) {" & vbNewLine
            TextEffects = TextEffects & "Xpos = (isIE) ? event.clientX : e.pageX;" & vbNewLine
            TextEffects = TextEffects & "Ypos = (nsValue) ? e.pageY : event.clientY;" & vbNewLine
            TextEffects = TextEffects & "if (check==" & """" & "no" & """" & ") {" & vbNewLine
            TextEffects = TextEffects & "diffX=XX-Xpos;" & vbNewLine
            TextEffects = TextEffects & "diffY=YY-Ypos;" & vbNewLine
            TextEffects = TextEffects & "check=" & """" & "yes" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "if (objectX==" & """" & "object11" & """" & ") check=" & """" & "no" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "Xpos+=diffX;" & vbNewLine
            TextEffects = TextEffects & "Ypos+=diffY;" & vbNewLine
            TextEffects = TextEffects & "if (OB==" & """" & "1" & """" & ") xpos1=Xpos,ypos1=Ypos;" & vbNewLine
            TextEffects = TextEffects & "moveOB();" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if (isNav) {" & vbNewLine
            TextEffects = TextEffects & "document.captureEvents(Event.CLICK);" & vbNewLine
            TextEffects = TextEffects & "document.captureEvents(Event.DBLCLICK);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "xpos1=50;" & vbNewLine
            TextEffects = TextEffects & "ypos1=50;" & vbNewLine
            TextEffects = TextEffects & "xpos11 = -50;" & vbNewLine
            TextEffects = TextEffects & "ypos11 = -50;" & vbNewLine
            TextEffects = TextEffects & "Xpos=5;" & vbNewLine
            TextEffects = TextEffects & "Ypos=5;" & vbNewLine
            TextEffects = TextEffects & "document.onmousemove = MoveHandler;" & vbNewLine
            TextEffects = TextEffects & "document.onclick = getObject;" & vbNewLine
            TextEffects = TextEffects & "document.ondblclick = releaseObject;" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine & vbNewLine
            TextEffects = TextEffects & "<BODY OnLoad=" & """" & "checkVersion4()" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<b>Click on " & """" & "Moveable Menu" & """" & " to pick<br>" & vbNewLine
            TextEffects = TextEffects & " it up and Double Click to drop it!</b>" & vbNewLine
            TextEffects = TextEffects & "<br>" & vbNewLine
            TextEffects = TextEffects & "<div id=" & """" & "object1" & """" & " style=" & """" & "position:absolute; visibility:show; left:50px; top:50px; z-index:2" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<table border=1 cellpadding=5>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td bgcolor=eeeeee><a href=" & """" & "javascript:void(0)" & """" & " onmousedown=" & """" & "setObject(1)" & """" & ">Movable Menu</a></td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td>" & vbNewLine
            TextEffects = TextEffects & "<br>" & vbNewLine
            TextEffects = TextEffects & "<A HREF=" & """" & "http://" & frmJava.txtMen(0) & """" & ">" & frmJava.txtMen(0) & "</a><br>" & vbNewLine
            TextEffects = TextEffects & "<A HREF=" & """" & "http://" & frmJava.txtMen(1) & """" & ">" & frmJava.txtMen(1) & "</a><br>" & vbNewLine
            TextEffects = TextEffects & "<A HREF=" & """" & "http://" & frmJava.txtMen(2) & """" & ">" & frmJava.txtMen(2) & "</a><br>" & vbNewLine
            TextEffects = TextEffects & "<A HREF=" & """" & "http://" & frmJava.txtMen(3) & """" & ">" & frmJava.txtMen(3) & "</a><br>" & vbNewLine
            TextEffects = TextEffects & "</td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "</table>" & vbNewLine
            TextEffects = TextEffects & "</div>" & vbNewLine
            TextEffects = TextEffects & "<div id=" & """" & "object11" & """" & " style=" & """" & "position:absolute; visibility:show; left:-70px; top:-70px; z-index:2" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "</div>" & vbNewLine
            TextEffects = TextEffects & "<p><center>" & vbNewLine
            TextEffects = TextEffects & "</center><p>" & vbNewLine
        Case 11
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "function setVariables() {" & vbNewLine
            TextEffects = TextEffects & "if (navigator.appName == " & """" & "Netscape" & """" & ") {" & vbNewLine
            TextEffects = TextEffects & "v=" & """" & ".top=" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "dS=" & """" & "document." & """" & ";" & vbNewLine
            TextEffects = TextEffects & "sD=" & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "y=" & """" & "window.pageYOffset" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "else {" & vbNewLine
            TextEffects = TextEffects & "v=" & """" & ".pixelTop=" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "dS=" & """" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "sD=" & """" & ".style" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "y=" & """" & "document.body.scrollTop" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function checkLocation() {" & vbNewLine
            TextEffects = TextEffects & "object=" & """" & "object1" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "yy=eval(y);" & vbNewLine
            TextEffects = TextEffects & "eval(dS+object+sD+v+yy);" & vbNewLine
            TextEffects = TextEffects & "setTimeout(" & """" & "checkLocation()" & """" & ",10);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine & vbNewLine
            TextEffects = TextEffects & "<BODY OnLoad=" & """" & "setVariables();checkLocation()" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<div id=" & """" & "object1" & """" & " style=" & """" & "position:absolute; visibility:show; left:0px; top:0px; z-index:2" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<table width=130 border=0 cellspacing=20 cellpadding=0>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td><CENTER>Menu Bar</CENTER></td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td><a href=" & """" & "http://" & frmJava.txtMen(0) & """" & " >" & frmJava.txtMen(0) & "</a></td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td><a href=" & """" & "http://" & frmJava.txtMen(1) & """" & " >" & frmJava.txtMen(1) & "</a></td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td><a href=" & """" & "http://" & frmJava.txtMen(2) & """" & " >" & frmJava.txtMen(2) & "</a></td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td><a href=" & """" & "http://" & frmJava.txtMen(3) & """" & " >" & frmJava.txtMen(3) & "</a></td>" & vbNewLine
            TextEffects = TextEffects & "</tr>" & vbNewLine
            TextEffects = TextEffects & "</table>" & vbNewLine
            TextEffects = TextEffects & "</div>" & vbNewLine
            TextEffects = TextEffects & "<table>" & vbNewLine
            TextEffects = TextEffects & "<tr>" & vbNewLine
            TextEffects = TextEffects & "<td width=130>" & vbNewLine
            TextEffects = TextEffects & "<font color=" & """" & "white" & """" & ">&nbsp; </font>" & vbNewLine
            TextEffects = TextEffects & "</td>" & vbNewLine
            TextEffects = TextEffects & "<td>" & vbNewLine
            TextEffects = TextEffects & "</td></tr>" & vbNewLine
            TextEffects = TextEffects & "</table>" & vbNewLine
            TextEffects = TextEffects & "<p><center>" & vbNewLine
            TextEffects = TextEffects & "</center><p>" & vbNewLine
        Case 12
            TextEffects = "<body>" & vbNewLine
            TextEffects = TextEffects & "<CENTER>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='green'" & """" & ">Green</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='greem'" & """" & ">Bright Green</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='seagreen'" & """" & ">Sea Green</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='red'" & """" & ">Red</a>]<BR>" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='magenta'" & """" & ">Magenta</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='fusia'" & """" & ">Fusia</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='pink'" & """" & ">Pink</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='purple'" & """" & ">Purple</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='cyan'" & """" & ">Cyan</a>]<BR>" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='navy'" & """" & ">Navy</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='blue'" & """" & ">Blue</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='royalblue'" & """" & ">Royal Blue</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='Skyblue'" & """" & ">Sky Blue</a>]<BR>" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='yellow'" & """" & ">Yellow</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='brown'" & """" & ">Brown</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='almond'" & """" & ">Almond</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='white'" & """" & ">White</a>]<BR>" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='black'" & """" & ">Black</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='coral'" & """" & ">Coral</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='olivedrab'" & """" & ">Olive Drab</a>]" & vbNewLine
            TextEffects = TextEffects & "[<a href=" & """" & "/" & """" & vbNewLine
            TextEffects = TextEffects & "onmouseover=" & """" & "document.bgColor='orange'" & """" & ">Orange</a>]" & vbNewLine
            TextEffects = TextEffects & "</CENTER>" & vbNewLine
        Case 13
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<!-- Begin" & vbNewLine
            TextEffects = TextEffects & "function jump2form() {" & vbNewLine
            TextEffects = TextEffects & "  document.dict_form.term.select();" & vbNewLine
            TextEffects = TextEffects & "  document.dict_form.term.focus();" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function isblank(s)" & vbNewLine
            TextEffects = TextEffects & "{" & vbNewLine
            TextEffects = TextEffects & "  for(var i = 0; i < s.length; i++) {" & vbNewLine
            TextEffects = TextEffects & "    var c = s.charAt(i);" & vbNewLine
            TextEffects = TextEffects & "    if ((c != ' ') && (c != '\n') && (c != '\t')) return false;" & vbNewLine
            TextEffects = TextEffects & "  }" & vbNewLine
            TextEffects = TextEffects & "  return true;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function formcheck() {" & vbNewLine
            TextEffects = TextEffects & "  var d = document.dict_form.db[1].checked;" & vbNewLine
            TextEffects = TextEffects & "  var e = document.dict_form.term.value;" & vbNewLine
            TextEffects = TextEffects & "  if ((e == null) || (e == " & """" & """" & ") || isblank(e)) {" & vbNewLine
            TextEffects = TextEffects & "    alert(" & """" & "Please enter a word to look up." & """" & ");" & vbNewLine
            TextEffects = TextEffects & "    jump2form();" & vbNewLine
            TextEffects = TextEffects & "  }" & vbNewLine
            TextEffects = TextEffects & "  else if (d == 1) {" & vbNewLine
            TextEffects = TextEffects & "    location.href = (" & """" & "http://www.thesaurus.com/cgi-bin/search?config=roget&words=" & """" & " + escape(e));" & vbNewLine
            TextEffects = TextEffects & "  }" & vbNewLine
            TextEffects = TextEffects & "  else {" & vbNewLine
            TextEffects = TextEffects & "    location.href = (" & """" & "http://www.dictionary.com/cgi-bin/dict.pl?term=" & """" & " + escape(e));" & vbNewLine
            TextEffects = TextEffects & "  }" & vbNewLine
            TextEffects = TextEffects & "  return false;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function ahdpop() {" & vbNewLine
            TextEffects = TextEffects & "win=window.open(" & """" & "http://www.dictionary.com/help/ahd3key.html" & """" & ",'AHDKey','width=500,height=330,toolbar=no,location=no,directories=no,menubar=no,scrollbars=yes,resizable=yes');" & vbNewLine
            TextEffects = TextEffects & "  return false;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "//  End -->" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine
            TextEffects = TextEffects & "<body>" & vbNewLine
            TextEffects = TextEffects & "<FORM NAME=" & """" & "dict_form" & """" & " METHOD=" & """" & "GET" & """" & " ACTION=" & """" & "/cgi-bin/dict.pl" & """" & " onsubmit=" & """" & "return formcheck();" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<center>" & vbNewLine
            TextEffects = TextEffects & "<h2>Look up:</h2>" & vbNewLine
            TextEffects = TextEffects & "<INPUT TYPE=" & """" & "text" & """" & " NAME=" & """" & "term" & """" & " SIZE=17 MAXLENGTH=48 VALUE=" & """" & """" & " style=" & """" & "font-size:11pt;" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "   <INPUT TYPE=" & """" & "submit" & """" & " VALUE=" & """" & "OK" & """" & " STYLE=" & """" & "font-size:11pt;" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "   <INPUT NAME=" & """" & "Reset" & """" & " type=" & """" & "reset" & """" & " value=" & """" & "Clear" & """" & " onClick=" & """" & "document.dict_form.term.focus();" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "   <h3>Search:</h3>" & vbNewLine
            TextEffects = TextEffects & "<INPUT TYPE=" & """" & "RADIO" & """" & " NAME=" & """" & "db" & """" & " VALUE=" & """" & "*" & """" & " CHECKED>Dictionary" & vbNewLine
            TextEffects = TextEffects & "<INPUT TYPE=" & """" & "RADIO" & """" & " NAME=" & """" & "db" & """" & " VALUE=" & """" & "roget" & """" & ">Thesaurus" & vbNewLine
            TextEffects = TextEffects & "</center>" & vbNewLine
            TextEffects = TextEffects & "</FORM>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "</body>" & vbNewLine
        Case 14
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<!-- Begin" & vbNewLine
            TextEffects = TextEffects & "function GetCookie (name) {" & vbNewLine
            TextEffects = TextEffects & "var arg = name + " & """" & "=" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var alen = arg.length;" & vbNewLine
            TextEffects = TextEffects & "var clen = document.cookie.length;" & vbNewLine
            TextEffects = TextEffects & "var i = 0;" & vbNewLine
            TextEffects = TextEffects & "while (i < clen) {" & vbNewLine
            TextEffects = TextEffects & "var j = i + alen;" & vbNewLine
            TextEffects = TextEffects & "if (document.cookie.substring(i, j) == arg)" & vbNewLine
            TextEffects = TextEffects & "return getCookieVal (j);" & vbNewLine
            TextEffects = TextEffects & "i = document.cookie.indexOf(" & """" & " " & """" & ", i) + 1;" & vbNewLine
            TextEffects = TextEffects & "if (i == 0) break;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "return null;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function SetCookie (name, value) {" & vbNewLine
            TextEffects = TextEffects & "var argv = SetCookie.arguments;" & vbNewLine
            TextEffects = TextEffects & "var argc = SetCookie.arguments.length;" & vbNewLine
            TextEffects = TextEffects & "var expires = (argc > 2) ? argv[2] : null;" & vbNewLine
            TextEffects = TextEffects & "var path = (argc > 3) ? argv[3] : null;" & vbNewLine
            TextEffects = TextEffects & "var domain = (argc > 4) ? argv[4] : null;" & vbNewLine
            TextEffects = TextEffects & "var secure = (argc > 5) ? argv[5] : false;" & vbNewLine
            TextEffects = TextEffects & "document.cookie = name + " & """" & "=" & """" & " + escape (value) +" & vbNewLine
            TextEffects = TextEffects & "((expires == null) ? " & """" & """" & " : (" & """" & "; expires=" & """" & " + expires.toGMTString())) +" & vbNewLine
            TextEffects = TextEffects & "((path == null) ? " & """" & """" & " : (" & """" & "; path=" & """" & " + path)) +" & vbNewLine
            TextEffects = TextEffects & "((domain == null) ? " & """" & """" & " : (" & """" & "; domain=" & """" & " + domain)) +" & vbNewLine
            TextEffects = TextEffects & "((secure == true) ? " & """" & "; secure" & """" & " : " & """" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function DeleteCookie (name) {" & vbNewLine
            TextEffects = TextEffects & "var exp = new Date();" & vbNewLine
            TextEffects = TextEffects & "exp.setTime (exp.getTime() - 1);" & vbNewLine
            TextEffects = TextEffects & "var cval = GetCookie (name);" & vbNewLine
            TextEffects = TextEffects & "document.cookie = name + " & """" & "=" & """" & " + cval + " & """" & "; expires=" & """" & " + exp.toGMTString();" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "var expDays = " & Val(Trim$(frmJava.txtInt)) & ";" & vbNewLine
            TextEffects = TextEffects & "var exp = new Date();" & vbNewLine
            TextEffects = TextEffects & "exp.setTime(exp.getTime() + (expDays*24*60*60*1000));" & vbNewLine
            TextEffects = TextEffects & "function amt(){" & vbNewLine
            TextEffects = TextEffects & "var count = GetCookie('count')" & vbNewLine
            TextEffects = TextEffects & "if(count == null) {" & vbNewLine
            TextEffects = TextEffects & "SetCookie('count','1')" & vbNewLine
            TextEffects = TextEffects & "return 1" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "else {" & vbNewLine
            TextEffects = TextEffects & "var newcount = parseInt(count) + 1;" & vbNewLine
            TextEffects = TextEffects & "DeleteCookie('count')" & vbNewLine
            TextEffects = TextEffects & "SetCookie('count',newcount,exp)" & vbNewLine
            TextEffects = TextEffects & "return count" & vbNewLine
            TextEffects = TextEffects & "   }" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function getCookieVal(offset) {" & vbNewLine
            TextEffects = TextEffects & "var endstr = document.cookie.indexOf (" & """" & ";" & """" & ", offset);" & vbNewLine
            TextEffects = TextEffects & "if (endstr == -1)" & vbNewLine
            TextEffects = TextEffects & "endstr = document.cookie.length;" & vbNewLine
            TextEffects = TextEffects & "return unescape(document.cookie.substring(offset, endstr));" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "// End -->" & vbNewLine
            TextEffects = TextEffects & "</SCRIPT>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<body>" & vbNewLine
            TextEffects = TextEffects & "<SCRIPT LANGUAGE=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<!-- Begin" & vbNewLine
            TextEffects = TextEffects & "document.write(" & """" & "You've been here <b>" & """" & " + amt() + " & """" & "</b> times." & """" & ")" & vbNewLine
            TextEffects = TextEffects & "// End -->" & vbNewLine
            TextEffects = TextEffects & "</SCRIPT>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "</body>" & vbNewLine
        Case 15
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<!-- Begin" & vbNewLine
            TextEffects = TextEffects & "function searchFunc() {" & vbNewLine
            TextEffects = TextEffects & "var newWindow = window.open(" & """" & "about:blank" & """" & ", " & """" & "searchValue" & """" & ", " & """" & "width=200, height=200, resizable=no, maximizable=no" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "var searchValue = document.search.queryField.value;" & vbNewLine
            TextEffects = TextEffects & "var yahooSearch = document.search.yahoo.value;" & vbNewLine
            TextEffects = TextEffects & "var altavistaSearch = document.search.altavista.value;" & vbNewLine
            TextEffects = TextEffects & "var webcrawlerSearch = document.search.webcrawler.value;" & vbNewLine
            TextEffects = TextEffects & "var exciteSearch = document.search.excite.value;" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.write(" & """" & "<html>\n<head>\n<title>Select Seach Engine</title>\n</head>" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.write(" & """" & "<body>\n" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.write(" & """" & "<a href='" & """" & " + yahooSearch + searchValue + " & """" & "' target = 'main'>Yahoo!</a><br><br>\n" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.write(" & """" & "<a href='" & """" & " + altavistaSearch + searchValue + " & """" & "' target = 'main'>Altavista</a><br><br>\n" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.write(" & """" & "<a href='" & """" & " + webcrawlerSearch + searchValue + " & """" & "' target = 'main'>WebCrawler</a><br><br>\n" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.write(" & """" & "<a href='" & """" & " + exciteSearch + searchValue + " & """" & "' target = 'main'>Excite</a><br><br>\n" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "newWindow.document.close();" & vbNewLine
            TextEffects = TextEffects & "self.name = " & """" & "main" & """" & ";" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "//  End -->" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<body>" & vbNewLine
            TextEffects = TextEffects & "<form name=search>" & vbNewLine
            TextEffects = TextEffects & "<input type=hidden name=yahoo value=" & """" & "http://search.yahoo.com/search?p=" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<input type=hidden name=altavista value=" & """" & "http://www.altavista.com/cgi-bin/query?pg=q&what=web&fmt=.&q=" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<input type=hidden name=webcrawler value=" & """" & "http://www.webcrawler.com/cgi-bin/WebQuery?searchText=" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<input type=hidden name=excite value=" & """" & "http://www.excite.com/search.gw?trace=a&search=" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "Search the internet for:" & vbNewLine
            TextEffects = TextEffects & "<input type=text size=25 name=queryField>" & vbNewLine
            TextEffects = TextEffects & "<br>" & vbNewLine
            TextEffects = TextEffects & "<input type=button value=" & """" & "Search" & """" & " onClick=" & """" & "searchFunc()" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<input type=reset value=" & """" & "Clear" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<br>" & vbNewLine
            TextEffects = TextEffects & "(A separate window will open where you can select a search engine.)" & vbNewLine
            TextEffects = TextEffects & "</form>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "</body>" & vbNewLine
        Case 16
            TextEffects = "<body>" & vbNewLine
            TextEffects = TextEffects & "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<!-- Begin" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "if ((navigator.appVersion.indexOf(" & """" & "4." & """" & ") != -1) && (navigator.appName.indexOf(" & """" & "Netscape" & """" & ") != -1)){" & vbNewLine
            TextEffects = TextEffects & "ip = " & """" & """" & " + java.net.InetAddress.getLocalHost().getHostAddress();" & vbNewLine
            TextEffects = TextEffects & "document.write(" & """" & "Your IP address is " & """" & " + ip);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "else {" & vbNewLine
            TextEffects = TextEffects & "document.write(" & """" & "IP Address only shown in Netscape with Java enabled!" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "//  End -->" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
        Case 17
            TextEffects = "<body>" & vbNewLine
            TextEffects = TextEffects & "<form method=" & """" & "POST" & """" & " name=" & """" & "log1" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "  <p align=" & """" & "center" & """" & "><font size=" & """" & "3" & """" & " color=" & """" & "#3366CC" & """" & " face=" & """" & "Arial" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "    <SCRIPT LANGUAGE=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<!-- Begin" & vbNewLine
            TextEffects = TextEffects & "var ch;" & vbNewLine
            TextEffects = TextEffects & "function search()" & vbNewLine
            TextEffects = TextEffects & "{" & vbNewLine
            TextEffects = TextEffects & "ch=document.log1.D1.value;" & vbNewLine
            TextEffects = TextEffects & "if(ch==" & """" & "no" & """" & ")" & vbNewLine
            TextEffects = TextEffects & "{" & vbNewLine
            TextEffects = TextEffects & "alert(" & """" & "First select a site to search!" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "document.log1.D1.focus();" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if(ch==" & """" & DetectEntry(frmJava.txtMen(0)) & """" & ")" & vbNewLine
            If Mode = False Then
                TextEffects = TextEffects & "window.location=" & """" & frmJava.txtMen(0) & """" & ";" & vbNewLine
            Else
                TextEffects = TextEffects & "window.location=" & """" & "http://" & frmJava.txtMen(0) & """" & ";" & vbNewLine
            End If
            TextEffects = TextEffects & "if(ch==" & """" & DetectEntry(frmJava.txtMen(1)) & """" & ")" & vbNewLine
            If Mode = False Then
                TextEffects = TextEffects & "window.location=" & """" & frmJava.txtMen(1) & """" & ";" & vbNewLine
            Else
                TextEffects = TextEffects & "window.location=" & """" & "http://" & frmJava.txtMen(1) & """" & ";" & vbNewLine
            End If
            TextEffects = TextEffects & "if(ch==" & """" & DetectEntry(frmJava.txtMen(2)) & """" & ")" & vbNewLine
            If Mode = False Then
                TextEffects = TextEffects & "window.location=" & """" & frmJava.txtMen(2) & """" & ";" & vbNewLine
            Else
                TextEffects = TextEffects & "window.location=" & """" & "http://" & frmJava.txtMen(2) & """" & ";" & vbNewLine
            End If
            TextEffects = TextEffects & "if(ch==" & """" & DetectEntry(frmJava.txtMen(3)) & """" & ")" & vbNewLine
            If Mode = False Then
                TextEffects = TextEffects & "window.location=" & """" & frmJava.txtMen(3) & """" & ";" & vbNewLine
            Else
                TextEffects = TextEffects & "window.location=" & """" & "http://" & frmJava.txtMen(3) & """" & ";" & vbNewLine
            End If
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "var sto = " & """" & "Click here to download the font required by my site!" & """" & vbNewLine
            TextEffects = TextEffects & "var sta = " & """" & "Done" & """" & vbNewLine
            TextEffects = TextEffects & "//  End -->" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "    </font> <font color=" & """" & "#3366CC" & """" & " face=" & """" & "Arial" & """" & "> </font> </span> </b> </i> <font size=" & """" & "3" & """" & " color=" & """" & "#3366CC" & """" & " face=" & """" & "Arial" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "    <b> <i>" & vbNewLine
            TextEffects = TextEffects & "    <select size=" & """" & "1" & """" & " name=" & """" & "D1" & """" & " style=" & """" & "text-transform:uppercase; text-align:center; font-family:Sylfaen" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "      <option selected value=" & """" & "no" & """" & ">Select Site</option>" & vbNewLine
            TextEffects = TextEffects & "      <option value=" & """" & DetectEntry(frmJava.txtMen(0)) & """" & ">" & DetectEntry(frmJava.txtMen(0)) & "</option>" & vbNewLine
            TextEffects = TextEffects & "      <option value=" & """" & DetectEntry(frmJava.txtMen(1)) & """" & ">" & DetectEntry(frmJava.txtMen(1)) & "</option>" & vbNewLine
            TextEffects = TextEffects & "      <option value=" & """" & DetectEntry(frmJava.txtMen(2)) & """" & ">" & DetectEntry(frmJava.txtMen(2)) & "</option>" & vbNewLine
            TextEffects = TextEffects & "      <option value=" & """" & DetectEntry(frmJava.txtMen(3)) & """" & ">" & DetectEntry(frmJava.txtMen(3)) & "</option>" & vbNewLine
            TextEffects = TextEffects & "    </select>" & vbNewLine
            TextEffects = TextEffects & "    </i></b></font><i><span style=" & """" & "font-size: 12pt" & """" & "><font face=" & """" & "Arial" & """" & "> </i>" & vbNewLine
            TextEffects = TextEffects & "    <input type=" & """" & "button" & """" & " value=" & """" & "Go" & """" & " onclick=" & """" & "search()" & """" & " style=" & """" & "cursor:hand" & """" & "></b></font>" & vbNewLine
            TextEffects = TextEffects & "  </p>" & vbNewLine
            TextEffects = TextEffects & "</form>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "</body>" & vbNewLine
        Case 18
            TextEffects = "<script language=" & """" & "JavaScript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "var adlink=" & """" & frmJava.txtMes & """" & ";" & vbNewLine
            TextEffects = TextEffects & "var timeout=60*60*24;" & vbNewLine
            TextEffects = TextEffects & "var showads = 1;" & vbNewLine
            TextEffects = TextEffects & "function adMessage(adcode) {" & vbNewLine
            TextEffects = TextEffects & "if (document.cookie == " & """" & """" & ") {" & vbNewLine
            TextEffects = TextEffects & "document.write(adcode);" & vbNewLine
            TextEffects = TextEffects & "} else {" & vbNewLine
            TextEffects = TextEffects & "var the_cookie = document.cookie;" & vbNewLine
            TextEffects = TextEffects & "the_cookie = unescape(the_cookie);" & vbNewLine
            TextEffects = TextEffects & "the_cookie_split = the_cookie.split(" & """" & ";" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "for (loop=0;loop<the_cookie_split.length;loop++) {" & vbNewLine
            TextEffects = TextEffects & "var part_of_split = the_cookie_split[loop];" & vbNewLine
            TextEffects = TextEffects & "var find_name = part_of_split.indexOf(" & """" & "ad" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "if (find_name!=-1) { break; } }" & vbNewLine
            TextEffects = TextEffects & "if (find_name==-1) {" & vbNewLine
            TextEffects = TextEffects & "document.write(adcode);" & vbNewLine
            TextEffects = TextEffects & "} else { var ad_split = part_of_split.split(" & """" & "=" & """" & ");" & vbNewLine
            TextEffects = TextEffects & "var last = ad_split[1];" & vbNewLine
            TextEffects = TextEffects & "if (last!=0) {" & vbNewLine
            TextEffects = TextEffects & "document.write(adcode);" & vbNewLine
            TextEffects = TextEffects & "} else { showads=0;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function writeCookie(show)" & vbNewLine
            TextEffects = TextEffects & "{" & vbNewLine
            TextEffects = TextEffects & "var today = new Date();" & vbNewLine
            TextEffects = TextEffects & "var the_date = new Date();" & vbNewLine
            TextEffects = TextEffects & "the_date.setTime(today.getTime() + 1000 * timeout);" & vbNewLine
            TextEffects = TextEffects & "var the_cookie_date = the_date.toGMTString();" & vbNewLine
            TextEffects = TextEffects & "var the_cookie = " & """" & "ad=" & """" & "+show;" & vbNewLine
            TextEffects = TextEffects & "var the_cookie = the_cookie + " & """" & ";expires=" & """" & " + the_cookie_date; document.cookie = the_cookie; location.reload(true);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "function handleClick(evnt)" & vbNewLine
            TextEffects = TextEffects & "{" & vbNewLine
            TextEffects = TextEffects & "var targetstring = new String(evnt.target);" & vbNewLine
            TextEffects = TextEffects & "if (targetstring.search(adlink) != -1) {" & vbNewLine
            TextEffects = TextEffects & "writeCookie(0);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "routeEvent(evnt);" & vbNewLine
            TextEffects = TextEffects & "return true;" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "if (window.Event) {" & vbNewLine
            TextEffects = TextEffects & "window.captureEvents(Event.CLICK);" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "window.onClick = handleClick; adMessage('');" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "// End -->" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "</HEAD>" & vbNewLine
            TextEffects = TextEffects & vbNullString & vbNewLine
            TextEffects = TextEffects & "<body>" & vbNewLine
            TextEffects = TextEffects & "<span onClick=" & """" & "writeCookie(0)" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<script language=" & """" & "javascript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "<!-- adMessage('<a href=" & """" & "<http://www.vinze.tk/new/>" & """" & " target=" & """" & "_blank" & """" & ">Example Ad</a>');" & vbNewLine
            TextEffects = TextEffects & "// -->" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "</span>" & vbNewLine
            TextEffects = TextEffects & "<br>" & vbNewLine
            TextEffects = TextEffects & "<script language=" & """" & "javascript" & """" & " TYPE=" & """" & "text/javascript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "if (showads) {" & vbNewLine
            TextEffects = TextEffects & "document.write(" & """" & "Remove the Ads by clicking it." & """" & ")" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "</script>" & vbNewLine
            TextEffects = TextEffects & "<!-- Optional To Show Ads Again -->" & vbNewLine
            TextEffects = TextEffects & "<script language=" & """" & "javascript" & """" & " TYPE=" & """" & "text/javascript" & """" & ">" & vbNewLine
            TextEffects = TextEffects & "if (!showads) {" & vbNewLine
            TextEffects = TextEffects & "document.write(" & """" & "<form><input type=button value='Ads back' onClick=writeCookie(1)><\/form>" & """" & ")" & vbNewLine
            TextEffects = TextEffects & "}" & vbNewLine
            TextEffects = TextEffects & "</script> <p>" & vbNewLine
            TextEffects = TextEffects & "<center>" & vbNewLine
            TextEffects = TextEffects & "</center><p>" & vbNewLine
    End Select
End Function

Public Function ClipboardWatcher() As Boolean
    ClipboardWatcher = Not (Clipboard.GetText) = vbNullString
End Function

Public Function ResetValue(ByVal Reset As Boolean)
    Repeat = Reset
    Wait = Reset
    MoreThan = Reset
    HideUnhide = Reset
    Process = 0
End Function

Public Function Watcher(FormName As Form)
    FormName.txtOutput = vbNullString
    If FormName.cmbOptions.ListIndex < 0 Then Exit Function
    If FormName.cmbOptions.ListIndex = 10 Or FormName.cmbOptions.ListIndex = 11 Then
        If Not FormName.txtMen(0) = vbNullString Or Not FormName.txtMen(1) = vbNullString Or Not _
                FormName.txtMen(2) = vbNullString Or Not FormName.txtMen(3) = vbNullString Then
            TextString = TextEffects(FormName.cmbOptions.ListIndex, FormName.txtMes, Val(Trim$(FormName.txtInt)))
            FormName.txtOutput = TextString
            FormName.cmdCopy.Enabled = True
            FormName.cmdClear.Enabled = True
        End If
        Exit Function
    End If
    If FormName.cmbOptions.ListIndex = 8 Or FormName.cmbOptions.ListIndex = 9 Then
        If Not FormName.txtInt = vbNullString Then
            TextString = TextEffects(FormName.cmbOptions.ListIndex, FormName.txtMes, Val(Trim$(FormName.txtInt)))
            FormName.txtOutput = TextString
            FormName.cmdCopy.Enabled = True
            FormName.cmdClear.Enabled = True
        End If
        Exit Function
    End If
    If FormName.cmbOptions.ListIndex < 6 Then
        If Not FormName.txtMes = vbNullString Then
            TextString = TextEffects(FormName.cmbOptions.ListIndex, FormName.txtMes, Val(Trim$(FormName.txtInt)))
            FormName.txtOutput = TextString
            FormName.cmdCopy.Enabled = True
            FormName.cmdClear.Enabled = True
        End If
    Else
        TextString = TextEffects(FormName.cmbOptions.ListIndex)
        FormName.txtOutput = TextString
        FormName.cmdCopy.Enabled = True
        FormName.cmdClear.Enabled = True
    End If
End Function
