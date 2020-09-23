Attribute VB_Name = "ProgressBar"
Option Explicit
Option Base 1
Option Private Module

Private Declare Function GetPixel Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal x As Long, _
    ByVal y As Long _
    ) As Long

Public Function ProgBar(PicX As PictureBox, _
    PercentIn As Long, _
    Optional BGcolor As Long = vbWhite, _
    Optional FGcolor As Long = vbBlue, _
    Optional TextColor = vbBlack, _
    Optional DisplayText As Boolean = True, _
    Optional Style As Integer = 0) As Boolean
    
    Dim OnePercent As Single
    Dim PBarWidth As Long
    Dim PBarHeight As Long
    Dim T As Long
    Dim i As Long
    Dim Temp As Long
    Dim StrProcess As String
    Static J As Long
    On Error GoTo Err_Handler
    If J > PercentIn Then PicX.Cls
    If J > 0 And J = PercentIn Then Exit Function
    With PicX
        .AutoRedraw = True
        .ScaleMode = 3
        .BackColor = BGcolor
        .ForeColor = FGcolor
        .Font.Bold = True
    End With
    PBarWidth = PicX.Width / Screen.TwipsPerPixelX
    PBarHeight = PicX.Height / Screen.TwipsPerPixelY
    OnePercent = PBarWidth / 100
    Select Case Style
        Case 1
            OnePercent = PBarHeight / 100
            For T = PBarHeight - (OnePercent * PercentIn) To PBarHeight
                PicX.Line (0, T)-(PBarWidth, T)
            Next T
            If DisplayText = True Then
                PicX.CurrentX = PBarWidth / 2 - 10
                PicX.CurrentY = PBarHeight / 2 - 8
                PicX.ForeColor = TextColor
                PicX.Print vbNullString & Format(PercentIn, "##%")
                For T = PBarHeight - (OnePercent * PercentIn) To PBarHeight
                    For i = 0 To PBarWidth
                        If GetPixel(PicX.hDC, i, T) = TextColor _
                            Then PicX.PSet (i, T), BGcolor
                    Next i
                    If T > OnePercent * 60 Then T = PBarHeight
                Next T
            End If
        Case 2
            For T = 0 To OnePercent * (PercentIn - 1)
                PicX.Line (T, 0)-(T, PBarHeight)
            Next T
            For T = 0 To OnePercent * (PercentIn - 1) Step (OnePercent * 7)
                PicX.ForeColor = BGcolor
                PicX.Line (0, 0)-(PBarWidth - 1, 0)
                PicX.Line (1, 1)-(1, PBarHeight - 1)
                PicX.Line (PBarWidth - 1, 0)-(PBarWidth - 1, PBarHeight - 1)
                PicX.Line (1, PBarHeight - 1)-(PBarWidth, PBarHeight - 1)
                PicX.Line (1, PBarHeight - 2)-(PBarWidth, PBarHeight - 2)
                PicX.Line (1, PBarHeight - (1 * 3))-(PBarWidth, PBarHeight - (1 * 3))
                PicX.Line (T - 1, 0)-(T - 1, PBarHeight)
                PicX.Line (T, 0)-(T, PBarHeight)
                PicX.ForeColor = FGcolor
            Next T
        Case 3
            Dim iRed As Integer, iBlue As Integer, iGreen As Integer
            Dim nRed As Integer, nBlue As Integer, nGreen As Integer
            Dim BlueRange As Long, RedRange As Long, GreenRange As Long
            Dim RedPcnt As Single, GreenPcnt As Single, BluePcnt As Single
            Dim Red1 As Long, Green1 As Long, Blue1 As Long
            Dim rTemp As Long, bTemp As Long, gTemp As Long
            Call ColorCodeToRGB(FGcolor, iRed, iGreen, iBlue)
            nRed = iBlue: nBlue = iRed: nGreen = 128
            RedRange = nRed - iRed
            BlueRange = nBlue - iBlue
            GreenRange = nGreen - iGreen
            RedPcnt = RedRange / 100
            GreenPcnt = GreenRange / 100
            BluePcnt = BlueRange / 100
            For T = 0 To OnePercent * (PercentIn - 1)
                Red1 = nRed - RedPcnt * (T / OnePercent + 1)
                If Red1 < 0 Then Red1 = 0
                Green1 = nGreen - GreenPcnt * (T / OnePercent + 1)
                If Green1 < 0 Then Green1 = 0
                Blue1 = nBlue - BluePcnt * (T / OnePercent + 1)
                If Blue1 < 0 Then Blue1 = 0
                PicX.ForeColor = RGB(Red1, Green1, Blue1)
                PicX.Line (T, 0)-(T, PBarHeight)
            Next T
        Case Else
            For T = 0 To OnePercent * (PercentIn - 1)
                PicX.Line (T, 0)-(T, PBarHeight)
            Next T
    End Select
    If DisplayText = True Then
        If Not Style = 1 Then
            PicX.CurrentX = PBarWidth / 2 - 7
            PicX.CurrentY = PBarHeight / 2 - 8
            PicX.ForeColor = TextColor
            If PercentIn <= 9 Then
                StrProcess = "0" & PercentIn
            Else
                StrProcess = PercentIn
            End If
            PicX.Print vbNullString & StrProcess & "%"
            If PercentIn > 40 Then
                For T = OnePercent * 40 To OnePercent * (PercentIn - 1)
                    For i = 0 To PBarHeight
                        If GetPixel(PicX.hDC, T, i) = TextColor Then
                            PicX.PSet (T, i), PicX.BackColor
                        End If
                    Next i
                    If T > OnePercent * 60 Then T = _
                        OnePercent * (PercentIn - 1)
                Next T
            End If
        End If
    End If
    J = PercentIn
    ProgBar = True
    Exit Function
    
Err_Handler:
    ProgBar = False
End Function

Public Function ColorCodeToRGB(lColorCode As Long, _
    iRed As Integer, _
    iGreen As Integer, _
    iBlue As Integer) As Boolean
    Dim lColor As Long
    lColor = lColorCode
    iRed = lColor Mod &H100
    lColor = lColor \ &H100
    iGreen = lColor Mod &H100
    lColor = lColor \ &H100
    iBlue = lColor Mod &H100
    ColorCodeToRGB = True
End Function
