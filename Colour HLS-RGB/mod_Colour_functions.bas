Attribute VB_Name = "mod_Colour_functions"
Option Explicit
Const HSLMAX As Long = 255
Const RGBMAX As Long = 255
Const UNDEFINED As Long = 0
Private R As Long
Private G As Long
Private B As Long
Public Hls_colour As HSLCol
Public Type HSLCol
    Hue As Long
    Sat As Long
    Lum As Long
End Type

Private Function iMax(a As Long, B As Long) As Long
    iMax = IIf(a > B, a, B)
End Function
Private Function iMin(a As Long, B As Long) As Long
    iMin = IIf(a < B, a, B)
End Function


Public Function RGBtoHSL(RGBCol As Long) As HSLCol
    Dim R As Long, G As Long, B As Long
    Dim cMax As Long, cMin As Long
    Dim RDelta As Double, GDelta As Double, BDelta As Double
    Dim H As Double, S As Double, L As Double
    Dim cMinus As Long, cPlus As Long
    
    R = Get_RED(RGBCol)
    G = Get_GREEN(RGBCol)
    B = Get_BLUE(RGBCol)
    
    cMax = iMax(iMax(R, G), B)
    cMin = iMin(iMin(R, G), B)
    
    cMinus = cMax - cMin
    cPlus = cMax + cMin
    
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then
        S = 0
        H = UNDEFINED
    Else

        If L <= (HSLMAX / 2) Then
            S = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            S = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If

        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus

        Select Case cMax
            Case CLng(R)
            H = BDelta - GDelta
            Case CLng(G)
            H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
            H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
    
    If H < 0 Then H = H + HSLMAX
End If

RGBtoHSL.Hue = CLng(H)
RGBtoHSL.Lum = CLng(L)
RGBtoHSL.Sat = CLng(S)
End Function


Public Function HSLtoRGB(HueLumSat As HSLCol) As Long
    Dim R As Long, G As Long, B As Long
    Dim H As Long, L As Long, S As Long
    Dim Magic1 As Long, Magic2 As Long
    H = HueLumSat.Hue
    L = HueLumSat.Lum
    S = HueLumSat.Sat


    If S = 0 Then
        R = (L * RGBMAX) / HSLMAX
        G = R
        B = R


    Else
        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + S) + (HSLMAX / 2)) / HSLMAX
        Else
            Magic2 = L + S - ((L * S) + (HSLMAX / 2)) / HSLMAX
        End If
        Magic1 = 2 * L - Magic2
        R = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) * RGBMAX + (HSLMAX / 2)) / HSLMAX
        G = (HuetoRGB(Magic1, Magic2, H) * RGBMAX + (HSLMAX / 2)) / HSLMAX
        B = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) * RGBMAX + (HSLMAX / 2)) / HSLMAX
    End If
    HSLtoRGB = RGB(CInt(R), CInt(G), CInt(B))
End Function


Private Function HuetoRGB(mag1 As Long, mag2 As Long, Hue As Long) As Long
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    Select Case Hue
        Case Is < (HSLMAX / 6)
        HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
        HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
        HuetoRGB = (mag1 + (((mag2 - mag1) * ((HSLMAX * 2 / 3) - Hue) + (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
        HuetoRGB = mag1
    End Select
End Function


Private Sub Get_Colors(COLOR As Long)
    Dim TEMP As Long
    TEMP = (COLOR And 255)
    R = TEMP And 255
    TEMP = Int(COLOR / 256)
    G = TEMP And 255
    TEMP = Int(COLOR / 65536)
    B = TEMP And 255
End Sub

Public Function Get_RED(COLOR As Long)
    Dim TEMP As Long
    TEMP = (COLOR And 255)
    Get_RED = TEMP And 255
End Function
Public Function Get_GREEN(COLOR As Long)
    Dim TEMP As Long
    TEMP = Int(COLOR / 256)
    Get_GREEN = TEMP And 255
End Function
Public Function Get_BLUE(COLOR As Long)
    Dim TEMP As Long
    TEMP = Int(COLOR / 65536)
    Get_BLUE = TEMP And 255
End Function

Public Function Get_HUE(COLOR As Long)
    Dim N As HSLCol
    N = RGBtoHSL(COLOR)
    Get_HUE = N.Hue
End Function

Public Function Get_LUM(COLOR As Long)
    Dim N As HSLCol
    N = RGBtoHSL(COLOR)
    Get_LUM = N.Lum
End Function

Public Function Get_SAT(COLOR As Long)
    Dim N As HSLCol
    N = RGBtoHSL(COLOR)
    Get_SAT = N.Sat
End Function
