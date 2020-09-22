VERSION 5.00
Begin VB.UserControl Colour_pickerHLS 
   Appearance      =   0  'Flat
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ScaleHeight     =   5415
   ScaleWidth      =   5895
   Begin VB.HScrollBar ScrollH 
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   5
      Top             =   3840
      Width           =   3855
   End
   Begin VB.VScrollBar ScrollS 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   4440
      Max             =   255
      TabIndex        =   4
      Top             =   0
      Value           =   255
      Width           =   255
   End
   Begin VB.VScrollBar ScrollL 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   0
      Max             =   255
      TabIndex        =   3
      Top             =   0
      Value           =   255
      Width           =   255
   End
   Begin VB.PictureBox Pic_picked 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      ScaleHeight     =   225
      ScaleWidth      =   585
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Pic_SEL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   0
      Left            =   240
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3825
   End
   Begin VB.PictureBox Pic_SEL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3825
      Index           =   1
      Left            =   4080
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "Colour_pickerHLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Event ChangeColour()

Private Sub Pic_SEL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If X = -1 Then Exit Sub
If Y = -1 Then Exit Sub

If Index = 0 Then
If Button = 0 Then Exit Sub
For S = 0 To 255 'y
Hls_colour.Hue = Get_HUE(GetPixel(Pic_SEL(0).hdc, X, Y))
Hls_colour.Lum = Get_LUM(GetPixel(Pic_SEL(0).hdc, X, Y))
Hls_colour.Sat = S
ScrollH.Value = X
ScrollL.Value = Y
Pic_SEL(1).Line (0, S)-(Pic_SEL(1).Width, S), HSLtoRGB(Hls_colour)
Next
Pic_SEL(1).Refresh
Pic_picked.BackColor = GetPixel(Pic_SEL(0).hdc, X, Y)

Hls_colour.Hue = Get_HUE(GetPixel(Pic_SEL(0).hdc, X, Y))
Hls_colour.Lum = Get_LUM(GetPixel(Pic_SEL(0).hdc, X, Y))
Hls_colour.Sat = CInt(ScrollS.Value)
Pic_picked.BackColor = HSLtoRGB(Hls_colour)
RaiseEvent ChangeColour

ElseIf Index = 1 Then
If Button = 0 Then Exit Sub
Pic_picked.BackColor = GetPixel(Pic_SEL(1).hdc, X, Y)
RaiseEvent ChangeColour
ScrollS.Value = Y
End If

End Sub

Public Function Get_colour()
Get_colour = Pic_picked.BackColor
End Function


Private Sub UserControl_Initialize()

For H = 0 To 255 'x
For L = 0 To 255 'y
Hls_colour.Hue = H
Hls_colour.Lum = L
Hls_colour.Sat = 255
SetPixel Pic_SEL(0).hdc, H, L, HSLtoRGB(Hls_colour)
Next
Next
Pic_SEL(0).Refresh

For S = 0 To 255 'y
Hls_colour.Hue = 0
Hls_colour.Lum = 128
Hls_colour.Sat = S
Pic_SEL(1).Line (0, S)-(Pic_SEL(1).Width, S), HSLtoRGB(Hls_colour)
Next
Pic_picked.BackColor = vbRed
Pic_SEL(1).Refresh
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Pic_picked.Left + Pic_picked.Width
UserControl.Height = Pic_picked.Top + Pic_picked.Height
End Sub
