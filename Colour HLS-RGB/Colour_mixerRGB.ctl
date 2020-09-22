VERSION 5.00
Begin VB.UserControl Colour_mixerRGB 
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   ScaleHeight     =   1950
   ScaleWidth      =   6060
   Begin VB.PictureBox Picture_res 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   4800
      ScaleHeight     =   735
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox Picture_col 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox Picture_mix 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   840
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3825
   End
   Begin VB.PictureBox Picture_col 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Colour_mixerRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Event ColourChanged()

Public Sub set_mix_col(Index As Integer, RGB_COL As Long)
Picture_col(Index).BackColor = RGB_COL
MIX_EXL
End Sub

Private Sub MIX_EXL()
On Error GoTo hhhkkk
Dim R1 As Double
Dim G1 As Double
Dim B1 As Double
Dim R2 As Double
Dim G2 As Double
Dim B2 As Double
Dim RD As Double
Dim GD As Double
Dim BD As Double
Dim RN As Double
Dim GN As Double
Dim BN As Double
Dim RI As Double
Dim GI As Double
Dim BI As Double
Dim N As Double

R1 = Get_RED(Picture_col(0).BackColor)
G1 = Get_GREEN(Picture_col(0).BackColor)
B1 = Get_BLUE(Picture_col(0).BackColor)

R2 = Get_RED(Picture_col(1).BackColor)
G2 = Get_GREEN(Picture_col(1).BackColor)
B2 = Get_BLUE(Picture_col(1).BackColor)

RD = R2 - R1
GD = G2 - G1
BD = B2 - B1

RI = RD / 255
GI = GD / 255
BI = BD / 255

Label1(0).Caption = RD
Label1(1).Caption = GD
Label1(2).Caption = BD

Label1(3).Caption = RI
Label1(4).Caption = GI
Label1(5).Caption = BI

Picture_mix.DrawWidth = 1
For N = 0 To 255 Step 1
RN = (RI * N) + R1
GN = (GI * N) + G1
BN = (BI * N) + B1

If RN < 0 Then RN = 0
If GN < 0 Then GN = 0
If BN < 0 Then BN = 0


Picture_mix.Line (N, 0)-(N, Picture_mix.Height), RGB(RN, GN, BN)


Next

Picture_res.BackColor = GetPixel(Picture_mix.hdc, 128, 0)
RaiseEvent ColourChanged
Exit Sub
hhhkkk:

End Sub
Public Function Get_colour()
Get_colour = Picture_res.BackColor
End Function

Private Sub Picture_mix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 0 Then Exit Sub
Picture_res.BackColor = GetPixel(Picture_mix.hdc, X, 0)
RaiseEvent ColourChanged
End Sub


Private Sub UserControl_Initialize()
UserControl_Resize
MIX_EXL
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Picture_res.Width + Picture_res.Left
UserControl.Height = Picture_res.Height + Picture_res.Top
End Sub
