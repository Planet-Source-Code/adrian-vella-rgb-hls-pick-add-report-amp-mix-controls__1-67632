VERSION 5.00
Begin VB.UserControl Add_HLS 
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   ScaleHeight     =   1950
   ScaleWidth      =   7965
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1065
      ScaleWidth      =   465
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   465
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Index           =   2
      LargeChange     =   15
      Left            =   1320
      Max             =   255
      Min             =   -255
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Index           =   1
      LargeChange     =   15
      Left            =   1320
      Max             =   255
      Min             =   -255
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Index           =   0
      LargeChange     =   15
      Left            =   1320
      Max             =   255
      Min             =   -255
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "S"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "L"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   18
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "H"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Lbl_out 
      Caption         =   "128"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   15
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Lbl_out 
      Caption         =   "255"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Lbl_out 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Lbl_in 
      Caption         =   "255"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Lbl_in 
      Caption         =   "128"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Lbl_in 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Out"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Plus"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "In"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lbl_vplus 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lbl_vplus 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Lbl_vplus 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Add_HLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Changed()
Public Sub set_colour(Colour As Long)
Picture1.BackColor = Colour
Lbl_in(0).Caption = Get_HUE(Picture1.BackColor)
Lbl_in(1).Caption = Get_LUM(Picture1.BackColor)
Lbl_in(2).Caption = Get_SAT(Picture1.BackColor)
HScroll_Change 0
End Sub

Public Function Get_colour() As Long
Get_colour = Picture2.BackColor

End Function




Private Sub HScroll_Change(Index As Integer)
'On Error Resume Next
For N = 0 To 2
Lbl_vplus(N).Caption = HScroll(N).Value
Next


Lbl_out(0).Caption = CInt(Lbl_in(0).Caption) + CInt(HScroll(0).Value)
Lbl_out(1).Caption = CInt(Lbl_in(1).Caption) + CInt(HScroll(1).Value)
Lbl_out(2).Caption = CInt(Lbl_in(2).Caption) + CInt(HScroll(2).Value)

If CInt(Lbl_out(0).Caption) < 0 Then Lbl_out(0).Caption = 255 - Abs(CInt(Lbl_out(0).Caption))
If CInt(Lbl_out(0).Caption) > 255 Then Lbl_out(0).Caption = Abs(CInt(Lbl_out(0).Caption)) - 255

If CInt(Lbl_out(1).Caption) < 0 Then Lbl_out(1).Caption = 0
If CInt(Lbl_out(1).Caption) > 255 Then Lbl_out(1).Caption = 255

If CInt(Lbl_out(2).Caption) < 0 Then Lbl_out(2).Caption = 0
If CInt(Lbl_out(2).Caption) > 255 Then Lbl_out(2).Caption = 255

Hls_colour.Hue = CLng(Lbl_out(0).Caption)
Hls_colour.Lum = CLng(Lbl_out(1).Caption)
Hls_colour.Sat = CLng(Lbl_out(2).Caption)
Picture2.BackColor = HSLtoRGB(Hls_colour)
RaiseEvent Changed
End Sub

Private Sub UserControl_Initialize()
Lbl_in(0).Caption = Get_HUE(Picture1.BackColor)
Lbl_in(1).Caption = Get_LUM(Picture1.BackColor)
Lbl_in(2).Caption = Get_SAT(Picture1.BackColor)
HScroll_Change 0
End Sub
Private Sub UserControl_Resize()
UserControl.Width = Picture2.Width + Picture2.Left
UserControl.Height = Picture2.Height + Picture2.Top
End Sub
