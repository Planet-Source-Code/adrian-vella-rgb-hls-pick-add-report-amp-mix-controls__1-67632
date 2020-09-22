VERSION 5.00
Begin VB.UserControl Results_RGB_HLS 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ScaleHeight     =   1200
   ScaleWidth      =   2505
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1680
      ScaleHeight     =   945
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label_v 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label_v 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label_v 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label_v 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label_v 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label_v 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Results_RGB_HLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const RsRed = 0
Const RsGreen = 1
Const RsBlue = 2
Const RsHue = 3
Const RsLum = 4
Const RsSat = 5

Public Sub find_values(Colour As Long)
Picture1.BackColor = Colour
Label_v(RsRed).Caption = Get_RED(Colour)
Label_v(RsGreen).Caption = Get_GREEN(Colour)
Label_v(RsBlue).Caption = Get_BLUE(Colour)
Label_v(RsHue).Caption = Get_HUE(Colour)
Label_v(RsLum).Caption = Get_LUM(Colour)
Label_v(RsSat).Caption = Get_SAT(Colour)
End Sub

Private Sub UserControl_Initialize()
find_values vbRed
End Sub
Private Sub UserControl_Resize()
UserControl.Width = Picture1.Width + Picture1.Left
UserControl.Height = Picture1.Height + Picture1.Top
End Sub


Public Function Get_RGB_HLS(Index As Integer)
Get_RGB_HLS = Label_v(Index).Caption
End Function
