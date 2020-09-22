VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   990
   ClientTop       =   -555
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin Project1.Add_HLS Add_HLS1 
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2143
   End
   Begin Project1.Results_RGB_HLS Results_RGB_HLS1 
      Height          =   975
      Left            =   6360
      TabIndex        =   3
      Top             =   4680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1720
   End
   Begin Project1.Colour_mixerRGB Colour_mixerRGB1 
      Height          =   765
      Left            =   600
      TabIndex        =   2
      Top             =   4680
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   1349
   End
   Begin Project1.Colour_pickerHLS Colour_pickerHLS1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7223
   End
   Begin Project1.Colour_pickerHLS Colour_pickerHLS2 
      Height          =   4095
      Left            =   5640
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7223
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_HLS1_Changed()
Me.Results_RGB_HLS1.find_values Me.Add_HLS1.Get_colour
Me.Caption = Me.Results_RGB_HLS1.Get_RGB_HLS(0) & ", " & Me.Results_RGB_HLS1.Get_RGB_HLS(1) & _
", " & Me.Results_RGB_HLS1.Get_RGB_HLS(2) & "; " & Me.Results_RGB_HLS1.Get_RGB_HLS(3) & ", " & _
Me.Results_RGB_HLS1.Get_RGB_HLS(4) & ", " & Me.Results_RGB_HLS1.Get_RGB_HLS(5)

End Sub

Private Sub Colour_mixerRGB1_ColourChanged()
Me.BackColor = Me.Colour_mixerRGB1.Get_colour
Me.Add_HLS1.set_colour Me.Colour_mixerRGB1.Get_colour

End Sub

Private Sub Colour_pickerHLS1_ChangeColour()
Colour_mixerRGB1.set_mix_col 0, Me.Colour_pickerHLS1.Get_colour
End Sub

Private Sub Colour_pickerHLS2_ChangeColour()
Colour_mixerRGB1.set_mix_col 1, Me.Colour_pickerHLS2.Get_colour
End Sub


