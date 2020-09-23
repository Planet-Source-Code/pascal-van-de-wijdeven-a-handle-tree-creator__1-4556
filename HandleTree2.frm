VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Press ESCAPE to go back"
      Height          =   195
      Left            =   4275
      TabIndex        =   0
      Top             =   2640
      Width           =   1905
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Form1.Visible = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If kk.Top = pk.Top And kk.Left = pk.Left Then
        pk.Top = 0
        pk.Left = 0
    End If
    Me.Top = (kk.Top + pk.Top) * Screen.TwipsPerPixelY
    Me.Left = (kk.Left + pk.Left) * Screen.TwipsPerPixelX
    Me.Width = (kk.Right - kk.Left) * Screen.TwipsPerPixelX
    Me.Height = (kk.Bottom - kk.Top) * Screen.TwipsPerPixelY
    Label1.Top = (Me.Height / 2) - (Label1.Height / 2)
    Label1.Left = 0
    Label1.Width = Me.Width
End Sub

