VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   5490
   ClientTop       =   2160
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   10560
      Top             =   5400
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmSplash1.frx":000C
      Top             =   0
      Width           =   9300
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\rainermariarilke147758.jpg")
End Sub

Private Sub Timer2_Timer()
a = a + 1
If a = 150 Then
End
End If
End Sub

