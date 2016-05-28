VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.PictureBox Picture1 
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2475
         ScaleWidth      =   3915
         TabIndex        =   1
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   2520
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_Load()
a = a + 1
If a = 100 Then
MsgBox "done!!!!"
End If
Picture1.Picture = LoadPicture("C:\Users\Bhaskar\Desktop\splashscreen.jpg")

End Sub

