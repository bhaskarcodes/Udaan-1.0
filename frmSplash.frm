VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5160
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   2760
      Top             =   5760
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   4890
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7785
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Reg. No-131040110031"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   4200
         Width           =   3405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Roll No-10400113031"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   3600
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   3060
         Left            =   4440
         Picture         =   "frmSplash.frx":000C
         Top             =   1560
         Width           =   3060
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Bhaskar Tejaswi"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Made By-"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   1755
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Udaan"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Flight Reservation System"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   5475
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\133331.jpg")
End Sub

Private Sub Timer2_Timer()
a = a + 1
If a = 200 Then
Form1.Show
Unload Me
End If
End Sub
