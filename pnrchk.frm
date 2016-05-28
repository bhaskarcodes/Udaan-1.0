VERSION 5.00
Begin VB.Form pnrchk 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ticket Details"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   2280
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "<---Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "ENTER PNR NUMBER "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "pnrchk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
mainpage.Show
Unload Me
End Sub
