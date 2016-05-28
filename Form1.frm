VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Enter Login Details"
   ClientHeight    =   3120
   ClientLeft      =   6510
   ClientTop       =   3525
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6975
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":008E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "User_Accounts"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
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
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   120
         Picture         =   "Form1.frx":011C
         ScaleHeight     =   1395
         ScaleWidth      =   1275
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "User Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Admin Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "PASSWORD         : "
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "USERNAME         : "
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
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

Dim a As Integer
a = 0
Adodc1.Recordset.MoveFirst
    Do Until Adodc1.Recordset.EOF
    
    If (Adodc1.Recordset(0) = Text1.Text And Adodc1.Recordset(1) = Text2.Text) Then
a = 1
        mainpage.Show
        Unload Me
     
    Else
        Adodc1.Recordset.MoveNext
    End If

Loop
If a = 0 Then
MsgBox "No such user found." & vbCrLf & "Please contact the Administrator to be able to access the software.", vbCritical, "User not found"
End If

End Sub

Private Sub Command3_Click()
frmSplash1.Show
End Sub

Private Sub Command4_Click()

If Text1.Text = "username" And Text2.Text = "passw0rd" Then
Unload Me
adminspace.Show
Else
MsgBox "You have entered wrong combination." & vbCrLf & "The program will now exit.", vbCritical, "Login Failed"
Text1.Text = ""
Text2.Text = ""
Unload Me
frmSplash1.Show
End If

End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
Picture1.Picture = LoadPicture(App.Path & "\key.jpg")
End Sub
