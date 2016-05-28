VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form new_user 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Create New User"
   ClientHeight    =   2880
   ClientLeft      =   7830
   ClientTop       =   3720
   ClientWidth     =   4650
   LinkTopic       =   "Form4"
   ScaleHeight     =   2880
   ScaleWidth      =   4650
   Begin VB.CommandButton cancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataSource      =   "adodc2"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataSource      =   "adodc2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton newuser 
      BackColor       =   &H00C0FFFF&
      Caption         =   "New User Registration"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adodc2 
      Height          =   330
      Left            =   600
      Top             =   3240
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=flight.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=flight.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "User_Accounts"
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "PASSWORD"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "USERNAME"
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "new_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"

End Sub

Private Sub newuser_Click()
Dim a As Integer
If (Text1.Text = "") Then
MsgBox "Username field is empty", vbCritical, "Error"
a = 1
Else
If (Text2.Text = "") Then
MsgBox "Password field is empty!", vbCritical, "Error"
Else
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Text1.Text
Adodc2.Recordset.Fields(1) = Text2.Text
Adodc2.Recordset.Update
MsgBox "New user created successfully", vbInformation, "Success !"
End If
End If
userlist.Show
Unload Me
End Sub

Private Sub cancel_Click()
Text1.Text = ""
Text2.Text = ""
adminspace.Show
Unload Me
End Sub

Private Sub new_user_Load()
Adodc2.Recordset.AddNew
End Sub
