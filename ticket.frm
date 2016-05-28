VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ticket 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Your Ticket"
   ClientHeight    =   7275
   ClientLeft      =   5760
   ClientTop       =   1800
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   8955
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   1575
      Left            =   9480
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2778
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   7440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"ticket.frx":0000
      OLEDBString     =   $"ticket.frx":008E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   8895
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Help"
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Back to My Zone"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Line Line16 
         X1              =   6120
         X2              =   6120
         Y1              =   2160
         Y2              =   120
      End
      Begin VB.Line Line15 
         X1              =   6120
         X2              =   6120
         Y1              =   2160
         Y2              =   4560
      End
      Begin VB.Line Line14 
         X1              =   0
         X2              =   4320
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
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
         Left            =   6600
         TabIndex        =   36
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "FARE (In Rs.)"
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
         Left            =   4800
         TabIndex        =   35
         Top             =   3480
         Width           =   975
      End
      Begin VB.Line Line13 
         X1              =   4320
         X2              =   8880
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line12 
         X1              =   4320
         X2              =   8880
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line11 
         X1              =   4320
         X2              =   8880
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line10 
         X1              =   4320
         X2              =   8880
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   4320
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   4320
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   4320
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   4440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   4320
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   4320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Have a Happy and Safe Journey....."
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         Top             =   4680
         Width           =   5175
      End
      Begin VB.Line Line3 
         X1              =   1800
         X2              =   1800
         Y1              =   120
         Y2              =   4560
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10320
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line1 
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   4560
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   6240
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label34 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6360
         TabIndex        =   32
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label33 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   31
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label32 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   30
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label31 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   29
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   28
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label30 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   27
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label29 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   26
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label28 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   25
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label26 
         BackColor       =   &H0080C0FF&
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
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "NAME    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "AGE              "
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
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "GENDER "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "PASSPORT NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "MEAL ORDER"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "PASSENGER ADDRESS     "
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
         Left            =   4680
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "FLIGHT DATE"
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
         TabIndex        =   16
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "FROM"
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
         Left            =   4800
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "TO"
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
         Left            =   4800
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "FLIGHT NO"
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
         TabIndex        =   13
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "CLASS"
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
         TabIndex        =   12
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "PNR NUMBER "
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
         Left            =   4800
         TabIndex        =   11
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
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
         Left            =   6600
         TabIndex        =   10
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label22 
         BackColor       =   &H0080C0FF&
         Caption         =   $"ticket.frx":011C
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
         Left            =   480
         TabIndex        =   6
         Top             =   5760
         Width           =   7335
      End
      Begin VB.Label Label23 
         BackColor       =   &H0080C0FF&
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
         Left            =   6600
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackColor       =   &H0080C0FF&
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
         Left            =   6600
         TabIndex        =   4
         Top             =   2880
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Ticket Details"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label Label27 
      Caption         =   "Label26"
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "ticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "The following Documents are required at the time of verification." & vbCrLf & "1. Passport" & vbCrLf & "2. Income Tax PAN Card " & vbCrLf & "3. Voter’s ID or Driving license", vbInformation, "Documents for Verification"
End Sub

Private Sub Command3_Click()
mainpage.Show
Unload Me
Unload bookticket1
Unload bookticket2
End Sub

Private Sub Form_Load()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
Adodc2.RecordSource = "Select * From Conf_Pass"
Adodc2.Refresh
Adodc2.Recordset.MoveLast
Dim X As Integer
X = Adodc2.Recordset.Fields(0)
Label26.Caption = UCase(bookticket2.Label30.Caption)
Label28.Caption = (bookticket2.Text2.Text)
Label29.Caption = (bookticket2.Combo1)
Label33.Caption = (bookticket2.Text3.Text)
Label30.Caption = (bookticket2.Combo2)
Label25.Caption = UCase(bookticket2.Label25.Caption)
Label31.Caption = UCase(bookticket2.Combo4)
Label32.Caption = UCase(bookticket2.Combo3)
Label34.Caption = UCase(bookticket2.Text4.Text)
Label23.Caption = UCase(bookticket2.Label23.Caption)
Label24.Caption = UCase(bookticket2.Label24.Caption)
Label26.Caption = UCase(bookticket2.Label30.Caption)
Label16.Caption = UCase(bookticket2.Label32.Caption)

Label15.Caption = X

Dim dtmdate As Date
dtmdate = DateValue(Now)
Dim dtmtime As Date
dtmtime = TimeValue(Now)


End Sub

