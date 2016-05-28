VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bookticket1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Plan Your Journey"
   ClientHeight    =   5400
   ClientLeft      =   5190
   ClientTop       =   2760
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   10695
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "FILL IN THE DETAILS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   4080
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   33023
         Appearance      =   0
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   44040193
         CurrentDate     =   42128
         MinDate         =   42064
      End
      Begin VB.CommandButton back 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Back"
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   4
         Text            =   "Start from here"
         Top             =   840
         Width           =   4695
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Text            =   "Go to "
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CommandButton enter 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enter"
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Reset"
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Height          =   615
         Left            =   960
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Height          =   615
         Left            =   960
         TabIndex        =   16
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         TabIndex        =   15
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label22 
         BackColor       =   &H0080C0FF&
         Caption         =   "DENOTES COMPULSORY FIELDS"
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
         Left            =   480
         TabIndex        =   14
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label Label5 
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
         Left            =   2400
         TabIndex        =   13
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label4 
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
         Left            =   2640
         TabIndex        =   12
         Top             =   1440
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
         Left            =   2280
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "SELECT SOURCE                 :"
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
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "SELECT DESTINATION         :"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "DATE OF JOURNEY              :"
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
         Left            =   600
         TabIndex        =   7
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "bookticket1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
bookticketfront.Show
Unload Me
End Sub


Private Sub Command1_Click()
Combo1.ListIndex = -1
Combo2.ListIndex = -1
End Sub

Private Sub enter_Click()

Dim a As Integer
a = 0
Dim today As Date
today = DateValue(Now)

Dim selected As Date
selected = MonthView1.Value

If Combo1.ListIndex < 0 Then
MsgBox "Source is a compulsory  field", vbCritical, "Error"
a = 1
Else
If Combo2.ListIndex < 0 Then
MsgBox "Destination is a compulsory  field", vbCritical, "Error"
a = 1
Else
If Combo1 = Combo2 And a <> 1 Then
MsgBox "Source and Destination can't be same.", vbExclamation, "Error"
Combo1.ListIndex = -1
Combo2.ListIndex = -1
a = 1
Else
If (DateDiff("d", today, selected) > 7) Then
MsgBox "Booking allowed only before a week", vbCritical, "Error"
' booking allowed only before 1 week
a = 1
End If
End If
End If
End If
If a <> 1 Then
bookticket1.Hide
bookticket2.Show
End If
End Sub

Private Sub Form_Load()

MonthView1.MinDate = MonthView1.ShowToday

Combo1.AddItem "Delhi", 0
Combo1.AddItem "Mumbai", 1
Combo1.AddItem "Bangalore", 2
Combo1.AddItem "Chennai", 3
Combo1.AddItem "Kolkata", 4
Combo1.AddItem "Hyderabad", 5
Combo1.AddItem "Kochi", 6
Combo1.AddItem "Ahmedabad", 7
Combo1.AddItem "Dabolim", 8
Combo1.AddItem "Pune", 9
Combo1.AddItem "Thiruvananthapuram", 10
Combo1.AddItem "Kozhikode", 11
Combo1.AddItem "Lucknow", 12
Combo1.AddItem "Srinagar", 13
Combo1.AddItem "Jaipur", 14
Combo1.AddItem "Guwahati", 15
Combo1.AddItem "Bhubaneswar", 16
Combo1.AddItem "Coimbatore", 17
Combo1.AddItem "Nagpur", 18
Combo1.AddItem "Mangalore", 19
Combo1.AddItem "Indore", 20
Combo1.AddItem "Tiruchirappalli", 21
Combo1.AddItem "Patna", 22
Combo1.AddItem "Chandigarh", 23
Combo1.AddItem "Amritsar", 24
Combo1.AddItem "Visakapatnam", 25
Combo1.AddItem "Bagdogra", 26
Combo1.AddItem "Varanasi", 27
Combo1.AddItem "Jammu", 28
Combo1.AddItem "Raipur", 29
Combo1.AddItem "Agartala", 30
Combo1.AddItem "Port Blair", 31
Combo1.AddItem "Vadodara", 32
Combo1.AddItem "Madurai", 33
Combo1.AddItem "Imphal", 34
Combo1.AddItem "Ranchi", 35
Combo1.AddItem "Udaipur", 36
Combo1.AddItem "Aurangabad", 37
Combo1.AddItem "Leh", 38
Combo1.AddItem "Bhopal", 39
Combo1.AddItem "Dehradun", 40
Combo1.AddItem "Rajkot", 41
Combo1.AddItem "Jodhpur", 42
Combo1.AddItem "Dibrugarh", 43
Combo1.AddItem "Tirupati", 44
Combo1.AddItem "Gaya", 45


Combo2.AddItem "Delhi", 0
Combo2.AddItem "Mumbai", 1
Combo2.AddItem "Bangalore", 2
Combo2.AddItem "Chennai", 3
Combo2.AddItem "Kolkata", 4
Combo2.AddItem "Hyderabad", 5
Combo2.AddItem "Kochi", 6
Combo2.AddItem "Ahmedabad", 7
Combo2.AddItem "Dabolim", 8
Combo2.AddItem "Pune", 9
Combo2.AddItem "Thiruvananthapuram", 10
Combo2.AddItem "Kozhikode", 11
Combo2.AddItem "Lucknow", 12
Combo2.AddItem "Srinagar", 13
Combo2.AddItem "Jaipur", 14
Combo2.AddItem "Guwahati", 15
Combo2.AddItem "Bhubaneswar", 16
Combo2.AddItem "Coimbatore", 17
Combo2.AddItem "Nagpur", 18
Combo2.AddItem "Mangalore", 19
Combo2.AddItem "Indore", 20
Combo2.AddItem "Tiruchirappalli", 21
Combo2.AddItem "Patna", 22
Combo2.AddItem "Chandigarh", 23
Combo2.AddItem "Amritsar", 24
Combo2.AddItem "Visakapatnam", 25
Combo2.AddItem "Bagdogra", 26
Combo2.AddItem "Varanasi", 27
Combo2.AddItem "Jammu", 28
Combo2.AddItem "Raipur", 29
Combo2.AddItem "Agartala", 30
Combo2.AddItem "Port Blair", 31
Combo2.AddItem "Vadodara", 32
Combo2.AddItem "Madurai", 33
Combo2.AddItem "Imphal", 34
Combo2.AddItem "Ranchi", 35
Combo2.AddItem "Udaipur", 36
Combo2.AddItem "Aurangabad", 37
Combo2.AddItem "Leh", 38
Combo2.AddItem "Bhopal", 39
Combo2.AddItem "Dehradun", 40
Combo2.AddItem "Rajkot", 41
Combo2.AddItem "Jodhpur", 42
Combo2.AddItem "Dibrugarh", 43
Combo2.AddItem "Tirupati", 44
Combo2.AddItem "Gaya", 45

Dim dtmdate As Date
dtmdate = DateValue(Now)
Dim dtmtime As Date
dtmtime = TimeValue(Now)
Label6.Caption = dtmdate & " " & dtmtime


End Sub



Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim a As String
a = Weekday(Me.MonthView1.SelStart, vbSunday)
Label8.Caption = MonthView1.Value
Label9.Caption = a
End Sub

