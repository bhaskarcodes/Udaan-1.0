VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form flightsearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Search your Flight"
   ClientHeight    =   3870
   ClientLeft      =   4065
   ClientTop       =   2565
   ClientWidth     =   11385
   LinkTopic       =   "Form2"
   ScaleHeight     =   3870
   ScaleWidth      =   11385
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Find"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=flight.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=flight.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Flight_Chart"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   7200
      Top             =   6000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=flight.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=flight.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc6"
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
      Left            =   480
      TabIndex        =   8
      Text            =   "Select Source"
      Top             =   840
      Width           =   3375
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
      Left            =   480
      TabIndex        =   7
      Text            =   "Select Destination"
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11175
      Begin VB.ComboBox Combo5 
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
         Left            =   7440
         TabIndex        =   3
         Text            =   "Select"
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Flights between locations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3975
         Begin VB.CommandButton Command2 
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
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Find"
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
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Height          =   1935
         Left            =   -2640
         TabIndex        =   6
         Top             =   4320
         Width           =   3975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search by Flight Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   7080
         TabIndex        =   9
         Top             =   120
         Width           =   3975
         Begin VB.CommandButton Command5 
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
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1440
            Width           =   1455
         End
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "flight12.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      BackColor       =   8438015
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "Flight_Num"
         Caption         =   "Flight_Num"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Airlines"
         Caption         =   "Airlines"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Source"
         Caption         =   "Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Destination"
         Caption         =   "Destination"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Day"
         Caption         =   "Day"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Time"
         Caption         =   "Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Time_arr"
         Caption         =   "Time_arr"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "fare_eco"
         Caption         =   "fare_eco"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "seats_eco"
         Caption         =   "seats_eco"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "fare_busi"
         Caption         =   "fare_busi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "seats_busi"
         Caption         =   "seats_busi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Seat_eco_now"
         Caption         =   "Seat_eco_now"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Seat_busi_now"
         Caption         =   "Seat_busi_now"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "flightsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
bookticketfront.Show
Unload Me
End Sub

Private Sub Command2_Click()
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Adodc1.RecordSource = "Flight_Chart"
Adodc1.Caption = Adodc1.RecordSource
DataGrid1.Visible = False

End Sub


Private Sub Command3_Click()
Dim a, b As String
a = Combo2.Text
b = Combo1.Text
If Combo2.ListIndex < 0 Then
MsgBox "Enter Source", vbCritical, "Field Vacant"
End If
If Combo1.ListIndex < 0 Then
MsgBox "Enter Destination", vbCritical, "Field Vacant"
End If
If a = b Then
MsgBox "Source and Destination cant be same", vbCritical, "Field Vacant"
End If
If (a <> b) Then
DataGrid1.Visible = True
Adodc1.RecordSource = "select * from Flight_Chart where Source='" & a & "' AND Destination = '" & b & "'"
Adodc1.Refresh
        If Adodc1.Recordset.EOF Then
        s = MsgBox("Sorry, there are no flights available between these two locations .... ", vbCritical, "No Flight Available")
        
        End If
End If

End Sub



Private Sub Command5_Click()
Combo5.ListIndex = -1
DataGrid1.Visible = False

End Sub

Private Sub Command6_Click()
Dim txt As String
txt = Combo5
DataGrid1.Visible = True
Adodc1.RecordSource = "select * from Flight_Chart where Flight_Num='" & Combo5 & "'"
Adodc1.Refresh
End Sub

Private Sub Form_Load()

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

Adodc6.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"

Combo5.Clear
Adodc6.RecordSource = "select distinct Flight_Num from Flight_Chart"
Adodc6.Refresh

With Adodc6.Recordset
Do Until .EOF
Combo5.AddItem ![Flight_Num]
.MoveNext
Loop
End With

Adodc6.RecordSource = "Flight_Chart"
Adodc6.Refresh

End Sub

