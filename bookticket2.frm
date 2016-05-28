VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bookticket2 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Enter Passenger Details"
   ClientHeight    =   7425
   ClientLeft      =   5565
   ClientTop       =   1980
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10275
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "bookticket2.frx":0000
      Height          =   270
      Left            =   9480
      TabIndex        =   48
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   476
      _Version        =   393216
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
         Weight          =   400
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8160
      Top             =   7680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"bookticket2.frx":0015
      OLEDBString     =   $"bookticket2.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
      Height          =   330
      Left            =   7200
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "bookticket2.frx":0131
      Height          =   375
      Left            =   9360
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "PNR_NO"
         Caption         =   "PNR_NO"
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
         DataField       =   "Name"
         Caption         =   "Name"
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
         DataField       =   "Age"
         Caption         =   "Age"
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
         DataField       =   "Gender"
         Caption         =   "Gender"
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
         DataField       =   "Class"
         Caption         =   "Class"
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
         DataField       =   "PP_No"
         Caption         =   "PP_No"
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
         DataField       =   "DOB"
         Caption         =   "DOB"
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
         DataField       =   "DOJ"
         Caption         =   "DOJ"
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
         DataField       =   "Meal_pref"
         Caption         =   "Meal_pref"
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
      BeginProperty Column10 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   3000
      Top             =   7560
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
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
      Caption         =   "Adodc11"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   10095
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "     RESERVATION DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CHECK AVALIABILITY"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5040
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   2280
         TabIndex        =   35
         Text            =   "Select Flight Number"
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "RESET"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5040
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2280
         TabIndex        =   25
         Text            =   "Select Class"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox Text4 
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
         Left            =   6960
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CANCEL"
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "BOOK"
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
         TabIndex        =   12
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "REPLAN"
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
         TabIndex        =   11
         Top             =   5040
         Width           =   1695
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
         Left            =   2280
         TabIndex        =   10
         Text            =   "What should we serve you with ?"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   8
         Top             =   1800
         Width           =   2775
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
         Left            =   2280
         TabIndex        =   6
         Text            =   "Enter Gender"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   3
         Top             =   360
         Width           =   2775
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
         Left            =   6960
         TabIndex        =   47
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label28 
         BackColor       =   &H0080C0FF&
         Caption         =   "FARE(in Rs.)"
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
         Left            =   5520
         TabIndex        =   45
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label29 
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
         Left            =   840
         TabIndex        =   43
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label27 
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
         Left            =   1320
         TabIndex        =   42
         Top             =   3360
         Width           =   375
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
         TabIndex        =   38
         Top             =   2880
         Width           =   2775
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
         Left            =   6960
         TabIndex        =   37
         Top             =   3240
         Width           =   2775
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
         Left            =   6960
         TabIndex        =   36
         Top             =   2760
         Width           =   2775
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
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   6000
         Width           =   3255
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
         TabIndex        =   32
         Top             =   5880
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
         Left            =   6600
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label19 
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
         Left            =   1560
         TabIndex        =   30
         Top             =   1680
         Width           =   375
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
         TabIndex        =   29
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label17 
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
         Left            =   600
         TabIndex        =   28
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label16 
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
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
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
         Height          =   375
         Left            =   6960
         TabIndex        =   26
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "SEATS LEFT"
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
         Left            =   5520
         TabIndex        =   24
         Top             =   4200
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
         TabIndex        =   23
         Top             =   4080
         Width           =   1695
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
         TabIndex        =   22
         Top             =   3480
         Width           =   1695
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
         Left            =   5520
         TabIndex        =   21
         Top             =   3240
         Width           =   975
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
         Left            =   5520
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
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
         TabIndex        =   18
         Top             =   2880
         Width           =   1695
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
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   1695
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
         TabIndex        =   9
         Top             =   2400
         Width           =   1935
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
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
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
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
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
         TabIndex        =   2
         Top             =   840
         Width           =   1695
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
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3000
      Top             =   7920
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
      RecordSource    =   "Conf_Pass"
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
   Begin VB.Label Label31 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label23"
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
      Left            =   6960
      TabIndex        =   46
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label30"
      Height          =   255
      Left            =   360
      TabIndex        =   44
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label26"
      Height          =   375
      Left            =   360
      TabIndex        =   40
      Top             =   7440
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "bookticket2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo3_Change()
If Combo4.ListIndex = -1 Then
MsgBox "select flight number first", vbCritical, "Error"
End If
End Sub

Private Sub Combo3_Click()
Label15.Caption = ""
End Sub

Private Sub Command1_Click()
Unload Me
bookticket1.Show
End Sub

Private Sub Command2_Click()

'Adodc1.RecordSource = "select * from  Flight_Chart where Flight_Num = '" & Combo4 & "' "
Dim ak As Integer
ak = MsgBox("Please note :" & vbCrLf & "You will not be able to make any changes in the ticket after booking. Are you sure you want to proceed ?", vbQuestion + vbOKCancel, "Confirm Booking")
If ak = 1 Then

Dim a As Integer
a = 0

If (Text1.Text = "") Then
MsgBox "Name is a compulsory  field", vbCritical, "Error"
Else
If (Text2.Text = "") Then
MsgBox "Age is a compulsory  field", vbCritical, "Error"
Else
If (IsNumeric(Text2.Text) = False) Then
MsgBox "Age is a numeric  field", vbCritical, "Error"
Else
If (Combo1.ListIndex < 0) Then
MsgBox "Gender is a compulsory  field", vbCritical, "Error"
Else
If (Text3.Text = "") Then
MsgBox "Passport No is a compulsory  field", vbCritical, "Error"
Else
If (Combo4.ListIndex < 0) Then
MsgBox "Flight No is a compulsory  field", vbCritical, "Error"
Else
If (Combo3.ListIndex < 0) Then
MsgBox "Class is a compulsory  field", vbCritical, "Error"
Else
If (Text4.Text = "") Then
MsgBox "Adress is a compulsory  field", vbCritical, "Error"
Else
                 Label30.Caption = Text1.Text
                 If Label15.Caption = "0" Then
                 MsgBox "Sorry,We are full. Try another flight", vbCritical, "No Seat Available"

            Else

            'this is where the booking procedure starts
            Dim dtmdate As Date
            dtmdate = DateValue(Now)
            Adodc2.Recordset.AddNew
            Adodc2.Recordset.Fields(1) = Text1.Text
            Adodc2.Recordset.Fields(2) = Text2.Text
            Adodc2.Recordset.Fields(3) = Combo1
            Adodc2.Recordset.Fields(4) = Combo3
            Adodc2.Recordset.Fields(5) = Text3.Text
            Adodc2.Recordset.Fields(6) = dtmdate
            Adodc2.Recordset.Fields(7) = Label25.Caption
            Dim fn As String
            fn = Combo4.Text
            Adodc2.Recordset.Fields(8) = Combo2
            Adodc2.Recordset.Fields(9) = Label26.Caption
            Adodc2.Recordset.Fields(10) = Combo4
            Adodc2.Recordset.Update
            MsgBox "Ticket Booked Successfully", vbInformation, "Success !"
            
            
            If (Combo3 = "Economy") Then
            Adodc11.RecordSource = "select Seat_eco_now from Flight_Chart where Flight_Num = '" & fn & "' "
            Adodc11.Refresh
            Adodc11.Recordset.Fields("Seat_eco_now") = Val(Label15.Caption) - 1
                                    Adodc11.Recordset.Update
            
                                    
            End If
  '          MsgBox (Adodc1.Recordset.Fields(12) - 1)
            If (Combo3 = "Business") Then
            Adodc11.RecordSource = "select Seat_busi_now from Flight_Chart where Flight_Num = '" & fn & "' "
            Adodc11.Refresh
            Adodc11.Recordset.Fields("Seat_busi_now") = Val(Label15.Caption) - 1
                                  Adodc11.Recordset.Update
            
            End If
            
           
                       

            ticket.Show
            bookticket2.Hide
            End If
            
End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Command3_Click()
bookticketfront.Show
Unload Me
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Combo1.ListIndex = -1
Text3.Text = ""
Combo2.ListIndex = -1
Combo4.ListIndex = -1
Combo3.ListIndex = -1
Text4.Text = ""

End Sub

Private Sub Command5_Click()
Dim s As VbMsgBoxResult
Dim fn, cl As String
fn = Combo4.Text
cl = Combo3.Text


If Combo3.Text = "Economy" Then

    If (c1 <> fn) Then
Adodc11.RecordSource = "select Seat_eco_now from Flight_Chart where Flight_Num = '" & fn & "' "
Adodc11.Refresh
adodc3.RecordSource = "select Airlines from Flight_Chart where Flight_Num = '" & fn & "' "
adodc3.Refresh
        
        If Adodc11.Recordset.EOF Then
        s = MsgBox("Sorry, no entry found." & vbCrLf & "Possible Reasons:" & vbCrLf & "1.Either you have not filled all the fields" & vbCrLf & "2.There are no flights available between these two locations .... ", vbCritical, "Something's not right !")
        End If
Adodc11.Refresh
                Label15.Caption = Adodc11.Recordset.Fields("Seat_eco_now")
                                Label26.Caption = adodc3.Recordset.Fields("Airlines")
adodc3.RecordSource = "select fare_eco from Flight_Chart where Flight_Num = '" & fn & "' "
adodc3.Refresh
                                Label32.Caption = adodc3.Recordset.Fields("fare_eco")
                
            
    End If
End If

If Combo3 = "Business" Then
  If (fn <> c1) Then

Adodc11.RecordSource = "select Seat_busi_now from Flight_Chart where Flight_Num = '" & fn & "' "
Adodc11.Refresh
  adodc3.RecordSource = "select Airlines from Flight_Chart where Flight_Num = '" & fn & "' "
adodc3.Refresh
        
        If Adodc11.Recordset.EOF Then
        s = MsgBox("Sorry, no entry found." & vbCrLf & "Possible Reasons:" & vbCrLf & "1.Either you have not filled all the fields" & vbCrLf & "2.There are no flights available between these two locations .... ", vbCritical, "Something's not right !")
     
        End If
        Label15.Caption = Adodc11.Recordset.Fields("Seat_busi_now")
        Label26.Caption = adodc3.Recordset.Fields("Airlines")
adodc3.RecordSource = "select fare_busi from Flight_Chart where Flight_Num = '" & fn & "' "
adodc3.Refresh
                                Label32.Caption = adodc3.Recordset.Fields("fare_busi")
                
    End If
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem "Male", 0
Combo1.AddItem "Female", 1

Combo2.AddItem "Continental", 0
Combo2.AddItem "Indian Vegetarian", 1
Combo2.AddItem "French", 2
Combo2.AddItem "Italian", 3
Combo2.AddItem "Sea Food", 4
Combo2.AddItem "Fruit Platter", 5
Combo2.AddItem "Chinese", 6

Combo3.AddItem "Economy", 0
Combo3.AddItem "Business", 1


Label25.Caption = bookticket1.Label8.Caption  ' date assignment

Label23.Caption = UCase(bookticket1.Combo1)
Label24.Caption = UCase(bookticket1.Combo2)

Dim dtmdate As Date
dtmdate = DateValue(Now)
Dim dtmtime As Date
dtmtime = TimeValue(Now)
Label9.Caption = "                                      " & dtmdate & " " & dtmtime

Adodc11.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\flight.mdb;"
Combo4.Clear
Dim a, k, b As String
a = bookticket1.Combo1
b = bookticket1.Combo2

k = bookticket1.Label9.Caption
If k = "0" Then
k = "SATURDAY"
End If

If k = "1" Then
k = "SUNDAY"
End If
If k = "2" Then
k = "MONDAY"
End If
If k = "3" Then
k = "TUESDAY"
End If
If k = "4" Then
k = "WEDNUSDAY"
End If
If k = "5" Then
k = "THURSDAY"
End If
If k = "6" Then
k = "FRIDAY"
End If

Adodc11.RecordSource = "select Flight_Num from Flight_Chart where Source='" & a & "' AND Destination = '" & b & "' AND Day = '" & k & "'"
Adodc11.Refresh

With Adodc11.Recordset
Do Until .EOF
Combo4.AddItem ![Flight_Num]
.MoveNext
Loop
End With

Adodc11.RecordSource = "Flight_Chart"
Adodc11.Refresh

End Sub

