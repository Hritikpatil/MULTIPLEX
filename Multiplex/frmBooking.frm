VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBooking 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Booking"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   10680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   3720
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   3360
         Top             =   120
      End
      Begin VB.Label lblHouseFull 
         Alignment       =   2  'Center
         Caption         =   "HOUSE FULL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1560
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   2175
      Left            =   8280
      TabIndex        =   44
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   840
         TabIndex        =   62
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   1680
         X2              =   1680
         Y1              =   1560
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   240
         Left            =   1800
         TabIndex        =   55
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblServiceTax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   54
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Tax (4%)"
         Height          =   240
         Left            =   720
         TabIndex        =   53
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label lblETax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   52
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Entertainment Tax (10%)"
         Height          =   480
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   2280
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   240
         Left            =   1560
         TabIndex        =   49
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lblRate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   240
         Left            =   1800
         TabIndex        =   47
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Height          =   2175
      Left            =   240
      TabIndex        =   37
      Top             =   0
      Width           =   3975
      Begin VB.ComboBox cmbShowTime 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbClass 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         ItemData        =   "frmBooking.frx":0000
         Left            =   1560
         List            =   "frmBooking.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbTheaterNo 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         ItemData        =   "frmBooking.frx":002C
         Left            =   1560
         List            =   "frmBooking.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSeatsAvail 
         BackStyle       =   0  'Transparent
         Caption         =   "672"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   58
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seats Available"
         Height          =   240
         Left            =   1920
         TabIndex        =   57
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label lblSeats 
         BackStyle       =   0  'Transparent
         Caption         =   "672"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   46
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Seats"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Time"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Select Class"
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Theater No."
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame frmClass 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Royal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   6120
      Width           =   11175
      Begin VB.CheckBox chkL 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkM 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox chkN 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkO 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkP 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "L"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "M"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "N"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "O"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "P"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.Frame frmClass 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Executive"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   11175
      Begin VB.CheckBox chkK 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkF 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox chkG 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkH 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkI 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkJ 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "K"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "F"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "G"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "H"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "I"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "J"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   255
      End
   End
   Begin VB.Frame frmClass 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Premium"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   11175
      Begin VB.CheckBox chkE 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkD 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkC 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkB 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkA 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "E"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "D"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "C"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "B"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "A"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   1575
      Left            =   4320
      ScaleHeight     =   1515
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin VB.Image Image1 
         Height          =   15
         Left            =   0
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label lblMovieName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SCREEN"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         Left            =   840
         TabIndex        =   14
         Top             =   0
         Width           =   2205
      End
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4320
      Picture         =   "frmBooking.frx":0046
      Top             =   1920
      Width           =   3750
   End
   Begin VB.Label lblToday 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      Height          =   255
      Left            =   4560
      TabIndex        =   63
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblBind 
      Caption         =   "Bind"
      Height          =   255
      Left            =   3960
      TabIndex        =   60
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim SQL As String
Dim Srch As String
Dim Booking As Integer

Private Sub chkA_Click(Index As Integer)
    Srch = Combine("A", Index)
    If chkA(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
    
End Sub

Private Sub chkB_Click(Index As Integer)
    Srch = Combine("B", Index)
    If chkB(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkC_Click(Index As Integer)
    Srch = Combine("C", Index)
    If chkC(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkD_Click(Index As Integer)
    Srch = Combine("D", Index)
    If chkD(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkE_Click(Index As Integer)
    Srch = Combine("E", Index)
    If chkE(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkF_Click(Index As Integer)
    Srch = Combine("F", Index)
    If chkF(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkG_Click(Index As Integer)
    Srch = Combine("G", Index)
    If chkG(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkH_Click(Index As Integer)
    Srch = Combine("H", Index)
    If chkH(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkI_Click(Index As Integer)
    Srch = Combine("I", Index)
    If chkI(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkJ_Click(Index As Integer)
    Srch = Combine("J", Index)
    If chkJ(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkK_Click(Index As Integer)
    Srch = Combine("K", Index)
    If chkK(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkL_Click(Index As Integer)
    Srch = Combine("L", Index)
    If chkL(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkM_Click(Index As Integer)
    Srch = Combine("M", Index)
    If chkM(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkN_Click(Index As Integer)
    Srch = Combine("N", Index)
    If chkN(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkO_Click(Index As Integer)
    Srch = Combine("O", Index)
    If chkO(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub chkP_Click(Index As Integer)
    Srch = Combine("P", Index)
    If chkP(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Call CalAmt
End Sub

Private Sub cmbClass_Click()
    
    Select Case cmbClass.ListIndex
        Case 0:
            frmClass(0).Enabled = True
            frmClass(1).Enabled = False
            frmClass(2).Enabled = False
            Call ClearExecutive
            Call ClearRoyal
        Case 1:
            frmClass(1).Enabled = True
            frmClass(0).Enabled = False
            frmClass(2).Enabled = False
            Call ClearPremium
            Call ClearRoyal
        Case 2:
            frmClass(2).Enabled = True
            frmClass(1).Enabled = False
            frmClass(0).Enabled = False
            Call ClearExecutive
            Call ClearPremium
    End Select
    
    SQL = "Select * from Class"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    
    Seats = 672
    lblSeatsAvail.Caption = 672
    lblAmount.Caption = ""
    lblServiceTax.Caption = ""
    lblETax.Caption = ""
    lblTotal.Caption = ""
    
    With Adodc1.Recordset
        .MoveFirst
        While .EOF <> True
            If .Fields(0) = cmbClass.Text Then
                lblRate.Caption = Adodc1.Recordset.Fields(1)
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
    End With
    
End Sub

Private Sub cmbShowTime_Click()
    Call ClearAllSeats
    SQL = "Select * from Booking"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    With Adodc1.Recordset
        .MoveFirst
        While .EOF <> True
            If .Fields(1) = cmbTheaterNo.ListIndex + 1 Then
                If Format(.Fields(2), "DD/MM/YYYY") = Format(Date, "DD/MM/YYYY") Then
                    If .Fields(3) = cmbShowTime.Text Then
                        Call MarkReserved(.Fields(0))
                    End If
                End If
                .MoveNext
            Else
                .MoveNext
            End If
        Wend
    End With
    Booking = 0
End Sub

Private Sub cmbTheaterNo_Click()

    Dim i As Integer
    
    SQL = "Select * from Movie where TheaterNo=" & cmbTheaterNo.ListIndex + 1 & " order by inDate Asc"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    Adodc1.Recordset.MoveLast
    lblMovieName.Caption = Adodc1.Recordset.Fields(2)
    
    cmbShowTime.Clear
    SQL = "Select * from Theater where TheaterNo=" & cmbTheaterNo.ListIndex + 1
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    
    For i = 2 To 6
        If Adodc1.Recordset.Fields(i) <> "12:00:00 AM" Then cmbShowTime.AddItem Adodc1.Recordset.Fields(i)
    Next i
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    Dim i As Integer
    Dim s As String
    
    SQL = "Select * from Booking"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    
    For i = 0 To X - 1
        s = s & SeatNos(i)
    Next
    
    With Adodc1.Recordset
        .AddNew
        .Fields(0) = .Fields(0) & s
        .Fields(1) = cmbTheaterNo.ListIndex + 1
        .Fields(2) = Format(Date, "DD/MM/YYYY")
        .Fields(3) = cmbShowTime.Text
        .Fields(4) = "C" 'Current Booking
        .Fields(5) = lblAmount.Caption
        .Fields(6) = lblETax.Caption
        .Fields(7) = lblServiceTax.Caption
        .Fields(8) = lblTotal.Caption
        .Save
    End With
    
    SQL = "Select * from Movie"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    With Adodc1.Recordset
        .MoveFirst
        While .EOF <> True
            If .Fields(2) = lblMovieName.Caption Then
                .Fields(4) = Val(lblTotal.Caption)
                .Update
                .MoveNext
            Else
                .MoveNext
            End If
        Wend
    End With
    
    SQL = "Select * from Ticket"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
        
    For i = X - Booking To X
        With Adodc1.Recordset
            .AddNew
            .Fields(0) = SeatNos(i)
            .Fields(1) = lblMovieName.Caption
            .Fields(2) = Format(Date, "DD/MM/YYYY")
            .Fields(3) = cmbShowTime.Text
            .Fields(4) = lblRate.Caption
            .Fields(5) = Val(lblRate.Caption) * 10 / 100
            .Fields(6) = Val(lblRate.Caption) * 4 / 100
            .Fields(7) = Val(lblRate.Caption) + Val(lblRate.Caption) * 4 / 100 + Val(lblRate.Caption) * 10 / 100
            .Save
        End With
    Next i
    
    'Call Datareport4 to print tickets
    If isLoad = True Then
        Unload DataReport4
        Unload DataEnvironment1
    Else
        isLoad = True
    End If

    DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    DataEnvironment1.cmdTicket
    Load DataReport4
    DataReport4.Show
    
    SQL = "Select * from Ticket"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    
    'Delete all records from the table
    Adodc1.Recordset.MoveFirst
    While Adodc1.Recordset.EOF <> True
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Wend
    
    
    SQL = "Select * from Collection"
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    
    With Adodc1.Recordset
        .MoveFirst
        While .EOF <> True
            If Format(.Fields(0), "DD/MM/YYYY") = Format(Date, "DD/MM/YYYY") Then
                .Fields(cmbTheaterNo.ListIndex + 1) = .Fields(cmbTheaterNo.ListIndex + 1) + Val(lblTotal.Caption)
                .Fields(4) = Val(.Fields(1)) + Val(.Fields(2)) + Val(.Fields(3))
                .Update
                .MoveNext
                Unload Me
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
        
        'If not found then add new record
        .AddNew
        .Fields(0) = Format(Date, "DD/MM/YYYY")
        .Fields(cmbTheaterNo.ListIndex + 1) = .Fields(cmbTheaterNo.ListIndex + 1) + Val(lblTotal.Caption)
        .Fields(4) = Val(.Fields(1)) + Val(.Fields(2)) + Val(.Fields(3))
        .Save
    End With
    
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    lblToday.Caption = Format(Date, "dddd, dd-MMMM-YYYY")
    Call DispSeats
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "Select * from Movie"
    Set lblBind.DataSource = Adodc1
    Seats = 672
End Sub

Public Sub DispSeats()
    Dim i As Integer
    For i = 1 To 41
        Load chkA(i)
        chkA(i).Left = chkA(i - 1).Left + chkA(i - 1).Width
        chkA(i).Visible = True
        
        Load chkB(i)
        chkB(i).Left = chkB(i - 1).Left + chkB(i - 1).Width
        chkB(i).Visible = True

        Load chkC(i)
        chkC(i).Left = chkC(i - 1).Left + chkC(i - 1).Width
        chkC(i).Visible = True
        
        Load chkD(i)
        chkD(i).Left = chkD(i - 1).Left + chkD(i - 1).Width
        chkD(i).Visible = True
        
        Load chkE(i)
        chkE(i).Left = chkE(i - 1).Left + chkE(i - 1).Width
        chkE(i).Visible = True
        
        Load chkF(i)
        chkF(i).Left = chkF(i - 1).Left + chkF(i - 1).Width
        chkF(i).Visible = True
        
        Load chkG(i)
        chkG(i).Left = chkG(i - 1).Left + chkG(i - 1).Width
        chkG(i).Visible = True
        
        Load chkH(i)
        chkH(i).Left = chkH(i - 1).Left + chkH(i - 1).Width
        chkH(i).Visible = True
        
        Load chkI(i)
        chkI(i).Left = chkI(i - 1).Left + chkI(i - 1).Width
        chkI(i).Visible = True
        
        Load chkJ(i)
        chkJ(i).Left = chkJ(i - 1).Left + chkJ(i - 1).Width
        chkJ(i).Visible = True
        
        Load chkK(i)
        chkK(i).Left = chkK(i - 1).Left + chkK(i - 1).Width
        chkK(i).Visible = True
        
        Load chkL(i)
        chkL(i).Left = chkL(i - 1).Left + chkL(i - 1).Width
        chkL(i).Visible = True
        
        Load chkM(i)
        chkM(i).Left = chkM(i - 1).Left + chkM(i - 1).Width
        chkM(i).Visible = True
        
        Load chkN(i)
        chkN(i).Left = chkN(i - 1).Left + chkN(i - 1).Width
        chkN(i).Visible = True
        
        Load chkO(i)
        chkO(i).Left = chkO(i - 1).Left + chkO(i - 1).Width
        chkO(i).Visible = True
        
        Load chkP(i)
        chkP(i).Left = chkP(i - 1).Left + chkP(i - 1).Width
        chkP(i).Visible = True
    Next i
End Sub

Public Sub CalAmt()
    Booking = Booking + 1
    lblSeatsAvail.Caption = Seats
    lblAmount.Caption = (672 - Seats) * Val(lblRate.Caption)
    lblETax.Caption = Val(lblAmount.Caption) * 10 / 100
    lblServiceTax.Caption = Val(lblAmount.Caption) * 4 / 100
    lblTotal.Caption = Val(lblAmount.Caption) + Val(lblETax.Caption) + Val(lblServiceTax.Caption)
End Sub

Public Sub ClearPremium()
    Dim i As Integer
    For i = 0 To 41
        chkA(i).Value = 0
        chkB(i).Value = 0
        chkC(i).Value = 0
        chkD(i).Value = 0
        chkE(i).Value = 0
    Next i
End Sub

Public Sub ClearExecutive()
    Dim i As Integer
    For i = 0 To 41
        chkF(i).Value = 0
        chkG(i).Value = 0
        chkH(i).Value = 0
        chkI(i).Value = 0
        chkJ(i).Value = 0
        chkK(i).Value = 0
    Next i
End Sub

Public Sub ClearRoyal()
    Dim i As Integer
    For i = 0 To 41
        chkL(i).Value = 0
        chkM(i).Value = 0
        chkN(i).Value = 0
        chkO(i).Value = 0
        chkP(i).Value = 0
    Next i
End Sub

Public Sub removeItem(src As String)
    Dim i As Integer
    Dim j As Integer
    For i = 0 To X - 1
        If SeatNos(i) = src Then
            For j = i + 1 To X - 1
                SeatNos(j - 1) = SeatNos(j)
            Next j
            SeatNos(j) = ""
            X = X - 1
        End If
    Next i
End Sub

Public Function Combine(s As String, Index As Integer) As String
    Dim comb As String
    Dim ind As String
    If Index < 10 Then
        ind = "00" & Index
    ElseIf Index < 100 Then
        ind = "0" & Index
    Else
        ind = Index
    End If
    comb = s & ind
    Combine = comb
End Function

Public Sub MarkReserved(SNo As String)
    Dim L As Integer
    Dim Series As String
    Dim Index As Integer
    Dim i As Integer
    i = 1
    L = Len(SNo)
    While i <= L
        Series = Mid(SNo, i, 4)
        i = i + 4
        lblSeatsAvail.Caption = Seats
        Index = Val(Mid(Series, 2, 3))
        Series = Mid(Series, 1, 1)
        Select Case Series
            Case "A":
                chkA(Index).Enabled = False
                chkA(Index).Value = 1
            Case "B":
                chkB(Index).Enabled = False
                chkB(Index).Value = 1
            Case "C":
                chkC(Index).Enabled = False
                chkC(Index).Value = 1
            Case "D":
                chkD(Index).Enabled = False
                chkD(Index).Value = 1
            Case "E":
                chkE(Index).Enabled = False
                chkE(Index).Value = 1
            Case "F":
                chkF(Index).Enabled = False
                chkF(Index).Value = 1
            Case "G":
                chkG(Index).Enabled = False
                chkG(Index).Value = 1
            Case "H":
                chkH(Index).Enabled = False
                chkH(Index).Value = 1
            Case "I":
                chkI(Index).Enabled = False
                chkI(Index).Value = 1
            Case "J":
                chkJ(Index).Enabled = False
                chkJ(Index).Value = 1
            Case "K":
                chkK(Index).Enabled = False
                chkK(Index).Value = 1
            Case "L":
                chkL(Index).Enabled = False
                chkL(Index).Value = 1
            Case "M":
                chkM(Index).Enabled = False
                chkM(Index).Value = 1
            Case "N":
                chkN(Index).Enabled = False
                chkN(Index).Value = 1
            Case "O":
                chkO(Index).Enabled = False
                chkO(Index).Value = 1
            Case "P":
                chkP(Index).Enabled = False
                chkP(Index).Value = 1
        End Select
        Booking = Booking + 1
    Wend
    
    lblAmount.Caption = ""
    lblTotal.Caption = ""
    lblServiceTax.Caption = ""
    lblETax.Caption = ""
End Sub

Public Sub ClearAllSeats()
    Dim i As Integer
    For i = 0 To 41
        chkA(i).Value = 0
        chkB(i).Value = 0
        chkC(i).Value = 0
        chkD(i).Value = 0
        chkE(i).Value = 0
        chkF(i).Value = 0
        chkG(i).Value = 0
        chkH(i).Value = 0
        chkI(i).Value = 0
        chkJ(i).Value = 0
        chkK(i).Value = 0
        chkL(i).Value = 0
        chkM(i).Value = 0
        chkN(i).Value = 0
        chkO(i).Value = 0
        chkP(i).Value = 0
        chkA(i).Enabled = True
        chkB(i).Enabled = True
        chkC(i).Enabled = True
        chkD(i).Enabled = True
        chkE(i).Enabled = True
        chkF(i).Enabled = True
        chkG(i).Enabled = True
        chkH(i).Enabled = True
        chkI(i).Enabled = True
        chkJ(i).Enabled = True
        chkK(i).Enabled = True
        chkL(i).Enabled = True
        chkM(i).Enabled = True
        chkN(i).Enabled = True
        chkO(i).Enabled = True
        chkP(i).Enabled = True
    Next i
End Sub

Private Sub lblSeatsAvail_Change()
    If Val(lblSeatsAvail.Caption) = 0 Then
        Picture2.Visible = True
    Else
        Picture2.Visible = False
    End If
End Sub

Private Sub Timer1_Timer()
    lblHouseFull.Visible = Not lblHouseFull.Visible
End Sub
