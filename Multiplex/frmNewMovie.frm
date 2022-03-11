VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewMovie 
   BackColor       =   &H00C0E0FF&
   Caption         =   "New Movie"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   9780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   7095
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   10935
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   240
         Top             =   6600
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Navigation"
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
         Height          =   1215
         Left            =   1680
         TabIndex        =   26
         Top             =   5520
         Width           =   8055
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            Height          =   495
            Left            =   4320
            TabIndex        =   28
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   495
            Left            =   1080
            TabIndex        =   27
            Top             =   480
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Timings"
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
         Height          =   2895
         Left            =   5880
         TabIndex        =   9
         Top             =   2160
         Width           =   3855
         Begin VB.ComboBox cmbAMPM 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            Index           =   4
            ItemData        =   "frmNewMovie.frx":0000
            Left            =   2760
            List            =   "frmNewMovie.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbAMPM 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            Index           =   3
            ItemData        =   "frmNewMovie.frx":001A
            Left            =   2760
            List            =   "frmNewMovie.frx":0024
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbAMPM 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            Index           =   2
            ItemData        =   "frmNewMovie.frx":0034
            Left            =   2760
            List            =   "frmNewMovie.frx":003E
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbAMPM 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            Index           =   1
            ItemData        =   "frmNewMovie.frx":004E
            Left            =   2760
            List            =   "frmNewMovie.frx":0058
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbAMPM 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            Index           =   0
            ItemData        =   "frmNewMovie.frx":0068
            Left            =   2760
            List            =   "frmNewMovie.frx":0072
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   855
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   11
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   14
            Top             =   960
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   17
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   20
            Top             =   1920
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1680
            TabIndex        =   23
            Top             =   2400
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show 05"
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
            Height          =   240
            Index           =   4
            Left            =   360
            TabIndex        =   22
            Top             =   2400
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show 04"
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
            Height          =   240
            Index           =   3
            Left            =   360
            TabIndex        =   19
            Top             =   1920
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show 03"
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
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   16
            Top             =   1440
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show 02"
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
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Top             =   960
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show 01"
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
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Theater and Shows"
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
         Height          =   2895
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   4215
         Begin VB.ComboBox cmbShows 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            ItemData        =   "frmNewMovie.frx":0082
            Left            =   1800
            List            =   "frmNewMovie.frx":0095
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1440
            Width           =   2295
         End
         Begin VB.ComboBox cmbTheater 
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            ItemData        =   "frmNewMovie.frx":00AD
            Left            =   1800
            List            =   "frmNewMovie.frx":00BA
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Shows"
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
            Height          =   240
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Theater"
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
            Height          =   240
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   10215
         Begin VB.TextBox txtMovieName 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   1920
            TabIndex        =   3
            Top             =   240
            Width           =   7935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of the Movie"
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
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1680
         End
      End
      Begin VB.Label lblBind 
         Caption         =   "Label4"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   6360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Today's Date"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd dd MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   9600
         TabIndex        =   25
         Top             =   240
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmNewMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbShows_Click()
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To cmbShows.Text - 1
        lblShow(i).Visible = True
        MaskEdBox1(i).Visible = True
        cmbAMPM(i).Visible = True
    Next
    
    For j = i To 4
        lblShow(j).Visible = False
        MaskEdBox1(j).Visible = False
        cmbAMPM(j).Visible = False
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim showTime As String
    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    
    With Adodc1.Recordset
        .MoveLast
        While .BOF <> True
            If .Fields(3) = cmbTheater.ListIndex + 1 Then
                .Fields(1) = Format(Date, "DD/MM/YYYY")
                .Update
                GoTo a:
            Else
                .MovePrevious
            End If
        Wend
    End With
a:
    With Adodc1.Recordset
        .AddNew
        .Fields(0) = Format(Date, "DD/MM/YYYY")
        .Fields(1) = Format(Date, "DD/MM/YYYY")
        .Fields(2) = txtMovieName.Text
        .Fields(3) = cmbTheater.ListIndex + 1
        .Fields(4) = Val("0")
        .Save
    End With
    
    Adodc1.Recordset.Close
    Adodc1.Recordset.Open "Select * from Theater", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Set lblBind.DataSource = Adodc1
    
    With Adodc1.Recordset
        .MoveFirst
        While .EOF <> True
            If .Fields(0) = cmbTheater.ListIndex + 1 Then
                For i = 0 To cmbShows.ListIndex
                    If cmbAMPM(i).ListIndex = 1 Then
                        If Mid(MaskEdBox1(i).ClipText, 1, 2) <> 12 Then
                            showTime = (12 + Mid(MaskEdBox1(i).ClipText, 1, 2))
                        Else
                            showTime = "12"
                        End If
                        showTime = showTime & ":" & Mid(MaskEdBox1(i).ClipText, 3, 2)
                    Else
                        showTime = MaskEdBox1(i).Text
                    End If
                    
                    .Fields(i + 2) = FormatDateTime(showTime, vbLongTime)
                    showTime = ""
                Next i
                For j = i To 4
                    .Fields(j + 2) = FormatDateTime("00:00:00", vbLongTime)
                Next
                .Update
                Unload Me
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
    End With
    
                    
End Sub

Private Sub Form_Load()
    lblToday.Caption = Format$(Date, "dddd, dd-MMMM-yyyy")
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "Select * from Movie"
    Set lblBind.DataSource = Adodc1
End Sub

Private Sub txtMovieName_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub
