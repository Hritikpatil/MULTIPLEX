VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBookDate 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Advance Booking Date"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Booking Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdContinue 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Continue >>>>>"
         Default         =   -1  'True
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   16711680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmBookDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DateError As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdContinue_Click()
    Dim Date1 As Date
    Dim Date2 As Date
    
    Date1 = Format(Date, "DD/MM/YYYY")
    
    If IsDate(MaskEdBox1.Text) = False Then
        MsgBox "Entered date is not a valid date", vbCritical + vbOKOnly, "Date Error"
        Exit Sub
    End If
    
    Date2 = Format(MaskEdBox1.Text, "DD/MM/YYYY")
    If DateError = True Then
        MsgBox "Please enter valid date", vbCritical + vbOKOnly, "Invalid Date"
        MaskEdBox1.SetFocus
        DateError = False
    ElseIf (Date2 - Date1) < 1 Then
        MsgBox "Booking date must be greater than today's date"
        MaskEdBox1.SetFocus
    Else
        AdvBookDate = Format(MaskEdBox1.Text, "DD/MM/YYYY")
        Unload Me
        Load frmAdvBookiing
        frmAdvBookiing.Show
    End If
    
End Sub

Private Sub Form_Load()
    lblToday.Caption = Format(Date, "dddd, DD/MMM/YYYY")
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If IsDate(MaskEdBox1.Text) = False Then
        DateError = True
        MsgBox "Entered date in not a valid date", vbCritical + vbOKOnly, "Date Error"
    End If
End Sub
