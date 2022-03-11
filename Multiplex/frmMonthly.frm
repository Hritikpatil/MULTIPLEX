VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMonthly 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0E0FF&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdGenerate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Generate Report"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   16711680
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbMonth 
         BackColor       =   &H0000FFFF&
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
         Height          =   360
         ItemData        =   "frmMonthly.frx":0000
         Left            =   1800
         List            =   "frmMonthly.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick a Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    Dim s As String
    s = "01/" & cmbMonth.ListIndex + 1 & "/" & MaskEdBox1.Text
    dt = Format(s, "DD/MM/YYYY")
    
    If isLoad = True Then
        Unload DataReport3
        Unload DataEnvironment1
    Else
        isLoad = True
    End If

    DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    DataEnvironment1.cmdMonthly (dt)
    Unload Me
    Load DataReport3
    DataReport3.Show

End Sub
