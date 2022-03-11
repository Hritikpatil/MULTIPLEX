VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInputMonth 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0E0FF&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdGenerate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Generate Movie wise Collection"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   65535
         ForeColor       =   16711680
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbMonth 
         BackColor       =   &H0080FFFF&
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
         Height          =   315
         ItemData        =   "frmInputMonth.frx":0000
         Left            =   1680
         List            =   "frmInputMonth.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmInputMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    Dim str As String
    str = "01/" & cmbMonth.ListIndex + 1 & "/" & MaskEdBox1.Text
    dt = Format(str, "DD/MM/YYYY")
    
    If isLoad = True Then
        Unload DataReport1
        Unload DataEnvironment1
    Else
        isLoad = True
    End If
    
    DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Multiplex.mdb;Persist Security Info=False"
    DataEnvironment1.cmdMovie (dt)
    Unload Me
    Load DataReport1
    DataReport1.Show
    
End Sub

