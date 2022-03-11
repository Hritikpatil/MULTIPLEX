VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Multiplex System"
   ClientHeight    =   3390
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6615
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAddNewMovie 
      Caption         =   "&Add New Movie"
   End
   Begin VB.Menu mnuCBooking 
      Caption         =   "Current &Booking"
   End
   Begin VB.Menu mnuAdvBooking 
      Caption         =   "Advance B&ooking"
   End
   Begin VB.Menu mnuCollection 
      Caption         =   "&Collection"
      Begin VB.Menu mnuMovieWiseCollection 
         Caption         =   "Movie &Wise"
      End
      Begin VB.Menu mnuTheaterWise 
         Caption         =   "Theater &Wise"
         Begin VB.Menu mnuDaily 
            Caption         =   "&Daily"
         End
         Begin VB.Menu mnuMonthly 
            Caption         =   "&Monthly"
         End
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAddNewMovie_Click()
    Load frmNewMovie
    frmNewMovie.Show
End Sub

Private Sub mnuAdvBooking_Click()
    Load frmBookDate
    frmBookDate.Show (1)
End Sub

Private Sub mnuCBooking_Click()
    Load frmBooking
    frmBooking.Show
End Sub

Private Sub mnuDaily_Click()
    Load frmDailyDt
    frmDailyDt.Show (1)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuMonthly_Click()
    Load frmMonthly
    frmMonthly.Show (1)
End Sub

Private Sub mnuMovieWiseCollection_Click()
    Load frmInputMonth
    frmInputMonth.Show (1)
End Sub
