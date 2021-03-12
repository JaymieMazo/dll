VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00400000&
   Caption         =   "Employee System"
   ClientHeight    =   8985
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuEmployees 
         Caption         =   "Employees"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "Master"
      Begin VB.Menu mnuDept 
         Caption         =   "Departments"
      End
      Begin VB.Menu mnuSections 
         Caption         =   "Sections"
      End
      Begin VB.Menu mnuPositions 
         Caption         =   "Positions"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
frmMSH.Show
End Sub

Private Sub mnuDept_Click()
frmDepartment.Show

End Sub

Private Sub mnuEmployees_Click()
frmMSH.Show
End Sub

Private Sub mnuPositions_Click()
frmPositions.Show
End Sub

Private Sub mnuSections_Click()
frmSections.Show

End Sub
