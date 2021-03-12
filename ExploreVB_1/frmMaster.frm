VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDepartment 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Department"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Delete"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1400
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Print"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1400
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Edit"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Add"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDept 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393216
      BackColor       =   4194304
      ForeColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   12648447
      BackColorSel    =   16777088
      BackColorBkg    =   4194304
      GridColor       =   16777215
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim strsql As String
Dim rsdept As Object
Dim clsnew As New clsMaster


With clsnew
    strsql = "Select * from Departments"
    Set rsdept = clsnew.Master(strsql)
End With

With mshDept
.Rows = rsdept.RecordCount + 1
.Cols = 5
.TextMatrix(0, 0) = "No."
.TextMatrix(0, 1) = "Department"
.TextMatrix(0, 2) = "Created Date"
.TextMatrix(0, 3) = "Deleted Date"
.TextMatrix(0, 4) = "Updated Date"

.ColWidth(0) = 500
.ColWidth(1) = 2000
.ColWidth(2) = 1100
.ColWidth(3) = 1100
.ColWidth(4) = 1100

For A = 1 To rsdept.RecordCount

.TextMatrix(A, 0) = rsdept.Fields("DepartmentID").Value
.TextMatrix(A, 1) = rsdept.Fields("DepartmentName").Value
.TextMatrix(A, 2) = Format(rsdept.Fields("CreatedDate").Value, "YYYY/MM/DD")
.TextMatrix(A, 3) = Format(rsdept.Fields("DeletedDate").Value, "YYYY/MM/DD")
.TextMatrix(A, 4) = Format(rsdept.Fields("UpdatedDate").Value, "YYYY/MM/DD")

rsdept.MoveNext
Next

End With
        


End Sub



Public Sub Master(ByVal name As String)

End Sub

