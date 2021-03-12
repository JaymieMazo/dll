VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPositions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Positions"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6255
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPos 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5530
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim strsql As String
Dim rsPositions As Object
Dim clsnew As New clsMaster


With clsnew
    strsql = "Select * from Positions"
    Set rsPositions = clsnew.Master(strsql)
End With

With mshPos
.Rows = rsPositions.RecordCount + 1
.Cols = 5
.TextMatrix(0, 0) = "No."
.TextMatrix(0, 1) = "Department"
.TextMatrix(0, 2) = "Created Date"
.TextMatrix(0, 3) = "Deleted Date"
.TextMatrix(0, 4) = "Updated Date"

.ColWidth(0) = 500
.ColWidth(1) = 3000
.ColWidth(2) = 1200
.ColWidth(3) = 1200
.ColWidth(4) = 1200

For A = 1 To rsPositions.RecordCount
    .TextMatrix(A, 0) = rsPositions.Fields("PositionID").Value
    .TextMatrix(A, 1) = rsPositions.Fields("PositionName").Value
    .TextMatrix(A, 2) = Format(rsPositions.Fields("PositionID").Value, "YYYY/MM/DD")
    .TextMatrix(A, 3) = Format(rsPositions.Fields("DeletedDate").Value, "YYYY/MM/DD")
    .TextMatrix(A, 4) = Format(rsPositions.Fields("UpdatedDate").Value, "YYYY/MM/DD")
    
rsPositions.MoveNext
Next

End With
End Sub
