VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSections 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sections"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7065
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshsect 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5741
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim strsql As String
Dim rsSections As Object
Dim clsnew As New clsMaster


With clsnew
    strsql = "Select * from Sections"
    Set rsSections = clsnew.Master(strsql)
End With

With mshsect
.Rows = rsSections.RecordCount + 1
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

For A = 1 To rsSections.RecordCount

.TextMatrix(A, 0) = rsSections.Fields("SectionID").Value
.TextMatrix(A, 1) = rsSections.Fields("SectionName").Value
.TextMatrix(A, 2) = Format(rsSections.Fields("CreatedDate").Value, "YYYY/MM/DD")
.TextMatrix(A, 3) = Format(rsSections.Fields("DeletedDate").Value, "YYYY/MM/DD")
.TextMatrix(A, 4) = Format(rsSections.Fields("UpdatedDate").Value, "YYYY/MM/DD")

rsSections.MoveNext
Next

End With

End Sub
