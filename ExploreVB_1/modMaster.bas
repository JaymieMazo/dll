Attribute VB_Name = "modMaster"
Option Explicit

Public Sub ComboRecords()
Dim a As Integer
Dim Newcls As New clsMaster
Dim rsdept As Object  'New ADODB.Recordset
Dim rsSec As New ADODB.Recordset
Dim rsPos As New ADODB.Recordset

Set rsdept = Newcls.rsdept
Set rsSec = Newcls.rsSec
Set rsPos = Newcls.rsPos

For a = 1 To rsdept.RecordCount
frmAddEmp.cboDept.AddItem rsdept.Fields("DepartmentName").Value
rsdept.MoveNext
Next


 For a = 1 To rsSec.RecordCount
        frmAddEmp.cboSection.AddItem rsSec.Fields(0).Value
        rsSec.MoveNext
Next

  For a = 1 To rsPos.RecordCount
            frmAddEmp.cboPos.AddItem rsPos.Fields(0).Value
            rsPos.MoveNext
        Next
End Sub

Public Sub SearchbyComboBox()
Dim a As Integer
Dim Newcls As New clsMaster
Dim rsdept As Object  'New ADODB.Recordset
Dim rsSec As New ADODB.Recordset
Dim rsPos As New ADODB.Recordset

Set rsdept = Newcls.rsdept
Set rsSec = Newcls.rsSec
Set rsPos = Newcls.rsPos

For a = 1 To rsdept.RecordCount
frmMSH.cboDept.AddItem rsdept.Fields("DepartmentName").Value
rsdept.MoveNext
Next


' For a = 1 To rsSec.RecordCount
'        frmMSH.cboSec.AddItem rsSec.Fields(0).Value
'        rsSec.MoveNext
'Next

  For a = 1 To rsPos.RecordCount
            frmMSH.cboPos.AddItem rsPos.Fields(0).Value
            rsPos.MoveNext
        Next
End Sub


Public Sub getsection()
Dim a As New clsDetails
Dim b As New clsMaster
Dim deptid As Integer
Dim rs As New ADODB.Recordset
Dim num As Integer


frmMSH.cboSec.Clear
deptid = a.ViewDept(frmMSH.cboDept.Text)

Set rs = b.ViewSectionLink(deptid)

For num = 1 To rs.RecordCount
    frmMSH.cboSec.AddItem rs.Fields("SectionName").Value
    rs.MoveNext
Next

End Sub







