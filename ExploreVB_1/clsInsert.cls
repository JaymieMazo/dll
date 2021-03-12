VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Employees(ByVal isAdd As Boolean) As String

Dim rs As New ADODB.Recordset
Dim strsql As New clsDetails
    
        If isAdd = True Then
        Employees = strsql.insertEmployee
        Else
         Employees = strsql.UpdateEmployee
        
        End If
        
Call connect

Set rs = cn.Execute(Employees)

If isAdd = True Then MsgBox "Added successfully", vbInformation, "Success"
frmMSH.Show
Call search(True)

frmMSH.MSHFlexGrid1.Refresh
frmAddEmp.Hide
cn.Close

End Function


Public Function Delete(ByVal empid As String) As String
Dim strdel As String
Dim rs As New ADODB.Recordset

strdel = "Delete from employees where empid='" & empid & "'"

Set rs = cn.Execute(strdel)


End Function

