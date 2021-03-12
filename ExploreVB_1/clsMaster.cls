VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function Master(ByVal strsql As String) As Object
Dim rsDept As New ADODB.Recordset

Call connect
rsDept.Open strsql, cn, adOpenDynamic, adLockReadOnly
Set Master = rsDept
End Function

Public Function rsDept() As Object
Dim depsql As String
Dim rs As New ADODB.Recordset

depsql = "Select *  from Departments"
Call connect
rs.Open depsql, cn, adOpenDynamic, adLockReadOnly
Set rsDept = rs
End Function

Public Function rsSec() As Object
Dim secsql As String
Dim rs As New ADODB.Recordset

secsql = "Select Sectionname from Sections"
Call connect
rs.Open secsql, cn, adOpenDynamic, adLockReadOnly
Set rsSec = rs
End Function

Public Function rsPos() As Object
Dim possql As String
Dim rs As New ADODB.Recordset

possql = "Select positionname from Positions"
Call connect
rs.Open possql, cn, adOpenDynamic, adLockReadOnly
Set rsPos = rs
End Function

Public Function ViewSectionLink(ByVal Dept As Integer) As Object
Dim secLinksql As String
Dim rs As New ADODB.Recordset

secLinksql = "Select * from SectionLink " & _
            " inner join sections on sections.sectionid=sectionLink.SectionID " & _
            " where departmentid =" & Dept  'Sectionid
Call connect
rs.Open secLinksql, cn, adOpenDynamic, adLockReadOnly
 
 Set ViewSectionLink = rs
 
End Function
