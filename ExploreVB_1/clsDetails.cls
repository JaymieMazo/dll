VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sql As String

Public Function ViewEmployee() As String
 
 sql = " SELECT * FROM dbo.Employees  " & _
    " INNER JOIN Departments ON Employees.EmpDept =Departments.DepartmentID " & _
    " INNER JOIN Sections ON Employees.Section =Sections.SectionID " & _
    " INNER JOIN Positions ON  Employees.Position =Positions.PositionID "
ViewEmployee = sql
End Function

Public Function ViewDept(ByVal deptName As String) As Integer
Dim rs As New ADODB.Recordset
Dim sqlDept As String

sqlDept = "Select * from Departments"
If deptName <> "" Then sqlDept = sqlDept & " where Departmentname='" & deptName & "'"

Call connect
rs.Open sqlDept, cn, adOpenDynamic, adLockReadOnly
ViewDept = rs.Fields("DepartmentID").Value
End Function

Public Function ViewSection(ByVal SecName As String) As Integer
Dim rs As New ADODB.Recordset
Dim sqlsec As String

sqlsec = "Select * from Sections "
If SecName <> "" Then sqlsec = sqlsec & " where SectionName='" & SecName & "'"
Call connect

rs.Open sqlDept, cn, adOpenDynamic, adLockReadOnly
ViewSection = rs.Fields("sectionid").Value
End Function

Public Function ViewPosition(ByVal PosName As String) As Integer
Dim rs As New ADODB.Recordset
Dim sqlPos As String

sqlPos = "Select * from Positions where PositionName='" & PosName & "'"

Call connect
rs.Open sqlPos, cn, adOpenDynamic, adLockReadOnly
ViewPosition = rs.Fields("Positionid").Value
End Function

Public Function Age(ByVal DOB As Date) As Byte
    Dim Birthday As Date
    Birthday = DateSerial(Year(Now()), Month(DOB), Day(DOB))
    Age = Year(Now()) - Year(DOB)
    If Birthday > Now() Then Age = Age - 1
End Function


Public Function insertEmployee() As String
Dim clsLink As New clsDetails
Dim DepartID As Integer
Dim Sectid As Integer
Dim posid As Integer

With frmAddEmp

DepartID = ViewDept(.cboDept)
Sectid = ViewSection(.cboSection)
posid = ViewPosition(.cboPos)

insertEmployee = "INSERT INTO dbo.Employees (EmpId, EmpLname, EmpFname, EmpMname, Status, " & _
"Sex, CivilStatus, Age, ContactNo, Email, PresAddress, ProvAddress, DateHired, EmpDept," & _
"Section, Position, SSSno, TINno, ContactLname, ContactCellno, " & _
"ContactAddress, Suffix, Relation, CreatedDate, DeletedDate, UpdatedDate, BirthDate, Photo)" & vbCrLf & _
"VALUES('" & .txtID & "','" & .txtLName & "','" & .txtFname & "','" & .txtMname & "','" & .cboStatus & "','" & .cboGender & "','" & _
.cboCivilStatus & "'," & .txtAge & ",'" & .txtContactNo & "','" & .txtEmail & "','" & .txtAddress & "','" & .txtProvince & "','" & _
.dtHired.Value & "'," & DepartID & "," & Sectid & "," & posid & _
",'" & .txtSSS & "','" & .txtTIN & "','" & .txtRelative & "','" & .txtContact & "','" & .txtConAddress & _
"','" & .txtSuffix & "','" & .txtrelation & "',GETDATE(),NULL, GETDATE(), '" & .dtBday & "','" & _
.cdPic.FileName & "')"
End With

End Function

Public Function UpdateEmployee()
Dim clsnew As New clsDetails
Dim strsql As String
Dim depid As Integer
Dim secid As Integer
Dim posid As Integer


With frmAddEmp
       strsql = "UPDATE dbo.Employees " & vbCrLf & "Set EmpLname ='" & .txtLName & "',EmpFname = '" & .txtFname & "',empmname='" & _
       .txtMname & "',  Status ='" & .cboStatus & "', Sex='" & .cboGender & "', CivilStatus='" & .cboCivilStatus & "', Age=" & .txtAge & _
        ", ContactNo='" & .txtContact & "',  Email='" & Trim(.txtEmail) & "', PresAddress='" & .txtAddress & "', ProvAddress='" & .txtProvince & "', " & _
        "DateHired='" & .dtHired.Value & "', EmpDept="
    
        depid = clsnew.ViewDept(.cboDept)
        secid = clsnew.ViewSection(.cboSection)
        posid = clsnew.ViewPosition(.cboPos)
        
        
        strsql = strsql & depid & ", Section=" & secid & ", Position=" & posid & _
        ", SSSno='" & .txtSSS & "', Tinno='" & .txtTIN & "', ContactLname='" & .txtRelative & "' , ContactCellno='" & .txtContactNo & "', ContactAddress='" & _
        .txtConAddress & "',Suffix='" & .txtSuffix & "', Relation='" & .txtrelation.Text & "', updateddate=GETDATE(), birthdate='" & .dtBday & "',Photo='" & .cdPic.FileName & "'" & _
        "where empid='" & .txtID & "'"
         
    UpdateEmployee = strsql
    
End With

End Function


Public Function FieldValidation() As Boolean

With frmAddEmp
        If .txtLName <> "" And .txtMname <> "" And .txtFname <> "" _
         And .cboStatus <> "" And .txtAge <> "" And .cboCivilStatus <> "" _
        And .cboGender <> "" And .cdPic.FileName <> "" And .txtContactNo <> "" And .txtAddress <> "" And .txtProvince <> "" And _
        .txtEmail <> "" And .txtRelative <> "" And .txtContact <> "" And .txtConAddress <> "" And .txtSSS <> "" And .txtTIN <> "" _
        And .cboDept <> "" And .cboSection <> "" And .cboPos <> "" And .cdPic.FileName <> "" Then

        FieldValidation = True
        Else
        
        FieldValidation = False
        
        End If

End With

End Function
