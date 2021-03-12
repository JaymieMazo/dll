Attribute VB_Name = "modConnection"

Public cn As New ADODB.Connection
Public deptid As Integer
Public Const strConnectionString As String = "Provider=SQLOLEDB.1;Password=h56r13d;Persist Security Info=True;User ID=rhrdap;Initial Catalog=WorkOrderMaintenance;Data Source=hrdsql6"
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Public Sub connect()
Set cn = New ADODB.Connection

    With cn
        .CursorLocation = adUseClient
        .ConnectionString = "Provider =SQLOLEDB ;Data Source=SD_SQL_TRAINING;Initial Catalog=Jai;UID=sa;PWD=81at84"
        .Open
    End With

End Sub

Public Sub search(ByVal all As Boolean)
Dim clsnew As New clsDetails
Dim empsql As String
Dim rs As Object

Call connect

    If all = True Then
    empsql = clsnew.ViewEmployee
    Else
    empsql = clsnew.ViewEmployee & " WHERE EmpId LIKE '%" & frmMSH.Text2.Text & "%' "
    End If

Set rs = New ADODB.Recordset
rs.Open empsql, cn, adOpenDynamic, adLockReadOnly


With frmMSH.MSHFlexGrid1
    
    If rs.EOF Then
        MsgBox "No Record Found!!!!!!", vbExclamation, "Message"
        .Rows = 1
        Exit Sub
    End If

    lncnt = rs.RecordCount
    .Cols = 6
    .Rows = lncnt + 1
    .ColWidth(1) = 3000
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Status"
    .TextMatrix(0, 3) = "Department"
    .TextMatrix(0, 4) = "Section"
    .TextMatrix(0, 5) = "Position"

    For a = 1 To lncnt
                .TextMatrix(a, 0) = rs("EmpID").Value
                .TextMatrix(a, 1) = rs("EmpFname").Value & " " & rs("EmpLname").Value
                .TextMatrix(a, 2) = rs("Status").Value
                .TextMatrix(a, 3) = rs("DepartmentName").Value
                .TextMatrix(a, 4) = rs("SectionName").Value
                .TextMatrix(a, 5) = rs("PositionName").Value
        rs.MoveNext
    Next
    
End With
End Sub


Public Sub AddEmp()
Dim empsql As String
Dim Add As New clsInsert
Dim clsIsSave As New clsDetails
Dim isSave As Integer

isSave = clsIsSave.FieldValidation

If isSave = True Then
empsql = Add.Employees(True)
Else
MsgBox "Please Complete all fields!", vbCritical, "Warning"
End If

End Sub

Public Sub UpdateEmp(ByVal empid As String)
Dim clsnew As New clsInsert
Dim aa As String

Dim strsql As String

    If MsgBox("Are you sure you want to save?", vbYesNo, "Confirmation") = vbYes Then
 
       aa = clsnew.Employees(False)
       
     Else: End If

End Sub

Public Sub getDeptID()
Dim a As New clsDetails
Dim b As New clsMaster
Dim Dept As Integer
Dim rs As Object

Dept = a.ViewDept(frmAddEmp.cboDept.Text)

frmAddEmp.cboSection.Clear

Set rs = b.ViewSectionLink(Dept)

    For num = 1 To rs.RecordCount
        frmAddEmp.cboSection.AddItem rs.Fields("sectionname").Value
        rs.MoveNext
    Next
End Sub

Public Sub EditEmp(ByVal empid As String)
Dim clsnew As New clsDetails
Dim rs As New ADODB.Recordset
Dim editsql As String

Call connect
editsql = clsnew.ViewEmployee & "where empid='" & empid & "'"

rs.Open editsql, cn, adOpenDynamic, adLockReadOnly
With frmAddEmp
   .cboDept.Text = Trim(rs.Fields("Departmentname").Value)
    .cboSection.Text = rs.Fields("Sectionname").Value
    .cboPos.Text = rs.Fields("Positionname").Value
    .imgPic.Picture = LoadPicture(rs.Fields("photo").Value)
        .txtID = rs.Fields("empid").Value
        .txtLName = rs.Fields("emplname").Value
        .txtFname = rs.Fields("empfname").Value
        .txtMname = rs.Fields("empmname").Value
        .cboStatus = rs.Fields("status").Value
        .cboGender = rs.Fields("sex").Value
        .cboCivilStatus = rs.Fields("civilstatus").Value
        .txtAge = rs.Fields("Age").Value
        .txtEmail = rs.Fields("email").Value
        .txtAddress = rs.Fields("presaddress").Value
        .txtProvince = rs.Fields("provaddress").Value
        .dtHired = rs.Fields("datehired").Value
        .txtSSS = rs.Fields("sssno").Value
        .txtTIN = rs.Fields("tinno").Value
        .txtRelative = rs.Fields("contactlname").Value
        .txtrelation = rs.Fields("relation").Value
        .txtContact = rs.Fields("contactcellno").Value
        .txtContactNo = rs.Fields("contactno").Value
        .txtConAddress = rs.Fields("contactaddress").Value
       
        
    .txtrelation = rs.Fields("Relation").Value
    .dtBday = rs.Fields("BirthDate").Value
    
    If rs.Fields("Suffix").Value <> "" Then .txtSuffix = rs.Fields("Suffix").Value
            If StrReverse(Mid(StrReverse(rs.Fields("photo").Value), 1, 4)) = ".bmp" Then
            .imgPic.Picture = LoadPicture(rs.Fields("photo").Value)
            Else: End If

    
End With
End Sub



Public Sub DefaultValue()
With frmAddEmp
.cboCivilStatus = "Single"
.cboStatus = "Contractual"
.dtHired = Now
.dtBday = DateSerial(Year("1995/05/29"), Month(Now), Day(Now))

End With
End Sub

Sub Main()
    
    On Error GoTo lnError
    '--- check if application is already open
    If App.PrevInstance = True Then MsgBox "System is already open!", vbInformation, "System": Exit Sub
    frmSplash.Show
    
    DoEvents
    
    'Call Connect
    
    Sleep 2000
    
    
    frmMSH.Show
    
    Unload frmSplash
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical
End Sub




