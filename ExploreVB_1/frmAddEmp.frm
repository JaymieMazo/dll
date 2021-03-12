VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEmp 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Add Employees"
   ClientHeight    =   7770
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "Other Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   34
      Top             =   5520
      Width           =   10215
      Begin VB.ComboBox cboSection 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":0000
         Left            =   6960
         List            =   "frmAddEmp.frx":0002
         TabIndex        =   22
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cboPos 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":0004
         Left            =   6960
         List            =   "frmAddEmp.frx":0006
         TabIndex        =   23
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cboDept 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":0008
         Left            =   6960
         List            =   "frmAddEmp.frx":000A
         TabIndex        =   21
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtTIN 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtSSS 
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000000&
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   48
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000000&
         Caption         =   "Section:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000000&
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   46
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000000&
         Caption         =   "TIN No.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         Caption         =   "SSS No.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Contact Information"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   10335
      Begin VB.ComboBox txtrelation 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":000C
         Left            =   7080
         List            =   "frmAddEmp.frx":0031
         TabIndex        =   51
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtConAddress 
         Height          =   285
         Left            =   7080
         TabIndex        =   18
         Top             =   1200
         Width           =   2700
      End
      Begin VB.TextBox txtContact 
         Height          =   285
         Left            =   7080
         TabIndex        =   17
         Top             =   840
         Width           =   2700
      End
      Begin VB.TextBox txtRelative 
         Height          =   285
         Left            =   7080
         TabIndex        =   16
         Top             =   480
         Width           =   2700
      End
      Begin VB.TextBox txtContactNo 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   2700
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   840
         Width           =   2700
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   2700
      End
      Begin VB.TextBox txtProvince 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   1560
         Width           =   2700
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000000&
         Caption         =   "Relation:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   44
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label txt 
         BackColor       =   &H80000000&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   43
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000000&
         Caption         =   "Contact No."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   42
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label LBLRelative 
         BackColor       =   &H80000000&
         Caption         =   "Relative:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   41
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000000&
         Caption         =   "Contact No."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000000&
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000000&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000000&
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame frapersonal 
      BackColor       =   &H80000000&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   10335
      Begin MSComDlg.CommonDialog cdPic 
         Left            =   9720
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtBday 
         Height          =   285
         Left            =   5880
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   99811329
         CurrentDate     =   43612
      End
      Begin VB.TextBox txtSuffix 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   2520
         Width           =   2460
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":008E
         Left            =   5880
         List            =   "frmAddEmp.frx":0098
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cboCivilStatus 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":00B2
         Left            =   6360
         List            =   "frmAddEmp.frx":00BC
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmAddEmp.frx":00D1
         Left            =   6360
         List            =   "frmAddEmp.frx":00DB
         TabIndex        =   10
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   7320
         TabIndex        =   8
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtLName 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   2460
      End
      Begin VB.TextBox txtFname 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   2460
      End
      Begin VB.TextBox txtMname 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   2040
         Width           =   2460
      End
      Begin MSComCtl2.DTPicker dtHired 
         Height          =   285
         Left            =   5880
         TabIndex        =   6
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   99811329
         CurrentDate     =   43612
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000000&
         Caption         =   "Hired:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   50
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "Suffix:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image imgPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000000&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   45
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000000&
         Caption         =   "Birthday:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   31
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000000&
         Caption         =   "Civil status:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   30
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000000&
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblID 
         BackColor       =   &H80000000&
         Caption         =   "EmployeeID:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuCancel 
      Caption         =   "Cancel"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "frmAddEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSecton_DropDown()
Dim a As New clsDetails
Dim b As New clsMaster
Dim dept As Integer
Dim rs As New ADODB.Recordset

cboSection.Clear
Set rs = b.ViewSectionLink(dept)

dept = a.ViewDept(cboDept)
Me.cboSection.AddItem rs.Fields("SectionName").Value
End Sub

Private Sub Form_Load()
With frmMSH.MSHFlexGrid1
    txtID.Text = .TextMatrix(.Rows - 1, 0) + 1
End With

Call ComboRecords

End Sub
Private Sub cboDept_Click()
Call getDeptID
End Sub

Private Sub cmdOpen_Click()
On Error GoTo chk
cdPic.ShowOpen
    If cdPic.FileName <> "" Then
        imgPic.Picture = LoadPicture(cdPic.FileName)
    Else: End If

chk:
    If Err.Number = 481 Then
    MsgBox "Use .bmp image ", vbInformation, "Warning"
    Exit Sub
    End If
End Sub

Private Sub dtBday_LostFocus()
Dim a As New clsDetails
Dim DOB As Byte
    DOB = a.Age(Me.dtBday)
    Me.txtAge = DOB
End Sub

Private Sub mnuCancel_Click()
Unload Me
frmMSH.Show
End Sub

Private Sub mnuSave_Click()
    If mnuSave.Caption = "Save" Then
    Call AddEmp
    Else
    Call UpdateEmp(Me.txtID)
    End If
End Sub
