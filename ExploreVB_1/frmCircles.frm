VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMSH 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees"
   ClientHeight    =   6060
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8910
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00FFFFC0&
      Caption         =   "E&xport"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox cboPos 
         Height          =   315
         Left            =   4680
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cboSec 
         Height          =   315
         Left            =   4680
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cboDept 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "Position:"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         Caption         =   "Section:"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400000&
         Caption         =   "Department:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         Caption         =   "Employee:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   8160
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   4194304
      ForeColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   16777152
      ForeColorFixed  =   0
      BackColorBkg    =   4194304
      GridColor       =   -2147483648
      GridColorFixed  =   -2147483648
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   7560
      Width           =   1335
   End
End
Attribute VB_Name = "frmMSH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Selected  As String



Private Sub cboDept_Click()
Call getsection
End Sub

Private Sub cmdDelete_Click()
Dim clsnew As New clsInsert
Dim strdel As String

If Selected <> "" Then
    If MsgBox("Are you sure you want to delete?", vbYesNo, "Confirmation") = vbYes Then
    strdel = clsnew.Delete(Selected)
    Call search(True)
    Me.MSHFlexGrid1.Refresh
    Else: End If
Else
MsgBox "Select which employee to Delete", vbCritical, "Warning"
End If

End Sub

Private Sub Form_Load()
Call search(True)
Selected = ""
Call SearchbyComboBox
End Sub

Private Sub cmdAdd_Click()
frmAddEmp.Show
frmAddEmp.Caption = "Add New Employee"
frmAddEmp.mnuSave.Caption = "Save"
End Sub

Private Sub cmdEdit_Click()
    If Selected <> "" Then
    
        Me.Hide
        frmAddEmp.Show
        Call EditEmp(Selected)
        frmAddEmp.Caption = "Update Employee Information"
        frmAddEmp.mnuSave.Caption = "Update"
    Else
        MsgBox "Please Select employee to update information", vbCritical, "Warning"
    End If
    
End Sub

Private Sub MSHFlexGrid1_Click()
    Selected = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
End Sub

Private Sub Text1_Change()
With Me.MSHFlexGrid1
    .TextMatrix(.Row, .Col) = Me.Text1.Text
End With
End Sub

Private Sub Text1_GotFocus()
Me.Text1.SelStart = 0
Me.Text1.SelLength = Len(Me.Text1.Text)
End Sub

Private Sub Text2_Change()
Call search(False)
End Sub

Private Sub cmdExport_Click()
'EXCEL
Dim xlApp As Excel.Application
Dim xlworkbook As Excel.Workbook
Dim xlworksheet As Excel.Worksheet
Dim pass As Integer
Dim rLetter As String

Set xlApp = CreateObject("Excel.Application")
Set xlworkbook = xlApp.Workbooks.Add
Set xlworksheet = xlworkbook.Sheets("Sheet1")

With xlworksheet
    .name = "TEST"
    .Range("A1").Formula = "EMPLOYEES"
    .Range("A1").Font.name = "ARIAL"
    .Range("A1").Font.Size = 12
    .Range("A1").Font.Bold = True
    .Range("A1:C1").Merge
    
End With

With Me.MSHFlexGrid1
num = 3
        For C = 0 To .Rows - 1
        num = num + 1
            For b = 0 To .Cols - 1
                
               If b = 0 Then
               rLetter = "A" & num
               
               ElseIf b = 1 Then
               rLetter = "B" & num
               
               ElseIf b = 2 Then
               rLetter = "C" & num
            
               ElseIf b = 3 Then
               
               rLetter = "D" & num
               ElseIf b = 4 Then
               
               rLetter = "E" & num
               ElseIf b = 5 Then
              
               rLetter = "F" & num
               End If
               
                'MsgBox rLetter & " = " & .TextMatrix(C, B)
                 xlworksheet.Range("A4:F4").Font.Bold = True
                 xlworksheet.Range("D4").ColumnWidth = 13
                 xlworksheet.Range("B4:E4").ColumnWidth = 18
                 xlworksheet.Range(rLetter).Formula = .TextMatrix(C, b)
                 
            Next b
        Next C
End With
xlApp.Visible = True

'=====EXPORT TO  NOTEPAD
'Dim fs As Object
'Dim A As Object
'
'Set fs = CreateObject("Scripting.FileSystemObject")
'Set A = fs.CreateTextFile("C:\Users\smd084\Desktop\aa.txt", True)
'A.WriteLine ("")
'For Z = 1 To Me.MSHFlexGrid1.Rows - 1
'    With Me.MSHFlexGrid1
'        A.WriteLine ("" & .TextMatrix(Z, 0) & Space$(5) & .TextMatrix(Z, 1) & "")
'    End With
'
'Next Z
'Call Shell("explorer.exe C:\Users\smd084\Desktop\test\aa.txt", vbNormalFocus)
'A.Close
End Sub
