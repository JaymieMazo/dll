VERSION 5.00
Begin VB.Form frmArrays 
   Caption         =   "Array"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdPrint_Click()
'SAMPLE1
'Dim n(9) As Integer
'Dim x As Integer
'
'Call Cls
'
'Print "Index" & Space$(3) & "Value"
'    For x = LBound(n) To UBound(n)
'    Print Space$(2) & x & Space$(7) & n(x)
'    Next

'SAMPLE2
Dim n(9) As Integer
Dim x As Integer

Call Cls
'Array to even elements
    For x = LBound(n) To UBound(n)
    n(x) = 2 + 2 * x
    Next x

  

  Print "Index" & Space$(3) & "Value"

     For x = LBound(n) To UBound(n)
    Print Space$(2) & x & Space$(7) & n(x)
     Next x

End Sub

