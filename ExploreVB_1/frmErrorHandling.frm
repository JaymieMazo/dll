VERSION 5.00
Begin VB.Form frmOnError 
   Caption         =   "On Error"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIserror 
      Caption         =   "Iserror"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmd13 
      Caption         =   "Type mismatch"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtden 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtenum 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "76"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label ans 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "click me"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "frmOnError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ans_Click()
Dim num As Double
Dim denum As Double

On Error GoTo inputhandler
num = txtenum.Text
denum = txtden.Text


On Error GoTo dividebyzerohandler
ans = num / denum
Exit Sub


dividebyzerohandler:
ans = "attempt divided by zero"
Exit Sub

inputhandler:
ans = "attempt non-numeric input"


End Sub

Private Sub denum_Change()





End Sub

Private Sub cmd13_Click()

Dim b As String
Dim sum  As Integer

    On Error GoTo bb
    
    sum = txtenum.Text + txtden.Text
    ans.Caption = sum
       
bb:
    
     ans.Caption = ""
   b = Error$
    
   'MsgBox b
    
   If b = "Type mismatch" Then
    
    MsgBox "Input numbers only!", vbExclamation, "Oops"
    
    Else
    sum = txtenum + ans
    
    MsgBox txtenum.Text & " + " & txtden.Text & " = " & sum
    ans.Caption = sum
    
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdIserror_Click()

Dim ReturnVal, MyCheck
Dim userfunction As Variant
Dim num1, num2, sum As Integer
Dim d As Integer
Dim a As Integer


On Error GoTo aa

 

num1 = txtenum.Text
num2 = txtden.Text
sum = num1 + num2


aa:
a = 11
For a = 13 To 15
MsgBox "a= " & a

d = a
Next
MsgBox Error(d)
userfunction = CVErr(a)
MsgBox CVErr(a)
ReturnVal = userfunction
MyCheck = IsError(ReturnVal)
MsgBox MyCheck

End Sub
