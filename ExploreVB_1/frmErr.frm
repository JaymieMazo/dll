VERSION 5.00
Begin VB.Form frmErrObj 
   Caption         =   "Error Object"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdErr 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmErrObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdErr_Click()
Dim num As Integer

Call Cls
Font.Size = 12
 
 
 On Error GoTo errorHandler
 
 Select Case Int(Rnd() * 3)
 
 Case 0
 num = 8888888888#
 
 Case 1
 num = 88 / 2
 
Case 2
num = "An Error"

End Select

Exit Sub


errorHandler:

Print "Source: " & Err.Source
Print "Error:";


Select Case Err.Number
Case 6
ForeColor = vbYellow


Case 11
ForeColor = vbWhite


Case 13
ForeColor = vbBlue

Case Else
Print "Unexpected Error!!!"
End Select

Print Err.Description
ForeColor = vbBlack
End Sub

