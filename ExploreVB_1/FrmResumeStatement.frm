VERSION 5.00
Begin VB.Form FrmResumeStatement 
   Caption         =   "Resume Statement"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Error Handling Response"
      Height          =   2415
      Left            =   6600
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      Begin VB.OptionButton optresume 
         Caption         =   "Resume"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optResumeLabel 
         Caption         =   "Resume at Resume Label"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.OptionButton optResumeNxt 
         Caption         =   "Resume Next"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   8760
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "FrmResumeStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
Dim x As Integer
Dim s As String

resumeLabel:


s = "Visual Basic How to Program"

Print s

On Error GoTo Handler
x = InputBox("Enter an integer:", s)

Print "Value of x is " & x
Exit Sub




Handler:

    If (optresume.Value) Then
    Print "Resume: ";
    Resume ' Repeat Line that raised error
    ElseIf (optResumeNxt.Value) Then
     Print "Resume Next: ";
    Resume Next ' Resume at next  Line  after error-raising line
    Else
    Print "Label:";
    Resume resumeLabel  'Resume resumeLabel
    End If
    
    
    
End Sub

Private Sub optresume_Click()
Call Cls

End Sub

Private Sub optResumeLabel_Click()
Call Cls
End Sub

Private Sub optResumeNxt_Click()
Call Cls
End Sub
