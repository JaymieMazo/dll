VERSION 5.00
Begin VB.Form frmFW 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmTimer 
      Interval        =   500
      Left            =   3480
      Top             =   2400
   End
End
Attribute VB_Name = "frmFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub drawshape()
Dim x As Single, y As Single
Dim totalRadians As Single, r As Single
Dim a As Single, theta As Single

Call Randomize

Scale (3, -3)-(-3, 3)
totalRadians = 8 * Atn(1)

ForeColor = QBColor(Rnd() * 15)

a = 3 * Rnd() 'offset used in equation

For theta = 0 To totalRadians Step 0.01
r = a * Sin(12 * theta)  ' multi-Leaved Rose
x = r * Cos(theta)      'y coordinate
y = r * Sin(theta)       'x coordinate
PSet (x, y)
Next theta
End Sub




Private Sub tmTimer_Timer()
Call drawshape
End Sub
