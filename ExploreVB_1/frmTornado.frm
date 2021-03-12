VERSION 5.00
Begin VB.Form frmSpiral 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmTimer 
      Interval        =   500
      Left            =   3720
      Top             =   2400
   End
End
Attribute VB_Name = "frmSpiral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub drawshape()
Dim x As Single, y As Single
Dim totalRadians As Single, r As Single
Dim a As Single, theta As Single
Dim x1 As Single, y1 As Single


Call Randomize

Scale (3, -3)-(-3, 3)     'change scale

totalRadians = 8 * Atn(1) 'circle in radians


ForeColor = RGB(Rnd() * 256, Rnd() * 256, Rnd() * 256)

a = 3 * Rnd() 'offset used in equation

For theta = 0.001 To totalRadians Step 0.01
    r = Sqr(a ^ 2 / theta)
    x = r * Cos(theta)     'y coordinate
    y = r * Sin(theta)     'x coordinate
    x1 = -r * Cos(theta)   'y coordinate
    y1 = -r * Sin(theta)   'x coordinate
PSet (x, y)
PSet (x1, y1)
Next theta
End Sub

Private Sub tmTimer_Timer()
Call drawshape   'call time interval expires

End Sub
