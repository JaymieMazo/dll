VERSION 5.00
Begin VB.Form frmLinesRectangle 
   Caption         =   "Lines"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "ok"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Width           =   615
   End
End
Attribute VB_Name = "frmLinesRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOk_Click()

Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single

Call Randomize

'pick values
x1 = Rnd() * Width
y1 = Rnd() * Height
x2 = Rnd() * Width
y2 = Rnd() * Height

'Shape`
If (Rnd() > 0.5) Then

Line (x1, y1)-(x2, y2), vbYellow, B

Else

Line (x1, y1)-(x2, y2), vbRed, B

End If



End Sub

