VERSION 5.00
Begin VB.Form frmCircles 
   Caption         =   "Circles"
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
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "frmCircles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Single, y As Single
Dim r As Single



Private Sub cmdOk_Click()



Call Randomize
x = Rnd() * 5
y = Rnd() * 5
r = Rnd() * 2.5

Scale (-1, 5)-(5, -5)

Select Case Int(Rnd() * 5)

Case 0 'circle
    Circle (x, y), r

Case 1 'Ellipse
    Circle (x, y), r, , , , 1.5

Case 2 'Ellipse
    Circle (x, y), r, , , , 0.5

Case 3 'Sector
    Circle (x, y), r, vbRed, -2 * Atn(-1), -4 * Atn(1)

Case 4 'Arc

    Circle (x, y), r, vbRed, 2 * Atn(1), 4 * Atn(1)
End Select
PSet (100, 200)
PSet (120, 220)
End Sub

Private Sub Form_Load()
' PI is equi. to 4 * Atn(1)


'Circle (50, 50), 25                          'circle
'Circle (0, 0), 1, vbRed, -PI / 2, -PI         'Sector
'Circle (1.5, 1.2), 1, vbRed, 3 * PI / 2, 0   'arc
'Circle (10, 10), 5, vbRed, , , 1.5           'ellipse\

End Sub
