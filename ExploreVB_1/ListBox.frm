VERSION 5.00
Begin VB.Form frmList 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDateFormat 
      Caption         =   "DateFormat"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdDatediff 
      Caption         =   "Datediff"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdDateAdd 
      Caption         =   "DateAdd"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdDtPart 
      Caption         =   "date part"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox lst1 
      Height          =   2400
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()

End Sub

Private Sub cmdDateAdd_Click()
lst1.Clear
lst1.AddItem "Now + 3 Years: " & DateAdd("yyyy", 2, Now)
End Sub

Private Sub cmdDatediff_Click()
lst1.Clear
lst1.AddItem "Days bet. Now and 2018/12/25: " & DateDiff("d", Now, "2018/12/25")
lst1.AddItem "Days bet.  2018/08/21 and Now :" & DateDiff("m", "2018/08/21", Now)
End Sub


Private Sub cmdDateFormat_Click()
lst1.Clear
lst1.AddItem "d/m/y : " & Format(Now, "d/m/yy")
lst1.AddItem "Nov. 12, 2018 : " & Format(Now, "mmm. dd , yyyy")
lst1.AddItem "November 12, 2018 : " & Format(Now, "mmmm/dd/yyyy")

End Sub

Private Sub cmdDtPart_Click()
lst1.Clear
lst1.AddItem "Year: " & DatePart("yyyy", Now)
lst1.AddItem "Quarter: " & DatePart("q", Now)
lst1.AddItem "Month: " & DatePart("m", Now)
lst1.AddItem "day of Year: " & DatePart("y", Now)
lst1.AddItem "Day: " & DatePart("d", Now)
lst1.AddItem "Weekday: " & DatePart("w", Now)
lst1.AddItem "Week: " & DatePart("ww", Now)
lst1.AddItem "Hour: " & DatePart("h", Now)
lst1.AddItem "Minute: " & DatePart("n", Now)
lst1.AddItem "Second: " & DatePart("s", Now)
End Sub

Private Sub Form_Load()
lst1.AddItem "Today is: " & Date
lst1.AddItem "Day: " & Day(Date)
lst1.AddItem "Weekday: " & Weekday(Date) 'Start from Sunday
lst1.AddItem "WeekdayName: " & WeekdayName(Weekday(Date))
lst1.AddItem "WeekdayName Abbr.: " & WeekdayName(Weekday(Date), True)
lst1.AddItem "Month: " & Month(Date)
lst1.AddItem "MonthName: " & MonthName(Month(Date))
lst1.AddItem "MonthName: " & MonthName(Month(Date), True)
lst1.AddItem "Year: " & Year(Date)
lst1.AddItem (#1/13/1996#)
End Sub




