VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulation time variables"
   ClientHeight    =   2775
   ClientLeft      =   9150
   ClientTop       =   2640
   ClientWidth     =   4110
   Icon            =   "Form7_sim_time.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4110
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "time dependant variables"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   1440
         Width           =   255
      End
      Begin VB.ListBox List1 
         Height          =   450
         ItemData        =   "Form7_sim_time.frx":0442
         Left            =   2040
         List            =   "Form7_sim_time.frx":045B
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.ListBox List2 
         Height          =   255
         ItemData        =   "Form7_sim_time.frx":0488
         Left            =   2040
         List            =   "Form7_sim_time.frx":048F
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Flexible time step"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "simulation time in years"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "time step"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form7.Hide

End Sub

Private Sub List1_Click()
f = Form7.List1.ListIndex
If f = 0 Then
    end_of_times = 100
ElseIf f = 1 Then
    end_of_times = 200
ElseIf f = 2 Then
    end_of_times = 500
ElseIf f = 3 Then
    end_of_times = 1000
ElseIf f = 4 Then
    end_of_times = 2000
ElseIf f = 5 Then
    end_of_times = 5000
ElseIf f = 6 Then
    end_of_times = 10000
End If

End Sub
