VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "discharge input"
   ClientHeight    =   2820
   ClientLeft      =   8655
   ClientTop       =   4125
   ClientWidth     =   4890
   Icon            =   "Form6_discharge.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4890
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "apply changes"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Discharge variables"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "Form6_discharge.frx":0442
         Left            =   2520
         List            =   "Form6_discharge.frx":045E
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Text            =   "2000"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Average discharge [m3/sec]"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "discharge volatility"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form6.Hide
End Sub

Private Sub Command2_Click()
Form6.Hide
End Sub

Private Sub List1_Click()
f = Form6.List1.ListIndex
If f = 0 Then
    Q_volatility = 0.1
ElseIf f = 1 Then
    Q_volatility = 0.2
ElseIf f = 2 Then
    Q_volatility = 0.3
ElseIf f = 3 Then
   Q_volatility = 0.4
   k_er_fluvC = 0.00001
   Form8.List2.ListIndex = 4
ElseIf f = 4 Then
   Q_volatility = 0.5
   k_er_fluvC = 0.00001
   Form8.List2.ListIndex = 4
ElseIf f = 5 Then
   Q_volatility = 0.6
   k_er_fluvC = 0.000005
   Form8.List2.ListIndex = 5
ElseIf f = 6 Then
   Q_volatility = 0.7
   k_er_fluvC = 0.000005
   Form8.List2.ListIndex = 5
ElseIf f = 7 Then
   Q_volatility = 0
End If
End Sub



Private Sub Text7_Change()
Q_average = Form6.Text7.Text
End Sub

