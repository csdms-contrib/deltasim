VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "erosion control"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form8_erosie.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "apply change"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   450
      ItemData        =   "Form8_erosie.frx":0442
      Left            =   2520
      List            =   "Form8_erosie.frx":045B
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "Form8_erosie.frx":049C
      Left            =   2520
      List            =   "Form8_erosie.frx":04A3
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "erosion capacity fluvial"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "erosion capacity marine"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form8.Hide

End Sub

Private Sub List2_Click()

f = Form8.List2.ListIndex
If f = 0 Then
    k_er_fluvC = 0.001
ElseIf f = 1 Then
    k_er_fluvC = 0.0005
ElseIf f = 2 Then
    k_er_fluvC = 0.0001
ElseIf f = 3 Then
   k_er_fluvC = 0.00005
ElseIf f = 4 Then
   k_er_fluvC = 0.00001
ElseIf f = 5 Then
   k_er_fluvC = 0.000005
ElseIf f = 6 Then
   k_er_fluvC = 0.000001
End If

End Sub
