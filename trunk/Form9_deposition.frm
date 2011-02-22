VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "deposition control"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Form9_deposition.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "resore default"
      Height          =   495
      Left            =   5280
      TabIndex        =   28
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   27
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "apply change"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   10080
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll12 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   19
      Top             =   9240
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll11 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   18
      Top             =   8520
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll10 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   17
      Top             =   7800
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll9 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   16
      Top             =   7080
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll8 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   15
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Travel distance marine domain"
      Height          =   4815
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   7095
      Begin VB.TextBox Text12 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text12"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text11"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text10"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text9"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text8"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text7"
         Top             =   480
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   14
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   4
      Top             =   3360
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   3
      Top             =   2640
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      Max             =   10
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Travel distance fluvial domain"
      Height          =   4575
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin VB.TextBox Text6 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   6
         Top             =   3960
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim base As Double
'Form9.Text1 = 1600
'Form9.Text2 = 1200
'Form9.Text3 = 600
'Form9.Text4 = 480
'Form9.Text5 = 320
'Form9.Text6 = 160
'Form9.Text7 = 70
'Form9.Text8 = 60
'Form9.Text9 = 50
'Form9.Text10 = 40
'Form9.Text11 = 10
'Form9.Text12 = 5

Private Sub Command1_Click()
Form9.Hide
End Sub

Private Sub Command2_Click()
Unload Form9
End Sub

Private Sub Command3_Click()
Module1.Parameter_and_constants_default
Form9.Hide
End Sub

Private Sub HScroll1_Change()
base = 1625
traveldist_fluvial(1) = base + (((2 * (Form9.HScroll1.Value - 5) / 10) * base))
Form9.Text1 = traveldist_fluvial(1)
End Sub


Private Sub HScroll10_Change()
base = 22.5
traveldist_marine(4) = base + (((2 * (Form9.HScroll10.Value - 5) / 10) * base))
Form9.Text10 = traveldist_marine(4)
End Sub

Private Sub HScroll11_Change()
base = 18#
traveldist_marine(5) = base + (((2 * (Form9.HScroll11.Value - 5) / 10) * base))
Form9.Text11 = traveldist_marine(5)
End Sub

Private Sub HScroll12_Change()
base = 5#
traveldist_marine(6) = base + (((2 * (Form9.HScroll12.Value - 5) / 10) * base))
Form9.Text12 = traveldist_marine(6)
End Sub

Private Sub HScroll2_Change()
base = 700
traveldist_fluvial(2) = base + (((2 * (Form9.HScroll2.Value - 5) / 10) * base))
Form9.Text2 = traveldist_fluvial(2)
End Sub

Private Sub HScroll3_Change()
base = 425
traveldist_fluvial(3) = base + (((2 * (Form9.HScroll3.Value - 5) / 10) * base))
Form9.Text3 = traveldist_fluvial(3)
End Sub

Private Sub HScroll4_Change()
base = 325
traveldist_fluvial(4) = base + (((2 * (Form9.HScroll4.Value - 5) / 10) * base))
Form9.Text4 = traveldist_fluvial(4)
End Sub

Private Sub HScroll5_Change()
base = 275
traveldist_fluvial(5) = base + (((2 * (Form9.HScroll5.Value - 5) / 10) * base))
Form9.Text5 = traveldist_fluvial(5)
End Sub

Private Sub HScroll6_Change()
base = 250
traveldist_fluvial(6) = base + (((2 * (Form9.HScroll6.Value - 5) / 10) * base))
Form9.Text6 = traveldist_fluvial(6)
End Sub

Private Sub HScroll7_Change()
base = 100#
traveldist_marine(1) = base + (((2 * (Form9.HScroll7.Value - 5) / 10) * base))
Form9.Text7 = traveldist_marine(1)
End Sub

Private Sub HScroll8_Change()
base = 45#
traveldist_marine(2) = base + (((2 * (Form9.HScroll8.Value - 5) / 10) * base))
Form9.Text8 = traveldist_marine(2)
End Sub

Private Sub HScroll9_Change()
base = 27#
traveldist_marine(3) = base + (((2 * (Form9.HScroll9.Value - 5) / 10) * base))
Form9.Text9 = traveldist_marine(3)
End Sub

