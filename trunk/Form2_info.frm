VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "program information"
   ClientHeight    =   4395
   ClientLeft      =   10155
   ClientTop       =   6495
   ClientWidth     =   4590
   Icon            =   "Form2_info.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4590
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Show well"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3960
      Width           =   2055
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   3825
      ScaleWidth      =   4530
      TabIndex        =   1
      Top             =   0
      Width           =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

 Unload Form2

End Sub

Private Sub Command2_Click()
 Load Frm_WELL
 Frm_WELL.Show
 Module2.draw_well1
End Sub

