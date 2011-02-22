VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_WELL 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Well "
   ClientHeight    =   10320
   ClientLeft      =   9765
   ClientTop       =   1110
   ClientWidth     =   5550
   Icon            =   "Form_WELL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   5550
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Well"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   9720
      Width           =   1935
   End
   Begin VB.PictureBox View_WellX 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   465
      ScaleWidth      =   4665
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.PictureBox View_WellY 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8025
      ScaleWidth      =   585
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin MSComctlLib.Slider Slider_WellPos 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      OLEDropMode     =   1
      Max             =   500
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin VB.PictureBox View_Well 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   720
      ScaleHeight     =   8025
      ScaleWidth      =   4665
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3240
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label2 
      Caption         =   "thickness [meters]"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Grain Size (microns)"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Frm_WELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Frm_WELL
well_counter = 0
End Sub

Private Sub Command2_Click()
   On Error GoTo ErrorHandle
    
    
    CommonDialog1.Filter = "Pictures (*.bmp)|*.bmp"
    CommonDialog1.ShowSave
    Call SavePicture(Frm_WELL.View_Well.Image, CommonDialog1.FileName)
    
ErrorHandle:
  
  If Err.Number <> 0 Then
  MsgBox Err.Number & " - " & Err.Description
  End If

End Sub

Private Sub Slider_WellPos_Change()
    If well_counter = 1 Then
    Module2.draw_well1
    End If
End Sub
    
    
