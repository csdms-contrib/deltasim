VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "PROCESS RESPONSE SIMULATION of a FLUVIAL DELTAIC ENVIRONMENT"
   ClientHeight    =   16200
   ClientLeft      =   2205
   ClientTop       =   1290
   ClientWidth     =   19485
   FontTransparent =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "Form1_delta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   16200
   ScaleWidth      =   19485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6960
      Left            =   0
      ScaleHeight     =   6930
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   9240
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   19425
      TabIndex        =   4
      Top             =   8160
      Width           =   19455
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   19425
      TabIndex        =   1
      Top             =   9240
      Width           =   19455
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7095
         Left            =   17880
         ScaleHeight     =   7065
         ScaleWidth      =   1545
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox pic1 
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
      DragMode        =   1  'Automatic
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
      Height          =   7800
      Left            =   0
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   7770
      ScaleWidth      =   19455
      TabIndex        =   0
      Top             =   240
      Width           =   19485
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   17880
         Top             =   6840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRun 
         Caption         =   "&Run"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSavepicture 
         Caption         =   "&Save picture"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuInput 
         Caption         =   "&Sediment Input "
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSealevel 
         Caption         =   "&Sea Level "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDischarge 
         Caption         =   "&Discharge Input"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTime 
         Caption         =   "&Simualtion Time"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "&Grid control"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuErosie 
         Caption         =   "&Ersosion control"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDeposition 
         Caption         =   "&Deposition control"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuWell 
         Caption         =   "&Show Well"
         Enabled         =   0   'False
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuAbout_Click()
frmAbout.Show

End Sub

Private Sub mnuDeposition_Click()
Parameter_and_constants
Form9.Show
End Sub

Private Sub mnuDischarge_Click()
Parameter_and_constants
Form6.Show
End Sub


Private Sub mnuErosie_Click()
Parameter_and_constants

Load Form8
Form8.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuGrid_Click()
Form5_Show = 0
Parameter_and_constants

Load Form5
Form5.Show
Form5.Refresh
Form5.Picture1.Cls
Form5.update_picture1_form5
Form5_Show = 1
End Sub

Private Sub mnuHorizont_Click()
'Form1.Arrange vbTileHorizontal
End Sub

Private Sub mnuInput_Click()

Parameter_and_constants

Load Form3
Form3.Show
Form3.Refresh
Form3.Picture1.Cls
Form3.update_a_picture_1

End Sub

Private Sub mnuRun_Click()

Form1.pic1.Height = (0.45 * Form1.Height)
Form1.pic1.Width = Form1.Width

Form1.Picture1.Height = (0.4 * Form1.Height)
Form1.Picture1.Width = Form1.Width

Form1.Picture4.Height = (0.05 * Form1.Height)
Form1.Picture4.Width = Form1.Width

Shell1
End Sub


Private Sub mnuSavepicture_Click()

    'save to file
    
    CommonDialog1.Filter = "Pictures (*.bmp)|*.bmp"
    CommonDialog1.ShowSave
   Call SavePicture(Form1.pic1.Image, CommonDialog1.FileName)
    

End Sub


Private Sub mnuSealevel_Click()
Parameter_and_constants

Load Form4
Form4.Show
Form4.Refresh
Form4.Picture1.Cls

If Form4.Option1.Value = True Then
    Form4.update_picture1_option1
ElseIf Form4.Option2.Value = True Then
    Form4.update_picture1_option2
ElseIf Form4.Option3.Value = True Then
    Form4.update_picture1_option3
End If


End Sub

Private Sub mnuTime_Click()
Parameter_and_constants
Form7.Show
End Sub

Private Sub mnuVertical_Click()
'Form1.Arrange vbTileVertical
End Sub


Private Sub mnuWell_Click()
 Load Frm_WELL
 Frm_WELL.Show
 Module2.draw_well1
 
End Sub

