VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRID variables"
   ClientHeight    =   8100
   ClientLeft      =   7995
   ClientTop       =   1980
   ClientWidth     =   6225
   Icon            =   "Form5_grid.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   6225
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3720
      TabIndex        =   15
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "grid variables"
      Height          =   4095
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   5415
      Begin VB.ListBox List2 
         Height          =   450
         ItemData        =   "Form5_grid.frx":0442
         Left            =   3120
         List            =   "Form5_grid.frx":0452
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "user defined profile"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "equilibrium profile"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3480
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   3000
         Width           =   255
      End
      Begin VB.ListBox List1 
         Height          =   450
         ItemData        =   "Form5_grid.frx":0470
         Left            =   3120
         List            =   "Form5_grid.frx":0486
         TabIndex        =   13
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Height          =   255
         Index           =   1
         ItemData        =   "Form5_grid.frx":04A8
         Left            =   3120
         List            =   "Form5_grid.frx":04AF
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List4 
         Height          =   450
         ItemData        =   "Form5_grid.frx":04B8
         Left            =   3120
         List            =   "Form5_grid.frx":04C8
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ListBox List5 
         Height          =   255
         ItemData        =   "Form5_grid.frx":04E2
         Left            =   3120
         List            =   "Form5_grid.frx":04E9
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ListBox List6 
         Height          =   450
         ItemData        =   "Form5_grid.frx":04F3
         Left            =   3120
         List            =   "Form5_grid.frx":0500
         TabIndex        =   4
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "initial onshore gradient"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "knick point cq oflap  break"
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "initial height"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "initial offshore gradient"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "cell length [m]"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "number of cells"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "apply changes"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   7320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2625
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   4320
      Width           =   5775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f As Integer
Dim topo_temp_GUI() As Single
Dim dy_onshore  As Double
Dim Y, X, yy As Double
Dim l, l1 As Double


Private Sub Check2_Click()
If Form5.Check2.Value = 1 Then
    Form5.List1.Enabled = True
    f = Form5.List1.ListIndex
    If f = 0 Then
        nickpoint = 150
    ElseIf f = 1 Then
        nickpoint = 160
    ElseIf f = 2 Then
        nickpoint = 170
    ElseIf f = 3 Then
        nickpoint = 180
    ElseIf f = 4 Then
        nickpoint = 190
    ElseIf f = 5 Then
        nickpoint = 200
    End If
Else
    nickpoint = 1
    Form5.List1.Enabled = False
End If
If Form5_Show = 1 Then
    update_picture1_form5
End If
End Sub

Private Sub Command1_Click()
    'Module1.Parameter_and_variabeles
    Form5.Hide
End Sub

Private Sub Command2_Click()
    Module1.Parameter_and_variabeles
    Form5.Hide
End Sub

Private Sub List1_Click()
f = Form5.List1.ListIndex
If f = 0 Then
    nickpoint = 150
ElseIf f = 1 Then
    nickpoint = 160
ElseIf f = 2 Then
    nickpoint = 170
ElseIf f = 3 Then
    nickpoint = 180
ElseIf f = 4 Then
    nickpoint = 190
ElseIf f = 5 Then
    nickpoint = 200
End If
If Form5_Show = 1 Then
    update_picture1_form5
End If

End Sub

Private Sub List2_Click()
f = Form5.List2.ListIndex
If f = 0 Then
   dy_onshore = 0.01
ElseIf f = 1 Then
    dy_onshore = 0.02
ElseIf f = 2 Then
    dy_onshore = 0.03
ElseIf f = 3 Then
    dy_onshore = 0.05

End If
If Form5_Show = 1 Then
    update_picture1_form5
End If
End Sub

Private Sub List4_Click()
f = Form5.List4.ListIndex
If f = 0 Then
   initial_gradient = 0.1
ElseIf f = 1 Then
    initial_gradient = 0.15
ElseIf f = 2 Then
    initial_gradient = 0.2
ElseIf f = 3 Then
    initial_gradient = 0.25
ElseIf f = 4 Then
    initial_gradient = 0.3
End If
If Form5_Show = 1 Then
    update_picture1_form5
End If

End Sub

Sub update_picture1_form5()




ReDim topo_temp_GUI(0 To nx)                     '//topographical height at start of profile

       topo_temp_GUI(0) = initial_height
        
       'dy_onshore = 0.05
       dy = initial_gradient                            '// initial gradient in m/dx
               
       i = 0
       For i = 0 To nickpoint 'nx - 1 'nx = 500
        
            topo_temp_GUI(i + 1) = (topo_temp_GUI(i) - (dy_onshore)) '+ noise
         
       Next i
       
       For i = nickpoint + 1 To nx - 1
            
            topo_temp_GUI(i + 1) = (topo_temp_GUI(i) - (dy)) '+ noise
            
       Next i
       i = 0

X = nx
Y = initial_height + 5
yy = 0

Form5.Picture1.Cls
Form5.Picture1.Scale (-1, Y + 1)-(X + 1, yy - 1)

Form5.Picture1.Line (0, yy - 1)-(0, Y + 1)
Form5.Picture1.Line (0, yy)-(X, yy)

For i = 1 To nx - 1
    l = topo_temp_GUI(i)
    l1 = topo_temp_GUI(i + 1)
    Form5.Picture1.Line (i, l)-(i + 1, l1)
Next i
       
End Sub

Private Sub List6_Click()
f = Form5.List6.ListIndex
If f = 0 Then
   nx = 500
ElseIf f = 1 Then
    nx = 1000
ElseIf f = 2 Then
    nx = 1500
End If
    
    If Form5_Show = 1 Then
    update_picture1_form5
    End If
End Sub

Private Sub Option1_Click()
GRID_Option = 2
Form5.List3(1).Enabled = False
Form5.List4.Enabled = False
Form5.List5.Enabled = False
Form5.List6.Enabled = False
Form5.List1.Enabled = False
Form5.List2.Enabled = False
Form5.Check2.Enabled = False

  
  update_picture1_form5_option2

End Sub
Private Sub update_picture1_form5_option2()
X = nx
Y = initial_height + 5
yy = 0

Form5.Picture1.Cls
Form5.Picture1.Scale (-1, Y + 1)-(X + 1, yy - 1)

Form5.Picture1.Line (0, yy - 1)-(0, Y + 1)
Form5.Picture1.Line (0, yy)-(X, yy)

For i = 1 To nx - 1
    l = Module3.load_EQ_profile(i)
    l1 = Module3.load_EQ_profile(i + 1)
    Form5.Picture1.Line (i, l)-(i + 1, l1)
Next i
End Sub


Private Sub Option2_Click()
GRID_Option = 1

Form5.List3(1).Enabled = True
Form5.List4.Enabled = True
Form5.List5.Enabled = True
Form5.List6.Enabled = True
Form5.List1.Enabled = True
Form5.List2.Enabled = True
Form5.Check2.Enabled = True

  update_picture1_form5
  
End Sub
