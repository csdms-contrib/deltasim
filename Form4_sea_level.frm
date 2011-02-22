VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEA LEVEL CONTROLS"
   ClientHeight    =   8910
   ClientLeft      =   8490
   ClientTop       =   2310
   ClientWidth     =   7965
   Icon            =   "Form4_sea_level.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   7965
   Begin VB.OptionButton Option7 
      Caption         =   "Sea level fall accelerated"
      Height          =   495
      Left            =   360
      TabIndex        =   22
      Top             =   4920
      Width           =   3015
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Sea level fall linear"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Text            =   "5"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Text            =   "5"
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Sea level rise accelerated"
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   4320
      Width           =   2775
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Sea level rise linear"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   2880
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   840
      ScaleHeight     =   1665
      ScaleWidth      =   6705
      TabIndex        =   9
      Top             =   6240
      Width           =   6735
   End
   Begin VB.TextBox Text17 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "5"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text16 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "200"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text15 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Oscilating Sea Level  "
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Steady Sea level "
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      Caption         =   $"Form4_sea_level.frx":0442
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label9 
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "value between -20 and 20"
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Label6"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   7920
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "Amplitude 1-50"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label21 
      Caption         =   "Frequency 10-10.000 yr"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.14159265359
Dim Y As Double
Dim z As Double
Dim l, l1 As Double
Dim X As Double
Sub update_picture1_option1()
X = end_of_times
Y = -18
yy = -40
Form4.Picture1.Cls
Form4.Picture1.Scale (-1, Y + 1.2)-(X + 1, yy - 1.2)

Form4.Label6.Caption = "Caspina Sea level graph for the last " & X & " years"
Form4.Label1.Caption = Y
Form4.Label2.Caption = ""
Form4.Label3.Caption = yy

Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"

Form4.Picture1.Line (0, yy - 1)-(0, Y + 1)
Form4.Picture1.Line (0, yy)-(X, yy)

Module4.Load_SL_Time
Module4.load_sl_value

For k = 0 To end_of_times
    
    l = Module3.CSL(k)
    l1 = Module3.CSL(k + 1)
    Form4.Picture1.Line (k, l)-(k + 1, l1), RGB(0, 0, 255)
Next k

End Sub
Sub update_picture1_option3()


    rsl = Form4.Text15
    Y = Form4.Text17
    X = end_of_times
    z = Form4.Text16


Form4.Picture1.Cls
Form4.Picture1.Scale (-1, Y + 1.2)-(X + 1, -Y - 1.2)


Form4.Label6.Caption = "Oscilating Sea level graph "


Form4.Picture1.Line (0, -Y - 1)-(0, Y + 1)
Form4.Picture1.Line (0, 0)-(X, 0)

Form4.Label1.Caption = rsl + Y
Form4.Label2.Caption = rsl
Form4.Label3.Caption = rsl - Y
Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"



For k = 0 To end_of_times
         
        l = Y * Sin((k * (PI / (0.5 * z)) + (0.5 * PI)))
        l1 = Y * Sin(((k + 1) * (PI / (0.5 * z)) + (0.5 * PI)))
        Form4.Picture1.Line (k, l)-(k + 1, l1), RGB(0, 0, 255)
       
Next k
       

Form4.Picture1.Refresh
End Sub
Sub update_picture1_option5()


    rsl = Form4.Text15
    Y = Form4.Text2
    X = end_of_times
    z = end_of_times * 4


Form4.Picture1.Cls
Form4.Picture1.Scale (-1, (Y) + 1.2)-(X + 1, rsl - 1.2)


Form4.Label6.Caption = "Accelerated Sea level rise graph "


Form4.Picture1.Line (0, rsl - 1)-(0, (Y) + 1)
Form4.Picture1.Line (0, 0)-(X, 0)

Form4.Label1.Caption = Y
Form4.Label2.Caption = ""
Form4.Label3.Caption = rsl
Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"


Y = Y
For k = 0 To end_of_times
         
        l = Y + (Y * Sin((k * (PI / (0.5 * z)) - (0.5 * PI))))
        l1 = Y + (Y * Sin(((k + 1) * (PI / (0.5 * z)) - (0.5 * PI))))
        Form4.Picture1.Line (k, l)-(k + 1, l1), RGB(0, 0, 255)
       
Next k
       

Form4.Picture1.Refresh
End Sub
Sub update_picture1_option7()
    rsl = Form4.Text15
    Y = Form4.Text2 * -1
    X = end_of_times
    z = end_of_times * 4


Form4.Picture1.Cls
Form4.Picture1.Scale (0.1, rsl + 1.2)-(X + 1, (Y) - 1.2)


Form4.Label6.Caption = "Accelerated Sea level fall graph "


Form4.Picture1.Line (0, rsl - 1)-(0, (Y) + 1)
Form4.Picture1.Line (0, 0)-(X, 0)

Form4.Label1.Caption = rsl
Form4.Label2.Caption = ""
Form4.Label3.Caption = Y
Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"


Y = Y
For k = 0 To end_of_times
         
        l = Y + (Y * Sin((k * (PI / (0.5 * z)) - (0.5 * PI))))
        l1 = Y + (Y * Sin(((k + 1) * (PI / (0.5 * z)) - (0.5 * PI))))
        Form4.Picture1.Line (k, l)-(k + 1, l1), RGB(0, 0, 255)
       
Next k
       

Form4.Picture1.Refresh
End Sub

Sub update_picture1_option2()

Y = Form4.Text15
X = end_of_times

Form4.Picture1.Cls
Form4.Picture1.Scale (-1, Y + 1.2)-(X + 1, Y - 1.2)

Form4.Picture1.PSet (X / 2.6, Y - 0.6)
Form4.Label6.Caption = " Steady Sea level graph "

Form4.Picture1.Line (0, Y)-(X, Y), RGB(0, 0, 255), BF
Form4.Picture1.Line (0, Y - 1)-(0, Y + 1)
Form4.Picture1.Line (0, Y - 1)-(X, Y - 1)

Form4.Label1.Caption = ""
Form4.Label2.Caption = Y
Form4.Label3.Caption = ""
Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"


Form4.Picture1.Refresh
End Sub
Sub update_picture1_option6()
Y = Form4.Text1 * -1
X = end_of_times

Form4.Picture1.Cls
Form4.Picture1.Scale (-0.25, 0.25)-(X + 1, Y - 0.25)

Form4.Picture1.PSet (X / 2.6, Y + 0.6)
Form4.Label6.Caption = " linear Sea level fall graph "

Form4.Picture1.Line (0, 0)-(X, Y), RGB(0, 0, 255) ', BF
Form4.Picture1.Line (0, 0)-(0, Y)
Form4.Picture1.Line (0, 0)-(X, 0)

Form4.Label1.Caption = 0
Form4.Label2.Caption = ""
Form4.Label3.Caption = Y
Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"


Form4.Picture1.Refresh
End Sub
Sub update_picture1_option4()

Y = Form4.Text1
X = end_of_times

Form4.Picture1.Cls
Form4.Picture1.Scale (-0.25, Y + 0.25)-(X + 1, -0.25)

Form4.Picture1.PSet (X / 2.6, Y - 0.6)
Form4.Label6.Caption = " linear Sea level rise graph "

Form4.Picture1.Line (0, 0)-(X, Y), RGB(0, 0, 255) ', BF
Form4.Picture1.Line (0, 0)-(0, Y + 1)
Form4.Picture1.Line (0, 0)-(X, 0)

Form4.Label1.Caption = Y
Form4.Label2.Caption = ""
Form4.Label3.Caption = 0
Form4.Label4.Caption = 0
Form4.Label5.Caption = end_of_times & " years"


Form4.Picture1.Refresh
End Sub


Private Sub Command1_Click()
Form4.Hide
End Sub



Private Sub Option1_Click()
Form4.Text1.Locked = True
Form4.Text2.Locked = True
Form4.Text15.Locked = True
Form4.Text16.Locked = True
Form4.Text17.Locked = True
SLoption = 1

update_picture1_option1

End Sub

Private Sub Option2_Click()
Dim Y As Double
Dim X As Double
Form4.Text1.Locked = True
Form4.Text2.Locked = True
Form4.Text15.Locked = False
Form4.Text16.Locked = True
Form4.Text17.Locked = True
SLoption = 2

update_picture1_option2



End Sub

Private Sub Option3_Click()
Form4.Text1.Locked = True
Form4.Text2.Locked = True
Form4.Text15.Locked = True
Form4.Text16.Locked = False
Form4.Text17.Locked = False
SLoption = 3

update_picture1_option3


End Sub


Private Sub Option4_Click()
Form4.Text1.Locked = False
Form4.Text2.Locked = True
Form4.Text15.Locked = True
Form4.Text16.Locked = True
Form4.Text17.Locked = True
SLoption = 4
Form4.Label8.Caption = "positive"
update_picture1_option4

End Sub

Private Sub Option5_Click()
Form4.Text1.Locked = True
Form4.Text2.Locked = False
Form4.Text15.Locked = True
Form4.Text16.Locked = True
Form4.Text17.Locked = True
SLoption = 5
Form4.Label9.Caption = "positive"
update_picture1_option5
End Sub

Private Sub Option6_Click()
Form4.Text1.Locked = False
Form4.Text2.Locked = True
Form4.Text15.Locked = True
Form4.Text16.Locked = True
Form4.Text17.Locked = True
SLoption = 6
Form4.Label8.Caption = "negative"
update_picture1_option6
End Sub

Private Sub Option7_Click()
Form4.Text1.Locked = True
Form4.Text2.Locked = False
Form4.Text15.Locked = True
Form4.Text16.Locked = True
Form4.Text17.Locked = True
SLoption = 7
Form4.Label9.Caption = "negative"
update_picture1_option7
End Sub

Private Sub Text1_Change()
sealevel = Format$(Form4.Text1, 0)

If sealevel < 1 Then        'Or sealevel = ""
    Form4.Text1 = 1
ElseIf sealevel > 200 Then
    Form4.Text1 = 200
End If

If Option4.Value = True Then
update_picture1_option4
Else
update_picture1_option6
End If

End Sub

Private Sub Text15_Change()

sealevel = Format$(Form4.Text15, 0)

If sealevel < -20 Then        'Or sealevel = ""
    Form4.Text15 = -20
ElseIf sealevel > 120 Then
    Form4.Text15 = 120
End If

'sealevel = sealevel + 80
update_picture1_option2

End Sub

Private Sub Text16_Change()
sealevel_frequency = Format$(Form4.Text16, 0)

If sealevel_frequency < 20 Or sealevel_frequency = "" Then
   Form4.Text16 = 20
ElseIf sealevel_frequency > 20000 Then
    Form4.Text16 = 20000
End If



update_picture1_option3

End Sub

Private Sub Text17_Change()
sealevel_amplitude = Format$(Form4.Text17, 0)

If sealevel_amplitude < 1 Or sealevel_amplitude = "" Or sealevel_amplitude = " " Then
    Form4.Text17 = 1
ElseIf sealevel_amplitude > 15 Then
    Form4.Text17 = 15
End If
update_picture1_option3

End Sub

Private Sub Text2_Change()
sealevel = Format$(Form4.Text2, 0)

If sealevel < 1 Then        'Or sealevel = ""
    Form4.Text2 = 1
ElseIf sealevel > 200 Then
    Form4.Text2 = 200
End If
If Option5.Value = True Then
update_picture1_option5
Else
update_picture1_option7
End If

End Sub
