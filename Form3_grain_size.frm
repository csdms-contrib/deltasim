VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "input control"
   ClientHeight    =   9960
   ClientLeft      =   10305
   ClientTop       =   1980
   ClientWidth     =   5370
   Icon            =   "Form3_grain_size.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   5370
   Begin VB.CommandButton Command6 
      Caption         =   "Kura "
      Height          =   495
      Left            =   3840
      TabIndex        =   31
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sand"
      Height          =   495
      Left            =   2640
      TabIndex        =   30
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Silt"
      Height          =   495
      Left            =   1440
      TabIndex        =   29
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clay"
      Height          =   495
      Left            =   240
      TabIndex        =   28
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "restore default values"
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   9360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply changes"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   9360
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   2505
      ScaleWidth      =   4785
      TabIndex        =   19
      Top             =   6000
      Width           =   4815
   End
   Begin VB.Frame Frame4 
      Caption         =   "GRAIN SIZE CLASSES in mm"
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.HScrollBar HScroll6 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   27
         Top             =   5040
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   26
         Top             =   4200
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   25
         Top             =   3360
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   24
         Top             =   2520
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   23
         Top             =   1680
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   22
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text14"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text13"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text12"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text11"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text10"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text9 
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text9"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "0.5"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "0.35"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "0.23"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "0.177"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "0.088"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "0.0044"
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "grain size 6"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "grain size 5"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "grain size 4"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "grain size 3"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "grain size 2"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "grain size 1"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub update_a_picture_1()

Form3.Picture1.Cls
Form3.Picture1.Scale (0, 1.2)-(8, -0.2)
Form3.Picture1.PSet (2, 1.1)
Form3.Picture1.Print "graph of grain size distribution"
Form3.Picture1.Line (1, 0)-(1, 1)
Form3.Picture1.Line (1, 0)-(7, 0)
Form3.Picture1.PSet (0.75, 0.1)
Form3.Picture1.Print "0"
Form3.Picture1.PSet (0.3, 1)
Form3.Picture1.Print "100%"

For n = 1 To 6
    Form3.Picture1.Line ((n), sed_cont_pct(n))-((n) + 0.9, 0), RGB(255, 0, 0), BF
Next n
Form3.Picture1.Refresh

End Sub
Function formtotal()

   total = Form3.HScroll1.Value
   total = total + Form3.HScroll2.Value
   total = total + Form3.HScroll3.Value
   total = total + Form3.HScroll4.Value
   total = total + Form3.HScroll5.Value
   total = total + Form3.HScroll6.Value

formtotal = total
End Function
Sub update_text()
total = formtotal

If total <= 0 Then
    Module1.Parameter_and_variabeles
    total = 10
End If

sed_cont_pct(1) = (Form3.HScroll1.Value / total) '/ 10
sed_cont_pct(2) = (Form3.HScroll2.Value / total) '/ 10
sed_cont_pct(3) = (Form3.HScroll3.Value / total) '/ 10
sed_cont_pct(4) = (Form3.HScroll4.Value / total) '/ 10
sed_cont_pct(5) = (Form3.HScroll5.Value / total) '/ 10
sed_cont_pct(6) = (Form3.HScroll6.Value / total) '/ 10

Form3.Text9 = FormatNumber(sed_cont_pct(1), 2)
Form3.Text10 = FormatNumber(sed_cont_pct(2), 2)
Form3.Text11 = FormatNumber(sed_cont_pct(3), 2)
Form3.Text12 = FormatNumber(sed_cont_pct(4), 2)
Form3.Text13 = FormatNumber(sed_cont_pct(5), 2)
Form3.Text14 = FormatNumber(sed_cont_pct(6), 2)

'Form3.Text9.Refresh
update_a_picture_1
End Sub

Private Sub Command1_Click()
Form3.HScroll1.Value = 0.6 * 10
Form3.HScroll2.Value = 0.5 * 10
Form3.HScroll3.Value = 0.4 * 10
Form3.HScroll4.Value = 0.3 * 10
Form3.HScroll5.Value = 0.2 * 10
Form3.HScroll6.Value = 0.1 * 10
End Sub

Private Sub Command2_Click()
Form3.Hide
End Sub

Private Sub Command3_Click()
Module1.Parameter_and_variabeles
Form3.Hide
End Sub

Private Sub Command4_Click()
Form3.HScroll1.Value = 0.2 * 10
Form3.HScroll2.Value = 0.3 * 10
Form3.HScroll3.Value = 0.4 * 10
Form3.HScroll4.Value = 0.4 * 10
Form3.HScroll5.Value = 0.3 * 10
Form3.HScroll6.Value = 0.2 * 10
End Sub
Private Sub Command5_Click()
Form3.HScroll1.Value = 0.1 * 10
Form3.HScroll2.Value = 0.2 * 10
Form3.HScroll3.Value = 0.3 * 10
Form3.HScroll4.Value = 0.4 * 10
Form3.HScroll5.Value = 0.5 * 10
Form3.HScroll6.Value = 0.6 * 10
End Sub

Private Sub Command6_Click()
Form3.HScroll1.Value = 0.5 * 10
Form3.HScroll2.Value = 0.6 * 10
Form3.HScroll3.Value = 0.5 * 10
Form3.HScroll4.Value = 0.4 * 10
Form3.HScroll5.Value = 0.3 * 10
Form3.HScroll6.Value = 0.3 * 10
End Sub

Private Sub HScroll1_Change()
update_text

End Sub

Private Sub HScroll2_Change()
update_text

End Sub

Private Sub HScroll3_Change()
update_text

End Sub

Private Sub HScroll4_Change()
update_text

End Sub

Private Sub HScroll5_Change()
update_text

End Sub

Private Sub HScroll6_Change()
update_text

End Sub









