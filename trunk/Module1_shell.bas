Attribute VB_Name = "Module1"
Option Explicit

'''''''''''''''''''''''''''''''Constanten''''''''''''''''''''''''''''''''''''''''
Public Const PI = 3.14159265359
Public Const EPS = 0.000001

Public Const well1 = 175
Public Const well2 = 225
Public Const well3 = 275
'''''''''''''''''''''''''''''''AUX parameters''''''''''''''''''''''''''''''''''''
Public j As Integer
Public h As Integer
Public z As Integer
Public f As Integer
Public p As Integer
Public i As Integer 'gridcell altego
Public n As Integer 'grainsize altego
Public k As Integer 'sim_time altego

Public equal, Value As Variant
Public fcount, inpcount  As Integer
Public run_no As Integer
Public well_counter As Integer
Public Xpos As Integer
Public inteN As Single 'Public i As Single
Public hueD As Single 'Dim h As Single
Public r1, g1, b1 As Single
Public temp1, temp2, temp3R, temp3G, temp3B, s, l As Single
'Public Xpos As Integer
 
Public e_datamaxx As Single
Public e_datamax As Single
Public e_dataminn As Single
Public e_datamin As Single
Public tmpsum As Single
Public Hue, R, G, b, min_dh, max_dh As Single
Public Form5_Show As Integer
Public delta_grad(0 To 4) As Single
Public deltaGRAD_sum, deltaGRAD_sum2 As Single
Public DELTAGRAD_AV_old, deltagrad_AV, DELTAGRAD_AV_old2, deltagrad_AV2 As Single
Public kk As Single
Public FT As Integer
Public i_flx As Integer

Public tot_onshore, onshore As Single
Public offshore, tot_offshore As Single
'''''''''''''''''''''''''''''''GRID Varia''''''''''''''''''''''''''''''''''''''''
Public nickpoint As Integer
Public dy_onshore As Single
Public nx As Integer
Public initial_height As Single
Public initial_gradient As Single
Public dy As Single
Public flx, flx_old As Single
Public dx As Single
Public GRID_Option As Integer

'''''''''''''''''''''''''''''''SEA LEVEL Varia'''''''''''''''''''''''''''''''''''
Public SLC As Single
Public SLoption As Integer
Public sealevel_amplitude As Variant
Public sealevel_frequency As Variant
Public model_sealevel, model_sealevel_old As Single
Public sealevel As Single
Public SL_Time(1 To 101) As Single
Public SL_Value(1 To 101) As Single
Public SL_nrvalue As Integer
Public SLmaxx, SLmax, SLminn, SLmin As Long
Public RICO As Single

'''''''''''''''''''''''''''''''GRAIN SIZE Varia''''''''''''''''''''''''''''''''''
Public tot As Single
Public temp As Integer
Public num_of_gscl As Integer

'''''''''''''''''''''''''''''''DISCHARGE Varia'''''''''''''''''''''''''''''''''''
Public corr1, Q_volatility  As Single
Public LNprop, Q_average As Single
Public M_tmp, M_tmp2 As Single
Public av_dis As Single
Public som_dis As Single
Public riverindex As Single
'''''''''''''''''''''''''''''''SCREEN OUTPUT Varia'''''''''''''''''''''''''''''''
Public wheeler_increment, cum_time_wheeler As Single
Public wheeler_tot_thick() As Single

Public FILE_NAME10 As String
Public FILE_NAME11 As String
Public FILE_NAME12 As String

Public FILE_NAME13 As String
Public FILE_NAME14 As String

Public FILE_NAME22 As String


Public FILE_NAME1, FILE_NAME2, FILE_NAME3, FILE_NAME4 As String

Public xstrtps As Single
Public xstopps As Single
Public timeln_intrval As Single
Public Rmarg As Integer
Public Tmarg As Integer
Public Bmarg As Integer
Public Lmarg As Integer
Public datamaxx As Single
Public datamax As Single
Public dataminn As Single
Public datamin As Single
Public ddata As Single
Public huedata As Single

Public datamindraw As Single
Public hhmax As Single
Public hmax As Single
Public hhmin As Single
Public hmin As Single
Public scy As Single
Public scx As Single
Public xcolortik As Single
Public ycolortik As Single
Public xcolortik1 As Single
Public ycolortik1 As Single
Public xtik As Single
Public ytik As Single
Public Iteration As Single
Public m As Single
Public years_for_output, num_outp, wellpos1 As Variant

'''''''''''''''''''''''''''''''MAIN CALC varia''''''''''''''''''''''''''''''''''''
Public dh_depo() As Single
Public av_depo_rate(0 To 500) As Single
Public Medianprint As String

Public erosion_power As Single
Public streampower As Single
Public out_coast_distance As Single
Public stream_velocity As Single

Public percentage As Single

Public sedimentload_factor As Single
Public init_load As Single
Public initload As Single

'Public max_dh As Single
Public maxx_dh As Single

Public traveldist As Single
Public max_dh_gs As Single
Public settle_rate As Double

Public discharge_shapef As Single
Public wavebase_cell As Single
Public k_er_fluv, k_er_fluvC, k_er_marine, spreading_angle As Single

'''''''''''''''''''''''''''''''TIME Varia''''''''''''''''''''''''''''''''''''''''''
Public PauseTime, Start, Finish, TotalTime As Variant
Public sim_time As Integer
Public MCrun As Integer
Public end_MCrun As Integer
Public end_of_times As Integer
Public dt As Integer

'''''''''''''''''''''''''''''''STATIC ARRAY''''''''''''''''''''''''''''''''''''''''
Public sed_cont_pct(1 To 6) As Variant
Public grain_size(1 To 6) As Single
Public traveldist_fluvial(1 To 6) As Single
Public traveldist_marine(1 To 6) As Single

'''''''''''''''''''''''''''''''DYNAMIC ARRAY''''''''''''''''''''''''''''''''''''''
Public sediment() As Single
Public sedflux() As Single
Public SLinput() As Single
Public discharge() As Single
Public waterdepth() As Single
Public erosion_rate() As Single
Public PS_data() As Variant
Public thickness() As Single
Public StratNode() As Single
Public topo_old() As Single
Public topo_temp() As Single
Public topo_calc() As Single
Public topo_depo() As Single
Public dh_erosion() As Single
Public settle_rate_wheeler() As Single
Public slope() As Single
Public Median() As Single

Public Prob_Sand() As Single
Public Prob_Silt() As Single
Public Prob_Clay() As Single

Public sum_sediment() As Single
Public av_sediment() As Single
Public D(0 To 1000) As Single



''''''''''''''''''''''''''''
Public Sub Shell1()
    
    MCrun = 1
    end_MCrun = 2500
    
    FILE_NAME11 = "D:\Temp\vba_tl.ps"
    FILE_NAME12 = "D:\Temp\alpha.dat"
    
    FILE_NAME13 = "D:\Temp\erosie.dat"
    FILE_NAME14 = "D:\Temp\depositie.dat"
    
    
    Open FILE_NAME11 For Output As #11
    Open FILE_NAME12 For Output As #12
    
    Open FILE_NAME13 For Output As #13
    Open FILE_NAME14 For Output As #14
    
    
    
    Module3.clearpictures_Form1
    
    Module2.Main
    
    Form1.mnuWritePS.Enabled = True
  ' Form1.mnuDeltaGrad = True
    Module3.show_runtime_info
    
    Close #14
    Close #13
    Close #12
    Close #11
    
 End Sub

Public Sub Parameter_and_constants_default()

dx = 1000
nx = 500



initial_height = 115
initial_gradient = 0.2

num_of_gscl = 6
sed_cont_pct(1) = 0.25
sed_cont_pct(2) = 0.2
sed_cont_pct(3) = 0.2
sed_cont_pct(4) = 0.15
sed_cont_pct(5) = 0.1
sed_cont_pct(6) = 0.1

'FILE_NAME1 = "layer.GRD"
'FILE_NAME2 = "deposition.GRD"
'FILE_NAME3 = "median.GRD"
'FILE_NAME4 = "thickness.GRD"

end_of_times = 3200
dt = 1



'output_counter = 0
'n_cli = 1

num_of_gscl = 6

k_er_fluv = 0.0005
k_er_marine = 0.0000005

spreading_angle = 20
years_for_output = 1

num_outp = 1

wellpos1 = 260

grain_size(1) = 0.0044
grain_size(2) = 0.088
grain_size(3) = 0.177
grain_size(4) = 0.23
grain_size(5) = 0.35
grain_size(6) = 0.5

traveldist_fluvial(1) = 65000
traveldist_fluvial(2) = 28000
traveldist_fluvial(3) = 17000
traveldist_fluvial(4) = 13000
traveldist_fluvial(5) = 11000
traveldist_fluvial(6) = 10000

traveldist_marine(1) = 10500
traveldist_marine(2) = 4500
traveldist_marine(3) = 2700
traveldist_marine(4) = 2250
traveldist_marine(5) = 1800
traveldist_marine(6) = 500

' 0
fcount = 0
inpcount = 1
stream_velocity = 0

Q_volatility = 0.2
Q_average = 7000

'Open FILE_NAME1 For Output As #1
'Open FILE_NAME2 For Output As #2

'Open FILE_NAME3 For Output As #3
'Open FILE_NAME4 For Output As #4

End Sub
Public Sub Parameter_and_variabeles()




Form3.HScroll1.Value = 0.25 * 10
Form3.HScroll2.Value = 0.2 * 10
Form3.HScroll3.Value = 0.2 * 10
Form3.HScroll4.Value = 0.15 * 10
Form3.HScroll5.Value = 0.1 * 10
Form3.HScroll6.Value = 0.1 * 10

Form3.Text9 = Form3.HScroll1.Value / 10
Form3.Text10 = Form3.HScroll2.Value / 10
Form3.Text11 = Form3.HScroll3.Value / 10
Form3.Text12 = Form3.HScroll4.Value / 10
Form3.Text13 = Form3.HScroll5.Value / 10
Form3.Text14 = Form3.HScroll6.Value / 10

Form4.Option2.Value = True


Form5.Check2.Value = 1
Form5.List1.ListIndex = 1
Form5.List2.ListIndex = 2
Form5.List4.ListIndex = 1
Form5.List6.ListIndex = 0
Form5.Option2.Value = True

Form6.List1.ListIndex = 7

Form7.List1.ListIndex = 1
Form7.List2.ListIndex = 0

Form8.List2.ListIndex = 5
Form8.List3.ListIndex = 0

Form9.Text1 = 6500
Form9.Text2 = 2800
Form9.Text3 = 1700
Form9.Text4 = 1300
Form9.Text5 = 1100
Form9.Text6 = 1000

Form9.Text7 = 1000
Form9.Text8 = 450
Form9.Text9 = 270
Form9.Text10 = 225
Form9.Text11 = 180
Form9.Text12 = 50

Form9.HScroll1.Value = 5
Form9.HScroll2.Value = 5
Form9.HScroll3.Value = 5
Form9.HScroll4.Value = 5
Form9.HScroll5.Value = 5
Form9.HScroll6.Value = 5

Form9.HScroll7.Value = 5
Form9.HScroll8.Value = 5
Form9.HScroll9.Value = 5
Form9.HScroll10.Value = 5
Form9.HScroll11.Value = 5
Form9.HScroll12.Value = 5

run_no = run_no + 1

End Sub


Public Sub Parameter_and_constants()

If run_no = 0 Then
    Parameter_and_variabeles
End If


f = Form7.List1.ListIndex
If f = 0 Then
    end_of_times = 100
ElseIf f = 1 Then
    end_of_times = 3200
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


dt = Form7.List2.List(0)
initial_height = Form5.List3(1).List(0)

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

dx = Form5.List5.List(0)

f = Form5.List6.ListIndex
If f = 0 Then
    nx = 500
ElseIf f = 1 Then
    nx = 1000
ElseIf f = 2 Then
    nx = 1500
End If


xstrtps = 10          '// usually a initial effect exist which you do not want to display
xstopps = nx - 10
timeln_intrval = 20
Rmarg = 36
Tmarg = 10
Bmarg = 50
Lmarg = 36


If Form5.Check2.Value = 1 Then

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
End If

Q_average = Form6.Text7.Text
f = Form6.List1.ListIndex
If f = 0 Then
    Q_volatility = 0.1
ElseIf f = 1 Then
    Q_volatility = 0.2
ElseIf f = 2 Then
    Q_volatility = 0.3
ElseIf f = 3 Then
   Q_volatility = 0.4
ElseIf f = 4 Then
   Q_volatility = 0.5
ElseIf f = 5 Then
   Q_volatility = 0.6
ElseIf f = 6 Then
   Q_volatility = 0.7
ElseIf f = 7 Then
   Q_volatility = 0
End If

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

sealevel_amplitude = Form4.Text17
sealevel_frequency = Form4.Text16
sealevel = Form4.Text15

sed_cont_pct(1) = FormatNumber(Form3.Text9, 4)
sed_cont_pct(2) = FormatNumber(Form3.Text10, 4)
sed_cont_pct(3) = FormatNumber(Form3.Text11, 4)
sed_cont_pct(4) = FormatNumber(Form3.Text12, 4)
sed_cont_pct(5) = FormatNumber(Form3.Text13, 4)
sed_cont_pct(6) = FormatNumber(Form3.Text14, 4)

traveldist_fluvial(1) = FormatNumber(Form9.Text1, 0)
traveldist_fluvial(2) = FormatNumber(Form9.Text2, 0)
traveldist_fluvial(3) = FormatNumber(Form9.Text3, 0)
traveldist_fluvial(4) = FormatNumber(Form9.Text4, 0)
traveldist_fluvial(5) = FormatNumber(Form9.Text5, 0)
traveldist_fluvial(6) = FormatNumber(Form9.Text6, 0)

traveldist_marine(1) = FormatNumber(Form9.Text7, 0)
traveldist_marine(2) = FormatNumber(Form9.Text8, 0)
traveldist_marine(3) = FormatNumber(Form9.Text9, 0)
traveldist_marine(4) = FormatNumber(Form9.Text10, 0)
traveldist_marine(5) = FormatNumber(Form9.Text11, 0)
traveldist_marine(6) = FormatNumber(Form9.Text12, 0)

num_of_gscl = 6

FILE_NAME1 = "layer.GRD"
FILE_NAME2 = "deposition.GRD"
FILE_NAME3 = "median.GRD"
FILE_NAME4 = "thickness.GRD"

num_of_gscl = 6

k_er_marine = 0.0000005
spreading_angle = 20
years_for_output = 1
num_outp = 1
wellpos1 = 260
SLC = 0


grain_size(1) = 0.0044
grain_size(2) = 0.088
grain_size(3) = 0.177
grain_size(4) = 0.23
grain_size(5) = 0.35
grain_size(6) = 0.5

fcount = 0
inpcount = 1
stream_velocity = 0
riverindex = 200
sim_time = 0



End Sub

