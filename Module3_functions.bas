Attribute VB_Name = "Module3"
Public Sub clearpictures_Form1()

    Form2.Pic2.Cls
    Form1.pic1.Cls
    Form1.Picture1.Cls
    Form2.Pic2.Refresh
    Form1.pic1.Refresh
    Form1.Picture1.Refresh
    Form1.mnuWell.Enabled = True
    Form1.Picture1.Cls
    Form1.Picture2.Cls
    Form1.Picture3.Cls
    Form10.Picture4.Cls
    Form1.pic1.Cls
    Form1.pic1.Scale (0, 280)-(700, 0)
    Form1.Picture1.Cls
    
    Form1.Picture3.Cls
    Form10.Picture4.Cls
    
    kk = 0
    delta_grad(1) = 0
    delta_grad(2) = 0
    deltaGRAD_sum = 0
    DELTAGRAD_AV_old = 0
    deltagrad_AV = 0
    deltaGRAD_sum2 = 0
    DELTAGRAD_AV_old2 = 0
    deltagrad_AV2 = 0
    
    PauseTime = 10    ' Set duration.
    Start = Timer    ' Set start time.
    Form2.Pic2.Print end_of_times; " year simulation started at " & Time()
    
End Sub

Public Sub show_runtime_info()
    
    ReDim sum_sediment(0 To nx, 1 To num_of_gscl)
    ReDim av_sediment(0 To nx, 1 To num_of_gscl)
    
    Load Form2
    Form2.Show
    Form2.Refresh
    Form2.Pic2.Print " "
    Form2.Pic2.Print "simulation finished at " & Time()
    Form2.Pic2.Print " "
    Finish = Timer    ' Set end time.
    TotalTime = Finish - Start    ' Calculate total time.
    Form2.Pic2.Print "Calculations took " & FormatNumber(TotalTime, 2) & " seconds (equals " & FormatNumber((TotalTime / 60), 1) & " minutes)"
    Form2.Pic2.Print FormatNumber(TotalTime / end_of_times, 3) & " seconds for one simulation year"
    Form2.Pic2.Print " "
   ' Form2.Pic2.Print FormatNumber(deltagrad_AV / 100, 5) & " average delta gradient for this run " '& "(" & FormatNumber(deltagrad_AV / 100, 4, vbTrue, vbTrue, vbTrue); ")"
   ' Form2.Pic2.Print " "
   ' Form2.Pic2.Print FormatNumber(sediment(end_of_times - 1, flx + 1, 1), 12)
   ' Form2.Pic2.Print FormatNumber(sediment(end_of_times - 1, flx + 1, 2), 12)
   ' Form2.Pic2.Print FormatNumber(sediment(end_of_times - 1, flx + 1, 3), 12)
   ' Form2.Pic2.Print FormatNumber(sediment(end_of_times - 1, flx + 1, 4), 12)
   ' Form2.Pic2.Print FormatNumber(sediment(end_of_times - 1, flx + 1, 5), 12)
   ' Form2.Pic2.Print FormatNumber(sediment(end_of_times - 1, flx + 1, 6), 12)
    
    
    
    'For i = 0 To nx
    'For n = 1 To num_of_gscl
        
        'For k = 1 To end_of_times - 1
        
         'sum_sediment(i, n) = sum_sediment(i, n) + sediment(k, i, n)
         'av_sediment(i, n) = sum_sediment(i, n) / k
    
        'Next k
    
    'Next n
    'Next i
    
    
    'FILE_NAME1 =D:\temp\
    Open "D:\temp\grain.dat" For Output As #1
    Print #1, flx
    Print #1, " "
    For i = 0 To nx
    'If i > 10 Or i = flx Then
        'If i Mod 5 = 0 Or i = flx Then
         
         Print #1, topo_temp(i)
        'Print #1, av_sediment(i, 1)
        'Print #1, av_sediment(i, 2)
        'Print #1, av_sediment(i, 3)
        'Print #1, av_sediment(i, 4)
        'Print #1, av_sediment(i, 5)
        'Print #1, av_sediment(i, 6)
        'Print #1, "X"
     '   End If
    'End If
    Next i
    Close #1
End Sub
Public Function grain_percent(sim_time As Integer, i As Integer, n As Integer) As Variant

     percentage = sediment(k - 1, i, n) / sediment(k - 1, i, 0)
     grain_percent = percentage
                
End Function
Public Sub getdischarge() '''''''''''''''''''''''''''markov-ian stochastic sampling for significant discharge regime

Dim kurasim As Single

kurasim = 1
If kurasim = 1 Then

k = sim_time
Do While (k <= end_of_times)
    If k < 1001 Then
        Q_volatility = 0.25
        Q_average = 250
    
    corr1 = -0.5 * (Q_volatility ^ 2)
    LNprop = Log(Q_average) + corr1
    
    M_tmp = (gasdev * Q_volatility)
    M_tmp2 = LNprop + M_tmp
    discharge(k, 0) = (2.718281828 ^ M_tmp2)
    
    k = k + 1
    
    'ElseIf k < 8000 Then
    '    Q_volatility = 0
    ''      Q_average = 0
    '    discharge(k, 0) = 0
    '    k = k + 1
        
    Else
        Q_volatility = 0.5
        Q_average = 650
    
    
    corr1 = -0.5 * (Q_volatility ^ 2)
    LNprop = Log(Q_average) + corr1
    
    M_tmp = (gasdev * Q_volatility)
    M_tmp2 = LNprop + M_tmp
    discharge(k, 0) = (2.718281828 ^ M_tmp2)
    
    k = k + 1
    
    End If
Loop
    
        
Else


k = sim_time
av_dis = 0
som_dis = 0

Do While (k <= end_of_times)
    corr1 = -0.5 * (Q_volatility ^ 2)
    LNprop = Log(Q_average) + corr1
    
    M_tmp = (gasdev * Q_volatility)
    M_tmp2 = LNprop + M_tmp
    discharge(k, 0) = (2.718281828 ^ M_tmp2)
    
    k = k + 1
Loop


End If

sim_time = 1

End Sub

Public Function Minimum(ByVal a As Double, _
                         ByVal b As Double) As Double
Dim mi As Double
                          
  mi = a
  If mi = 0 Then
    mi = grain_size(1)
  End If
  If b < mi And b <> 0 Then
    mi = b
  End If
 
  Minimum = mi
                          
End Function

Public Function gasdev()

'/* Returns a normally distributed deviate with zero mean and unit variance,
'using rnd() as the source of uniform deviated.
    
   Dim gset, iset As Double
   Dim fac, rsq, v1, v2 As Double
   
   iset = 0

   If iset = 0 Then
       
       Do While (rsq >= 1# Or rsq = 0#)
            v1 = 2# * Rnd() - 1#
            v2 = 2# * Rnd() - 1#
            rsq = v1 * v1 + v2 * v2
       Loop
       
       fac = Sqr(-2# * Log(rsq) / rsq)
       gset = v1 * fac
       iset = 1
       gasdev = v2 * fac
        
    Else
        iset = 0
        gasdev = gset
   End If
End Function

Public Function HuetoColorVal(Hue, M1, M2)

Dim V As Double
Dim Vret As Integer

        If Hue < 0 Then
            Hue = Hue + 1
        End If
        If Hue > 1 Then
            Hue = Hue - 1
        End If


    If Hue < 0.1666666 Then
            V = M1 + (M2 - M1) * Hue * 6
    ElseIf Hue < 0.5 Then
        V = M2
    ElseIf Hue < 0.666666 Then
        V = M1 + (M2 - M1) * (0.666666 - Hue) * 6
    Else: V = M1
    End If
      
    HuetoColorVal = 255 * V
End Function


Public Function Maximum(ByVal a As Double, _
                         ByVal b As Double) As Double
Dim ma As Double
                          
  ma = a
  If b > ma Then
    ma = b
  End If
 
  Maximum = ma
                          
End Function
Public Function CSL(k)
Dim l As Integer
actual_t = k 'end_of_times + 1 - k

If k <= 202 Then
    
    If k = 202 Then
    CSL = -26
    ElseIf k = 201 Then CSL = -25.98243243
    ElseIf k = 200 Then CSL = -25.96486486
    ElseIf k = 199 Then CSL = -25.9472973
    ElseIf k = 198 Then CSL = -25.92972973
    ElseIf k = 197 Then CSL = -25.91216216
    ElseIf k = 196 Then CSL = -25.89459459
    ElseIf k = 195 Then CSL = -25.87702703
    ElseIf k = 194 Then CSL = -25.85945946
    ElseIf k = 193 Then CSL = -25.84189189
    ElseIf k = 192 Then CSL = -25.82432432
    ElseIf k = 191 Then CSL = -25.80675676
    ElseIf k = 190 Then CSL = -25.78918919
    ElseIf k = 189 Then CSL = -25.77162162
    ElseIf k = 188 Then CSL = -25.75405405
    ElseIf k = 187 Then CSL = -25.73648649
    ElseIf k = 186 Then CSL = -25.71891892
    ElseIf k = 185 Then CSL = -25.70135135
    ElseIf k = 184 Then CSL = -25.68378378
    ElseIf k = 183 Then CSL = -25.66621622
    ElseIf k = 182 Then CSL = -25.64864865
    ElseIf k = 181 Then CSL = -25.63108108
    ElseIf k = 180 Then CSL = -25.61351351
    ElseIf k = 179 Then CSL = -25.59594595
    ElseIf k = 178 Then CSL = -25.57837838
    ElseIf k = 177 Then CSL = -25.56081081
    ElseIf k = 176 Then CSL = -25.54324324
    ElseIf k = 175 Then CSL = -25.52567568
    ElseIf k = 174 Then CSL = -25.50810811
    ElseIf k = 173 Then CSL = -25.49054054
    ElseIf k = 172 Then CSL = -25.47297297
    ElseIf k = 171 Then CSL = -25.45540541
    ElseIf k = 170 Then CSL = -25.43783784
    ElseIf k = 169 Then CSL = -25.42027027
    ElseIf k = 168 Then CSL = -25.4027027
    ElseIf k = 167 Then CSL = -25.38513514
    ElseIf k = 166 Then CSL = -25.36756757
    ElseIf k = 165 Then CSL = -25.35
    ElseIf k = 164 Then CSL = -25.4
    ElseIf k = 163 Then CSL = -25.4
    ElseIf k = 162 Then CSL = -25.4
    ElseIf k = 161 Then CSL = -25.6
    ElseIf k = 160 Then CSL = -25.8
    ElseIf k = 159 Then CSL = -26#
    ElseIf k = 158 Then CSL = -26.1
    ElseIf k = 157 Then CSL = -26#
    ElseIf k = 156 Then CSL = -25.9
    ElseIf k = 155 Then CSL = -25.6
    ElseIf k = 154 Then CSL = -25.5
    ElseIf k = 153 Then CSL = -25.5
    ElseIf k = 152 Then CSL = -25.7
    ElseIf k = 151 Then CSL = -25.8
    ElseIf k = 150 Then CSL = -25.8
    ElseIf k = 149 Then CSL = -26#
    ElseIf k = 148 Then CSL = -26.2
    ElseIf k = 147 Then CSL = -26.1
    ElseIf k = 146 Then CSL = -26#
    ElseIf k = 145 Then CSL = -26.1
    ElseIf k = 144 Then CSL = -26.2
    ElseIf k = 143 Then CSL = -26.2
    ElseIf k = 142 Then CSL = -26.1
    ElseIf k = 141 Then CSL = -26.1
    ElseIf k = 140 Then CSL = -26.1
    ElseIf k = 139 Then CSL = -26#
    ElseIf k = 138 Then CSL = -26#
    ElseIf k = 137 Then CSL = -26#
    ElseIf k = 136 Then CSL = -26#
    ElseIf k = 135 Then CSL = -26.1
    ElseIf k = 134 Then CSL = -25.8
    ElseIf k = 133 Then CSL = -25.5
    ElseIf k = 132 Then CSL = -25.4
    ElseIf k = 131 Then CSL = -25.7
    ElseIf k = 130 Then CSL = -25.8
    ElseIf k = 129 Then CSL = -25.9
    ElseIf k = 128 Then CSL = -25.9
    ElseIf k = 127 Then CSL = -25.7
    ElseIf k = 126 Then CSL = -25.5
    ElseIf k = 125 Then CSL = -25.5
    ElseIf k = 124 Then CSL = -25.4
    ElseIf k = 123 Then CSL = -25.3
    ElseIf k = 122 Then CSL = -25.6
    ElseIf k = 121 Then CSL = -25.5
    ElseIf k = 120 Then CSL = -25.4
    ElseIf k = 119 Then CSL = -25.2
    ElseIf k = 118 Then CSL = -25.4
    ElseIf k = 117 Then CSL = -25.5
    ElseIf k = 116 Then CSL = -25.7
    ElseIf k = 115 Then CSL = -25.7
    ElseIf k = 114 Then CSL = -25.7
    ElseIf k = 113 Then CSL = -25.7
    ElseIf k = 112 Then CSL = -25.6
    ElseIf k = 111 Then CSL = -25.5
    ElseIf k = 110 Then CSL = -25.6
    ElseIf k = 109 Then CSL = -25.7
    ElseIf k = 108 Then CSL = -25.7
    ElseIf k = 107 Then CSL = -25.6
    ElseIf k = 106 Then CSL = -25.4
    ElseIf k = 105 Then CSL = -25.4
    ElseIf k = 104 Then CSL = -25.4
    ElseIf k = 103 Then CSL = -25.6
    ElseIf k = 102 Then CSL = -25.6
    ElseIf k = 101 Then CSL = -25.6
    ElseIf k = 100 Then CSL = -25.6
    ElseIf k = 99 Then CSL = -25.7
    ElseIf k = 98 Then CSL = -25.6
    ElseIf k = 97 Then CSL = -25.6
    ElseIf k = 96 Then CSL = -25.7
    ElseIf k = 95 Then CSL = -25.6
    ElseIf k = 94 Then CSL = -25.7
    ElseIf k = 93 Then CSL = -25.7
    ElseIf k = 92 Then CSL = -25.7
    ElseIf k = 91 Then CSL = -25.9
    ElseIf k = 90 Then CSL = -26.1
    ElseIf k = 89 Then CSL = -26.1
    ElseIf k = 88 Then CSL = -26.2
    ElseIf k = 87 Then CSL = -26.1
    ElseIf k = 86 Then CSL = -25.9
    ElseIf k = 85 Then CSL = -25.8
    ElseIf k = 84 Then CSL = -25.8
    ElseIf k = 83 Then CSL = -25.9
    ElseIf k = 82 Then CSL = -26#
    ElseIf k = 81 Then CSL = -26.1
    ElseIf k = 80 Then CSL = -26.3
    ElseIf k = 79 Then CSL = -26.3
    ElseIf k = 78 Then CSL = -26.4
    ElseIf k = 77 Then CSL = -26.4
    ElseIf k = 76 Then CSL = -26.5
    ElseIf k = 75 Then CSL = -26.4
    ElseIf k = 74 Then CSL = -26.2
    ElseIf k = 73 Then CSL = -26.1
    ElseIf k = 72 Then CSL = -25.9
    ElseIf k = 71 Then CSL = -26#
    ElseIf k = 70 Then CSL = -26.2
    ElseIf k = 69 Then CSL = -26.1
    ElseIf k = 68 Then CSL = -26.1
    ElseIf k = 67 Then CSL = -26.3
    ElseIf k = 66 Then CSL = -26.5
    ElseIf k = 65 Then CSL = -26.7
    ElseIf k = 64 Then CSL = -26.9
    ElseIf k = 63 Then CSL = -27.3
    ElseIf k = 62 Then CSL = -27.6
    ElseIf k = 61 Then CSL = -27.8
    ElseIf k = 60 Then CSL = -27.8
    ElseIf k = 59 Then CSL = -27.7
    ElseIf k = 58 Then CSL = -27.7
    ElseIf k = 57 Then CSL = -27.7
    ElseIf k = 56 Then CSL = -27.9
    ElseIf k = 55 Then CSL = -27.9
    ElseIf k = 54 Then CSL = -27.8
    ElseIf k = 53 Then CSL = -27.8
    ElseIf k = 52 Then CSL = -27.8
    ElseIf k = 51 Then CSL = -28#
    ElseIf k = 50 Then CSL = -28.1
    ElseIf k = 49 Then CSL = -28.1
    ElseIf k = 48 Then CSL = -28.3
    ElseIf k = 47 Then CSL = -28.3
    ElseIf k = 46 Then CSL = -28.4
    ElseIf k = 45 Then CSL = -28.4
    ElseIf k = 44 Then CSL = -28.3
    ElseIf k = 43 Then CSL = -28.2
    ElseIf k = 42 Then CSL = -28.2
    ElseIf k = 41 Then CSL = -28.2
    ElseIf k = 40 Then CSL = -28.5
    ElseIf k = 39 Then CSL = -28.5
    ElseIf k = 38 Then CSL = -28.5
    ElseIf k = 37 Then CSL = -28.4
    ElseIf k = 36 Then CSL = -28.5
    ElseIf k = 35 Then CSL = -28.3
    ElseIf k = 34 Then CSL = -28.4
    ElseIf k = 33 Then CSL = -28.5
    ElseIf k = 32 Then CSL = -28.5
    ElseIf k = 31 Then CSL = -28.4
    ElseIf k = 30 Then CSL = -28.5
    ElseIf k = 29 Then CSL = -28.5
    ElseIf k = 28 Then CSL = -28.7
    ElseIf k = 27 Then CSL = -28.6
    ElseIf k = 26 Then CSL = -28.8
    ElseIf k = 25 Then CSL = -29#
    ElseIf k = 24 Then CSL = -29#
    ElseIf k = 23 Then CSL = -29#
    ElseIf k = 22 Then CSL = -28.6
    ElseIf k = 21 Then CSL = -28.6
    ElseIf k = 20 Then CSL = -28.3
    ElseIf k = 19 Then CSL = -28.2
    ElseIf k = 18 Then CSL = -28.2
    ElseIf k = 17 Then CSL = -28.1
    ElseIf k = 16 Then CSL = -28#
    ElseIf k = 15 Then CSL = -27.9
    ElseIf k = 14 Then CSL = -27.8
    ElseIf k = 13 Then CSL = -27.6
    ElseIf k = 12 Then CSL = -27.7
    ElseIf k = 11 Then CSL = -27.6
    ElseIf k = 10 Then CSL = -27.3
    ElseIf k = 9 Then CSL = -27.1
    ElseIf k = 8 Then CSL = -27#
    ElseIf k = 7 Then CSL = -26.7
    ElseIf k = 6 Then CSL = -26.7
    ElseIf k = 5 Then CSL = -26.8
    ElseIf k = 4 Then CSL = -27#
    ElseIf k = 3 Then CSL = -26.9
    ElseIf k = 2 Then CSL = -27#
    ElseIf k = 1 Then CSL = -27#
    ElseIf k = 0 Then CSL = -27#
    End If
 Else
    CSL = cubic_spline(actual_t)
 End If
 
End Function

 Function load_EQ_profile(i)

If i = 0 Then
load_EQ_profile = 118.5498
ElseIf i = 1 Then load_EQ_profile = 118.5371
ElseIf i = 2 Then load_EQ_profile = 118.0603
ElseIf i = 3 Then load_EQ_profile = 118.0122
ElseIf i = 4 Then load_EQ_profile = 118.0122
ElseIf i = 5 Then load_EQ_profile = 118.0122
ElseIf i = 6 Then load_EQ_profile = 118.0122
ElseIf i = 7 Then load_EQ_profile = 118.0122
ElseIf i = 8 Then load_EQ_profile = 117.9899
ElseIf i = 9 Then load_EQ_profile = 117.6026
ElseIf i = 10 Then load_EQ_profile = 117.6026
ElseIf i = 11 Then load_EQ_profile = 117.6026
ElseIf i = 12 Then load_EQ_profile = 117.6026
ElseIf i = 13 Then load_EQ_profile = 117.6026
ElseIf i = 14 Then load_EQ_profile = 117.6026
ElseIf i = 15 Then load_EQ_profile = 117.6026
ElseIf i = 16 Then load_EQ_profile = 117.5975
ElseIf i = 17 Then load_EQ_profile = 117.3024
ElseIf i = 18 Then load_EQ_profile = 116.9961
ElseIf i = 19 Then load_EQ_profile = 116.9961
ElseIf i = 20 Then load_EQ_profile = 116.9961
ElseIf i = 21 Then load_EQ_profile = 116.7299
ElseIf i = 22 Then load_EQ_profile = 116.5081
ElseIf i = 23 Then load_EQ_profile = 116.5081
ElseIf i = 24 Then load_EQ_profile = 116.5081
ElseIf i = 25 Then load_EQ_profile = 116.5081
ElseIf i = 26 Then load_EQ_profile = 116.5081
ElseIf i = 27 Then load_EQ_profile = 116.3538
ElseIf i = 28 Then load_EQ_profile = 115.9826
ElseIf i = 29 Then load_EQ_profile = 115.9826
ElseIf i = 30 Then load_EQ_profile = 115.9826
ElseIf i = 31 Then load_EQ_profile = 115.9826
ElseIf i = 32 Then load_EQ_profile = 115.9826
ElseIf i = 33 Then load_EQ_profile = 115.9826
ElseIf i = 34 Then load_EQ_profile = 115.9826
ElseIf i = 35 Then load_EQ_profile = 115.9826
ElseIf i = 36 Then load_EQ_profile = 115.9826
ElseIf i = 37 Then load_EQ_profile = 115.9826
ElseIf i = 38 Then load_EQ_profile = 115.9826
ElseIf i = 39 Then load_EQ_profile = 115.9826
ElseIf i = 40 Then load_EQ_profile = 115.9826
ElseIf i = 41 Then load_EQ_profile = 115.9826
ElseIf i = 42 Then load_EQ_profile = 115.9826
ElseIf i = 43 Then load_EQ_profile = 115.6574
ElseIf i = 44 Then load_EQ_profile = 115.5137
ElseIf i = 45 Then load_EQ_profile = 115.5137
ElseIf i = 46 Then load_EQ_profile = 115.5133
ElseIf i = 47 Then load_EQ_profile = 115.2016
ElseIf i = 48 Then load_EQ_profile = 115.094
ElseIf i = 49 Then load_EQ_profile = 115.094
ElseIf i = 50 Then load_EQ_profile = 115.094
ElseIf i = 51 Then load_EQ_profile = 115.094
ElseIf i = 52 Then load_EQ_profile = 115.094
ElseIf i = 53 Then load_EQ_profile = 115.094
ElseIf i = 54 Then load_EQ_profile = 115.094
ElseIf i = 55 Then load_EQ_profile = 115.094
ElseIf i = 56 Then load_EQ_profile = 115.094
ElseIf i = 57 Then load_EQ_profile = 115.094
ElseIf i = 58 Then load_EQ_profile = 115.094
ElseIf i = 59 Then load_EQ_profile = 115.0893
ElseIf i = 60 Then load_EQ_profile = 114.6933
ElseIf i = 61 Then load_EQ_profile = 114.6313
ElseIf i = 62 Then load_EQ_profile = 114.6313
ElseIf i = 63 Then load_EQ_profile = 114.6313
ElseIf i = 64 Then load_EQ_profile = 114.602
ElseIf i = 65 Then load_EQ_profile = 114.2012
ElseIf i = 66 Then load_EQ_profile = 114.1987
ElseIf i = 67 Then load_EQ_profile = 114.197
ElseIf i = 68 Then load_EQ_profile = 114.1959
ElseIf i = 69 Then load_EQ_profile = 114.1951
ElseIf i = 70 Then load_EQ_profile = 114.1934
ElseIf i = 71 Then load_EQ_profile = 114.1164
ElseIf i = 72 Then load_EQ_profile = 113.7735
ElseIf i = 73 Then load_EQ_profile = 113.7735
ElseIf i = 74 Then load_EQ_profile = 113.7735
ElseIf i = 75 Then load_EQ_profile = 113.7735
ElseIf i = 76 Then load_EQ_profile = 113.556
ElseIf i = 77 Then load_EQ_profile = 113.2504
ElseIf i = 78 Then load_EQ_profile = 113.2504
ElseIf i = 79 Then load_EQ_profile = 113.2504
ElseIf i = 80 Then load_EQ_profile = 113.2504
ElseIf i = 81 Then load_EQ_profile = 113.2504
ElseIf i = 82 Then load_EQ_profile = 113.2504
ElseIf i = 83 Then load_EQ_profile = 113.2504
ElseIf i = 84 Then load_EQ_profile = 113.2504
ElseIf i = 85 Then load_EQ_profile = 113.2504
ElseIf i = 86 Then load_EQ_profile = 113.2504
ElseIf i = 87 Then load_EQ_profile = 113.1442
ElseIf i = 88 Then load_EQ_profile = 112.9154
ElseIf i = 89 Then load_EQ_profile = 112.9154
ElseIf i = 90 Then load_EQ_profile = 112.9154
ElseIf i = 91 Then load_EQ_profile = 112.9154
ElseIf i = 92 Then load_EQ_profile = 112.9154
ElseIf i = 93 Then load_EQ_profile = 112.9154
ElseIf i = 94 Then load_EQ_profile = 112.6494
ElseIf i = 95 Then load_EQ_profile = 112.424
ElseIf i = 96 Then load_EQ_profile = 112.424
ElseIf i = 97 Then load_EQ_profile = 112.424
ElseIf i = 98 Then load_EQ_profile = 112.424
ElseIf i = 99 Then load_EQ_profile = 112.424
ElseIf i = 100 Then load_EQ_profile = 112.424
ElseIf i = 101 Then load_EQ_profile = 112.424
ElseIf i = 102 Then load_EQ_profile = 112.263
ElseIf i = 103 Then load_EQ_profile = 112.0161
ElseIf i = 104 Then load_EQ_profile = 112.0161
ElseIf i = 105 Then load_EQ_profile = 111.8512
ElseIf i = 106 Then load_EQ_profile = 111.545
ElseIf i = 107 Then load_EQ_profile = 111.545
ElseIf i = 108 Then load_EQ_profile = 111.545
ElseIf i = 109 Then load_EQ_profile = 111.545
ElseIf i = 110 Then load_EQ_profile = 111.545
ElseIf i = 111 Then load_EQ_profile = 111.545
ElseIf i = 112 Then load_EQ_profile = 111.545
ElseIf i = 113 Then load_EQ_profile = 111.545
ElseIf i = 114 Then load_EQ_profile = 111.545
ElseIf i = 115 Then load_EQ_profile = 111.545
ElseIf i = 116 Then load_EQ_profile = 111.4482
ElseIf i = 117 Then load_EQ_profile = 111.1622
ElseIf i = 118 Then load_EQ_profile = 111.1622
ElseIf i = 119 Then load_EQ_profile = 111.1622
ElseIf i = 120 Then load_EQ_profile = 111.1622
ElseIf i = 121 Then load_EQ_profile = 111.1622
ElseIf i = 122 Then load_EQ_profile = 111.1615
ElseIf i = 123 Then load_EQ_profile = 110.8171
ElseIf i = 124 Then load_EQ_profile = 110.8171
ElseIf i = 125 Then load_EQ_profile = 110.6614
ElseIf i = 126 Then load_EQ_profile = 110.4043
ElseIf i = 127 Then load_EQ_profile = 110.4043
ElseIf i = 128 Then load_EQ_profile = 110.4043
ElseIf i = 129 Then load_EQ_profile = 110.4043
ElseIf i = 130 Then load_EQ_profile = 110.4043
ElseIf i = 131 Then load_EQ_profile = 110.4043
ElseIf i = 132 Then load_EQ_profile = 110.4043
ElseIf i = 133 Then load_EQ_profile = 110.4043
ElseIf i = 134 Then load_EQ_profile = 110.2632
ElseIf i = 135 Then load_EQ_profile = 110.0114
ElseIf i = 136 Then load_EQ_profile = 109.9252
ElseIf i = 137 Then load_EQ_profile = 109.9252
ElseIf i = 138 Then load_EQ_profile = 109.9252
ElseIf i = 139 Then load_EQ_profile = 109.9252
ElseIf i = 140 Then load_EQ_profile = 109.9252
ElseIf i = 141 Then load_EQ_profile = 109.9252
ElseIf i = 142 Then load_EQ_profile = 109.9252
ElseIf i = 143 Then load_EQ_profile = 109.9228
ElseIf i = 144 Then load_EQ_profile = 109.6059
ElseIf i = 145 Then load_EQ_profile = 109.6059
ElseIf i = 146 Then load_EQ_profile = 109.6059
ElseIf i = 147 Then load_EQ_profile = 109.6059
ElseIf i = 148 Then load_EQ_profile = 109.6059
ElseIf i = 149 Then load_EQ_profile = 109.6059
ElseIf i = 150 Then load_EQ_profile = 109.6059
ElseIf i = 151 Then load_EQ_profile = 109.5157
ElseIf i = 152 Then load_EQ_profile = 109.2885
ElseIf i = 153 Then load_EQ_profile = 109.2885
ElseIf i = 154 Then load_EQ_profile = 109.2885
ElseIf i = 155 Then load_EQ_profile = 109.2885
ElseIf i = 156 Then load_EQ_profile = 109.2885
ElseIf i = 157 Then load_EQ_profile = 109.2885
ElseIf i = 158 Then load_EQ_profile = 109.0136
ElseIf i = 159 Then load_EQ_profile = 108.8764
ElseIf i = 160 Then load_EQ_profile = 108.8764
ElseIf i = 161 Then load_EQ_profile = 108.8764
ElseIf i = 162 Then load_EQ_profile = 108.8764
ElseIf i = 163 Then load_EQ_profile = 108.8764
ElseIf i = 164 Then load_EQ_profile = 108.8764
ElseIf i = 165 Then load_EQ_profile = 108.8764
ElseIf i = 166 Then load_EQ_profile = 108.8764
ElseIf i = 167 Then load_EQ_profile = 108.8764
ElseIf i = 168 Then load_EQ_profile = 108.8764
ElseIf i = 169 Then load_EQ_profile = 108.8764
ElseIf i = 170 Then load_EQ_profile = 108.8763
ElseIf i = 171 Then load_EQ_profile = 108.6276
ElseIf i = 172 Then load_EQ_profile = 108.5299
ElseIf i = 173 Then load_EQ_profile = 108.5299
ElseIf i = 174 Then load_EQ_profile = 108.5299
ElseIf i = 175 Then load_EQ_profile = 108.5299
ElseIf i = 176 Then load_EQ_profile = 108.5299
ElseIf i = 177 Then load_EQ_profile = 108.4962
ElseIf i = 178 Then load_EQ_profile = 108.142
ElseIf i = 179 Then load_EQ_profile = 108.142
ElseIf i = 180 Then load_EQ_profile = 108.142
ElseIf i = 181 Then load_EQ_profile = 108.142
ElseIf i = 182 Then load_EQ_profile = 108.142
ElseIf i = 183 Then load_EQ_profile = 108.142
ElseIf i = 184 Then load_EQ_profile = 108.142
ElseIf i = 185 Then load_EQ_profile = 108.142
ElseIf i = 186 Then load_EQ_profile = 108.142
ElseIf i = 187 Then load_EQ_profile = 108.142
ElseIf i = 188 Then load_EQ_profile = 107.9623
ElseIf i = 189 Then load_EQ_profile = 107.8523
ElseIf i = 190 Then load_EQ_profile = 107.8523
ElseIf i = 191 Then load_EQ_profile = 107.8523
ElseIf i = 192 Then load_EQ_profile = 107.8523
ElseIf i = 193 Then load_EQ_profile = 107.8501
ElseIf i = 194 Then load_EQ_profile = 107.5209
ElseIf i = 195 Then load_EQ_profile = 107.5209
ElseIf i = 196 Then load_EQ_profile = 107.5209
ElseIf i = 197 Then load_EQ_profile = 107.5209
ElseIf i = 198 Then load_EQ_profile = 107.5209
ElseIf i = 199 Then load_EQ_profile = 107.5209
ElseIf i = 200 Then load_EQ_profile = 107.3188
ElseIf i = 201 Then load_EQ_profile = 107.3188
ElseIf i = 202 Then load_EQ_profile = 107.3188
ElseIf i = 203 Then load_EQ_profile = 107.2814
ElseIf i = 204 Then load_EQ_profile = 107.2811
ElseIf i = 205 Then load_EQ_profile = 107.1373
ElseIf i = 206 Then load_EQ_profile = 107.1373
ElseIf i = 207 Then load_EQ_profile = 107.0047
ElseIf i = 208 Then load_EQ_profile = 106.9786
ElseIf i = 209 Then load_EQ_profile = 106.5466
ElseIf i = 210 Then load_EQ_profile = 106.1304
ElseIf i = 211 Then load_EQ_profile = 105.7286
ElseIf i = 212 Then load_EQ_profile = 105.3405
ElseIf i = 213 Then load_EQ_profile = 104.9648
ElseIf i = 214 Then load_EQ_profile = 104.6009
ElseIf i = 215 Then load_EQ_profile = 104.2478
ElseIf i = 216 Then load_EQ_profile = 103.9048
ElseIf i = 217 Then load_EQ_profile = 103.5711
ElseIf i = 218 Then load_EQ_profile = 103.2461
ElseIf i = 219 Then load_EQ_profile = 102.9297
ElseIf i = 220 Then load_EQ_profile = 102.6207
ElseIf i = 221 Then load_EQ_profile = 102.319
ElseIf i = 222 Then load_EQ_profile = 102.0243
ElseIf i = 223 Then load_EQ_profile = 101.7357
ElseIf i = 224 Then load_EQ_profile = 101.4532
ElseIf i = 225 Then load_EQ_profile = 101.1764
ElseIf i = 226 Then load_EQ_profile = 100.9049
ElseIf i = 227 Then load_EQ_profile = 100.6382
ElseIf i = 228 Then load_EQ_profile = 100.3765
ElseIf i = 229 Then load_EQ_profile = 100.1192
ElseIf i = 230 Then load_EQ_profile = 99.86585
ElseIf i = 231 Then load_EQ_profile = 99.61658
ElseIf i = 232 Then load_EQ_profile = 99.37096
ElseIf i = 233 Then load_EQ_profile = 99.1293
ElseIf i = 234 Then load_EQ_profile = 98.89095
ElseIf i = 235 Then load_EQ_profile = 98.65593
ElseIf i = 236 Then load_EQ_profile = 98.42416
ElseIf i = 237 Then load_EQ_profile = 98.19514
ElseIf i = 238 Then load_EQ_profile = 97.96899
ElseIf i = 239 Then load_EQ_profile = 97.74547
ElseIf i = 240 Then load_EQ_profile = 97.52454
ElseIf i = 241 Then load_EQ_profile = 97.3059
ElseIf i = 242 Then load_EQ_profile = 97.0897
ElseIf i = 243 Then load_EQ_profile = 96.87566
ElseIf i = 244 Then load_EQ_profile = 96.66377
ElseIf i = 245 Then load_EQ_profile = 96.45399
ElseIf i = 246 Then load_EQ_profile = 96.24618
ElseIf i = 247 Then load_EQ_profile = 96.04011
ElseIf i = 248 Then load_EQ_profile = 95.8358
ElseIf i = 249 Then load_EQ_profile = 95.63338
ElseIf i = 250 Then load_EQ_profile = 95.43269
ElseIf i = 251 Then load_EQ_profile = 95.23344
ElseIf i = 252 Then load_EQ_profile = 95.0358
ElseIf i = 253 Then load_EQ_profile = 94.8395
ElseIf i = 254 Then load_EQ_profile = 94.64468
ElseIf i = 255 Then load_EQ_profile = 94.45122
ElseIf i = 256 Then load_EQ_profile = 94.25919
ElseIf i = 257 Then load_EQ_profile = 94.06839
ElseIf i = 258 Then load_EQ_profile = 93.87895
ElseIf i = 259 Then load_EQ_profile = 93.69085
ElseIf i = 260 Then load_EQ_profile = 93.50362
ElseIf i = 261 Then load_EQ_profile = 93.31721
ElseIf i = 262 Then load_EQ_profile = 93.13219
ElseIf i = 263 Then load_EQ_profile = 92.94787
ElseIf i = 264 Then load_EQ_profile = 92.76455
ElseIf i = 265 Then load_EQ_profile = 92.58223
ElseIf i = 266 Then load_EQ_profile = 92.40077
ElseIf i = 267 Then load_EQ_profile = 92.2203
ElseIf i = 268 Then load_EQ_profile = 92.04047
ElseIf i = 269 Then load_EQ_profile = 91.86156
ElseIf i = 270 Then load_EQ_profile = 91.68344
ElseIf i = 271 Then load_EQ_profile = 91.50627
ElseIf i = 272 Then load_EQ_profile = 91.32983
ElseIf i = 273 Then load_EQ_profile = 91.15386
ElseIf i = 274 Then load_EQ_profile = 90.97879
ElseIf i = 275 Then load_EQ_profile = 90.80433
ElseIf i = 276 Then load_EQ_profile = 90.63061
ElseIf i = 277 Then load_EQ_profile = 90.45732
ElseIf i = 278 Then load_EQ_profile = 90.28497
ElseIf i = 279 Then load_EQ_profile = 90.11357
ElseIf i = 280 Then load_EQ_profile = 89.94257
ElseIf i = 281 Then load_EQ_profile = 89.77187
ElseIf i = 282 Then load_EQ_profile = 89.60161
ElseIf i = 283 Then load_EQ_profile = 89.43185
ElseIf i = 284 Then load_EQ_profile = 89.26241
ElseIf i = 285 Then load_EQ_profile = 89.09368
ElseIf i = 286 Then load_EQ_profile = 88.92538
ElseIf i = 287 Then load_EQ_profile = 88.75743
ElseIf i = 288 Then load_EQ_profile = 88.59016
ElseIf i = 289 Then load_EQ_profile = 88.42307
ElseIf i = 290 Then load_EQ_profile = 88.25651
ElseIf i = 291 Then load_EQ_profile = 88.09048
ElseIf i = 292 Then load_EQ_profile = 87.92488
ElseIf i = 293 Then load_EQ_profile = 87.75951
ElseIf i = 294 Then load_EQ_profile = 87.59458
ElseIf i = 295 Then load_EQ_profile = 87.43003
ElseIf i = 296 Then load_EQ_profile = 87.26605
ElseIf i = 297 Then load_EQ_profile = 87.10266
ElseIf i = 298 Then load_EQ_profile = 86.93974
ElseIf i = 299 Then load_EQ_profile = 86.77675
ElseIf i = 300 Then load_EQ_profile = 86.61372
ElseIf i = 301 Then load_EQ_profile = 86.45104
ElseIf i = 302 Then load_EQ_profile = 86.28867
ElseIf i = 303 Then load_EQ_profile = 86.12657
ElseIf i = 304 Then load_EQ_profile = 85.96496
ElseIf i = 305 Then load_EQ_profile = 85.80348
ElseIf i = 306 Then load_EQ_profile = 85.64252
ElseIf i = 307 Then load_EQ_profile = 85.48132
ElseIf i = 308 Then load_EQ_profile = 85.32092
ElseIf i = 309 Then load_EQ_profile = 85.16047
ElseIf i = 310 Then load_EQ_profile = 85.00046
ElseIf i = 311 Then load_EQ_profile = 84.84048
ElseIf i = 312 Then load_EQ_profile = 84.68082
ElseIf i = 313 Then load_EQ_profile = 84.52179
ElseIf i = 314 Then load_EQ_profile = 84.36314
ElseIf i = 315 Then load_EQ_profile = 84.20419
ElseIf i = 316 Then load_EQ_profile = 84.04559
ElseIf i = 317 Then load_EQ_profile = 83.8867
ElseIf i = 318 Then load_EQ_profile = 83.72838
ElseIf i = 319 Then load_EQ_profile = 83.5702
ElseIf i = 320 Then load_EQ_profile = 83.41194
ElseIf i = 321 Then load_EQ_profile = 83.25409
ElseIf i = 322 Then load_EQ_profile = 83.0965
ElseIf i = 323 Then load_EQ_profile = 82.939
ElseIf i = 324 Then load_EQ_profile = 82.78181
ElseIf i = 325 Then load_EQ_profile = 82.6245
ElseIf i = 326 Then load_EQ_profile = 82.46755
ElseIf i = 327 Then load_EQ_profile = 82.31067
ElseIf i = 328 Then load_EQ_profile = 82.15427
ElseIf i = 329 Then load_EQ_profile = 81.99812
ElseIf i = 330 Then load_EQ_profile = 81.84188
ElseIf i = 331 Then load_EQ_profile = 81.68564
ElseIf i = 332 Then load_EQ_profile = 81.52945
ElseIf i = 333 Then load_EQ_profile = 81.37334
ElseIf i = 334 Then load_EQ_profile = 81.21747
ElseIf i = 335 Then load_EQ_profile = 81.06171
ElseIf i = 336 Then load_EQ_profile = 80.90606
ElseIf i = 337 Then load_EQ_profile = 80.75056
ElseIf i = 338 Then load_EQ_profile = 80.59517
ElseIf i = 339 Then load_EQ_profile = 80.43998
ElseIf i = 340 Then load_EQ_profile = 80.2849
ElseIf i = 341 Then load_EQ_profile = 80.12988
ElseIf i = 342 Then load_EQ_profile = 79.97495
ElseIf i = 343 Then load_EQ_profile = 79.82069
ElseIf i = 344 Then load_EQ_profile = 79.66607
ElseIf i = 345 Then load_EQ_profile = 79.51149
ElseIf i = 346 Then load_EQ_profile = 79.35696
ElseIf i = 347 Then load_EQ_profile = 79.20252
ElseIf i = 348 Then load_EQ_profile = 79.04806
ElseIf i = 349 Then load_EQ_profile = 78.89378
ElseIf i = 350 Then load_EQ_profile = 78.73956
ElseIf i = 351 Then load_EQ_profile = 78.58563
ElseIf i = 352 Then load_EQ_profile = 78.43161
ElseIf i = 353 Then load_EQ_profile = 78.27792
ElseIf i = 354 Then load_EQ_profile = 78.12398
ElseIf i = 355 Then load_EQ_profile = 77.97031
ElseIf i = 356 Then load_EQ_profile = 77.81668
ElseIf i = 357 Then load_EQ_profile = 77.66356
ElseIf i = 358 Then load_EQ_profile = 77.51019
ElseIf i = 359 Then load_EQ_profile = 77.3567
ElseIf i = 360 Then load_EQ_profile = 77.20324
ElseIf i = 361 Then load_EQ_profile = 77.0501
ElseIf i = 362 Then load_EQ_profile = 76.89663
ElseIf i = 363 Then load_EQ_profile = 76.74359
ElseIf i = 364 Then load_EQ_profile = 76.5904
ElseIf i = 365 Then load_EQ_profile = 76.43747
ElseIf i = 366 Then load_EQ_profile = 76.28442
ElseIf i = 367 Then load_EQ_profile = 76.13156
ElseIf i = 368 Then load_EQ_profile = 75.97888
ElseIf i = 369 Then load_EQ_profile = 75.82601
ElseIf i = 370 Then load_EQ_profile = 75.67332
ElseIf i = 371 Then load_EQ_profile = 75.52093
ElseIf i = 372 Then load_EQ_profile = 75.36835
ElseIf i = 373 Then load_EQ_profile = 75.21581
ElseIf i = 374 Then load_EQ_profile = 75.06322
ElseIf i = 375 Then load_EQ_profile = 74.91089
ElseIf i = 376 Then load_EQ_profile = 74.75838
ElseIf i = 377 Then load_EQ_profile = 74.60583
ElseIf i = 378 Then load_EQ_profile = 74.45362
ElseIf i = 379 Then load_EQ_profile = 74.30161
ElseIf i = 380 Then load_EQ_profile = 74.14921
ElseIf i = 381 Then load_EQ_profile = 73.99704
ElseIf i = 382 Then load_EQ_profile = 73.8451
ElseIf i = 383 Then load_EQ_profile = 73.69296
ElseIf i = 384 Then load_EQ_profile = 73.54076
ElseIf i = 385 Then load_EQ_profile = 73.38892
ElseIf i = 386 Then load_EQ_profile = 73.23691
ElseIf i = 387 Then load_EQ_profile = 73.08522
ElseIf i = 388 Then load_EQ_profile = 72.9333
ElseIf i = 389 Then load_EQ_profile = 72.78098
ElseIf i = 390 Then load_EQ_profile = 72.62951
ElseIf i = 391 Then load_EQ_profile = 72.47784
ElseIf i = 392 Then load_EQ_profile = 72.32605
ElseIf i = 393 Then load_EQ_profile = 72.17432
ElseIf i = 394 Then load_EQ_profile = 72.02273
ElseIf i = 395 Then load_EQ_profile = 71.87112
ElseIf i = 396 Then load_EQ_profile = 71.7192
ElseIf i = 397 Then load_EQ_profile = 71.5677
ElseIf i = 398 Then load_EQ_profile = 71.41631
ElseIf i = 399 Then load_EQ_profile = 71.26468
ElseIf i = 400 Then load_EQ_profile = 71.11315
ElseIf i = 401 Then load_EQ_profile = 70.96194
ElseIf i = 402 Then load_EQ_profile = 70.81059
ElseIf i = 403 Then load_EQ_profile = 70.6591
ElseIf i = 404 Then load_EQ_profile = 70.50766
ElseIf i = 405 Then load_EQ_profile = 70.35629
ElseIf i = 406 Then load_EQ_profile = 70.20524
ElseIf i = 407 Then load_EQ_profile = 70.05422
ElseIf i = 408 Then load_EQ_profile = 69.9028
ElseIf i = 409 Then load_EQ_profile = 69.75079
ElseIf i = 410 Then load_EQ_profile = 69.5995
ElseIf i = 411 Then load_EQ_profile = 69.44865
ElseIf i = 412 Then load_EQ_profile = 69.2977
ElseIf i = 413 Then load_EQ_profile = 69.14658
ElseIf i = 414 Then load_EQ_profile = 68.9952
ElseIf i = 415 Then load_EQ_profile = 68.84408
ElseIf i = 416 Then load_EQ_profile = 68.69327
ElseIf i = 417 Then load_EQ_profile = 68.54243
ElseIf i = 418 Then load_EQ_profile = 68.3914
ElseIf i = 419 Then load_EQ_profile = 68.24024
ElseIf i = 420 Then load_EQ_profile = 68.08936
ElseIf i = 421 Then load_EQ_profile = 67.93819
ElseIf i = 422 Then load_EQ_profile = 67.78665
ElseIf i = 423 Then load_EQ_profile = 67.6357
ElseIf i = 424 Then load_EQ_profile = 67.48502
ElseIf i = 425 Then load_EQ_profile = 67.33399
ElseIf i = 426 Then load_EQ_profile = 67.18317
ElseIf i = 427 Then load_EQ_profile = 67.03248
ElseIf i = 428 Then load_EQ_profile = 66.88178
ElseIf i = 429 Then load_EQ_profile = 66.73093
ElseIf i = 430 Then load_EQ_profile = 66.58031
ElseIf i = 431 Then load_EQ_profile = 66.42951
ElseIf i = 432 Then load_EQ_profile = 66.27889
ElseIf i = 433 Then load_EQ_profile = 66.128
ElseIf i = 434 Then load_EQ_profile = 65.97667
ElseIf i = 435 Then load_EQ_profile = 65.82552
ElseIf i = 436 Then load_EQ_profile = 65.67484
ElseIf i = 437 Then load_EQ_profile = 65.52421
ElseIf i = 438 Then load_EQ_profile = 65.3736
ElseIf i = 439 Then load_EQ_profile = 65.22289
ElseIf i = 440 Then load_EQ_profile = 65.07225
ElseIf i = 441 Then load_EQ_profile = 64.92178
ElseIf i = 442 Then load_EQ_profile = 64.77142
ElseIf i = 443 Then load_EQ_profile = 64.62083
ElseIf i = 444 Then load_EQ_profile = 64.47034
ElseIf i = 445 Then load_EQ_profile = 64.31989
ElseIf i = 446 Then load_EQ_profile = 64.16898
ElseIf i = 447 Then load_EQ_profile = 64.01867
ElseIf i = 448 Then load_EQ_profile = 63.8676
ElseIf i = 449 Then load_EQ_profile = 63.71692
ElseIf i = 450 Then load_EQ_profile = 63.56606
ElseIf i = 451 Then load_EQ_profile = 63.41536
ElseIf i = 452 Then load_EQ_profile = 63.26473
ElseIf i = 453 Then load_EQ_profile = 63.11445
ElseIf i = 454 Then load_EQ_profile = 62.96407
ElseIf i = 455 Then load_EQ_profile = 62.8138
ElseIf i = 456 Then load_EQ_profile = 62.6629
ElseIf i = 457 Then load_EQ_profile = 62.51227
ElseIf i = 458 Then load_EQ_profile = 62.3615
ElseIf i = 459 Then load_EQ_profile = 62.21124
ElseIf i = 460 Then load_EQ_profile = 62.06092
ElseIf i = 461 Then load_EQ_profile = 61.91075
ElseIf i = 462 Then load_EQ_profile = 61.76012
ElseIf i = 463 Then load_EQ_profile = 61.60945
ElseIf i = 464 Then load_EQ_profile = 61.45876
ElseIf i = 465 Then load_EQ_profile = 61.30821
ElseIf i = 466 Then load_EQ_profile = 61.15818
ElseIf i = 467 Then load_EQ_profile = 61.00793
ElseIf i = 468 Then load_EQ_profile = 60.85749
ElseIf i = 469 Then load_EQ_profile = 60.70691
ElseIf i = 470 Then load_EQ_profile = 60.55643
ElseIf i = 471 Then load_EQ_profile = 60.40578
ElseIf i = 472 Then load_EQ_profile = 60.2553
ElseIf i = 473 Then load_EQ_profile = 60.10526
ElseIf i = 474 Then load_EQ_profile = 59.95525
ElseIf i = 475 Then load_EQ_profile = 59.80466
ElseIf i = 476 Then load_EQ_profile = 59.65416
ElseIf i = 477 Then load_EQ_profile = 59.50357
ElseIf i = 478 Then load_EQ_profile = 59.35311
ElseIf i = 479 Then load_EQ_profile = 59.20279
ElseIf i = 480 Then load_EQ_profile = 59.05289
ElseIf i = 481 Then load_EQ_profile = 58.90265
ElseIf i = 482 Then load_EQ_profile = 58.75212
ElseIf i = 483 Then load_EQ_profile = 58.60155
ElseIf i = 484 Then load_EQ_profile = 58.45116
ElseIf i = 485 Then load_EQ_profile = 58.30079
ElseIf i = 486 Then load_EQ_profile = 58.15058
ElseIf i = 487 Then load_EQ_profile = 58.00056
ElseIf i = 488 Then load_EQ_profile = 57.85027
ElseIf i = 489 Then load_EQ_profile = 57.69992
ElseIf i = 490 Then load_EQ_profile = 57.5494
ElseIf i = 491 Then load_EQ_profile = 57.39896
ElseIf i = 492 Then load_EQ_profile = 57.24855
ElseIf i = 493 Then load_EQ_profile = 57.09853
ElseIf i = 494 Then load_EQ_profile = 56.9487
ElseIf i = 495 Then load_EQ_profile = 56.7982
ElseIf i = 496 Then load_EQ_profile = 56.64781
ElseIf i = 497 Then load_EQ_profile = 56.4974
ElseIf i = 498 Then load_EQ_profile = 56.34701
ElseIf i = 499 Then load_EQ_profile = 56.1967
ElseIf i = 500 Then load_EQ_profile = 55.99899
End If

End Function
