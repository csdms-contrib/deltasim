Attribute VB_Name = "Module2"
Option Explicit

Private Sub deposition_calc()


'//////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////// Deposition calculations ////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//initial sedimentload in m3/yr
        
 Dim alpha1 As Single
 Dim alpha2 As Single
 Dim alpha3 As Single
 Dim kappa As Single
         
       If sim_time < 1001 Then
        alpha1 = 0.00013
        alpha2 = 0.55
        alpha3 = 1.5
        kappa = 0.1
       Else
        alpha1 = 0.0011
        alpha2 = 0.53
        alpha3 = 1.1
        kappa = 0.06
       End If
       
        initload = (alpha1 * (discharge(sim_time, 0) ^ alpha2)) * (2700 ^ alpha3) * (2.71 ^ (kappa * Q_volatility * 5)) 'sedimentload_factor                    '0.001) 'init_load (clim_fac) / (discharge(sim_time, 0) * 0.001)
        initload = initload / 100 'riverindex  'van m3 naar m2 125 '

'sediment wordt bij het systeem gevoegd!
    For n = 1 To num_of_gscl
        sedflux(0, n) = sedflux(0, n) + (initload * sed_cont_pct(n))
        sedflux(0, 0) = sedflux(0, 0) + sedflux(0, n)
    Next n

'GRID LOOP
For i = 0 To nx - 1
 
    topo_temp(i) = topo_temp(i) + topo_calc(i)
    topo_depo(i) = 0
            
    '// to improve model stability;
    '// sedimentation higher than the previous gridcell is not allowed
    '// to prevent bulges that cannot be overcome in 2D (negative slopes)
                        
    If i = 0 Then
        max_dh = initial_gradient 'topo_temp(i)
    Else
        max_dh = topo_temp(i - 1) - topo_temp(i) - EPS
        
        If max_dh < EPS Then
            max_dh = EPS
        End If
        
    End If
   
    
    
    For n = 1 To num_of_gscl
            
        If i < flx Then
                     
                     
       ' streampower = (discharge(sim_time, i) * ((1 + slope(i)))) / 1000
        
       ' stream_velocity = ((discharge(sim_time, i) / 100)) * (1 / streampower) '(max_dh / initial_gradient)
        
       stream_velocity = (discharge(sim_time, i) / Q_average) * (1 + slope(i))
        
       traveldist = stream_velocity * traveldist_fluvial(n)
        'traveldist = traveldist_fluvial(n)
        
        Else
            traveldist = traveldist_marine(n)
       
        End If
        
        
                
        max_dh_gs = max_dh * sed_cont_pct(n) '((dx / traveldist) / tot)          '
        settle_rate = sedflux(i, n) / (traveldist) '
       
                
        ' ACTUAL DEPOSTION
         If (settle_rate * dt >= max_dh_gs) Then              '//no deposition higher than previous gridcell
             
             sediment(sim_time / dt, i, n) = max_dh_gs
             sediment(sim_time / dt, i, 0) = sediment(sim_time / dt, i, 0) + sediment(sim_time / dt, i, n)
                    
             sedflux(i + 1, n) = sedflux(i + 1, n) + (sedflux(i, n) - max_dh_gs)
             sedflux(i + 1, 0) = sedflux(i + 1, 0) + (sedflux(i, n) - max_dh_gs)
            
             topo_depo(i) = topo_depo(i) + max_dh_gs
                    
         Else
                      
             sediment(sim_time / dt, i, n) = settle_rate * dt
             sediment(sim_time / dt, i, 0) = sediment(sim_time / dt, i, 0) + sediment(sim_time / dt, i, n)
            '// here the sedflux that is left after settling is transported to the next cell
                    
             sedflux(i + 1, n) = sedflux(i + 1, n) + (sedflux(i, n) - (settle_rate))
             sedflux(i + 1, 0) = sedflux(i + 1, 0) + (sedflux(i, n) - settle_rate) * dt
                    
             topo_depo(i) = topo_depo(i) + (settle_rate * dt)
                        
         End If
         
         
         If (sedflux(i, n) < EPS) Then
             sedflux(i, n) = 0
         End If
         
         
         If (sedflux(i + 1, n) < EPS) Then
             sedflux(i + 1, n) = 0
         End If
         
         If sediment(sim_time / dt, i, n) < 0 Then
           sediment(sim_time / dt, i, n) = 0
         End If
         
    Next n
            
         topo_temp(i) = topo_temp(i) + topo_depo(i)
    
    
    If sim_time = 100 Then
              
              Print #14, topo_depo(i)
    End If
    
        
    If (i <= flx) And (i <> 0) Then
                
        For n = 1 To num_of_gscl
           sedflux(i, n) = sedflux(i, n) + sedflux(i - 1, n) '+ sedflux(0, n)
           sedflux(i, 0) = sedflux(i, 0) + sedflux(i, n)
        Next n
        '// offshore the plume divides the sediment over larger lateral area
        '// in pseudo 3d we follow the plume axis
                                
    ElseIf (i > flx) Then
                                
        For n = 1 To num_of_gscl
            out_coast_distance = i - flx
            sedflux(i, n) = sedflux(i, n) + (sedflux(i - 1, n) - (sedflux(i - 1, n) * ((0.002 * out_coast_distance * (Tan(spreading_angle * (PI / 180)))))))
            sedflux(i, 0) = sedflux(i, 0) + sedflux(i, n)
        Next n
                         
    End If
Next i

'end Grid loop
   
   If end_of_times <= 500 Then
    timeln_intrval = end_of_times / 10
   Else
    timeln_intrval = end_of_times / 10
   End If
   
    kk = sim_time
    
   'write_wheeler_to_screen
   
   If (sim_time) Mod timeln_intrval = 0 Or sim_time = 1 Or sim_time = end_of_times - 1 Then
   'write_timeline_to_screen
   'Module5.draw_real_time_line
   End If
   
   If sim_time = 1 Then
    tot = topo_temp(flx)
    i_flx = flx
    onshore = flx - 100
    tot_onshore = topo_temp(onshore)
    offshore = flx + 100
    tot_offshore = topo_temp(offshore)
    Print #12, onshore, i_flx, offshore
   End If
   If (sim_time) Mod timeln_intrval = 0 Or sim_time = 1 Or sim_time = end_of_times - 1 Then
   'Print #12, topo_temp(flx); " "; flx; " "; sim_time
        Print #12, (topo_temp(onshore) - tot_onshore), (topo_temp(i_flx) - tot), (topo_temp(offshore) - tot_offshore)
   End If
   'Print #12, " "
     
End Sub


Public Sub draw_well1()
    draw_well
End Sub

Sub write_timeline_to_screen()
    
    timeln_intrval = 20
    Rmarg = 36
    Tmarg = 10
    Bmarg = 50
    Lmarg = 36
 '/* 30 pts for vertical scale bar    */
   

 '/* Determine x & y scale */
    hhmax = -1000000
    hmax = hhmax
    hhmin = 1000000
    hmin = hhmin
    
        For j = xstrtps To xstopps   '500
            hmin = Minimum(topo_temp(j), hmin)
            hmax = Maximum(topo_temp(j), hmax)
        Next j
   

     scy = (280 - 60 - (Tmarg + Bmarg)) / (hmax - hmin) '//333
     scx = (700 - (Lmarg + Rmarg)) / (xstopps - xstrtps)

   '/* Draw sealevel
    
  
    Form1.pic1.Line ((Lmarg + scx * (flx_old - xstrtps)), (Bmarg + scy * (model_sealevel_old - hmin) + 1))-((Lmarg + scx * (xstopps)), (Bmarg + scy * (model_sealevel_old - hmin) - 1)), RGB(255, 255, 255), BF
    Form1.pic1.Line ((Lmarg + scx * (flx - xstrtps)), (Bmarg + scy * (model_sealevel - hmin)))-((Lmarg + scx * (xstopps)), (Bmarg + scy * (model_sealevel - hmin))), RGB(0, 0, 255), BF
     flx_old = flx
     model_sealevel_old = model_sealevel
    
    '/* draw some timelines in stratigraphy(if required) */
            For j = xstrtps To xstopps
            If j = flx Then
                flx = flx
            End If
                Form1.pic1.Line ((Lmarg + scx * (j - xstrtps)), (Bmarg + scy * (topo_temp(j) - hmin)))-((Lmarg + scx * (j + 1 - xstrtps)), (Bmarg + scy * (topo_temp(j + 1) - hmin)))

            Next j
            
      Form1.pic1.Refresh

End Sub
Private Sub write_wheeler_to_screen_orginal()
    s = 0.75
    l = 0.5

    Rmarg = 36
    Tmarg = 10
    Bmarg = 50
    Lmarg = 36
    
    datamaxx = -1000000
    datamax = datamaxx
    dataminn = 1000000
    datamin = dataminn
  
     datamin = 0.001
     e_datamax = 0.001
     e_datamin = 0
     datamax = 0.02
     
    
    'For i = xstrtps To xstopps ' xstopps
            'dataminn = Minimum(topo_depo(i), datamin)
            'datamin = dataminn
             
            'e_dataminn = Minimum(dh_erosion(i), e_datamin)
            'e_datamin = e_dataminn
            
            'datamaxx = Maximum(topo_depo(i), datamax)
            'datamax = datamaxx
          
            'e_datamaxx = Maximum(dh_erosion(i), e_datamax)
            'e_datamax = e_datamaxx
            
        'Next i
    'End If
    
    Form1.Picture1.Scale (xstrtps - 5, end_of_times)-(xstopps + Rmarg - 25, 0)
    Form1.Picture1.DrawWidth = 5 '450 / end_of_times
    
    For i = xstrtps To xstopps
       If i < flx Then
        
        topo_depo(i) = topo_depo(i) - dh_erosion(i)
        
        If topo_depo(i) > 0 Then
            dh_erosion(i) = 0
        Else
            dh_erosion(i) = -topo_depo(i)
            topo_depo(i) = 0
        End If
       
       End If
       
      
        
            If topo_depo(i) < 0 Then '0.001
                  i = i
            Else
                If dh_erosion(i) < 0.001 Then
            
                'If topo_depo(i) > dh_erosion(i) Then
             
                    s = 0.99 * ((topo_depo(i) - datamin) / (datamax - datamin))
                    l = 1 - (0.5 * ((topo_depo(i) - datamin) / (datamax - datamin)))
                    hueD = 0.7
            
                    'l = 0.5
    
                    If s < 0.001 Then
                        r1 = 255 * l
                        g1 = r1
                        b1 = r1
                    Else
                        If l < 0.5 Then
                            temp2 = l * (1 + s)
                        Else
                            temp2 = l + s - l * s
                        End If
    
                    temp1 = 2 * l - temp2
                    temp3R = hueD + 0.33333
         
                    If temp3R > 1 Then
                        temp3R = temp3R - 1
                    End If
         
                    temp3G = hueD
                    temp3B = hueD - 0.33333
       
                    If temp3B < 0 Then
                         temp3B = temp3B + 1
                    End If
    
                    r1 = HuetoColorVal(temp3R, temp1, temp2)
                    
                    If r1 < 0 Then r1 = 0
                    If r1 > 255 Then r1 = 255
                    g1 = HuetoColorVal(temp3G, temp1, temp2)
                    If g1 < 0 Then g1 = 0
                    If g1 > 255 Then g1 = 255
                    b1 = HuetoColorVal(temp3B, temp1, temp2)
                    If b1 < 0 Then b1 = 0
                    If b1 > 255 Then b1 = 255
                End If
    
    
                Form1.Picture1.Line (i, sim_time)-(i + 1, sim_time), RGB(r1, g1, b1)
        
             End If
            End If
        
        ' ElseIf i < flx Then
            If dh_erosion(i) < 0.0091 Then  '
      
            Else
                s = (0.99) * ((dh_erosion(i) - e_datamin) / (e_datamax - e_datamin))
                l = 0.5 '1 - ((0.5) * ((dh_erosion(i) - e_datamin) / (e_datamax - e_datamin)))
                hueD = 0
                'l = 0.5
                
                If s < 0.001 Then
                    r1 = 255 * l
                    g1 = r1
                    b1 = r1
                Else
                    If l < 0.5 Then
                        temp2 = l * (1 + s)
                    Else
                        temp2 = l + s - l * s
                    End If
    
                    temp1 = 2 * l - temp2
                    temp3R = hueD + 0.33333
         
                    If temp3R > 1 Then
                        temp3R = temp3R - 1
                    End If
         
                    temp3G = hueD
                    temp3B = hueD - 0.33333
       
                    If temp3B < 0 Then
                        temp3B = temp3B + 1
                    End If
    
                    r1 = HuetoColorVal(temp3R, temp1, temp2)
                    If r1 < 0 Then r1 = 0
                    If r1 > 255 Then r1 = 255
                    
                    g1 = HuetoColorVal(temp3G, temp1, temp2)
                    If g1 < 0 Then g1 = 0
                    If g1 > 255 Then g1 = 255
                    b1 = HuetoColorVal(temp3B, temp1, temp2)
                    If b1 < 0 Then b1 = 0
                    If b1 > 255 Then b1 = 255
              End If
    
            
            
             Form1.Picture1.Line (i, sim_time)-(i + 1, sim_time), RGB(r1, g1, b1)
            
        End If
     'End If
    Next i
            
      
    
    Form1.Picture1.Line (flx, sim_time)-(flx + 2, sim_time), RGB(0, 0, 0), BF
    
    Form1.Picture2.Scale (1, 3 * Q_average)-(end_of_times, 0)
    Form1.Picture2.Line (sim_time, 0)-(sim_time + 1, (discharge(sim_time, 1) - (0.8 * Q_average))), RGB(0, 0, 0), BF
   
   If SLoption <> 1 Then
        If SLoption = 4 Or SLoption = 5 Then
            Form1.Picture3.Scale (1, RICO + 1)-(end_of_times, 0)
            Form1.Picture3.Line (sim_time, model_sealevel - SLC)-(sim_time + 1, (model_sealevel - SLC + 1)), RGB(0, 0, 255), BF
        Else
         Form1.Picture3.Scale (1, sealevel_amplitude + 1)-(end_of_times, -sealevel_amplitude - 1)
         Form1.Picture3.Line (sim_time, model_sealevel - SLC)-(sim_time + 1, (model_sealevel - SLC + 1)), RGB(0, 0, 255), BF
        End If
   Else
        Form1.Picture3.Scale (0, SLmax + 1)-(end_of_times, SLmin - 1)
        Form1.Picture3.Line (sim_time, SLinput(sim_time))-(sim_time + 1, SLinput(sim_time + 1)), RGB(0, 0, 255)
   End If

Form1.Picture1.Refresh
Form1.Picture2.Refresh
Form1.Picture3.Refresh
Form10.Picture4.Refresh
End Sub

 Private Sub write_wheeler_to_screen()
   
    Dim R, G, b As Single
    Dim nett_dep(0 To 500) As Single
    
    
    Form1.Picture1.Scale (xstrtps - Lmarg, end_of_times)-(xstopps + Rmarg, 0)

    For i = xstrtps To xstopps
        nett_dep(i) = topo_depo(i) - erosion_rate(i)
        
         If (nett_dep(i) < 0) Then
                R = 0
                b = 100000 * -1 * nett_dep(i)
                    If b > 255 Then b = 255
                G = 100000 * -1 * nett_dep(i)
                     If G > 255 Then G = 255
                 
                 If i > xstrtps And i < xstopps Then
                    Form1.Picture1.Line (i, sim_time)-(i + 1, sim_time + 1), RGB(255 - R, 255 - G, 255 - b), BF
                 End If
                
                     
       ElseIf (nett_dep(i) > 0) Then
       
        If i < flx Then
                R = 5000 * nett_dep(i)
                    If R > 255 Then R = 255
                G = 5000 * nett_dep(i)
                    If G > 255 Then G = 255
                b = 0
        
        
       
        ElseIf i > flx Then
                R = 5000 * nett_dep(i)
                    If R > 255 Then R = 255
                G = 5000 * nett_dep(i)
                    If G > 255 Then G = 255
                b = 0
          
        End If
                If i >= xstrtps And i < xstopps Then
                    Form1.Picture1.Line (i, sim_time)-(i + 1, sim_time + 1), RGB(255 - R, 255 - G, 255 - b), BF
                
                End If
             
                       
        
      End If
                
        If i = flx Then
                
            Form1.Picture1.Line (flx, sim_time)-(flx, sim_time + 1), RGB(255, 255, 255), BF
        
        End If
        
       
    Next i
    'Form1.View2.Refresh
    
'End Sub
    
   ' s = 0.75
    'l = 0.5

   ' Rmarg = 36
   ' Tmarg = 10
   ' Bmarg = 50
  '  Lmarg = 36
  '
  '  datamaxx = -1000000
  '  datamax = datamaxx
  '  dataminn = 1000000
  '  datamin = dataminn
  
  '   datamin = 0
'     e_datamax = 0.0025
 '    e_datamin = 0
 '    datamax = 0.005
     
  '      For i = xstrtps To xstopps ' xstopps
            'dataminn = Minimum(topo_depo(i), datamin)
            'datamin = dataminn
             
           ' e_dataminn = Minimum(dh_erosion(i), e_datamin)
           ' e_datamin = e_dataminn
            
  '          datamaxx = Maximum(topo_depo(i), datamax)
  '          datamax = datamaxx
          
  '          e_datamaxx = Maximum(dh_erosion(i), e_datamax)
  '          e_datamax = e_datamaxx
            
 '       Next i
    
  
    
 '   Form1.Picture1.Scale (xstrtps - 5, end_of_times)-(xstopps + Rmarg - 25, 0)
 '   Form1.Picture1.DrawWidth = 5 '450 / end_of_times
    
 '   For i = xstrtps To xstopps
      
        
 '       topo_depo(i) = topo_depo(i) - dh_erosion(i)
        
        ''If i < flx Then
        ''    If dh_erosion(i) > topo_depo(i) * 0.5 Then
        ''        topo_depo(i) = 0
         ''       dh_erosion(i) = dh_erosion(i) * 0.5
         ''   Else
         ''       topo_depo(i) = topo_depo(i) * 0.25
         ''
         ''   End If
        
        
        
       '' Else
       ''     dh_erosion(i) = 0
       '' End If
    
        
 '
 '      If topo_depo(i) <> 0 Then
'            r1 = 255 - (255 * (topo_depo(i) - datamin) / (datamax - datamin))
 '           g1 = 255 - (255 * (topo_depo(i) - datamin) / (datamax - datamin))
 '           b1 = 255
       
 '      ElseIf dh_erosion(i) <> 0 Then
 '           r1 = 255
  '          g1 = 255 - (255 * (dh_erosion(i) - e_datamin) / (e_datamax - e_datamin))
 '           b1 = 255 - (255 * (dh_erosion(i) - e_datamin) / (e_datamax - e_datamin))
 '      End If
            
 '           Form1.Picture1.Line (i, sim_time)-(i + 1, sim_time), RGB(r1, g1, b1)
            
        
 '    'End If
 '   Next i
            
      
    
'    Form1.Picture1.Line (flx, sim_time)-(flx + 2, sim_time), RGB(0, 0, 0), BF
    
'    Form1.Picture2.Scale (1, 3 * Q_average)-(end_of_times, 0)
'    Form1.Picture2.Line (sim_time, 0)-(sim_time + 1, (discharge(sim_time, 1) - (0.8 * Q_average))), RGB(0, 0, 0), BF
   
'   If SLoption <> 1 Then
'        If SLoption = 4 Or SLoption = 5 Then
'            Form1.Picture3.Scale (1, RICO + 1)-(end_of_times, 0)
'            Form1.Picture3.Line (sim_time, model_sealevel - SLC)-(sim_time + 1, (model_sealevel - SLC + 1)), RGB(0, 0, 255), BF
'        Else
'         Form1.Picture3.Scale (1, sealevel_amplitude + 1)-(end_of_times, -sealevel_amplitude - 1)
'         Form1.Picture3.Line (sim_time, model_sealevel - SLC)-(sim_time + 1, (model_sealevel - SLC + 1)), RGB(0, 0, 255), BF
'        End If
 '  Else
'        Form1.Picture3.Scale (0, SLmax + 1)-(end_of_times, SLmin - 1)
 '       Form1.Picture3.Line (sim_time, SLinput(sim_time))-(sim_time + 1, SLinput(sim_time + 1)), RGB(0, 0, 255)
'   End If

Form1.Picture1.Refresh
Form1.Picture2.Refresh
Form1.Picture3.Refresh
Form10.Picture4.Refresh
End Sub

     
   
    
   
'---------------
Function HuetoColorVal(Hue, M1, M2)

Dim V As Single
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

Private Sub Initialize_simulation()
   
    
    sediment(0, 0, 0) = initial_height                      '//topographical height at start of profile
    topo_old(0) = sediment(0, 0, 0)
    topo_temp(0) = sediment(0, 0, 0)
            

dy = initial_gradient                            '// initial gradient in m/dx
               
i = 0
If GRID_Option = 1 Then
    For i = 0 To nickpoint - 1 'nx - 1 'nx = 500
        z = i + 1
        sediment(0, i + 1, 0) = (sediment(0, i, 0) - (dy_onshore)) '+ noise
        topo_old(i + 1) = sediment(0, i + 1, 0)
        topo_temp(i + 1) = sediment(0, i + 1, 0)
        topo_calc(i) = 0
        n = 1
           
        For n = 1 To num_of_gscl
            sediment(0, i, n) = sediment(0, i, 0) * sed_cont_pct(n)
        Next n
            
    Next i
       
    For i = nickpoint To nx - 1
        sediment(0, i + 1, 0) = (sediment(0, i, 0) - dy) '+ noise
        topo_old(i + 1) = sediment(0, i + 1, 0)
        topo_temp(i + 1) = sediment(0, i + 1, 0)
        topo_calc(i) = 0
        n = 1
           
        For n = 1 To num_of_gscl
            sediment(0, i, n) = sediment(0, i, 0) * sed_cont_pct(n)
        Next n
            
    Next i
ElseIf GRID_Option = 2 Then
   
        For i = 0 To nx
            sediment(0, i, 0) = Module3.load_EQ_profile(i) '+ noise
            topo_old(i) = sediment(0, i, 0)
            topo_temp(i) = sediment(0, i, 0)
            topo_calc(i) = 0
            
               
            For n = 1 To num_of_gscl
                sediment(0, i, n) = sediment(0, i, 0) * sed_cont_pct(n)
            Next n
        Next i
End If
   i = 0
End Sub

Private Sub ersosion_calc()

        i = 1
        Do Until i = flx + 1  '//onshore
                discharge(sim_time, i) = discharge(sim_time, i - 1)
        i = i + 1
        Loop
        
        inpcount = 0
        i = flx
        
        If i = 0 Then
            i = 1
        End If
        
        Do Until i = nx     '// offshore
        
        '//erosion decreases with increasing waterdepth
            discharge_shapef = 0.363970234 '(tan 20 degrees)

            waterdepth(i) = model_sealevel - topo_temp(i)
            D(i) = i - flx
            
            discharge(sim_time, i) = discharge(sim_time, i - 1) - (2 * discharge_shapef * D(i) * discharge(sim_time, i - 1)) '* dx
                
                If (discharge(sim_time, i) < EPS) Then
                
                   If (discharge(sim_time, i - 1) = EPS) Then
                        wavebase_cell = i
                   End If
                   discharge(sim_time, i) = 0
                End If
        
            i = i + 1
        Loop
        som_dis = som_dis + discharge(sim_time, 1)
        av_dis = (som_dis / sim_time)
'////////////////// Erosion calculations ///////////////////////////////////////////////////////
'//////////////// Assumption: erosion is equal for all grainsizes//////////////////////////////

        erosion_power = 1
        i = 0
        Do Until i = nx
        
        '//erosion is slope independent in the marine domain
            If (i >= flx) Then
               erosion_power = 0
            End If

            slope(i) = (topo_temp(i) - topo_temp(i + 1)) '/ dx
            If slope(i) = 0 Then
            slope(i) = EPS
            End If
            streampower = ((discharge(sim_time, i) / Q_average) * ((1 + slope(i)) ^ erosion_power)) '/ Q_average
            
            If (i < flx) Then
                        
                k_er_fluv = 0.0025 'k_er_fluvC '* ((discharge(sim_time, 1)) / (av_dis))
                erosion_rate(i) = k_er_fluv * streampower
                        
            Else
                        
                erosion_rate(i) = k_er_marine * streampower
                        
            End If
            If sim_time = 100 Then
              Print #13, erosion_rate(i)
            End If
            
            
            dh_erosion(i) = dt * erosion_rate(i)

           
            
            topo_calc(i) = topo_calc(i) - dh_erosion(i)
           
          
            '// erosion is substracted from sediment array (for the different grainsizes)

            k = (sim_time / dt)
           
            
            For k = (sim_time / dt) To 1 Step -1
                
                If sediment(k - 1, i, 0) - dh_erosion(i) < 0 And sediment(k - 1, i, 0) <> 0 Then
                
                    For n = 1 To num_of_gscl
                        sedflux(flx, n) = sedflux(flx, n) + sediment(k - 1, i, n)
                        sedflux(flx, 0) = sedflux(flx, 0) + sedflux(i, n)
                        sediment(k - 1, i, n) = 0
                    Next n
                 
                 dh_erosion(i) = dh_erosion(i) - sediment(k - 1, i, 0)
                 sediment(k - 1, i, 0) = 0
                 
                 ElseIf sediment(k - 1, i, 0) <> 0 Then
                    For n = 1 To num_of_gscl
                        sedflux(flx, n) = sedflux(flx, n) + (grain_percent(k - 1, i, n) * dh_erosion(i)) '(sediment(k - 1, i, n) * dh_erosion(i))
                        sediment(k - 1, i, n) = sediment(k - 1, i, n) - (grain_percent(k - 1, i, n) * dh_erosion(i)) '(sediment(k - 1, i, n) * dh_erosion(i))
                        sedflux(flx, 0) = sedflux(flx, 0) + sedflux(i, n)
                    Next n
                    sediment(k - 1, i, 0) = sediment(k - 1, i, 0) - dh_erosion(i)
                    k = 0
                  End If
            Next k
            
    i = i + 1
    Loop


'write_wheeler_to_screen
'write_wheeler_to_screen_orginal
End Sub

'
'
Private Sub initialize_sealevel()
'// fill for every time-step sedflux and topo_calc arrays with zero's


    ReDim sedflux(0 To nx, 0 To num_of_gscl)
    
    ReDim topo_calc(0 To nx)

'/* calculate sea level
   
    'model_sealevel_old = model_sealevel
  
If sim_time = 1 Then
            
    
            
    If SLoption = 1 Then
        Module4.Load_SL_Time
        Module4.load_sl_value
        
        For k = 0 To end_of_times
            SLinput(k) = Module3.CSL(end_of_times - k)
        Next k
        
        SLmaxx = -1000000
        SLmax = SLmaxx
        SLminn = 1000000
        SLmin = SLminn
        For k = 0 To end_of_times
           
            SLminn = Minimum(SLinput(k), SLmin)
            SLmin = SLminn
            SLmaxx = Maximum(SLinput(k), SLmax)
            SLmax = SLmaxx
    
        Next k
        
        If end_of_times <> 10000 Then
            flx = 300 'nickpoint
            SLC = topo_temp(flx) - SLinput(end_of_times)
        Else
            flx = nickpoint + SLinput(end_of_times)
            SLC = topo_temp(flx)
        End If
        
    ElseIf SLoption = 3 Then
        For k = 0 To end_of_times
            SLinput(k) = (sealevel_amplitude * Sin((k * (PI / (0.5 * sealevel_frequency)) + (0.5 * PI))) - sealevel_amplitude)
        Next k
    ElseIf SLoption = 4 Then
        RICO = Form4.Text1
        For k = 0 To end_of_times
            SLinput(k) = (RICO / end_of_times) * k
        Next k
    ElseIf SLoption = 6 Then
        
        RICO = Form4.Text1 * -1
        For k = 0 To end_of_times
            SLinput(k) = (RICO / end_of_times) * k
        Next k
        
    ElseIf SLoption = 5 Then
        RICO = Form4.Text2
        For k = 0 To end_of_times
            SLinput(k) = RICO + (RICO * Sin((k * (PI / (0.5 * 4 * end_of_times)) - (0.5 * PI))))
        Next k
    ElseIf SLoption = 7 Then
        RICO = Form4.Text2 * -1
        For k = 0 To end_of_times
            SLinput(k) = RICO + (RICO * Sin((k * (PI / (0.5 * 4 * end_of_times)) - (0.5 * PI))))
        Next k
        
    Else
        For k = 0 To end_of_times
             SLinput(k) = sealevel
        Next k
    End If
 
    
    If nickpoint <> 1 And SLoption <> 1 And GRID_Option = 1 Then
        
        flx = nickpoint
        SLC = topo_temp(flx)
        model_sealevel = SLinput(sim_time) + SLC '- 10
        
    ElseIf SLoption = 1 Then
        
        model_sealevel = SLinput(sim_time) + SLC
        
    ElseIf GRID_Option = 2 Then
        flx = 201
        SLC = topo_temp(flx)
        model_sealevel = SLinput(sim_time) + SLC
             
    Else
        flx = 200
        SLC = topo_temp(flx)
        model_sealevel = SLinput(sim_time) + SLC - 10
        
    End If
            
    If SLC = 0 Then
        SLC = 80
    End If
    
ElseIf sim_time <> 1 Then
            
    model_sealevel = SLinput(sim_time) + SLC
            
    '// determine coastline position
            
    flx = 0
    i = 0
    
    Do Until topo_temp(i) <= model_sealevel
        flx = flx + 1
        i = i + 1
    Loop

End If


End Sub

Private Sub store_strat()
i = 0
Do Until i = (end_of_times / dt)

    j = 0
    Do Until j = nx

        If i = 0 Then
            StratNode(i, j) = sediment(i, j, 0)
        Else
            StratNode(i, j) = (StratNode(i - 1, j) + sediment(i, j, 0))
            thickness(j) = StratNode(i, j) - StratNode(i - 1, j)
        End If
        j = j + 1
    
        'Print #12, sediment(i, well1, 0); sediment(i, well1, 1); sediment(i, well1, 2); sediment(i, well1, 3); sediment(i, well1, 4); sediment(i, well1, 5); sediment(i, well1, 6)
    
    Loop
        
    i = i + 1
Loop

End Sub '
Private Sub store_texture()
Dim Sorting() As String
Dim tempsort As String
ReDim Sorting(0 To end_of_times, 0 To nx)


k = 0
For k = 0 To (end_of_times / dt)
    For i = 0 To nx
        If i = nx Then
         
         If k = 100 Then
            k = k
         End If
         End If
        
        If sediment(k, i, 0) <> 0 Then
            tmpsum = 0
            For n = 1 To num_of_gscl
                tmpsum = tmpsum + ((sediment(k, i, n) / sediment(k, i, 0)) * grain_size(n))
            Next n
            Median(k, i) = tmpsum / 1 '((grain_size(n) + grain_size(n + 1)) / 2)
            
            tempsort = 0
            For n = 1 To num_of_gscl
                tempsort = tempsort + ((((sediment(k, i, n) / sediment(k, i, 0)) * grain_size(n)) - Median(k, i)) ^ 2) / (num_of_gscl - 1)
           Next n
            Sorting(k, i) = Sqr(tempsort)
                
                If i = well1 Then
                    'Print #12, Median(k, well1), Sorting(k, well1)
                End If
        Else
            Median(k, i) = 0
        End If
            
        If MCrun = 205 Then
            If Median(k, i) >= 0.15 Then
                Prob_Sand(k, i) = (Prob_Sand(k, i) + 1) / MCrun 'Median(k, i)
                Prob_Silt(k, i) = (Prob_Silt(k, i) + 0) / MCrun
                Prob_Clay(k, i) = (Prob_Clay(k, i) + 0) / MCrun
            ElseIf Median(k, i) <= 0.1 Then
                Prob_Sand(k, i) = (Prob_Sand(k, i) + 0) / MCrun
                Prob_Silt(k, i) = (Prob_Silt(k, i) + 0) / MCrun
                Prob_Clay(k, i) = (Prob_Clay(k, i) + 1) / MCrun 'Median(k, i)
            Else
                Prob_Sand(k, i) = (Prob_Sand(k, i) + 0) / MCrun
                Prob_Silt(k, i) = (Prob_Silt(k, i) + 1) / MCrun 'Median(k, i)
                Prob_Clay(k, i) = (Prob_Clay(k, i) + 0) / MCrun
            End If
                
        Else
                            
            If Median(k, i) >= 0.25 Then
                Prob_Sand(k, i) = (Prob_Sand(k, i) + 1)  'Median(k, i)
                'Prob_Silt(k, i) = (Prob_Silt(k, i) + 0) / MCrun
                'Prob_Clay(k, i) = (Prob_Clay(k, i) + 0) / MCrun
            ElseIf Median(k, i) <= 0.1 Then
                'Prob_Sand(k, i) = (Prob_Sand(k, i) + 0) / MCrun
                'Prob_Silt(k, i) = (Prob_Silt(k, i) + 0) / MCrun
                Prob_Clay(k, i) = (Prob_Clay(k, i) + 1)  'Median(k, i)
            Else
                'Prob_Sand(k, i) = (Prob_Sand(k, i) + 0) / MCrun
                Prob_Silt(k, i) = (Prob_Silt(k, i) + 1) 'Median(k, i)
                'Prob_Clay(k, i) = (Prob_Clay(k, i) + 0) / MCrun
            End If
        End If
        
    Next i
Next k

If MCrun = 5 And sim_time = end_of_times Then
'    Prob_Sand(k, i) = Prob_Sand(k, i)
'    Prob_Silt(k, i) = Prob_Silt(k, i)
'    Prob_Clay(k, i) = Prob_Clay(k, i)
End If

End Sub '
Private Sub write_data_to_screen()
    
    Form1.pic1.Cls
    Form1.pic1.Scale (0, 280)-(700, 0)
'// usually a initial effect exist which you do not want to display
    
    s = 0.75
    l = 0.5
    
    Rmarg = 36
    Tmarg = 10
    Bmarg = 50
    Lmarg = 36
    If end_of_times <= 500 Then
        timeln_intrval = end_of_times / 5
    ElseIf end_of_times <= 2000 Then
        timeln_intrval = end_of_times / 10
    Else
        timeln_intrval = end_of_times / 20
    End If
'* 30 pts for vertical scale bar    */
'    /* draw permeability color coding bar */
'    /* first determine perm range */
    datamaxx = -1000000
    datamax = datamaxx
    dataminn = 1000000
    datamin = dataminn
  
    For k = 0 To end_of_times - 1 '
        For i = 0 To nx - 1
            PS_data(k, i) = Median(k, i)
        Next i
    Next k
    For k = 0 To end_of_times
        For i = xstrtps To xstopps ' xstopps
            dataminn = Minimum(PS_data(k, i), datamin)
            datamin = dataminn
            datamaxx = Maximum(PS_data(k, i), datamax)
            datamax = datamaxx
        Next i
    Next k
    ddata = (datamaxx - dataminn) / 10
        
    For i = 1 To 10
    
        huedata = (0.7 - 0.7 * i / 10)            ';    // hue runs from 0(red) to 0.75(blue)
        
        If i <= 5 Then
            r1 = 0
        ElseIf i <= 8 Then
            r1 = (255 * ((i - 5) / 3))
        Else
            r1 = 255
        End If
        
        
        If i <= 4 Then
            g1 = 255 * ((i) / 4)
        ElseIf i >= 4 And i <= 8 Then
            g1 = 255
        Else
            g1 = 255 - (255 * ((i - 8) / 2))
        End If
        
        If i >= 5 Then
            b1 = 0
        ElseIf i <= 3 Then
            b1 = 255
        Else
           b1 = 255 - (255 * ((i - 2) / 3))
        End If
          
        Form1.pic1.Line (Lmarg + (i * 5) + 10, 35)-(Lmarg + (i * 5) + 15, 25), RGB(r1, g1, b1), BF
          
        'pic1.Line (((i + 1) * 25) + 10, 15)-((i + 1) * 25 + 10, 0)
               
        datamindraw = i * ddata
               
        datamindraw = FormatNumber((datamindraw), 2)
        
        If i = 1 Then
            Form1.pic1.PSet (Lmarg + (i * 5) + 5, 44)
            Form1.pic1.Print datamindraw
        'ElseIf i = 5 Then
        '   Form1.pic1.PSet (Lmarg + (i * 5) + 5, 44)
        '   Form1.pic1.Print datamindraw
        ElseIf i = 10 Then
            Form1.pic1.PSet (Lmarg + (i * 5) + 5, 44)
            Form1.pic1.Print datamindraw
            
            Form1.pic1.PSet (Lmarg, 25)
            Form1.pic1.Print "median grainsize classes in mm"
        End If
    Next i
    '/* Determine x & y scale */
    hhmax = -1000000
    hmax = hhmax
    hhmin = 1000000
    hmin = hhmin
    For i = 1 To (end_of_times - dt)  'nIterations - (nIterations - Iteration)
        For j = xstrtps To xstopps   '500
            hmin = Minimum(StratNode(i, j), hmin)
            hmax = Maximum(StratNode(i, j), hmax)
        Next j
    Next i
     scy = (280 - 60 - (Tmarg + Bmarg)) / (hmax - hmin) '//333
     scx = (700 - (Lmarg + Rmarg)) / (xstopps - xstrtps)
    
    '/* Draw vertical scale bar (ruler)  */
    hhmin = sediment(0, xstopps, 0) '- model_sealevel  'hmin
    hhmax = (0.5 + hmax) '- model_sealevel
    Form1.pic1.Line ((Lmarg + 2), (Bmarg + scy * ((topo_temp(xstopps) - hhmin))))-(Lmarg + 2, (Bmarg + scy * (topo_temp(xstrtps) - hhmin)))
    Form1.pic1.Refresh
    For i = hhmin To hhmax
       If (i Mod 5 = 0) Then
            Form1.pic1.Line ((Lmarg), (Bmarg + scy * ((topo_temp(xstopps) - hhmin) - (hhmin - i))))-(Lmarg + 5, (Bmarg + scy * ((topo_temp(xstopps) - hhmin - (hhmin - i)))))
            Form1.pic1.Refresh
            If (i Mod 10 = 0) Then
                Form1.pic1.PSet ((Lmarg - 10), (Bmarg + scy * ((topo_temp(xstopps) - hhmin - (hhmin - i))) + 2))
                Form1.pic1.Print FormatNumber(i - model_sealevel, 0)
            End If
        Else

        End If
    Next i
    '/* draw horizontal gridline */
    
    For k = 1 To (end_of_times - 1)
   
       For i = xstrtps To xstopps
            If Median(k, i) = 0 Then
            
            Else
                hueD = 0.7 - (0.7 * (Median(k, i) - grain_size(1)) / ((10 * ddata) - grain_size(1)))
                
                If s < 0.001 Then
                    r1 = 255 * l
                    g1 = r1
                    b1 = r1
                Else
                    If l < 0.5 Then
                        temp2 = l * (1 + s)
                    Else
                        temp2 = l + s - l * s
                    End If
    
                    temp1 = 2 * l - temp2
                    temp3R = hueD + 0.33333
            
                    If temp3R > 1 Then
                        temp3R = temp3R - 1
                    End If
            
                    temp3G = hueD
                    temp3B = hueD - 0.33333
          
                    If temp3B < 0 Then
                        temp3B = temp3B + 1
                    End If
    
                    r1 = HuetoColorVal(temp3R, temp1, temp2)
                    If r1 < 0 Then r1 = 0
                    If r1 > 255 Then r1 = 255
                    g1 = HuetoColorVal(temp3G, temp1, temp2)
                    If g1 < 0 Then g1 = 0
                    If g1 > 255 Then g1 = 255
                    b1 = HuetoColorVal(temp3B, temp1, temp2)
                    If b1 < 0 Then b1 = 0
                    If b1 > 255 Then b1 = 255
    
                End If
            
                xcolortik = (Lmarg + scx * (i - xstrtps))
                ycolortik = (Bmarg + scy * (StratNode(k, i) - hmin))
                xcolortik1 = (Lmarg + scx * (i + 1 - xstrtps))
                ycolortik1 = (Bmarg + scy * (StratNode(k - 1, i + 1) - hmin))
                
                Form1.pic1.Line (xcolortik, ycolortik)-(xcolortik1, ycolortik1), RGB(r1, g1, b1), BF
            End If
            '/* put tickmarks every x-gridcells and labels every other x-gridcells   */
            If (k = 1) Then
                xtik = (Lmarg + scx * (i - xstrtps))
                ytik = Bmarg - 5 + scy * (StratNode(k, i) - hmin)
                If (i Mod 50 = 0) Then
                    Form1.pic1.PSet (xtik - 5, ytik + 3)
                    Form1.pic1.Print (i / 10)
                ElseIf (i Mod 10 = 0) And (i Mod 50 <> 0) Then
                    Form1.pic1.Line (xtik, ytik)-(xtik, ytik + 5)
                End If
           
            End If
       
        Next i
        
       
           
         ' If k = end_of_times - 1 Then
                
         '       For z = 1 To 4
         '       Form10.Picture4.PSet (1 + 10 * z, 0.005)
         '       Form10.Picture4.Print FormatNumber(delta_grad(z), 5)
         '       Form10.Picture4.PSet (1 + 10 * z, 0.004)
         '       If z = 1 Then
         '             Form10.Picture4.Print z * 100 & "m."
         '       ElseIf z = 2 Then
         '            Form10.Picture4.Print z * 250 & "m."
         '       ElseIf z = 3 Then
         '           Form10.Picture4.Print "1000 m."
         '       Else
         '           Form10.Picture4.Print "5000 m."
         '       End If
         '   Next z
         ' End If
        
    Next k
   
   
   
   
   '/* Draw sealevel
            Form1.pic1.Line ((Lmarg + scx * (flx - xstrtps)), (Bmarg + scy * (model_sealevel - hmin)))-((Lmarg + scx * (xstopps)), (Bmarg + scy * (model_sealevel - hmin))), RGB(0, 0, 255), BF
            
    '/* draw some timelines in stratigraphy(if required) */
    
    For i = 0 To end_of_times '- (200 - i) 'nIterations - (nIterations - Iteration)
        'Form2.ProgressBar2(1).Value = 40 + ((i / end_of_times) * 60)
        If (timeln_intrval <> -1 And i <> 200 And i Mod timeln_intrval = 0) Then
            For j = xstrtps To xstopps
            
            '/*act_h_history(i,j)*/' Or i = 199
                
                Form1.pic1.Line ((Lmarg + scx * (j - xstrtps)), (Bmarg + scy * (StratNode(i, j) - hmin)))-((Lmarg + scx * (j + 1 - xstrtps)), (Bmarg + scy * (StratNode(i, j + 1) - hmin)))
                
            ' pic1.Refresh
            Next j
        End If
      
    Next i
    '/* always draw final timeline */
    
    For j = xstrtps To xstopps
    
           Form1.pic1.Line ((Lmarg + scx * (j - xstrtps)), (Bmarg + scy * (StratNode(end_of_times - 1, j) - hmin)))-((Lmarg + scx * (j + 1 - xstrtps)), (Bmarg + scy * (StratNode(end_of_times - 1, j + 1) - hmin)))
        
    Next j

End Sub '
Public Sub Main()

'/////////////////////////////////////////////////////////////////////////////////////////////////
'// input van Parameter_and_constants
'/////////////////////////////////////////////////////////////////////////////////////////////////]
Do While (MCrun < 205 + 1)
Parameter_and_constants
Randomize
load_Array

'/////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////
'//start of simualtion
'/////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////

Initialize_simulation

'/////////////////////////////////////////////////////////////////////////////////////////////////
'Form2.Pic2.Print "Initialize initial surface"
'/////////////////////////////////////////////////////////////////////////////////////////////////

getdischarge

'/////////////////////////////////////////////////////////////////////////////////////////////////
'// START TIME_LOOP
'/////////////////////////////////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////////////////////////////////////////
Do While (sim_time < end_of_times)
Form1.ProgressBar1(0).Value = (sim_time / end_of_times) * 100
initialize_sealevel

'////////////////// Erosion calculations ///////////////////////////////////////////////////////
'//////////////// Assumption: erosion is equal for all grainsizes//////////////////////////////
    ersosion_calc
    deposition_calc

'deposition_calc
'//////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////// Deposition calculations ////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////

'/ data_writing
 
sim_time = sim_time + dt '// time step gets increased here
'Close #14
Loop
'/////////////////////////////////////////////////////////////////////////////////////////////////
'// END TIME_LOOP   '
'/////////////////////////////////////////////////////////////////////////////////////////////////
'//  Write data
'/////////////////////////////////////////////////////////////////////////////////////////////////

 Beep
'/////////////////////////////////////////////////////////////////////////////////////////////////
store_strat
store_texture
If MCrun >= 205 Then
write_data_to_screen
End If
'

MCrun = MCrun + 1
Loop
'/////////////////////////////////////////////////////////////////////////////////////////////////
'Form2.Pic2.Print "end of programm"
'/////////////////////////////////////////////////////////////////////////////////////////////////
'//end of simulation
'//////////////////////////////////////////////////////////////////////////////////////////////
End Sub

Private Sub load_Array()

ReDim wheeler_tot_thick(0 To nx)
ReDim dh_depo(0 To nx, 0 To num_of_gscl)
'ReDim av_depo_rate(0 To nx)

ReDim sediment(0 To end_of_times, 0 To nx, 0 To num_of_gscl)
ReDim sedflux(0 To nx, 0 To num_of_gscl)
ReDim SLinput(0 To end_of_times)
ReDim discharge(0 To end_of_times, 0 To nx)
ReDim waterdepth(0 To nx)
ReDim erosion_rate(0 To nx)
ReDim PS_data(0 To end_of_times, 0 To nx)
ReDim thickness(0 To nx)
ReDim StratNode(0 To end_of_times, 0 To nx)
ReDim Median(0 To end_of_times, 0 To nx)

If MCrun = 1 Then
    ReDim Prob_Sand(0 To end_of_times, 0 To nx)
    ReDim Prob_Silt(0 To end_of_times, 0 To nx)
    ReDim Prob_Clay(0 To end_of_times, 0 To nx)
End If

ReDim topo_old(0 To nx)
ReDim topo_temp(0 To nx)
ReDim topo_calc(0 To nx)
ReDim topo_depo(0 To nx)
ReDim dh_erosion(0 To nx)
ReDim settle_rate_wheeler(0 To nx)
ReDim slope(0 To nx)

End Sub

'
Private Sub draw_well()
       
    Frm_WELL.Slider_WellPos.Min = xstrtps
    Frm_WELL.Slider_WellPos.Max = xstopps
    
    If well_counter = 0 Then
          
        '/* Determine x & y scale */
    hhmax = -1000000
    hmax = hhmax
    hhmin = 1000000
    hmin = hhmin
    For i = 1 To (end_of_times - dt)  'nIterations - (nIterations - Iteration)
        For j = xstrtps To xstopps   '500
            hmin = Minimum(StratNode(i, j), hmin)
            hmax = Maximum(StratNode(i, j), hmax)
        Next j
    Next i
     
    '/* Draw vertical scale bar (ruler)  */
    
    hhmin = sediment(0, xstopps, 0)     '- model_sealevel  'hmin
    hhmax = (0.5 + hmax)                '- model_sealevel
    
    Frm_WELL.View_WellY.Cls
    Frm_WELL.View_WellY.Scale (0, topo_temp(xstopps) - hhmin)-(25, topo_temp(xstrtps) - hhmin)
    Frm_WELL.View_WellY.Line (2, topo_temp(xstopps) - hhmin)-(2, (topo_temp(xstrtps) - hhmin))
    
    
    For i = hhmin To hhmax
       If (i Mod 5 = 0) Then
            Frm_WELL.View_WellY.Line ((2), ((topo_temp(xstopps) - hhmin) - (hhmin - i)))-(2 + 5, ((topo_temp(xstopps) - hhmin - (hhmin - i))))
            
            If (i Mod 10 = 0) Then
                Frm_WELL.View_WellY.PSet (10, (((topo_temp(xstopps) - hhmin - (hhmin - i))) + 2))
                Frm_WELL.View_WellY.Print FormatNumber(hhmin - i, 0)
            End If
        Else

        End If
    Next i
    Frm_WELL.View_WellY.Refresh
        
        
        
        Frm_WELL.View_WellX.Scale (0, 10)-(grain_size(num_of_gscl) + 20, 0)
        Frm_WELL.View_WellX.Line (0, 9)-(grain_size(num_of_gscl) + 20, 9)
         Frm_WELL.View_WellX.Refresh
        For i = 1 To num_of_gscl 'grain_size(1) To grain_size(num_of_gscl)
            
                Frm_WELL.View_WellX.Line ((i / num_of_gscl) * 20, 9)-((i / num_of_gscl) * 20, 6)
                
                    Frm_WELL.View_WellX.PSet (((i / num_of_gscl) * 20), 5)
                    Frm_WELL.View_WellX.Print grain_size(i)
                
            
        Next i
        
        Frm_WELL.View_Well.Scale (grain_size(1), topo_temp(xstrtps))-(grain_size(num_of_gscl), topo_temp(xstopps))
        Xpos = flx
        Frm_WELL.Slider_WellPos.Value = ((flx - xstrtps)) '/ (xstopps - xstrtps)) * 500
        well_counter = 1
        
    Else
        
        Xpos = xstrtps + ((Frm_WELL.Slider_WellPos.Value / nx) * (xstopps - xstrtps))
    End If
    
    Frm_WELL.Text1 = (Frm_WELL.Slider_WellPos.Value / 10)
    Frm_WELL.Caption = "Well No" & (Frm_WELL.Slider_WellPos.Value / 10) & " km."
        Frm_WELL.View_Well.Cls
        
        s = 0.9
        l = 0.5
    
        For k = 2 To end_of_times - 2
             If Median(k, Xpos) = 0 Then
             Else
             'hueD = 0.7 - (0.7 * ((Median(i, i) - grain_size(1)) / ((10 * ddata) - grain_size(1))))
             hueD = 0.7 - (0.7 * (Median(k, Xpos) - grain_size(1)) / ((10 * ddata) - grain_size(1)))
                If s < 0.001 Then
                    r1 = 255 * l
                    g1 = r1
                    b1 = r1
                Else
                    If l < 0.5 Then
                        temp2 = l * (1 + s)
                    Else
                        temp2 = l + s - l * s
                    End If
    
                    temp1 = 2 * l - temp2
                    temp3R = hueD + 0.33333
            
                    If temp3R > 1 Then
                        temp3R = temp3R - 1
                    End If
            
                    temp3G = hueD
                    temp3B = hueD - 0.33333
          
                    If temp3B < 0 Then
                        temp3B = temp3B + 1
                    End If
    
                    r1 = HuetoColorVal(temp3R, temp1, temp2)
                    If r1 < 0 Then r1 = 0
                    If r1 > 255 Then r1 = 255
                    g1 = HuetoColorVal(temp3G, temp1, temp2)
                    If g1 < 0 Then g1 = 0
                    If g1 > 255 Then g1 = 255
                    b1 = HuetoColorVal(temp3B, temp1, temp2)
                    If b1 < 0 Then b1 = 0
                    If b1 > 255 Then b1 = 255
    
                End If
            End If
    
            Frm_WELL.View_Well.Line (grain_size(1), StratNode(k - 1, Xpos))-(Median(k, Xpos), StratNode(k, Xpos)), RGB(r1, g1, b1), BF
        
        Next k

End Sub

