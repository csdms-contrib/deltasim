Attribute VB_Name = "Module5"
Option Explicit

Sub write_psfile()
'*******************************
'*** Get minimum of three values
'*******************************
      FILE_NAME10 = "D:\Temp\vba_ps.ps"
      Open FILE_NAME10 For Output As #10
   
   
   'FILE_NAMEA = "D:\Temp\alpha.dat"
   'Open FILE_NAMEA For Output As #a
   
    Print #10, "%!"
    Print #10, "gsave"
    
    '/* Determine x & y scale */
    hhmax = -1000000
    hmax = hhmax
    hhmin = 1000000
    hmin = hhmin
    For i = 1 To end_of_times '- dx 'nIterations - (nIterations - Iteration)
        For j = xstrtps To xstopps   '500
            hmin = Minimum(StratNode(i, j), hmin)
            hmax = Maximum(StratNode(i, j), hmax)
        Next j
    Next i

     scy = 280 '(280 - 60 - (Tmarg + Bmarg)) / (hmax - hmin) '//333
     scx = 700 '(700 - (Lmarg + Rmarg)) / (xstopps - xstrtps)

    
    '/* Draw vertical scale bar (ruler)  */
    
      
    hhmin = sediment(0, xstopps, 0) - sealevel  'hmin
    hhmax = (0.5 + hmax - sealevel)
     Print #10, "gsave"
     Print #10, "-20 0 translate"   ' /* draw scale bar left of Lmarg */
     Print #10, Format(Lmarg, "#####0"); " "; Format(Bmarg + 1, "#####0"); " "; "moveto"
     Print #10, "/Helvetica findfont 10 scalefont setfont"
    
   
        
     For i = hhmin To hhmax
       
        Print #10, Format(Lmarg, "#####0"); " "; Format(Bmarg + scy * (i - hhmin), "#####0"); " "; "moveto"
       
       If (i Mod 5 = 0) Then
             Print #10, "5 0 rlineto"    '/*scx was 10*/
             Print #10, "stroke"
    
        
                       
             If (i Mod 10 = 0) Then
             Print #10, "0 0 0 setrgbcolor"
              Print #10, Format(Lmarg, "#####0"); " "; Format(Bmarg + scy * (i - hhmin), "#####0"); " "; "moveto"
                 Print #10, "-20 -4 rmoveto"
             Print #10, "("; i; ")"; " "; "show"
      
               
            End If
        
        Else
        
    '//      fprintf(psfp,"3 0 rmoveto\n");   //small ticks
    '//      fprintf(psfp,"2 0 rlineto\n");
    '//      fprintf(psfp,"stroke\n");
         End If
       
        
     Next i
    
    Print #10, "grestore"
     '/* draw horizontal gridline
           
           Print #10, "gsave"
                Print #10, "0 0 0 setrgbcolor"
                Print #10, "0.5 setlinewidth"
           For j = xstrtps To xstopps
            i = 0
               '/* put tickmarks every x-gridcells and labels every other x-gridcells   */
         If (i = 0) Then
                
                
            
                    
                    If (j Mod 50 = 0) Then
      
                
                Print #10, Format((Lmarg + scx * (j - xstrtps)), "#####0.0"); " "; Format((Bmarg - 5 + scy * (StratNode(i, j) - hmin)), "#####0.0"); " "; " moveto"
                        
                        Print #10, "0 -5 rlineto"
                         Print #10, "-4 -12 rmoveto"
                         Print #10, "/Helvetica findfont 10 scalefont setfont"
                         Print #10, "("; j / 10; ") "; " "; "show" '* dx) / 1000)
                        
                ElseIf (j Mod 5 = 0) And (j Mod 50 <> 0) Then
                      Print #10, Format((Lmarg + scx * (j - xstrtps)), "#####0.0"); " "; Format((Bmarg - 5 + scy * (StratNode(i, j) - hmin)), "#####0.0"); " "; " moveto"
                       Print #10, "0 -5 rlineto"
                            Print #10, "stroke"
                           Print #10, "grestore"
                          'Print #10, "4 12 rmoveto"
                    End If
   
            End If
   
      
        Next j
    

    
    
    '/* draw some timelines in stratigraphy(if required) */
    Print #10, "0.1 setlinewidth"
    Print #10, "0 0 0 setrgbcolor"
    For i = 0 To end_of_times '- (200 - i) 'nIterations - (nIterations - Iteration)
        
        If (timeln_intrval <> -1 And i <> end_of_times And i Mod timeln_intrval = 0) Then
            For j = xstrtps To xstopps
            
            '/*act_h_history(i,j)*/' Or i = 199
                
                Print #10, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(i, j) - hmin), "#####0.0"); " "; "moveto"
                Print #10, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(i, j + 1) - hmin), "#####0.0"); " "; "lineto"
                Print #10, "stroke"
                
               
            ' pic1.Refresh
            Next j
        End If
      
    Next i
    '/* always draw final timeline */
    
    Print #10, "0.1 setlinewidth"
    Print #10, "0 0 0 setrgbcolor"
    For j = xstrtps To xstopps
    
        Print #10, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(end_of_times, j) - hmin), "#####0.0"); " "; "moveto"
        Print #10, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(end_of_times, j + 1) - hmin), "#####0.0"); " "; "lineto"
        Print #10, "stroke"
       
    Next j
    
    
   
   ' (Format(Lmarg + scx * (j - xstrtps), "#####0.0"),  Format(Bmarg + scy * (StratNode(Iteration, j) - hmin), "#####0.0")- _
    (Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0") , Format(Bmarg + scy * (StratNode(Iteration, j + 1) - hmin), "#####0.0")))
    
    '/* always draw first timeline */
    Print #10, "0.1 setlinewidth"
    Print #10, "0 0 0 setrgbcolor"
    For j = xstrtps To xstopps
    
        Print #10, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(1, j) - hmin), "#####0.0"); " "; " moveto"
        Print #10, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(1, j + 1) - hmin), "#####0.0"); " "; "lineto"
        Print #10, "stroke"
       Next j

  
    Print #10, "showpage"
    

    Close #10
' /*    end of write_perm_psfile()  */

'GradientSquare Picture1, 300, 300, RGB(0, 0, 255), RGB(255, 0, 0), 255, True

    
   'Form2.Pic2.Print " "

'Close #a
End Sub
'*******************************
'*** Get minimum of three values
'*******************************

Sub draw_real_time_line()
   
write_psfile




If FT = 0 Then
   
    Print #11, "%!"
    Print #11, "gsave"
    FT = 1
End If
    
   Print #11, "0.1 setlinewidth"
    Print #11, "0 0 0 setrgbcolor"
    For j = xstrtps To xstopps
    
        Print #11, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (topo_temp(j) - hmin), "#####0.0"); " "; "moveto"
        Print #11, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (topo_temp(j + 1) - hmin), "#####0.0"); " "; "lineto"
        Print #11, "stroke"
       
    Next j
   
End Sub

Sub write_perm_psfile()

    write_psfile
    
    FILE_NAME22 = "D:\Temp\ps_time_kura_2.ps" 'psfpvbclay.ps" '" & scenario & ".ps"
    Open FILE_NAME22 For Output As #22
    
    'char f_name[255];
    xstrtps = 250          '// usually a initial effect exist which you do not want to display
    xstopps = 950
    timeln_intrval = 20
    Rmarg = 16
    Tmarg = 10
    Bmarg = 50
    Lmarg = 36
              '          /* 30 pts for vertical scale bar    */
    Print #22, "%!"
    Print #22, "gsave"

'    /* draw permeability color coding bar */
'    /* first determine perm range */
    datamaxx = -1000000
    datamax = datamaxx
    dataminn = 1000000
    datamin = dataminn
    
    For i = 0 To end_of_times - 1 '- (200 - i) 'nIterations - (nIterations - Iteration)
        For j = 0 To 950
            PS_data(i, j) = Median(i, j) 'Prob_Silt(i, j) 'Prob_Clay(i, j) 'Prob_Sand(i, j) 'Median(i, j)
        Next j
    Next i
    
    For i = 0 To end_of_times 'nIterations - (nIterations - Iteration)
        For j = xstrtps To 950 ' xstopps
            dataminn = Minimum(PS_data(i, j), datamin)
            datamin = dataminn
            datamaxx = Maximum(PS_data(i, j), datamax)
            datamax = datamaxx
        Next j
    Next i
    'ddata = (datamaxx - dataminn) / 9
    ddata = 0.1
    
    Print #22, "gsave"
    'datamax = 0.5
    'ddata = (datamax - datamin) / 9
    Print #22, Format(Lmarg + 25, "#####0"); " "; Format(Bmarg - 15, "#####0"); " "; " translate"
   'Print #22,   '//positioning of the scalebar
    Print #22, "/Helvetica findfont 10 scalefont setfont"

    '/* Now draw scalebar 10 classes */
    Print #22, "0 0 0 setrgbcolor"
    Print #22, "0.2 setlinewidth"

    Print #22, "10 5 moveto"
    Print #22, "0 -10 rlineto"
    Print #22, "stroke"
    
    
    For i = 0 To 9
    
        huedata = (0.8 * i / 9)  '(0.7 - 0.7 * i / 9)            ';    // hue runs from 0(red) to 0.75(blue)
       
        Print #22, Format((i * 25) + 10, "#####0"); " "; "0 moveto"
        
        Print #22, "0 15 rlineto"
        Print #22, "25 0 rlineto"
        Print #22, "0 -15 rlineto"
        Print #22, "-25 0 rlineto"         ' //boxfill
                                                                             'fprintf(psfp,"%5.3f 1.0 1.0 sethsbcolor\n fill\n",hueperm); //COLOUR
        Print #22, Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; "0 setcmykcolor"
        
        'Print #22, Format(huedata, "#####0.000"); " "; "1.0 1.0 sethsbcolor" 'Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; "0 setcmykcolor" '
        'Print #22, Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; "0 setcmykcolor" '//B&W
        
        Print #22, "fill"
       
        Print #22, "0 0 0 setrgbcolor"
        Print #22, Format((i * 25) + 10, "#####0"); " "; "0 moveto"
        Print #22, "0 15 rlineto"
        Print #22, "25 0 rlineto"
        Print #22, "0 -15 rlineto"
        Print #22, "-25 0 rlineto"
        Print #22, "stroke " '//box
        
        Print #22, Format(((i + 1) * 25) + 10, "#####0"); " "; "0 moveto"
        Print #22, "0 -5 rlineto"
        Print #22, "stroke" ',;  //divider line"
        
        Print #22, Format((i * 25) + 10, "#####0"); " "; "0 moveto" '\n        -10 -15 rmoveto\         (%8.0f) show\n        ",(i*25)+10,Permmin+i*dperm)      ;//value
        Print #22, "-10 -15 rmoveto"
        datamin = i * ddata
        Print #22, "("; Format(datamin, "#####0.00"); ")"; " "; "show" '//value
       
    Next i
    
    Print #22, "grestore"
    Print #22, "0.4 setlinewidth"
    Print #22, "0.0 setgray"
    
    '/* Determine x & y scale */
    hhmax = -1000000
    hmax = hhmax
    hhmin = 1000000
    hmin = hhmin
    For i = 1 To end_of_times - 1 '- dx 'nIterations - (nIterations - Iteration)
        For j = xstrtps To xstopps   '500
            hmin = Minimum(StratNode(i, j), hmin)
            hmax = Maximum(StratNode(i, j), hmax)
        Next j
    Next i

     scy = (280 - 60 - (Tmarg + Bmarg)) / (hmax - hmin) '//333
     scx = (700 - (Lmarg + Rmarg)) / (xstopps - xstrtps)

    '/* Draw vertical scale bar (ruler)  */
    hhmin = hmin - sealevel
    hhmax = (0.5 + hmax - sealevel)
    Print #22, "gsave"
    Print #22, "-20 0 translate"   ' /* draw scale bar left of Lmarg */

    Print #22, Format(Lmarg, "#####0"); " "; Format(Bmarg + 1, "#####0"); " "; "moveto"
    Print #22, "/Helvetica findfont 10 scalefont setfont"
    For i = hhmin To hhmax
       Print #22, Format(Lmarg, "#####0"); " "; Format(Bmarg + scy * (i - hhmin), "#####0"); " "; "moveto"
       
       If (i Mod 5 = 0) Then
            Print #22, "5 0 rlineto"    '/*scx was 10*/
            Print #22, "stroke"
            
            If (i Mod 10 = 0) Then
                Print #22, "0 0 0 setrgbcolor"
                Print #22, Format(Lmarg, "#####0"); " "; Format(Bmarg + scy * (i - hhmin), "#####0"); " "; "moveto"
                Print #22, "-20 -4 rmoveto"
                Print #22, "("; i; ")"; " "; "show"
            End If
        
        Else
        
    '//      fprintf(psfp,"3 0 rmoveto\n");   //small ticks
    '//      fprintf(psfp,"2 0 rlineto\n");
    '//      fprintf(psfp,"stroke\n");
        End If
    
    Next i
    Print #22, "grestore"

    '/* draw horizontal gridline */

    For i = 0 To end_of_times - 1 '- (200 - i) 'nIterations - (nIterations - Iteration)
    
        For j = xstrtps To xstopps
        
          huedata = 0.8 * (PS_data(i, j) - dataminn) / (datamaxx - dataminn) '0.7 - 0.7 * (PS_data(i, j) - dataminn) / (datamaxx - dataminn) ';/* COLOUR 0.7=rood ..0.35=groen .. 0.0=blauw*/
            
          'huedata = 0.7 - 0.7 * (PS_data(i, j) - dataminn) / (datamax - dataminn)  '/*B&W*/
          
            
            If (i > 0) Then 'And (sediment(sim_time, i, j) < EPS And (sediment(sim_time, i, j + 1) < EPS)) Then
'//          if (i>0 && !(sediment[nsc][i][j] < EPS))
            
'                /*  Draw a cell. nodestrat[i][j] is upper left corner,nodestrat[i+1][j+1] the lower right. */
             '   If (sediment(i, j, 0) < EPS) Then  '/* zero thickness -> borrow perm from neighbour */
                '//  huedata = 0.70 - 0.70*(PS_data[i][j+1]-datamin)/(datamax-datamin);//COLOUR
           '         huedata = 0.8 - 0.6 * (PS_data(i, j + 1) - dataminn) / (datamax - dataminn) ';/*B&W*/
              '  End If
                   'Print #22, Format(huedata, "#####0.000"); " "; " 1.0 1.0 sethsbcolor"  ' //COLOUR
               '     Print #22, Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; Format(huedata, "#####0.000"); " "; "0 setcmykcolor" '//B&W
               '     Print #22, Format((Lmarg + scx * (j - xstrtps)), "#####0.0"); " "; Format((Bmarg + scy * (StratNode(i, j) - hmin)), "#####0.0"); " "; "moveto"
              ''       Print #22, Format((Lmarg + scx * (j + 1 - xstrtps)), "#####0.0"); " "; Format((Bmarg + scy * (StratNode(i, j + 1) - hmin)), "#####0.0"); " "; "lineto"
               '     Print #22, Format((Lmarg + scx * (j + 1 - xstrtps)), "#####0.0"); " "; Format((Bmarg + scy * (StratNode(i - 1, j + 1) - hmin)), "#####0.0"); " "; "lineto"
               '     Print #22, Format((Lmarg + scx * (j - xstrtps)), "#####0.0"); " "; Format((Bmarg + scy * (StratNode(i - 1, j) - hmin)), "#####0.0"); " "; "lineto"
               '     Print #22, Format((Lmarg + scx * (j - xstrtps)), "#####0.0"); " "; Format((Bmarg + scy * (StratNode(i, j) - hmin)), "#####0.0"); " "; "lineto"
               '     Print #22, "fill"
                    
                    'Worksheets("sheet6").Cells(j, i) = StratNode(i, j)
            
             '/* put tickmarks every x-gridcells and labels every other x-gridcells   */
            ElseIf (i = 0) Then
                Print #22, "gsave"
                Print #22, "0 0 0 setrgbcolor"
                Print #22, "0.5 setlinewidth"
                Print #22, Format((Lmarg + scx * (j - xstrtps)), "#####0.0"); " "; Format((Bmarg - 5 + scy * (StratNode(i, j) - hmin)), "#####0.0"); " "; " moveto"
                
                'If (i Mod 5 = 0) Then
                 '   Print #22, "0 -5 rlineto"    '/*scx was 10*/
                  '  Print #22, "stroke"
                    If (j Mod 50 = 0) Then
                        Print #22, "0 -5 rlineto"
                        Print #22, "-4 -12 rmoveto"
                        Print #22, "/Helvetica findfont 10 scalefont setfont"
                        Print #22, "("; j / 10; ") "; " "; "show" '* dx) / 1000)
                        
                    ElseIf (j Mod 5 = 0) And (j Mod 50 <> 0) Then
                           
                           Print #22, "0 -5 rlineto"
                           Print #22, "stroke"
                           Print #22, "grestore"
                          'Print #22, "4 12 rmoveto"
                    End If
                'Else '// fprintf(psfp,"0 -5 rlineto\n");

                'Print #22, "stroke"
                'Print #22, "grestore"
            End If
            
        
        Next j
    
    Next i
    '/* draw some timelines in stratigraphy(if required) */
    Print #22, "0.1 setlinewidth"
    Print #22, "0 0 0 setrgbcolor"
    For i = 0 To 3200 - 1 '- (200 - i) 'nIterations - (nIterations - Iteration)
        
       For j = xstrtps To xstopps
           If (timeln_intrval <> -1 And i Mod timeln_intrval = 0) Or i = 199 Then
            '/*act_h_history(i,j)*/'
               Print #22, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(i, j) - hmin), "#####0.0"); " "; "moveto"
               Print #22, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(i, j + 1) - hmin), "#####0.0"); " "; "lineto"
               Print #22, "stroke"
                
           End If
       Next j
    Next i
    '/* always draw final timeline */
    Print #22, "0.1 setlinewidth"
    Print #22, "0 0 0 setrgbcolor"
    For j = xstrtps To j < xstopps
    
        Print #22, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(end_of_times - 1, j) - hmin), "#####0.0"); " "; "moveto"
        Print #22, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(end_of_times - 1, j + 1) - hmin), "#####0.0"); " "; "lineto"
        Print #22, "stroke"
    
    Next j

    '/* always draw first timeline */
    Print #22, "0.1 setlinewidth"
    Print #22, "0 0 0 setrgbcolor"
    For j = xstrtps To xstopps
    
        Print #22, Format(Lmarg + scx * (j - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(1, j) - hmin), "#####0.0"); " "; " moveto"
        Print #22, Format(Lmarg + scx * (j + 1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(1, j + 1) - hmin), "#####0.0"); " "; "lineto"
        Print #22, "stroke"
    
    Next j



    Print #22, "showpage"
    

    Close #22
' /*    end of write_perm_psfile()  */



End Sub
Sub well()
    '/* Draw position of the three wells + perm log  */
 Print #22, "0 0 0 setrgbcolor"
 Print #22, "0.2 setlinewidth"
 
 If (well1 >= xstrtps And well1 <= xstopps) Then

     Print #22, Format(Lmarg + scx * (well1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(end_of_times - 1, well1) - hmin), "#####0.0"); " "; " moveto"
    'Print #22, "moveto", Lmarg + scx * (wellpos1 - xstrtps), Bmarg + scy * (StratNode(end_of_times - 1, wellpos1 - hmin)); ""
     Print #22, Format(Lmarg + scx * (well1 - xstrtps), "#####0.0"); " "; Format(Bmarg + scy * (StratNode(1, well1) - hmin), "#####0.0"); " "; " lineto"
    'print #22,"%5.1f %5.1f lineto\n",Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[0][wellpos1]-hmin));
     Print #22, Format(0.1 * (PS_data(1, well1) - 5) + Lmarg + scx * (well1 - xstrtps), "####0.0"); " "; Format(Bmarg + scy * (StratNode(1, well1) - hmin), "#####0.0"); " "; " lineto"
    'print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[1][wellpos1])-5)+Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[0][wellpos1]-hmin));
     Print #22, "stroke"
    'print #22,"stroke\n");// the above 3 line plot the well position and the lowerboundary -'L'line
     Print #22, Format(0.1 * (PS_data(1, well1) - 5) + Lmarg + scx * (well1 - xstrtps), "####0.0"); " "; Format(Bmarg + scy * (StratNode(0, well1) - hmin), "#####0.0"); " "; " moveto"
    'print #22,"%5.1f %5.1f moveto\n",(0.10*(PS_data[1][wellpos1])-5)+Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[0][wellpos1]-hmin)); //goto lower right position were the line has just been drawn

        For i = 1 To i = end_of_times - 1
'//      for (i=1;i<nIterations-1;i++)

'//      {
             thickness(0) = StratNode(i, wellpos1) - StratNode(i - 1, well1)

'//          thickness = StratNode[i][wellpos1]-StratNode[i-1][wellpos1];
             If thickness(0) < 0.01 Then
'//          if (thickness < 0.01)
                 PS_data(i, well1) = PS_data(i - 1, well1)
    '//          PS_data[i][wellpos1] = PS_data[i-1][wellpos1];
                 Print #22, Format(0.1 * (PS_data(i, well1) - 5) + Lmarg + scx * (well1 - xstrtps), "####0.0"); " "; Format(Bmarg + scy * (StratNode(i - 1, well1) - hmin), "#####0.0"); " "; "lineto"
    '//          print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i][wellpos1])-5)+Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[i-1][wellpos1]-hmin));
                 Print #22, Format(0.1 * (PS_data(i, well1) - 5) + Lmarg + scx * (well1 - xstrtps), "####0.0"); " "; Format(Bmarg + scy * (StratNode(i, well1) - hmin), "#####0.0"); " "; "lineto"
    '//          print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i][wellpos1])-5)+Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[i][wellpos1]-hmin));
                 Print #22, Format(Lmarg + scx * (wellpos1 - xstrtps), "####0.0"); " "; Format(Bmarg + scy * (StratNode(i, well1) - hmin), "####0.0"); " "; "lineto"
    '//          print #22,"%5.1f %5.1f lineto\n",Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[i][wellpos1]-hmin));
                 thickness(0) = thickness(0) + (StratNode(i + 1, wellpos1) - StratNode(i, well1))
    '//          thickness = StratNode[i+1][wellpos1]-StratNode[i][wellpos1];
             End If
             If thickness(0) >= 0.01 Then
'//          if (thickness >= 0.01)
                 Print #22, Format(0.1 * (PS_data(i + 1, well1) - 5) + Lmarg + scx * (well1 - xstrtps), "####0.0"); " "; Format(Bmarg + scy * (StratNode(i, well1) - hmin), "####0.0"); " "; "lineto"
'//              print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i+1][wellpos1])-5)+Lmarg+scx*(wellpos1-xstrtps),Bmarg+scy*(StratNode[i][wellpos1]-hmin));
             End If
         Next i
'//      }
         Print #22, "stroke"
'//      print #22,"stroke\n");
     End If
'//  }

'    if (wellpos2 >= xstrtps && wellpos2 <= xstopps)
'    {
'        print #22,"%5.1f %5.1f moveto\n",Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[nIterations-1][wellpos2]-hmin));
'        print #22,"%5.1f %5.1f lineto\n",Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[0][wellpos2]-hmin));
'        print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[1][wellpos2])-5)+Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[0][wellpos2]-hmin));
'        print #22,"stroke\n");// the above 3 line plot the well position and the lowerboundary -'L'line
'        print #22,"%5.1f %5.1f moveto\n",(0.10*(PS_data[1][wellpos2])-5)+Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[0][wellpos2]-hmin)); //goto lower right position were the line has just been drawn
'        for (i=1;i<nIterations-1;i++)
'        {
'            thickness = StratNode[i][wellpos2]-StratNode[i-1][wellpos2];
'            if (thickness < 0.01)
'                PS_data[i][wellpos2] = PS_data[i-1][wellpos2];
'            print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i][wellpos2])-5)+Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[i-1][wellpos2]-hmin));
'            print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i][wellpos2])-5)+Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[i][wellpos2]-hmin));
'            print #22,"%5.1f %5.1f lineto\n",Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[i][wellpos2]-hmin));
'            thickness = StratNode[i+1][wellpos2]-StratNode[i][wellpos2];
'            if (thickness >= 0.01)
'                print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i+1][wellpos2])-5)+Lmarg+scx*(wellpos2-xstrtps),Bmarg+scy*(StratNode[i][wellpos2]-hmin));
'        }
'        print #22,"stroke\n");
'    }
'    if (wellpos3 >= xstrtps && wellpos3 <= xstopps)
'    {
'        print #22,"%5.1f %5.1f moveto\n",Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[nIterations-1][wellpos3]-hmin));
'        print #22,"%5.1f %5.1f lineto\n",Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[0][wellpos3]-hmin));
'        print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[1][wellpos3])-5)+Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[0][wellpos3]-hmin));
'        print #22,"stroke\n");// the above 3 line plot the well position and the lowerboundary -'L'line
'        print #22,"%5.1f %5.1f moveto\n",(0.10*(PS_data[1][wellpos3])-5)+Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[0][wellpos3]-hmin)); //goto lower right position were the line has just been drawn
'        for (i=1;i<nIterations-1;i++)
'        {
'            thickness = StratNode[i][wellpos3]-StratNode[i-1][wellpos3];
'            if (thickness < 0.01)
'                PS_data[i][wellpos3] = PS_data[i-1][wellpos3];
'            print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i][wellpos3])-5)+Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[i-1][wellpos3]-hmin));
'            print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i][wellpos3])-5)+Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[i][wellpos3]-hmin));
'            print #22,"%5.1f %5.1f lineto\n",Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[i][wellpos3]-hmin));
'            thickness = StratNode[i+1][wellpos3]-StratNode[i][wellpos3];
'            if (thickness >= 0.01)
'                print #22,"%5.1f %5.1f lineto\n",(0.10*(PS_data[i+1][wellpos3])-5)+Lmarg+scx*(wellpos3-xstrtps),Bmarg+scy*(StratNode[i][wellpos3]-hmin));
'        }
'        print #22,"stroke\n");
'    }

    
    Print #22, "showpage"
    

    Close #22
' /*    end of write_perm_psfile()  */



End Sub


