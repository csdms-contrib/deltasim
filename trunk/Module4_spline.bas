Attribute VB_Name = "Module4"
Option Explicit

'********************  Cubic_Spline by SRS1 Software **************** ' ' ' Version 1.00


Function cubic_spline(actual_t)

Dim input_count As Integer
Dim Y As Single

'Purpose:   Given a data set consisting of a list of x values
'           and y values, this function will smoothly interpolate
'           a resulting output (y) value from a given input (x) value


' This counts how many points are in "input" and "output" set of data 'Dim input_count As Integer 'Dim output_count As Integer

'input_count = input_column.Rows.Count
'output_count = output_column.Rows.Count

'Next check to be sure that "input" # points = "output" # points 'If input_count <> output_count Then
'    cubic_spline = "Something's messed up!  The number of indeces number of output_columnues don't match!"
'    GoTo out
'End If
 
SL_nrvalue = 101
input_count = SL_nrvalue
 
ReDim xin(input_count) As Single
ReDim yin(input_count) As Single

Dim c As Integer

For c = 1 To input_count
xin(c) = SL_Time(c)  'xin(c) = input_column(c)
yin(c) = SL_Value(c) 'yin(c) = output_column(c)
Next c

'''''''''''''''''''''''''''''''''''''''
' values are populated
'''''''''''''''''''''''''''''''''''''''
Dim n As Integer 'n=input_count
Dim i, k As Integer 'these are loop counting integers
Dim p, qn, sig, un As Single
ReDim u(input_count - 1) As Single
ReDim yt(input_count) As Single 'these are the 2nd deriv values

n = input_count
yt(1) = 0
u(1) = 0

For i = 2 To n - 1
    sig = (xin(i) - xin(i - 1)) / (xin(i + 1) - xin(i - 1))
    p = sig * yt(i - 1) + 2
    yt(i) = (sig - 1) / p
    u(i) = (yin(i + 1) - yin(i)) / (xin(i + 1) - xin(i)) - (yin(i) - yin(i - 1)) / (xin(i) - xin(i - 1))
    u(i) = (6 * u(i) / (xin(i + 1) - xin(i - 1)) - sig * u(i - 1)) / p
    
    Next i
    
qn = 0
un = 0

yt(n) = (un - qn * u(n - 1)) / (qn * yt(n - 1) + 1)

For k = n - 1 To 1 Step -1
    yt(k) = yt(k) * yt(k + 1) + u(k)
Next k


''''''''''''''''''''
'now eval spline at one point
'''''''''''''''''''''
Dim klo, khi As Integer
Dim h, b, a As Single

' first find correct interval
klo = 1
khi = n
Do
k = khi - klo
If xin(k) > actual_t Then
khi = k
Else
klo = k
End If
k = khi - klo
Loop While k > 1
h = xin(khi) - xin(klo)
a = (xin(khi) - actual_t) / h
b = (actual_t - xin(klo)) / h
Y = a * yin(klo) + b * yin(khi) + ((a ^ 3 - a) * yt(klo) + (b ^ 3 - b) * yt(khi)) * (h ^ 2) / 6


cubic_spline = Y

out:
End Function

Sub Load_SL_Time()

SL_Time(1) = 0
SL_Time(2) = 100
SL_Time(3) = 200
SL_Time(4) = 300
SL_Time(5) = 400
SL_Time(6) = 500
SL_Time(7) = 600
SL_Time(8) = 700
SL_Time(9) = 800
SL_Time(10) = 900
SL_Time(11) = 1000
SL_Time(12) = 1100
SL_Time(13) = 1200
SL_Time(14) = 1300
SL_Time(15) = 1400
SL_Time(16) = 1500
SL_Time(17) = 1600
SL_Time(18) = 1700
SL_Time(19) = 1800
SL_Time(20) = 1900
SL_Time(21) = 2000
SL_Time(22) = 2100
SL_Time(23) = 2200
SL_Time(24) = 2300
SL_Time(25) = 2400
SL_Time(26) = 2500
SL_Time(27) = 2600
SL_Time(28) = 2700
SL_Time(29) = 2800
SL_Time(30) = 2900
SL_Time(31) = 3000
SL_Time(32) = 3100
SL_Time(33) = 3200
SL_Time(34) = 3300
SL_Time(35) = 3400
SL_Time(36) = 3500
SL_Time(37) = 3600
SL_Time(38) = 3700
SL_Time(39) = 3800
SL_Time(40) = 3900
SL_Time(41) = 4000
SL_Time(42) = 4100
SL_Time(43) = 4200
SL_Time(44) = 4300
SL_Time(45) = 4400
SL_Time(46) = 4500
SL_Time(47) = 4600
SL_Time(48) = 4700
SL_Time(49) = 4800
SL_Time(50) = 4900
SL_Time(51) = 5000
SL_Time(52) = 5100
SL_Time(53) = 5200
SL_Time(54) = 5300
SL_Time(55) = 5400
SL_Time(56) = 5500
SL_Time(57) = 5600
SL_Time(58) = 5700
SL_Time(59) = 5800
SL_Time(60) = 5900
SL_Time(61) = 6000
SL_Time(62) = 6100
SL_Time(63) = 6200
SL_Time(64) = 6300
SL_Time(65) = 6400
SL_Time(66) = 6500
SL_Time(67) = 6600
SL_Time(68) = 6700
SL_Time(69) = 6800
SL_Time(70) = 6900
SL_Time(71) = 7000
SL_Time(72) = 7100
SL_Time(73) = 7200
SL_Time(74) = 7300
SL_Time(75) = 7400
SL_Time(76) = 7500
SL_Time(77) = 7600
SL_Time(78) = 7700
SL_Time(79) = 7800
SL_Time(80) = 7900
SL_Time(81) = 8000
SL_Time(82) = 8100
SL_Time(83) = 8200
SL_Time(84) = 8300
SL_Time(85) = 8400
SL_Time(86) = 8500
SL_Time(87) = 8600
SL_Time(88) = 8700
SL_Time(89) = 8800
SL_Time(90) = 8900
SL_Time(91) = 9000
SL_Time(92) = 9100
SL_Time(93) = 9200
SL_Time(94) = 9300
SL_Time(95) = 9400
SL_Time(96) = 9500
SL_Time(97) = 9600
SL_Time(98) = 9700
SL_Time(99) = 9800
SL_Time(100) = 9900
SL_Time(101) = 10000


End Sub
Sub load_sl_value()
SL_Value(1) = -27.6
SL_Value(2) = -29.3
SL_Value(3) = -26.4
SL_Value(4) = -25.4
SL_Value(5) = -27
SL_Value(6) = -25.8
SL_Value(7) = -27
SL_Value(8) = -28
SL_Value(9) = -31.6
SL_Value(10) = -31.8
SL_Value(11) = -31.9
SL_Value(12) = -31.9
SL_Value(13) = -31.6
SL_Value(14) = -31.4
SL_Value(15) = -31.2
SL_Value(16) = -30.8
SL_Value(17) = -30.6
SL_Value(18) = -30.2
SL_Value(19) = -30
SL_Value(20) = -28.9
SL_Value(21) = -28.2
SL_Value(22) = -26
SL_Value(23) = -24
SL_Value(24) = -23.8
SL_Value(25) = -48
SL_Value(26) = -48
SL_Value(27) = -48
SL_Value(28) = -48
SL_Value(29) = -48
SL_Value(30) = -48
SL_Value(31) = -48
SL_Value(32) = -48
SL_Value(33) = -48
SL_Value(34) = -48
SL_Value(35) = -48
SL_Value(36) = -48
SL_Value(37) = -48
SL_Value(38) = -48
SL_Value(39) = -48
SL_Value(40) = -48
SL_Value(41) = -28.2
SL_Value(42) = -28.4
SL_Value(43) = -28.6
SL_Value(44) = -28.7
SL_Value(45) = -28.9
SL_Value(46) = -29
SL_Value(47) = -29
SL_Value(48) = -28.9
SL_Value(49) = -28.5
SL_Value(50) = -28.2
SL_Value(51) = -28.1
SL_Value(52) = -27.8
SL_Value(53) = -26.4
SL_Value(54) = -25.7
SL_Value(55) = -24.5
SL_Value(56) = -23.6
SL_Value(57) = -22.4
SL_Value(58) = -21.8
SL_Value(59) = -21.6
SL_Value(60) = -21.5
SL_Value(61) = -22
SL_Value(62) = -23.2
SL_Value(63) = -25.2
SL_Value(64) = -28
SL_Value(65) = -28.5
SL_Value(66) = -28.1
SL_Value(67) = -25.5
SL_Value(68) = -23.3
SL_Value(69) = -21.7
SL_Value(70) = -20.8
SL_Value(71) = -20.2
SL_Value(72) = -20
SL_Value(73) = -20.2
SL_Value(74) = -20.7
SL_Value(75) = -21.4
SL_Value(76) = -22.7
SL_Value(77) = -24
SL_Value(78) = -25.2
SL_Value(79) = -26.1
SL_Value(80) = -27.8
SL_Value(81) = -25.3
SL_Value(82) = -25.7
SL_Value(83) = -26.1
SL_Value(84) = -26.3
SL_Value(85) = -25.8
SL_Value(86) = -25.6
SL_Value(87) = -25.5
SL_Value(88) = -25.5
SL_Value(89) = -25.6
SL_Value(90) = -25.8
SL_Value(91) = -26
SL_Value(92) = -26.7
SL_Value(93) = -27.9
SL_Value(94) = -28.8
SL_Value(95) = -29.7
SL_Value(96) = -30.8
SL_Value(97) = -32
SL_Value(98) = -33.5
SL_Value(99) = -34.9
SL_Value(100) = -36.8
SL_Value(101) = -38

End Sub
