Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 17. Average time tester. The tester is composed                       #
'# of a simulator that generates 10,000 observations. These observations                      #
'# are then analyzed and the frequencies of “A”, “B”, “C” and “D” letters                     #
'# are determined. These frequencies represent the average time spent in                      #
'# each state.                                                                                #
'##############################################################################################

Dim P(0 To 4, 0 To 3) As Variant
Dim Jar(1 To 4) As Variant
Dim f(0 To 3) As Variant

Private Sub main()

P(0, 0) = "A"
P(0, 1) = "B"
P(0, 2) = "C"
P(0, 3) = "D"

P(1, 0) = 0
P(1, 1) = 1
P(1, 2) = 0
P(1, 3) = 0

P(2, 0) = 0.33
P(2, 1) = 0
P(2, 2) = 0.33
P(2, 3) = 0.33

P(3, 0) = 0
P(3, 1) = 1
P(3, 2) = 0
P(3, 3) = 0

P(4, 0) = 0
P(4, 1) = 0
P(4, 2) = 1
P(4, 3) = 0

For j = 1 To 4
    Jar(j) = Fill_Jar(j)
Next j

draws = 10000
a = Draw(2)

For i = 1 To draws
    For j = 0 To 3
        If a = P(0, j) Then
            a = Draw(j + 1)
            z = z & a
            GoTo 1
        End If
    Next j
1:
Next i

For i = 1 To Len(z)
    g = Mid(z, i, 1)
    If g = "A" Then f(0) = f(0) + 1
    If g = "B" Then f(1) = f(1) + 1
    If g = "C" Then f(2) = f(2) + 1
    If g = "D" Then f(3) = f(3) + 1
Next i

For i = 0 To 3
pro = pro & P(0, i) & "=" & Int((100 / Len(z)) * f(i)) & "%" & Chr(9)
Next i

MsgBox pro
End Sub

Function Fill_Jar(ByVal S As Variant) As Variant
Ltot = 100
For i = 0 To 3
    a = Int(Ltot * P(S, i))
        For j = 1 To a
            b = b & P(0, i)
        Next j
Next i
Fill_Jar = b
End Function

Function Draw(ByVal S As Variant) As Variant
    Randomize
    randomly_choose = Int(Rnd * Len(Jar(S)))
    ball = Mid(Jar(S), randomly_choose + 1, 1)
    Draw = ball
End Function
