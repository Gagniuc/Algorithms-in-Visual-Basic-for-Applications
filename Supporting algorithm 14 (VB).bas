Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 14. A 3-states Markov Chain simulator. The probability                #
'# values present inside a 3x3 transition matrix (P) are directly used for an                 #
'# automatic generation of the letter combination that make up the representation             #
'# of the jars. Thus, the three letter sequences have a calculated proportion of              #
'# “A”, “B” and “C” letters. The chance of a letter chosen at random from one of              #
'# the three sequences is directly dictated by the proportions of “A”, “B” and                #
'# “C” letters.                                                                               #
'##############################################################################################

Dim P(0 To 3, 0 To 2) As Variant
Dim Jar(1 To 3) As Variant

Private Sub main()

P(0, 0) = "A"
P(0, 1) = "B"
P(0, 2) = "C"

P(1, 0) = 0.33
P(1, 1) = 0.33
P(1, 2) = 0.33

P(2, 0) = 0
P(2, 1) = 0.5
P(2, 2) = 0.5

P(3, 0) = 1
P(3, 1) = 0
P(3, 2) = 0

For j = 1 To 3
    Jar(j) = Fill_Jar(j)
    MsgBox Jar(j)
Next j

draws = 20
a = Draw(1)

For i = 1 To draws
    For j = 0 To 3
        If a = P(0, j) Then
            a = Draw(j + 1)
            q = q & P(0, j)
            z = z & ", Jar " & P(0, j) & "[" & a & "]"
            GoTo 1
        End If
    Next j
1:

Next i

MsgBox q
MsgBox z
End Sub

Function Fill_Jar(ByVal S As Variant) As Variant

Ltot = 27

For i = 0 To 2
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
