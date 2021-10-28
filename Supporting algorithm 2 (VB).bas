Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 2. A two states Markov Chain simulator based on probability values.   #
'# The probability values present inside the transition matrix are directly used for an       #
'# automatic generation of the letter combination that make up the representation of the      #
'# jars. Thus, the two letter sequences have a calculated proportion of “W” and “B”           #
'# letters. The chance of a letter chosen at random from one of the two sequences             #
'# is directly dictated by the proportions of “W” and “B” letters.                            #
'##############################################################################################


Dim Jar(0 To 1) As String

Private Sub main()

Call Fill_Jar(0, 0.2) 'W
Call Fill_Jar(1, 0.6) 'B

draws = 17

a = Draw(0) ' Draws start from jar "W"

z = z & " Jar W[" & a & "],"

For i = 1 To draws
    If a = "W" Then
        a = Draw(0)
        z = z & " Jar W[" & a & "],"
    Else
        a = Draw(1)
        z = z & " Jar B[" & a & "],"
    End If
MsgBox z
Next i

End Sub

Function Draw(ByVal S As Variant) As String
    Randomize
    randomly_choose = Int(Rnd * Len(Jar(S)))
    ball = Mid(Jar(S), randomly_choose + 1, 1)
    Draw = ball
End Function

Function Fill_Jar(ByVal S As Integer, ByVal p As Variant)
Balls_W = Int(100 * p)
Balls_B = 100 - Balls_W

For i = 1 To Balls_W
    Jar(S) = Jar(S) & "W"
Next i

For i = 1 To Balls_B
    Jar(S) = Jar(S) & "B"
Next i

End Function
