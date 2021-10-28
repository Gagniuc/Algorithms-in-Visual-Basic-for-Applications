Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 1. A two states Markov Chain simulator based on the proportions of    #
'# letters. Two letter sequences with predetermined proportions of “W” and “B” letters are    #
'# used for the representation of two jars. The chance of a letter chosen at random from      #
'# one of the two sequences is directly dictated by the proportions of “W” and “B” letters.   #
'##############################################################################################

Dim Jar(0 To 1) As String

Private Sub main()

draws = 17

Jar(0) = "WWBBBBBBBB"
Jar(1) = "WWWWWBBBBB"

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

Function Draw(ByVal S As Integer) As String
    Randomize
    randomly_choose = Int(Rnd * Len(Jar(S)))
    ball = Mid(Jar(S), randomly_choose + 1, 1)
    Draw = ball
End Function
