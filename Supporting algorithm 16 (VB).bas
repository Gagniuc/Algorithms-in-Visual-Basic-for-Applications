Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 16. Transition probability tester. Previously, a sequence             #
'# of observations has been provided by a simulator. To test the accuracy of the              #
'# simulator, the sequence of observations is used for creating a transition matrix,          #
'# which is then compared with the original.                                                  #
'##############################################################################################

Dim P(1 To 4, 1 To 4) As String

Private Sub main()

Call ExtractProb("BABABDCBABDCBCBDCBDCBABDCBCBCBCBDCBCBDCBDCBDCBABA" & _
"BABDCBDCBABABDCBABCBABCBABDCBDCBABABABABCBABCBDCBDC")

For i = 1 To 4
    For j = 1 To 4
       z = z & Chr(9) & Round(P(i, j), 2)
    Next j
    z = z & vbCrLf
Next i

MsgBox z
End Sub

Function ExtractProb(ByVal s As String)

Ea = "A"
Eb = "B"
Ec = "C"
Ed = "D"

For i = 1 To 4
    For j = 1 To 4
      P(i, j) = 0
    Next j
Next i

Ta = 0
Tb = 0
Tc = 0
Td = 0

For i = 2 To Len(s) - 1

        DI1 = Mid(s, i, 1)
        DI2 = Mid(s, i + 1, 1)

        If DI1 = Ea Then r = 1
        If DI1 = Eb Then r = 2
        If DI1 = Ec Then r = 3
        If DI1 = Ed Then r = 4
        
        If DI2 = Ea Then c = 1
        If DI2 = Eb Then c = 2
        If DI2 = Ec Then c = 3
        If DI2 = Ed Then c = 4

        P(r, c) = Val(P(r, c)) + 1

        If DI1 = Ea Then Ta = Ta + 1
        If DI1 = Eb Then Tb = Tb + 1
        If DI1 = Ec Then Tc = Tc + 1
        If DI1 = Ed Then Td = Td + 1

Next i

For i = 1 To 4
    For j = 1 To 4
        If i = 1 Then P(i, j) = Val(P(i, j)) / Ta
        If i = 2 Then P(i, j) = Val(P(i, j)) / Tb
        If i = 3 Then P(i, j) = Val(P(i, j)) / Tc
        If i = 4 Then P(i, j) = Val(P(i, j)) / Td
    Next j
Next i

End Function
