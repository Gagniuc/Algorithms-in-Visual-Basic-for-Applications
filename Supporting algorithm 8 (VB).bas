Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 8. Step-by-step prediction using a sequence of observations           #
'# made by a 3-state Markov machine. First, a 3x3 matrix is used for counting all             #
'# the combinations of pairs of letters (Da->b) in the sequence (Da->b is represented         #
'# by joining two string variables, namely DI1 and DI2). In parallel, the first               #
'# letter (Na) of each pair is counted inside the sequence (Na is represented by              #
'# variable DI1). Secondly, the transition probabilities are computed. The values             #
'# from each element of the matrix are divided by their corresponding Na. In the              #
'# final phase, a probability vector is repeatedly multiplied by the new transition           #
'# matrix. The vectors obtained from these repetitions show the probability of each           #
'# outcome on a particular step.                                                              #
'##############################################################################################

Dim M(1 To 3, 1 To 3) As String

Private Sub main()
Dim v(0 To 2) As Variant

Call ExtractProb("SRCCRRSSCSRCSR")

chain = 5
v(0) = 0
v(1) = 1
v(2) = 0

For i = 1 To chain

    x = (v(0) * M(1, 1)) + (v(1) * M(2, 1)) + (v(2) * M(3, 1))
    y = (v(0) * M(1, 2)) + (v(1) * M(2, 2)) + (v(2) * M(3, 2))
    z = (v(0) * M(1, 3)) + (v(1) * M(2, 3)) + (v(2) * M(3, 3))
    
    v(0) = x
    v(1) = y
    v(2) = z
    
    MsgBox "Day (" & i & ")=[" & v(0) & " | " & v(1) & " | " & v(2) & "]"

Next i

End Sub

Function ExtractProb(ByVal s As String)

Eb = "S"
Es = "R"
Ec = "C"

For i = 1 To 3
    For j = 1 To 3
      M(i, j) = 0
    Next j
Next i

TB = 0
TS = 0
TC = 0

For i = 2 To Len(s) - 1

        DI1 = Mid(s, i, 1)
        DI2 = Mid(s, i + 1, 1)

        If DI1 = Eb Then r = 1
        If DI1 = Es Then r = 2
        If DI1 = Ec Then r = 3
        If DI2 = Eb Then c = 1
        If DI2 = Es Then c = 2
        If DI2 = Ec Then c = 3

        M(r, c) = Val(M(r, c)) + 1

        If DI1 = Eb Then TB = TB + 1
        If DI1 = Es Then TS = TS + 1
        If DI1 = Ec Then TC = TC + 1

Next i

For i = 1 To 3
    For j = 1 To 3
       If i = 1 Then M(i, j) = Val(M(i, j)) / TB
       If i = 2 Then M(i, j) = Val(M(i, j)) / TS
       If i = 3 Then M(i, j) = Val(M(i, j)) / TC
    Next j
Next i

End Function
