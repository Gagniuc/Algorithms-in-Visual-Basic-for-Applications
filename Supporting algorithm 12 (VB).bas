Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 12. Prediction based on a 3x3 transition matrix.                      #
'# Known transition probability values are directly used from a transition                    #
'# matrix for highlighting the behavior of an absorbing Markov Chain.                         #
'##############################################################################################

Dim Jar(1 To 3, 1 To 3) As String

Private Sub main()
Dim v(0 To 2, 0 To 1) As Variant

Jar(1, 1) = 0.33
Jar(1, 2) = 0.33
Jar(1, 3) = 0.33

Jar(2, 1) = 0.5
Jar(2, 2) = 0.5
Jar(2, 3) = 0

Jar(3, 1) = 0
Jar(3, 2) = 0
Jar(3, 3) = 1

chain = 5

v(0, 0) = 1
v(1, 0) = 0
v(2, 0) = 0

v(0, 1) = 0
v(1, 1) = 0
v(2, 1) = 0

For k = 1 To chain
    
    For i = 0 To 2
        For j = 0 To 2
            v(i, 1) = v(i, 1) + (v(j, 0) * Jar(j + 1, i + 1))
        Next j
    Next i

    For i = 0 To 2
        v(i, 0) = v(i, 1)
        v(i, 1) = 0
    Next i

    A = v(0, 0)
    B = v(1, 0)
    C = v(2, 0)

    MsgBox "Step(" & k & ")=[" & A & " | " & B & " | " & C & "]"

Next k

End Sub
