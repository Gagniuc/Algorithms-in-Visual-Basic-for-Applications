Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 10. Predictions based on sequences produced by n-state                #
'# Markov machines. This example also uses a DNA sequence as a model. However,                #
'# the algorithm allows for an unlimited number of letters (observations).                    #
'# Previously, the vector – matrix multiplication cycle was declared manually                 #
'# with a range of expressions. Here, the multiplication cycle is made                        #
'# iteratively. For a prediction on more than 4 states, the matrix elements                   #
'# and the number of vector components can be increased to cover a new                        #
'# prediction requirement. Note that “ExtractProb” function is not shown.                     #
'# However, when the above algorithm is used the “ExtractProb” function                       #
'# must be present.                                                                           #
'##############################################################################################

Dim M(1 To 4, 1 To 4) As String

Private Sub main()

Dim v(0 To 3, 0 To 1) As Variant

Call ExtractProb("TACTTCGATTTAAGCGCGGCGGCCTATATTA")

chain = 5

v(0, 0) = 1
v(1, 0) = 0
v(2, 0) = 0
v(3, 0) = 0

v(0, 1) = 0
v(1, 1) = 0
v(2, 1) = 0
v(3, 1) = 0

For k = 1 To chain
    
    For i = 0 To 3
        For j = 0 To 3
            v(i, 1) = v(i, 1) + (v(j, 0) * M(j + 1, i + 1))
        Next j
    Next i

    For i = 0 To 3
        v(i, 0) = v(i, 1)
        v(i, 1) = 0
    Next i

    A = v(0, 0)
    T = v(1, 0)
    c = v(2, 0)
    G = v(3, 0)
    
    MsgBox "V(" & k & ")=[" & A & " | " & T & " | " & c & " | " & G & "]"

Next k

End Sub


Function ExtractProb(ByVal s As String)

Ea = "A"
Et = "T"
Eg = "G"
Ec = "C"

For i = 1 To 4
    For j = 1 To 4
      M(i, j) = 0
    Next j
Next i

Ta = 0
Tt = 0
Tg = 0
Tc = 0

For i = 2 To Len(s) - 1

        DI1 = Mid(s, i, 1)
        DI2 = Mid(s, i + 1, 1)

        If DI1 = Ea Then r = 1
        If DI1 = Et Then r = 2
        If DI1 = Eg Then r = 3
        If DI1 = Ec Then r = 4
        
        If DI2 = Ea Then c = 1
        If DI2 = Et Then c = 2
        If DI2 = Eg Then c = 3
        If DI2 = Ec Then c = 4

        M(r, c) = Val(M(r, c)) + 1

        If DI1 = Ea Then Ta = Ta + 1
        If DI1 = Et Then Tt = Tt + 1
        If DI1 = Eg Then Tg = Tg + 1
        If DI1 = Ec Then Tc = Tc + 1

Next i

For i = 1 To 4
    For j = 1 To 4
        If i = 1 Then M(i, j) = Val(M(i, j)) / Ta
        If i = 2 Then M(i, j) = Val(M(i, j)) / Tt
        If i = 3 Then M(i, j) = Val(M(i, j)) / Tg
        If i = 4 Then M(i, j) = Val(M(i, j)) / Tc
    Next j
Next i

End Function

