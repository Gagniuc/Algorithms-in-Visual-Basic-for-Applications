Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 9. Step-by-step prediction by using a DNA sequence.                   #
'# The letters that make up a DNA sequence are: “A”, “T”, “G” and “C”. Thus,                  #
'# the observations present in a DNA sequence are suitable for exemplifications               #
'# involving a 4-state Markov machine. As before, a 4x4 matrix is used for                    #
'# counting all the combinations of pairs of letters (Da->b) in the DNA                       #
'# sequence (Da->b is represented by joining two string variables, namely                     #
'# DI1 and DI2). In parallel, the first letter (Na) of each pair is counted                   #
'# inside the DNA sequence (Na is represented by variable DI1). Secondly,                     #
'# the transition probabilities are computed. The values from each element                    #
'# of the matrix are divided by their corresponding Na. In the final phase,                   #
'# a probability vector is repeatedly multiplied by the new transition matrix.                #
'# The vectors obtained from these repetitions show the probability of each                   #
'# outcome on a particular step.                                                              #
'##############################################################################################


Dim M(1 To 4, 1 To 4) As String

Private Sub main()

Dim v(0 To 3) As Variant

Call ExtractProb("TACTTCGATTTAAGCGCGGCGGCCTATATTA")

chain = 3

v(0) = 1
v(1) = 0
v(2) = 0
v(3) = 0

For i = 1 To chain
    
x = (v(0) * M(1, 1)) + (v(1) * M(2, 1)) + (v(2) * M(3, 1)) + (v(3) * M(4, 1))
y = (v(0) * M(1, 2)) + (v(1) * M(2, 2)) + (v(2) * M(3, 2)) + (v(3) * M(4, 2))
z = (v(0) * M(1, 3)) + (v(1) * M(2, 3)) + (v(2) * M(3, 3)) + (v(3) * M(4, 3))
w = (v(0) * M(1, 4)) + (v(1) * M(2, 4)) + (v(2) * M(3, 4)) + (v(3) * M(4, 4))

v(0) = x
v(1) = y
v(2) = z
v(3) = w

out = Empty

For c = 0 To 3
      out = out & v(c) & "|"
Next c

MsgBox "Base(" & i & ")=[" & out & "]"
    
BaseBy = BaseBy & Base(v())

MsgBox BaseBy
    
Next i

End Sub


Function Base(ByRef v() As Variant)

For i = 0 To UBound(v)

    If v(i) > old Then
        x = v(i)
        h = i
    End If
    
    old = x

Next i

    If h = 0 Then n = "A"
    If h = 1 Then n = "T"
    If h = 2 Then n = "G"
    If h = 3 Then n = "C"
        
Base = n

End Function

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
