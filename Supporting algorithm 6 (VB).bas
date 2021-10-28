Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 6. The computation of the steady state vector. The above formulas are #
'# used for computing the x and y components of the steady state vector. Note that iterations #
'# are not required.                                                                          #
'##############################################################################################

Dim M(1 To 2, 1 To 2) As String

Private Sub main()
Dim v(0 To 1) As Variant

Call ExtractProb("SRRSRSRRSRSRRSS")

chain = 5

v(0) = 1
v(1) = 0

For i = 1 To chain

    x = (v(0) * M(1, 1)) + (v(1) * M(2, 1))
    y = (v(0) * M(1, 2)) + (v(1) * M(2, 2))
    
    v(0) = x
    v(1) = y
    
    MsgBox "Day (" & i & ")=[" & v(0) & " - " & v(1) & "]"

Next i
End Sub

Function ExtractProb(ByVal s As String)

Eb = "S"
Es = "R"

For i = 1 To 2
    For j = 1 To 2
      M(i, j) = 0
    Next j
Next i

TB = 0
TS = 0

For i = 2 To Len(s) - 1

        DI1 = Mid(s, i, 1)
        DI2 = Mid(s, i + 1, 1)

        If DI1 = Eb Then r = 1
        If DI1 = Es Then r = 2
        If DI2 = Eb Then c = 1
        If DI2 = Es Then c = 2

        M(r, c) = Val(M(r, c)) + 1

        If DI1 = Eb Then TB = TB + 1
        If DI1 = Es Then TS = TS + 1

Next i

For i = 1 To 2
    For j = 1 To 2
       If i = 1 Then M(i, j) = Val(M(i, j)) / TB
       If i = 2 Then M(i, j) = Val(M(i, j)) / TS
    Next j
Next i

End Function

