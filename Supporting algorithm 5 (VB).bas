Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 5. Step-by-step prediction for 50 discreet steps.                     #
'# A probability vector is repeatedly multiplied by a 2x2 transition matrix.                  #
'# The vectors obtained from these repeated multiplications show the probability              #
'# of each outcome on a particular step. Furthermore, at every cycle the old vector           #
'# is compared to the new vector in order to detect the first occurrence of                   #
'# equilibrium, namely the steady state vector.                                               #
'##############################################################################################

Private Sub main()
Dim P(1 To 2, 1 To 2) As Variant
Dim v(0 To 1) As Variant

chain = 50

P(1, 1) = 0.2
P(1, 2) = 0.625
P(2, 1) = 0.8
P(2, 2) = 0.375

v(0) = 1
v(1) = 0

For i = 1 To chain

    x = (v(0) * P(1, 1)) + (v(1) * P(1, 2))
    y = (v(0) * P(2, 1)) + (v(1) * P(2, 2))
    
    If v(0) = x And v(1) = y Then
        MsgBox "Steady state vector at day [" & i & "]!"
        Exit Sub
    Else
        MsgBox "Day[" & i & "], v=[" & x & " | " & y & "]"
    End If
    
    v(0) = x
    v(1) = y
    
Next i

End Sub
