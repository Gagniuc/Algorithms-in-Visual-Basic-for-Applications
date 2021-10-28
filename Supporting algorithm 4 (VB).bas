Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 4. Step-by-step prediction using a 2x2 transition matrix.             #
'# A probability vector is repeatedly multiplied by a transition matrix. The                  #
'# vectors obtained from these repetitions show the probability of each                       #
'# outcome on a particular step.                                                              #
'##############################################################################################

Private Sub main()
Dim P(1 To 2, 1 To 2) As Variant
Dim v(0 To 1) As Variant

chain = 5

P(1, 1) = 0.2
P(1, 2) = 0.625
P(2, 1) = 0.8
P(2, 2) = 0.375

v(0) = 1
v(1) = 0

For i = 1 To chain

    x = (v(0) * P(1, 1)) + (v(1) * P(1, 2))
    y = (v(0) * P(2, 1)) + (v(1) * P(2, 2))
    
    v(0) = x
    v(1) = y
    
    MsgBox v(0) & " | " & v(1)

Next i

End Sub

