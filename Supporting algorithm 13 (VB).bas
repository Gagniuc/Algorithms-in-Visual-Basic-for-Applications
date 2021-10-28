Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 13. Prediction framework based on a 4x4 transition matrix.            #
'# Known transition probability values are directly used from a transition matrix             #
'# for highlighting the behavior of a Markov Chain. A variable number of states               #
'# and configurations can be tested. For instance, a diagram with three states can            #
'# be molded on the algorithm. However, since the framework is made for a total of            #
'# 4 states the following modifications are required: all Jar(n, 4) elements are              #
'# set to zero (Jar(1, 4)=0; Jar(2, 4)=0; Jar(3, 4)=0; and Jar(4, 4)=0). Furthermore,         #
'# the molding of a diagram with two states on this framework involves two                    #
'# modifications: all Jar(n, 4) elements are set to zero and all Jar(n, 3) elements           #
'# are set to zero. Any type of diagram configuration can be tested by following two          #
'# rules: 1) the absence of an arrow is indicated by zero, and 2) any value greater           #
'# than zero and less than or equal to 1 indicates an arrow.                                  #
'##############################################################################################

Dim Jar(1 To 4, 1 To 4) As String
Private Sub main()

Dim v(0 To 3, 0 To 1) As Variant

Jar(1, 1) = 1
Jar(1, 2) = 0
Jar(1, 3) = 0
Jar(1, 4) = 0

Jar(2, 1) = 0.5
Jar(2, 2) = 0
Jar(2, 3) = 0.5
Jar(2, 4) = 0

Jar(3, 1) = 0
Jar(3, 2) = 0.5
Jar(3, 3) = 0
Jar(3, 4) = 0.5

Jar(4, 1) = 0
Jar(4, 2) = 0
Jar(4, 3) = 1
Jar(4, 4) = 0

chain = 5

v(0, 0) = 0
v(1, 0) = 0
v(2, 0) = 0
v(3, 0) = 1

v(0, 1) = 0
v(1, 1) = 0
v(2, 1) = 0
v(3, 1) = 0

For k = 1 To chain
    For i = 0 To 3
        For j = 0 To 3
            v(i, 1) = v(i, 1) + (v(j, 0) * Jar(j + 1, i + 1))
        Next j
    Next i

    For i = 0 To 3
        v(i, 0) = v(i, 1)
        v(i, 1) = 0
    Next i

    A = v(0, 0)
    B = v(1, 0)
    C = v(2, 0)
    D = v(3, 0)
    
    MsgBox "Step(" & k & ")=[" & A & "|" & B & "|" & C & "|" & D & "]"
Next k

End Sub
