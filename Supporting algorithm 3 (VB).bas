Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 3. The conversion of a sequence of observations to a transition       #
'# matrix. A 2x2 matrix is used for counting all the combinations of pairs of letters         #
'# (Da->b) in the sequence (Da->b is represented by joining two string variables, namely      #
'# DI1 and DI2). In parallel, the first letter of each pair (Na) is counted inside the        #
'# sequence (Na is represented by variable DI1). Next, the values from each element of the    #
'# 2x2 matrix are divided by the number of first letters found in the sequence. Depending     #
'# on the type of values (counts or probability values) stored inside, the same matrix is     #
'# then shown twice in a graphical format.                                                    #
'##############################################################################################

Dim M(1 To 2, 1 To 2) As String

Private Sub main()
Call ExtractProb("SRRSRSRRSRSRRSS")
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

        If DI1 = Eb Then R = 1
        If DI1 = Es Then R = 2
        If DI2 = Eb Then c = 1
        If DI2 = Es Then c = 2

        M(R, c) = Val(M(R, c)) + 1

        If DI1 = Eb Then TB = TB + 1
        If DI1 = Es Then TS = TS + 1
Next i

MsgBox DrowMatrix(2, 2, M, "(C)", "Count:")

For i = 1 To 2
    For j = 1 To 2
       If i = 1 Then M(i, j) = Val(M(i, j)) / TB
       If i = 2 Then M(i, j) = Val(M(i, j)) / TS
    Next j
Next i

MsgBox DrowMatrix(2, 2, M, "(P)", "Transition matrix M:")
End Function

Function DrowMatrix(ib, jb, ByVal M As Variant, ByVal model As String, ByVal msg As String) As String

Eb = "S"
Es = "R"

y = "|___|___|___|"
ct = ct & vbCrLf & "____________"
ct = ct & vbCrLf & "| " & model & " |  " & Eb & "  |  " & Es & "  | "
ct = ct & vbCrLf & y & vbCrLf

For i = 1 To ib
    For j = 1 To jb
    
    v = Round(M(i, j), 2)
    
        If Len(v) = 0 Then u = "|     "
        If Len(v) = 1 Then u = "|    "
        If Len(v) = 2 Then u = "|   "
        If Len(v) = 3 Then u = "|  "
        If Len(v) = 4 Then u = "| "
        If Len(v) = 5 Then u = "|"
        
        If j = jb Then o = "|" Else o = ""
        If j = 1 And i = 1 Then ct = ct & "|  " & Eb & "  "
        If j = 1 And i = 2 Then ct = ct & "|  " & Es & "  "
        
        ct = ct & u & v & o
    Next j
ct = ct & vbCrLf & y & vbCrLf
Next i

DrowMatrix = msg & " M[" & Val(jb) & "," & Val(ib) & "]" & vbCrLf & ct & vbCrLf
End Function

