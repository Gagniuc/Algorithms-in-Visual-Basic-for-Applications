Attribute VB_Name = "Module1"
'##############################################################################################
'# John Wiley & Sons, Inc.                                                                    #
'#                                                                                            #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                   #
'# Author: Dr. Paul Gagniuc                                                                   #
'# Data:   01/09/2016                                                                         #
'#                                                                                            #
'# Description:                                                                               #
'# Supporting algorithm 11. The conversion of measurements to states.                         #
'# A range of values is divided into 4 equal regions. Each region                             #
'# corresponds to a state: “A”, “B”, “C”, and “D”. The numeric values                         #
'# are associated with a representative letter based on their position                        #
'# over the regions. Thus, the initial values are listed as letters (observations).           #
'##############################################################################################

Private Sub main()
Dim Inp() As String
Dim R As String

R = "159,82,187,194,179,115,197,102,105,104,95,126,74,143,143,127,98," & _
"70,92,170,168,182,149,85,137,100,170,180,61,177,86,195,198,182,150," & _
"197,103,103,186,100,96,196"

Inp = Split(R, ",")
Lu = 200
Ld = 60
n = 4
Pr = (Lu - Ld) / n

For i = 0 To UBound(Inp)
    s = (Inp(i) - Ld) / Pr
    s = Split(s, ".")(0)

    If s = 0 Then l = "A"
    If s = 1 Then l = "B"
    If s = 2 Then l = "C"
    If s = 3 Then l = "D"

    Obs = Obs & l
    Reg = Reg & s & ","
Next i

MsgBox "Reg=" & Reg & vbCrLf & "Obs=" & Obs & vbCrLf
End Sub
