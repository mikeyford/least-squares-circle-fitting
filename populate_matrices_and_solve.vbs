Sub Solve()

f = Sheets("Input").Cells(1, 14) ' Set variable for number of lines in functional model
l = Sheets("Input").Cells(2, 14) ' Set variable for number of measurements

For m = 1 To l
    For n = 1 To l
        Sheets("C").Cells(m, n) = "0" ' Generate empty C matrix
    Next n
Next m

m = 1 ' set C matrix position variable
For n = 1 To f
    If Sheets("Input").Cells(n, 1).Value = "Left Tangent" Then
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) ' Insert X-error into C
        Sheets("C").Cells(m, m + 1) = Sheets("Input").Cells(n, 7) ' Insert X-row covariance into C
        m = m + 1
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 6) ' Insert Y-error into C
        Sheets("C").Cells(m, m - 1) = Sheets("Input").Cells(n, 7) ' Insert Y-row covariance into C
        m = m + 1
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 8) ' Insert Angle-error into C
        m = m + 1
        Sheets("A").Cells(n, 1) = "1" ' Insert radius coefficient into A col 1
        Sheets("A").Cells(n, 2) = "=-COS(RADIANS(Input!D" & n & "))" '  Insert -cos(alpha) into A col 2
        Sheets("A").Cells(n, 3) = "=SIN(RADIANS(Input!D" & n & "))" ' Insert sin(alpha) into A col 3
        Sheets("b").Cells(n, 1) = "=-(Input!$M$10-Input!$M$11*COS(RADIANS(Input!D" & n & "))+Input!$M$12*SIN(RADIANS(Input!D" & n & "))+Input!B" & n & "*COS(RADIANS(Input!D" & n & "))-Input!C" & n & "*SIN(RADIANS(Input!D" & n & ")))" ' Insert b equation
    End If
    If Sheets("Input").Cells(n, 1).Value = "Right Tangent" Then
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) ' Insert X-error into C
        Sheets("C").Cells(m, m + 1) = Sheets("Input").Cells(n, 7) ' Insert X-row covariance into C
        m = m + 1
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) ' Insert Y-error into C
        Sheets("C").Cells(m, m - 1) = Sheets("Input").Cells(n, 7) ' Insert Y-row covariance into C
        m = m + 1
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 8) ' Insert Angle-error into C
        m = m + 1
        Sheets("A").Cells(n, 1) = "-1" ' Insert radius coefficient into A col 1
        Sheets("A").Cells(n, 2) = "=-COS(RADIANS(Input!D" & n & "))" ' Insert -cos(alpha) into A col 2
        Sheets("A").Cells(n, 3) = "=SIN(RADIANS(Input!D" & n & "))" ' Insert sin(alpha) into A col 3
        Sheets("b").Cells(n, 1) = "=-(-Input!$M$10-Input!$M$11*COS(RADIANS(Input!D" & n & "))+Input!$M$12*SIN(RADIANS(Input!D" & n & "))+Input!B" & n & "*COS(RADIANS(Input!D" & n & "))-Input!C" & n & "*SIN(RADIANS(Input!D" & n & ")))" ' Insert b equation
    End If
    If Sheets("Input").Cells(n, 1).Value = "Point" Then
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) ' Insert X-error into C
        Sheets("C").Cells(m, m + 1) = Sheets("Input").Cells(n, 7) ' Insert X-row covariance into C
        m = m + 1
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 6) ' Insert Y-error into C
        Sheets("C").Cells(m, m - 1) = Sheets("Input").Cells(n, 7) ' Insert Y-row covariance into C
        m = m + 1
        Sheets("A").Cells(n, 1) = "=-2*Input!$M$10" ' Insert -2r into A col 1
        Sheets("A").Cells(n, 2) = "=-2*(Input!B" & n & "-Input!$M$11)" ' Insert -2*(x_p-x_c) into A col 2"
        Sheets("A").Cells(n, 3) = "=-2*(Input!C" & n & "-Input!$M$12)" ' Insert -2*(y_p-y_c) into A col 3"
        Sheets("b").Cells(n, 1) = "=-(-(Input!$M$10^2)+(Input!B" & n & "-Input!$M$11)^2+(Input!C" & n & "-Input!$M$12)^2)" ' Insert equation into b
    End If
Next n

For m = 1 To f
    For n = 1 To l
        Sheets("M").Cells(m, n) = "0" ' Generate empty M matrix
    Next n
Next m

m = 1 ' reset as M matrix position variable
For n = 1 To f
    If Sheets("Input").Cells(n, 1).Value = "Left Tangent" Then
        Sheets("M").Cells(n, m).Value = "=COS(RADIANS(Input!D" & n & "))" ' set cos(alpha) in M
        m = m + 1
        Sheets("M").Cells(n, m).Value = "=-SIN(RADIANS(Input!D" & n & "))" ' set -sin(alpha) in M
        m = m + 1
        Sheets("M").Cells(n, m).Value = "=-Input!$M$11*-SIN(RADIANS(Input!D" & n & "))+Input!$M$12*COS(RADIANS(Input!D" & n & "))+Input!B" & n & "*-SIN(RADIANS(Input!D" & n & "))-Input!C" & n & "*COS(RADIANS(Input!D" & n & "))" ' Add partial with respect to alpha to M
        m = m + 1
    End If
    If Sheets("Input").Cells(n, 1).Value = "Right Tangent" Then
        Sheets("M").Cells(n, m).Value = "=COS(RADIANS(Input!D" & n & "))" ' set cos(alpha) in M
        m = m + 1
        Sheets("M").Cells(n, m).Value = "=-SIN(RADIANS(Input!D" & n & "))" ' set -sin(alpha) in M
        m = m + 1
        Sheets("M").Cells(n, m).Value = "=-Input!$M$11*-SIN(RADIANS(Input!D" & n & "))+Input!$M$12*COS(RADIANS(Input!D" & n & "))+Input!B" & n & "*-SIN(RADIANS(Input!D" & n & "))-Input!C" & n & "*COS(RADIANS(Input!D" & n & "))" ' Add partial with respect to alpha to M
        m = m + 1
    End If
    If Sheets("Input").Cells(n, 1).Value = "Point" Then
        Sheets("M").Cells(n, m).Value = "=2*(Input!B" & n & "-Input!$M$11)" ' Insert 2(x_p - x_c) into M
        m = m + 1
        Sheets("M").Cells(n, m).Value = "=2*(Input!C" & n & "-Input!$M$12)" ' Insert 2(y_p - y_c) into M
        m = m + 1
    End If
Next n


' Next section perform the solution iteration
n = Sheets("Input").Cells(15, 13) ' Get max number of iterations
d = Sheets("Input").Cells(16, 13) ' Get decimal points required
c = 0
For i = 1 To n
    r_input = Round(Sheets("Input").Cells(10, 13).Value, d)
    x_input = Round(Sheets("Input").Cells(11, 13).Value, d)
    y_input = Round(Sheets("Input").Cells(12, 13).Value, d)
    r_output = Round(Sheets("Input").Cells(19, 13).Value, d)
    x_output = Round(Sheets("Input").Cells(20, 13).Value, d)
    y_output = Round(Sheets("Input").Cells(21, 13).Value, d)
    If r_input <> r_output Or x_input <> x_output Or y_input <> y_output Then ' check if the input and output values agree
        Sheets("Iterations").Cells(1, i) = "Iteration " & i
        Sheets("Iterations").Cells(2, i) = Sheets("Input").Cells(19, 13).Value
        Sheets("Iterations").Cells(3, i) = Sheets("Input").Cells(20, 13).Value
        Sheets("Iterations").Cells(4, i) = Sheets("Input").Cells(21, 13).Value
        c = c + 1
    End If
    Sheets("Input").Cells(10, 13).Value = Sheets("Input").Cells(19, 13).Value
    Sheets("Input").Cells(11, 13).Value = Sheets("Input").Cells(20, 13).Value
    Sheets("Input").Cells(12, 13).Value = Sheets("Input").Cells(21, 13).Value
Next i

' Populate residual standard errors from Cv
For i = 1 To l
    Sheets("Residuals").Cells(i, 2).Value = "=SQRT(" & Sheets("Cv").Cells(i, i).Value & ")"
Next i

' Populate residuals standard variances from error covariance matrix
For i = 1 To l
    Sheets("Residuals").Cells(i, 5).Value = Sheets("C").Cells(i, i).Value
Next i

' Populate MDE Matrix
For i = 1 To l
    Sheets("MDE").Cells(i, i).Value = Sheets("Residuals").Cells(i, 6).Value
Next i


' Give user warning if the solution did not converge or W-test exceeded
If r_input = r_output And x_input = x_output And y_input = y_output Then
    If Sheets("Residuals").Cells(2, 8).Value > Sheets("Test").Cells(13, 3).Value Then
        MsgBox "Warning! The solution converged but the W-test statistic was exceeded. Check data for outliers.", vbCritical
    Else
        MsgBox "The solution converged in " & c & " iterations."
    End If
Else
    MsgBox "Warning! The solution did not converge in " & n & " iterations.", vbCritical
End If

End Sub
