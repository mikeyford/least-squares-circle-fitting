{\rtf1\ansi\ansicpg1252\cocoartf1504\cocoasubrtf820
{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue128;\red255\green255\blue255;\red0\green0\blue0;
\red0\green128\blue0;}
{\*\expandedcolortbl;;\csgenericrgb\c0\c0\c50196;\csgenericrgb\c100000\c100000\c100000;\csgenericrgb\c0\c0\c0;
\csgenericrgb\c0\c50196\c0;}
\paperw11900\paperh16840\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx560\tx1120\tx1680\tx2240\tx2800\tx3360\tx3920\tx4480\tx5040\tx5600\tx6160\tx6720\pardirnatural\partightenfactor0

\f0\fs22 \cf2 \cb3 Sub\cf4  Solve()\
\cf0 \
\cf4 f = Sheets("Input").Cells(1, 14) \cf5 ' Set variable for number of lines in functional model\
\cf4 l = Sheets("Input").Cells(2, 14) \cf5 ' Set variable for number of measurements\
\cf0 \
\cf2 For\cf4  m = 1 \cf2 To\cf4  l\
    \cf2 For\cf4  n = 1 \cf2 To\cf4  l\
        Sheets("C").Cells(m, n) = "0" \cf5 ' Generate empty C matrix\
\cf4     \cf2 Next\cf4  n\
\cf2 Next\cf4  m\
\cf0 \
\cf4 m = 1 \cf5 ' set C matrix position variable\
\cf2 For\cf4  n = 1 \cf2 To\cf4  f\
    \cf2 If\cf4  Sheets("Input").Cells(n, 1).Value = "Left Tangent" \cf2 Then\
\cf4         Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) \cf5 ' Insert X-error into C\
\cf4         Sheets("C").Cells(m, m + 1) = Sheets("Input").Cells(n, 7) \cf5 ' Insert X-row covariance into C\
\cf4         m = m + 1\
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 6) \cf5 ' Insert Y-error into C\
\cf4         Sheets("C").Cells(m, m - 1) = Sheets("Input").Cells(n, 7) \cf5 ' Insert Y-row covariance into C\
\cf4         m = m + 1\
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 8) \cf5 ' Insert Angle-error into C\
\cf4         m = m + 1\
        Sheets("A").Cells(n, 1) = "1" \cf5 ' Insert radius coefficient into A col 1\
\cf4         Sheets("A").Cells(n, 2) = "=-COS(RADIANS(Input!D" & n & "))" \cf5 '  Insert -cos(alpha) into A col 2\
\cf4         Sheets("A").Cells(n, 3) = "=SIN(RADIANS(Input!D" & n & "))" \cf5 ' Insert sin(alpha) into A col 3\
\cf4         Sheets("b").Cells(n, 1) = "=-(Input!$M$10-Input!$M$11*COS(RADIANS(Input!D" & n & "))+Input!$M$12*SIN(RADIANS(Input!D" & n & "))+Input!B" & n & "*COS(RADIANS(Input!D" & n & "))-Input!C" & n & "*SIN(RADIANS(Input!D" & n & ")))" \cf5 ' Insert b equation\
\cf4     \cf2 End\cf4  \cf2 If\
\cf4     \cf2 If\cf4  Sheets("Input").Cells(n, 1).Value = "Right Tangent" \cf2 Then\
\cf4         Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) \cf5 ' Insert X-error into C\
\cf4         Sheets("C").Cells(m, m + 1) = Sheets("Input").Cells(n, 7) \cf5 ' Insert X-row covariance into C\
\cf4         m = m + 1\
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) \cf5 ' Insert Y-error into C\
\cf4         Sheets("C").Cells(m, m - 1) = Sheets("Input").Cells(n, 7) \cf5 ' Insert Y-row covariance into C\
\cf4         m = m + 1\
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 8) \cf5 ' Insert Angle-error into C\
\cf4         m = m + 1\
        Sheets("A").Cells(n, 1) = "-1" \cf5 ' Insert radius coefficient into A col 1\
\cf4         Sheets("A").Cells(n, 2) = "=-COS(RADIANS(Input!D" & n & "))" \cf5 ' Insert -cos(alpha) into A col 2\
\cf4         Sheets("A").Cells(n, 3) = "=SIN(RADIANS(Input!D" & n & "))" \cf5 ' Insert sin(alpha) into A col 3\
\cf4         Sheets("b").Cells(n, 1) = "=-(-Input!$M$10-Input!$M$11*COS(RADIANS(Input!D" & n & "))+Input!$M$12*SIN(RADIANS(Input!D" & n & "))+Input!B" & n & "*COS(RADIANS(Input!D" & n & "))-Input!C" & n & "*SIN(RADIANS(Input!D" & n & ")))" \cf5 ' Insert b equation\
\cf4     \cf2 End\cf4  \cf2 If\
\cf4     \cf2 If\cf4  Sheets("Input").Cells(n, 1).Value = "Point" \cf2 Then\
\cf4         Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 5) \cf5 ' Insert X-error into C\
\cf4         Sheets("C").Cells(m, m + 1) = Sheets("Input").Cells(n, 7) \cf5 ' Insert X-row covariance into C\
\cf4         m = m + 1\
        Sheets("C").Cells(m, m) = Sheets("Input").Cells(n, 6) \cf5 ' Insert Y-error into C\
\cf4         Sheets("C").Cells(m, m - 1) = Sheets("Input").Cells(n, 7) \cf5 ' Insert Y-row covariance into C\
\cf4         m = m + 1\
        Sheets("A").Cells(n, 1) = "=-2*Input!$M$10" \cf5 ' Insert -2r into A col 1\
\cf4         Sheets("A").Cells(n, 2) = "=-2*(Input!B" & n & "-Input!$M$11)" \cf5 ' Insert -2*(x_p-x_c) into A col 2"\
\cf4         Sheets("A").Cells(n, 3) = "=-2*(Input!C" & n & "-Input!$M$12)" \cf5 ' Insert -2*(y_p-y_c) into A col 3"\
\cf4         Sheets("b").Cells(n, 1) = "=-(-(Input!$M$10^2)+(Input!B" & n & "-Input!$M$11)^2+(Input!C" & n & "-Input!$M$12)^2)" \cf5 ' Insert equation into b\
\cf4     \cf2 End\cf4  \cf2 If\
Next\cf4  n\
\cf0 \
\cf2 For\cf4  m = 1 \cf2 To\cf4  f\
    \cf2 For\cf4  n = 1 \cf2 To\cf4  l\
        Sheets("M").Cells(m, n) = "0" \cf5 ' Generate empty M matrix\
\cf4     \cf2 Next\cf4  n\
\cf2 Next\cf4  m\
\cf0 \
\cf4 m = 1 \cf5 ' reset as M matrix position variable\
\cf2 For\cf4  n = 1 \cf2 To\cf4  f\
    \cf2 If\cf4  Sheets("Input").Cells(n, 1).Value = "Left Tangent" \cf2 Then\
\cf4         Sheets("M").Cells(n, m).Value = "=COS(RADIANS(Input!D" & n & "))" \cf5 ' set cos(alpha) in M\
\cf4         m = m + 1\
        Sheets("M").Cells(n, m).Value = "=-SIN(RADIANS(Input!D" & n & "))" \cf5 ' set -sin(alpha) in M\
\cf4         m = m + 1\
        Sheets("M").Cells(n, m).Value = "=-Input!$M$11*-SIN(RADIANS(Input!D" & n & "))+Input!$M$12*COS(RADIANS(Input!D" & n & "))+Input!B" & n & "*-SIN(RADIANS(Input!D" & n & "))-Input!C" & n & "*COS(RADIANS(Input!D" & n & "))" \cf5 ' Add partial with respect to alpha to M\
\cf4         m = m + 1\
    \cf2 End\cf4  \cf2 If\
\cf4     \cf2 If\cf4  Sheets("Input").Cells(n, 1).Value = "Right Tangent" \cf2 Then\
\cf4         Sheets("M").Cells(n, m).Value = "=COS(RADIANS(Input!D" & n & "))" \cf5 ' set cos(alpha) in M\
\cf4         m = m + 1\
        Sheets("M").Cells(n, m).Value = "=-SIN(RADIANS(Input!D" & n & "))" \cf5 ' set -sin(alpha) in M\
\cf4         m = m + 1\
        Sheets("M").Cells(n, m).Value = "=-Input!$M$11*-SIN(RADIANS(Input!D" & n & "))+Input!$M$12*COS(RADIANS(Input!D" & n & "))+Input!B" & n & "*-SIN(RADIANS(Input!D" & n & "))-Input!C" & n & "*COS(RADIANS(Input!D" & n & "))" \cf5 ' Add partial with respect to alpha to M\
\cf4         m = m + 1\
    \cf2 End\cf4  \cf2 If\
\cf4     \cf2 If\cf4  Sheets("Input").Cells(n, 1).Value = "Point" \cf2 Then\
\cf4         Sheets("M").Cells(n, m).Value = "=2*(Input!B" & n & "-Input!$M$11)" \cf5 ' Insert 2(x_p - x_c) into M\
\cf4         m = m + 1\
        Sheets("M").Cells(n, m).Value = "=2*(Input!C" & n & "-Input!$M$12)" \cf5 ' Insert 2(y_p - y_c) into M\
\cf4         m = m + 1\
    \cf2 End\cf4  \cf2 If\
Next\cf4  n\
\cf0 \
\
\cf5 ' Next section perform the solution iteration\
\cf4 n = Sheets("Input").Cells(15, 13) \cf5 ' Get max number of iterations\
\cf4 d = Sheets("Input").Cells(16, 13) \cf5 ' Get decimal points required\
\cf4 c = 0\
\cf2 For\cf4  i = 1 \cf2 To\cf4  n\
    r_input = Round(Sheets("Input").Cells(10, 13).Value, d)\
    x_input = Round(Sheets("Input").Cells(11, 13).Value, d)\
    y_input = Round(Sheets("Input").Cells(12, 13).Value, d)\
    r_output = Round(Sheets("Input").Cells(19, 13).Value, d)\
    x_output = Round(Sheets("Input").Cells(20, 13).Value, d)\
    y_output = Round(Sheets("Input").Cells(21, 13).Value, d)\
    \cf2 If\cf4  r_input <> r_output \cf2 Or\cf4  x_input <> x_output \cf2 Or\cf4  y_input <> y_output \cf2 Then\cf4  \cf5 ' check if the input and output values agree\
\cf4         Sheets("Iterations").Cells(1, i) = "Iteration " & i\
        Sheets("Iterations").Cells(2, i) = Sheets("Input").Cells(19, 13).Value\
        Sheets("Iterations").Cells(3, i) = Sheets("Input").Cells(20, 13).Value\
        Sheets("Iterations").Cells(4, i) = Sheets("Input").Cells(21, 13).Value\
        c = c + 1\
    \cf2 End\cf4  \cf2 If\
\cf4     Sheets("Input").Cells(10, 13).Value = Sheets("Input").Cells(19, 13).Value\
    Sheets("Input").Cells(11, 13).Value = Sheets("Input").Cells(20, 13).Value\
    Sheets("Input").Cells(12, 13).Value = Sheets("Input").Cells(21, 13).Value\
\cf2 Next\cf4  i\
\cf0 \
\cf5 ' Populate residual standard errors from Cv\
\cf2 For\cf4  i = 1 \cf2 To\cf4  l\
    Sheets("Residuals").Cells(i, 2).Value = "=SQRT(" & Sheets("Cv").Cells(i, i).Value & ")"\
\cf2 Next\cf4  i\
\cf0 \
\cf5 ' Populate residuals standard variances from error covariance matrix\
\cf2 For\cf4  i = 1 \cf2 To\cf4  l\
    Sheets("Residuals").Cells(i, 5).Value = Sheets("C").Cells(i, i).Value\
\cf2 Next\cf4  i\
\cf0 \
\cf5 ' Populate MDE Matrix\
\cf2 For\cf4  i = 1 \cf2 To\cf4  l\
    Sheets("MDE").Cells(i, i).Value = Sheets("Residuals").Cells(i, 6).Value\
\cf2 Next\cf4  i\
\cf0 \
\
\cf5 ' Give user warning if the solution did not converge or W-test exceeded\
\cf2 If\cf4  r_input = r_output \cf2 And\cf4  x_input = x_output \cf2 And\cf4  y_input = y_output \cf2 Then\
\cf4     \cf2 If\cf4  Sheets("Residuals").Cells(2, 8).Value > Sheets("Test").Cells(13, 3).Value \cf2 Then\
\cf4         MsgBox "Warning! The solution converged but the W-test statistic was exceeded. Check data for outliers.", vbCritical\
    \cf2 Else\
\cf4         MsgBox "The solution converged in " & c & " iterations."\
    \cf2 End\cf4  \cf2 If\
Else\
\cf4     MsgBox "Warning! The solution did not converge in " & n & " iterations.", vbCritical\
\cf2 End\cf4  \cf2 If\
\cf0 \
\cf2 End\cf4  \cf2 Sub}