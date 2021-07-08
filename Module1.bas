Attribute VB_Name = "Module1"
Function colebrook(ks, dd, re)
''developed by Fernando Nunez
''fcnunezp@gmail.com

colebrook = 0.014
f = 0

Do
f = colebrook
colebrook = (-(2 / Log(10)) * Log(0.27 * ks / dd + 2.51 / (re * Sqr(f)))) ^ (-2)
Loop Until Abs(f - colebrook) < 0.000001

End Function



Sub time_gravity()
''developed by Fernando Nunez
''fcnunezp@gmail.com

    Sheets("Sheet1").Range("B14:K14").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

visc = Sheets("Sheet1").Range("B8").Value
D1 = Sheets("Sheet1").Range("D1").Value
D2 = Sheets("Sheet1").Range("D2").Value
ro = Sheets("Sheet1").Range("B7").Value
dt = Sheets("Sheet1").Range("B5").Value
Slope = Sheets("Sheet1").Range("E3").Value
fss = Sheets("Sheet1").Range("B10").Value
dens = Sheets("Sheet1").Range("F5").Value

a1 = 3.1416 * (D1 * D1) * 0.25
a2 = 3.1416 * (D2 * D2) * 0.25
g = 9.81


For i = 14 To 10000

Sheets("Sheet1").Range("B" & CStr(i)).Value = Sheets("Sheet1").Range("B" & CStr(i - 1)).Value + dt

L = Sheets("Sheet1").Range("G" & CStr(i - 1)).Value
H = Sheets("Sheet1").Range("H" & CStr(i - 1)).Value


vel2 = 30 ''estimación inicial m/s

Do
vel = vel2

Rey = dens * vel * D1 / (visc / 1000)

If Rey > 4000 Then

ff = colebrook(ro, D1, Rey)
Else
ff = 64 / Rey
End If

caudal = H * 2 * g / (((1 + k) / (a2 * a2)) - ((1 - (ff * (1 + fss) * L / D1)) / (a1 * a1)))

vel2 = caudal / a1

Loop Until Abs(vel2 - vel) < 0.00001

''Sheets("Sheet1").Range("F13").Value = vel2





Sheets("Sheet1").Range("C" & CStr(i)).Value = caudal
Sheets("Sheet1").Range("D" & CStr(i)).Value = vel2
Sheets("Sheet1").Range("E" & CStr(i)).Value = Sheets("Sheet1").Range("I" & CStr(i - 1)).Value
Sheets("Sheet1").Range("F" & CStr(i)).Value = caudal * dt
Sheets("Sheet1").Range("G" & CStr(i)).Value = L - ((4 * caudal * dt) / (3.1416 * (D1 * D1)))
Sheets("Sheet1").Range("H" & CStr(i)).Value = Sheets("Sheet1").Range("G" & CStr(i)).Value * Slope

Sheets("Sheet1").Range("I" & CStr(i)).Value = Sheets("Sheet1").Range("E" & CStr(i)).Value - Sheets("Sheet1").Range("F" & CStr(i)).Value
Sheets("Sheet1").Range("J" & CStr(i)).Value = Rey
Sheets("Sheet1").Range("K" & CStr(i)).Value = caudal / a2

If Sheets("Sheet1").Range("I" & CStr(i)).Value < 0.1 Then

Exit For

End If


Next i

MsgBox "El tiempo de drenaje es: " & CStr(Round((i - 13) / 60, 2)) & " min"

End Sub
