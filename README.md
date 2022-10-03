# Gerak-lurus-beraturan-
Pada praktikumkali ini yaitu membuat pemodelan menggunakan Microsoft Excel, dimana pemodelan yang diambil adalah Gerak Lurus Beraturan (GLB). Konsep yang diterapkan yaitu menggunakan solusi dengan persamaan untuk jarak, yang melibatkan nilai
kecepatan dan waktu. Pada GLB ini kecepatan nya akan tetap konstan dan berada pada
lintasan yang berupa garis lurus
'
' Macro1 Macro
'

'
    ActiveCell.FormulaR1C1 = "=RC[2]*RC[-1]"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("F4").Select
End Sub
Sub Macro2()
'
' Macro2 Macro
'

'
    Range("H5").Select
    ActiveSheet.ScrollBars.Add(145.5, 120.75, 135, 45).Select
    With Selection
        .Value = 0
        .Min = 0
        .Max = 100
        .SmallChange = 1
        .LargeChange = 10
        .LinkedCell = "$F$3"
        .Display3DShading = True
    End With
    Range("E17").Select
End Sub
