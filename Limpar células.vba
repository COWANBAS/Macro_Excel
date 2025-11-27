Sub LIMPAR()
    Dim celulas As Variant
    Dim i As Integer
    celulas = Array( _
        "G12", "J12", "M12", "P12", _
        "G17", "J17", "M17", "P17", _
        "G18", "J18", "M18", "P18", _
        "G21", "J21", "M21", "P21", _
        "G22", "J22", "M22", "P22", _
        "G25", "J25", "M25", "P25", _
        "G26", "J26", "M26", "P26", _
        "G31", "J31", "M31", "P31", _
        "G32", "J32", "M32", "P32", _
        "G35", "J35", "M35", "P35", _
        "G36", "J36", "M36", "P36", _
        "G39", "J39", "M39", "P39", _
        "G40", "J40", "M40", "P40", _
        "G45", "J45", "M45", "P45", _
        "G46", "J46", "M46", "P46", _
        "G49", "J49", "M49", "P49", _
        "G50", "J50", "M50", "P50", _
        "G53", "J53", "M53", "P53", _
        "G54", "J54", "M54", "P54")
    For i = LBound(celulas) To UBound(celulas)
        Range(celulas(i)).Value = "-"
    Next i
End Sub

