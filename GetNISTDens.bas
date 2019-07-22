REM This Basic script only supports Microsoft Office Excel for now.

Function GetDensNIST(fluid As String, pressure As Double, temperature As Double) As Double

    Dim request As Object
    Set request = CreateObject("MSXML2.XMLHTTP")

    Dim URL As String
    Dim FluidID As String

    If (StrComp(fluid, "N2", vbTextCompare) = 0) Or (StrComp(fluid, "Nitrogen", vbTextCompare) = 0) Then
        FluidID = "C7727379"
    Else
        If (StrComp(fluid, "H2O", vbTextCompare) = 0) Or (StrComp(fluid, "Water", vbTextCompare) = 0) Then
            FluidID = "C7732185"
        Else
            If (StrComp(fluid, "C3H8", vbTextCompare) = 0) Or (StrComp(fluid, "Propane", vbTextCompare) = 0) Then
                FluidID = "C74986"
            Else
                If (StrComp(fluid, "C7H16", vbTextCompare) = 0) Or (StrComp(fluid, "Heptane", vbTextCompare) = 0) Then
                    FluidID = "C142825"
                End If
            End If
        End If
    End If

    Dim Tstr As String
    Tstr = Str(temperature)
    Tstr = Trim(Replace(Tstr, ",", "."))

    Dim Pstr As String
    Pstr = Str(pressure + 14.7)
    Pstr = Trim(Replace(Pstr, ",", "."))

    URL = "http://webbook.nist.gov/cgi/fluid.cgi?Action=Data&Wide=on&ID=" & FluidID & "&Type=IsoTherm&Digits=5&PLow=" & Pstr & "&PHigh=" & Pstr & "&PInc=&T=" & Tstr & "&RefState=DEF&TUnit=C&PUnit=psia&DUnit=kg%2Fm3&HUnit=kJ%2Fmol&WUnit=m%2Fs&VisUnit=uPa*s&STUnit=N%2Fm"

    Dim SDens As String


    With request
        .Open "GET", URL, False
        .send
        Line = Split(.responseText, vbLf)(1)
        Debug.Print (.responseText)
        SDens = Split(Line, vbTab)(2)
        Debug.Print (SDens)
    End With

    GetDensNIST = 0.001 * Val(SDens)

End Function
