Public Class jap
    Public Function otros(ByVal h As String) As String
        Dim result As String = ""
        For x As Integer = 0 To CInt(h.Length / 2) - 1
            result = result + Chr(CInt("&H" & h.Substring((x * 2), 2)))
        Next
        Return Trim(result)
    End Function

    Public Sub fcjap(ByRef x1 As String, ByRef x2 As String, ByRef x3 As String, ByRef x4 As String, ByRef x5 As String)
        x1 = otros("20204775657465303120202020")
        x2 = otros("20204775657465303220202020")
        x3 = otros("20204775657465303320202020")
        x4 = otros("20204775657465303420202020")
        x5 = otros("20204775657465303520202020")
    End Sub

    Public Sub fjap(ByRef x1 As String, ByRef x2 As String, ByRef x3 As String, ByRef x4 As String, ByRef x5 As String)
        x1 = otros("2020204D6170613030202020202020")
        x2 = otros("202020204D61706130332020202020")
        x3 = otros("20204D617061303120202020202020")
        x4 = otros("2020204D6170613032202020202020")
        x5 = otros("204C75636173303120202020202020")
    End Sub

    Public Sub fmjap(ByRef x1 As String, ByRef x2 As String, ByRef x3 As String, ByRef x4 As String, ByRef x5 As String)
        x1 = otros("2020506170617961202020")
        x2 = otros("2020526963616D61202020")
        x3 = otros("2020506174617461202020")
        x4 = otros("2020506F7A75656C6F2020")
        x5 = otros("2020456C6F746520202020")
    End Sub

    Public Shared Function b2b(ByVal cnst As String) As String
        ' The input String.
        Dim i As Integer
        Dim str As String = ""
        For i = 1 To cnst.Length Step 3
            str = str + ChrW(Mid(cnst, i, 3))
        Next

        Return str
    End Function
    Public Function b2h(ByVal cnst As String) As String
        ' The input String.
        Dim i As Integer
        Dim str As String = ""
        For i = 1 To cnst.Length Step 3
            str += ChrW(Mid(cnst, i, 3))
        Next

        Return str
    End Function

    'Public Sub crea_cajas(ByRef obj As Object, ByRef hola SqlConnection)

    'End Sub
End Class
