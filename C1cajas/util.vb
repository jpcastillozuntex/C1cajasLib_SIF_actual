Public Class util
    Public Function lector_dc() As String
        Dim dc As String = "Data Source=192.9.201.1\\JT;Initial Catalog=lector;Persist Security Info=True;User ID=user_l;Password=Lector_01"
        Return dc
    End Function
    Public Function get_mtp() As String
        Dim dc As String = "Server=192.9.201.1\\TEXF_D;Database=MTP;User Id=cop_f;Password=Copito_t12;"
        Return dc
    End Function

    Public Function con_string(ByVal e As Integer) As String
        Dim database As String = "inventarios"
        Dim constr As String
        Select Case e
            Case 0
                database = "inventarios"
            Case 1
                database = "TRECENTO"
            Case 2
                database = "zuntex"
        End Select
        constr = jap.b2b("083101114118101114061049057050046057046050048049046049092084069088070095068059100097116097098097115101061") + database + jap.b2b("059085115101114032073100061117115117097114105111059080097115115119111114100061076101099115113116095052054059")

        Return constr
    End Function

    Public Function con_ole(ByVal e As Integer) As String
        Dim database As String = "inventarios"
        Dim conole As String = ""
        Select Case e
            Case 0
                database = "inventarios"
            Case 1
                database = "TRECENTO"
            Case 2
                database = "zuntex"
        End Select
        Return conole
    End Function

End Class
