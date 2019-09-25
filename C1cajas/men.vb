Public Class men
    Dim cn As New cnv.cvr
    Public Sub menu(ByVal password As String, ByRef nombre As String, ByVal ip As String, ByRef sec As String, ByRef dt As DataTable, ByRef ok As Boolean)
        Dim strsql As String
        Dim dr As DataRow
        Dim corre As String = ""
        Dim ipn As String
        Dim pw As New DataTable
        Dim o As New jap
        Dim p1 As String = ""
        Dim p2 As String = ""
        Dim p3 As String = ""
        Dim p4 As String = ""
        Dim p5 As String = ""
        o.fmjap(p3, p2, p1, p4, p5)
        dt = New DataTable
        ok = False
        strsql = "SELECT CORRELATIVO,CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',PASSWORD)) AS PASSWORD, CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p2 & "',NOMBRE)) AS NOMBRE, CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p1 & "',IP)) AS IP ,CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p4 & "',SECCION)) AS SECCION FROM CAJAS99 WHERE CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',PASSWORD)) = '" & password & "'"
        llena_tablas(pw, strsql)
        If pw.Rows.Count > 0 Then
            dr = pw.Rows(0)
            ipn = dr("IP")
            If Trim(ipn) <> "" Then
                If Trim(ipn) <> Trim(ip) Then
                    MsgBox("Acceso Denegado a esta terminal !!!!", MsgBoxStyle.Critical, "Utilice terminales autorizadas.")
                    Exit Sub
                End If
            End If
            nombre = dr("NOMBRE")
            ip = ipn
            sec = dr("SECCION")
            corre = dr("CORRELATIVO")
            o.fjap(p1, p2, p3, p4, p5)
            strsql = "SELECT CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',CAJAS97.MENU)) AS MENU , CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p4 & "',DESCRIPCION)) AS DESCRIPCION,CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p2 & "',PROGRAMA)) AS PROGRAMA,ALTO,ANCHO,COLOR,CONTROL FROM CAJAS97 LEFT JOIN CAJAS98 ON CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',CAJAS97.MENU)) = CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p5 & "',CAJAS98.MENU)) WHERE USUARIO = '" & corre & "' AND LEFT(CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',CAJAS97.MENU)),1) = 'P' ORDER BY MENU"
            llena_tablas(dt, strsql)
            ok = True
        End If
    End Sub
    Public Sub menu_m(ByVal password As String, ByRef nombre As String, ByVal ip As String, ByRef dt As DataTable, ByRef ok As Boolean)
        Dim strsql As String
        Dim dr As DataRow
        Dim o As New jap
        Dim corre As String = ""
        Dim ipn As String
        Dim pw As New DataTable
        Dim p1 As String = ""
        Dim p2 As String = ""
        Dim p3 As String = ""
        Dim p4 As String = ""
        Dim p5 As String = ""
        o.fcjap(p3, p2, p1, p4, p5)
        dt = New DataTable
        ok = False
        strsql = "SELECT CORRELATIVO,CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',PASSWORD)) AS PASSWORD, CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p2 & "',NOMBRE)) AS NOMBRE, CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p1 & "',IP)) AS IP FROM MENUS_SI_P WHERE CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',PASSWORD)) = '" & password & "'"
        llena_tablas(pw, strsql)
        If pw.Rows.Count > 0 Then
            dr = pw.Rows(0)
            ipn = dr("IP")
            If Trim(ipn) <> "" Then
                If Trim(ipn) <> Trim(ip) Then
                    MsgBox("Acceso Denegado a esta terminal !!!!", MsgBoxStyle.Critical, "Utilice terminales autorizadas.")
                    Exit Sub
                End If
            End If
            nombre = dr("NOMBRE")
            ip = ipn
            corre = dr("CORRELATIVO")
            o.fjap(p1, p2, p3, p4, p5)
            strsql = "SELECT CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p5 & "',CAJAS97.MENU)) AS MENU , CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p4 & "',DESCRIPCION)) AS DESCRIPCION,CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p2 & "',PROGRAMA)) AS PROGRAMA,ALTO,ANCHO,COLOR,CONTROL FROM CAJAS97 LEFT JOIN CAJAS98 ON CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p5 & "',CAJAS97.MENU)) = CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p3 & "',CAJAS97.MENU)) WHERE USUARIO = '" & corre & "' AND LEFT(CONVERT(VARCHAR(300),DECRYPTBYPASSPHRASE('" & p5 & "',CAJAS98.MENU)),1) = 'P' ORDER BY MENU"
            llena_tablas(dt, strsql)
            ok = True
        End If
    End Sub
        
    Public Sub llena_tablas(ByRef dt As DataTable, ByVal strSql As String)
        Dim cnn As New System.Data.SqlClient.SqlConnection()
        Dim da As System.Data.SqlClient.SqlDataAdapter
        cnn.ConnectionString = cn.b2b("083101114118101114061049057050046057046050048049046049092084069088070095068059100097116097098097115101061105110118101110116097114105111115059085115101114032073100061117115117097114105111059080097115115119111114100061076101099115113116095052054059")
        Dim ds As New DataSet()
        da = New System.Data.SqlClient.SqlDataAdapter(strSql, cnn)
        Try
            da.Fill(ds)
            dt = ds.Tables(0)
        Catch
        End Try
    End Sub
End Class
