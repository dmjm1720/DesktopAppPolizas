
Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Public Class InsertarSignoMenos
    Public Shared Sub InsertMenos()
        Using ConnEntrada As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaEntrada As New SqlDataAdapter("SELECT DISTINCT CVE_FOLIO FROM MINVE01 WHERE CVE_CPTO <>51 AND CVE_CPTO<>52 AND CVE_CPTO<>56 AND CVE_CPTO<>57 AND CVE_CPTO<>58 AND CVE_CPTO<>64 AND SIGNO=-1 AND YEAR(FECHA_DOCU)>2015", ConnEntrada)
            Dim dtEntrada As New DataTable
            adaEntrada.Fill(dtEntrada)
            ConnEntrada.Open()
            Dim i As Integer = (dtEntrada.Rows.Count - 1)
            Dim j As Integer = 0
            Do While (j <= i)
                FOLIOMINVE01 = dtEntrada.AsEnumerable().ElementAtOrDefault(j).Item(0)
                Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConAux.Open()
                    Dim cmdAux As New SqlCommand(("INSERT INTO INSOFTEC_MINVE_OUT (CVE_FOLIO) VALUES ('" & FOLIOMINVE01 & "')"), ConAux)
                    Try
                        cmdAux.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                    ConAux.Close()
                End Using
                FOLIOMINVE01 = ""
                j += 1
            Loop
            ConnEntrada.Close()
        End Using
    End Sub
End Class
