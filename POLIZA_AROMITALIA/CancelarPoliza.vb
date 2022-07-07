Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos

Public Class CancelarPoliza
    Public Shared Sub FacturaPolizaCancelda()
        Using cn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim ada As New SqlDataAdapter("SELECT CVE_DOC, FECHA_DOC FROM  " & FACTF & " INNER JOIN INSOFTEC_FACTF_CLIB01 ON " & FACTF & ".CVE_DOC = INSOFTEC_FACTF_CLIB01.CLAVE_DOC WHERE (INSOFTEC_FACTF_CLIB01.CAMPLIB1 >'" & TIPOFACTF & "' AND INSOFTEC_FACTF_CLIB01.CAMPLIB2 IS NULL) AND (" & FACTF & ".STATUS = 'C')", cn)
            Dim dt As New DataTable
            ada.Fill(dt)
            cn.Open()
            Dim j As Integer = (dt.Rows.Count)
            Dim i As Integer = 0
            Try
                Do While (i <= j)
                    CVE_DOC = dt.AsEnumerable().ElementAtOrDefault(i).Item(0).ToString
                    FECHA_DOC = dt.AsEnumerable().ElementAtOrDefault(i).Item(1)
                    'PARA TOMAR EL AÑO DEL EJERCICIO FISCAL
                    Dim extrae_per As String = FECHA_DOC
                    Dim Testarray() As String = extrae_per.Split("/")
                    EJERCICIO = Testarray(2).Trim
                    PERIODO = Testarray(1).Trim
                    PERIODO1 = Testarray(1).Trim
                    Dim recortar As String = EJERCICIO
                    Dim EJER_FISCAL As String = Microsoft.VisualBasic.Right(recortar, 2)
                    AUXILIAR = AUXI & EJER_FISCAL
                    POLIZA = POLIZ & EJER_FISCAL
                    Using cn1 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        cn1.Open()
                        Dim cmd1 As New SqlCommand(("UPDATE " & AUXILIAR & " set MONTOMOV = 0 WHERE TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" + CVE_DOC + "%'"), cn1)
                        cmd1.ExecuteNonQuery()
                        cn1.Close()
                        Using cn2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            cn2.Open()
                            Dim command6611 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01 SET CAMPLIB1 ='CANCELADO', CAMPLIB3='PROCESO 2' WHERE CLAVE_DOC= '" + CVE_DOC + "'"), cn2)
                            command6611.ExecuteNonQuery()
                            cn2.Close()
                        End Using
                        Using cn3 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            cn3.Open()
                            Dim adapter3 As SqlDataReader = New SqlCommand("SELECT CONCEP_PO FROM  " & POLIZA & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" + CVE_DOC + "%'", cn3).ExecuteReader
                            adapter3.Read()
                            DATO = adapter3.Item(0)
                        End Using

                        Using cn4 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            cn4.Open()
                            'ACTUALIZA EL FOLIO
                            Dim command6611 As New SqlCommand(("UPDATE " & POLIZA & "  SET CONCEP_PO='" & DATO + "  **CANCELADO**" & "', CONTABILIZ='N' WHERE TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO  = '" + DATO + "'"), cn4)
                            command6611.ExecuteNonQuery()
                            cn4.Close()

                        End Using

                    End Using

                    i += 1
                Loop
            Catch ex As Exception
                Exit Try
            End Try
        End Using
        CVE_DOC = ""
    End Sub
End Class
