Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Public Class ProcesarNotaCreditoFacte

    Public Shared Sub NotaCreditoFacte()
        Using conexFacte As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaFacte As New SqlDataAdapter("SELECT FACTE01.TIP_DOC, FACTE01.CVE_DOC, FACTE01.CVE_CLPV, FACTE01.STATUS, FACTE01.FECHA_DOC, FACTE01.FECHA_ENT, FACTE01.FECHA_VEN, FACTE01.CAN_TOT, FACTE01.IMP_TOT1, FACTE01.IMP_TOT4, FACTE01.CVE_OBS, FACTE01.NUM_ALMA, FACTE01.ACT_CXC, FACTE01.ACT_COI, FACTE01.ENLAZADO, FACTE01.TIP_DOC_E, FACTE01.NUM_MONED, FACTE01.TIPCAMB, FACTE01.NUM_PAGOS, FACTE01.PRIMERPAGO, FACTE01.RFC, FACTE01.CTLPOL, FACTE01.ESCFD, FACTE01.AUTORIZA, FACTE01.SERIE, FACTE01.FOLIO, FACTE01.DAT_ENVIO, FACTE01.CONTADO, FACTE01.CVE_BITA, FACTE01.BLOQ, FACTE01.FORMAENVIO, FACTE01.DES_FIN_PORC, FACTE01.DES_TOT_PORC, FACTE01.IMPORTE, FACTE01.COM_TOT_PORC, FACTE01.TIP_DOC_ANT,FACTE01.DOC_ANT FROM  FACTE01 INNER JOIN INSOFTEC_FACTE_CLIB01 ON FACTE01.CVE_DOC = INSOFTEC_FACTE_CLIB01.CLAVE_DOC WHERE (INSOFTEC_FACTE_CLIB01.CAMPLIB1 IS NULL AND INSOFTEC_FACTE_CLIB01.CAMPLIB4 IS NULL) AND (FACTE01.STATUS <> 'C')", conexFacte)
            Dim dtFacte As New DataTable
            adaFacte.Fill(dtFacte)
            conexFacte.Open()
            Dim i As Integer = (dtFacte.Rows.Count - 1)
            Dim j As Integer = 0
            Do While (j <= i)
                TIP_DOC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(0).ToString
                CVE_DOC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(1).ToString
                CVE_CLPV = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(2).ToString
                STATUS = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(3).ToString
                FECHA_DOC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(4)
                FECHA_ENT = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(5)
                FECHA_VEN = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(6)
                CAN_TOT = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(7)
                IMP_TOT1 = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(8)
                IMP_TOT4 = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(9)
                CVE_OBS = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(10)
                NUM_ALMA = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(11)
                ACT_CXC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(12).ToString
                ACT_COI = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(13).ToString
                ENLAZADO = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(14).ToString
                TIP_DOC_E = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(15).ToString
                NUM_MONED = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(16)
                TIPCAMB = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(17)
                NUM_PAGOS = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(18)
                PRIMERPAGO = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(19)
                RFC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(20).ToString
                CTLPOL = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(21)
                ESCFD = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(22).ToString
                AUTORIZA = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(23)
                SERIE = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(24).ToString
                FOLIO = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(25)
                DAT_ENVIO = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(26)
                CONTADO = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(27).ToString
                CVE_BITA = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(28)
                BLOQ = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(29).ToString
                FORMAENVIO = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(30).ToString
                DES_FIN_PORC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(31)
                DES_TOT_PORC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(32)
                IMPORTE = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(33)
                COM_TOT_PORC = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(34)
                TIP_DOC_ANT = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(35).ToString
                DOC_ANT = dtFacte.AsEnumerable().ElementAtOrDefault(j).Item(36).ToString

                Dim extrae_per As String = FECHA_DOC
                Dim Testarray() As String = extrae_per.Split("/")
                EJERCICIO = Testarray(2).Trim
                PERIODO = Testarray(1).Trim
                PERIODO1 = Testarray(1).Trim
                Dim recortar As String = EJERCICIO
                Dim EJER_FISCAL As String = Microsoft.VisualBasic.Right(recortar, 2)
                AUXILIAR = AUXI & EJER_FISCAL
                POLIZA = POLIZ & EJER_FISCAL
                CUENTA = CUENT & EJER_FISCAL
                Dim FECH As String = FECHA_DOC.ToString("yyyy-MM-dd")

                Using ConPol As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConPol.Open()
                    Dim adaPol As SqlDataReader = New SqlCommand("SELECT MAX(NUM_POLIZ) FROM " & POLIZA & " WHERE PERIODO ='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "' AND TIPO_POLI='" & TIPOFACTD & "'", ConPol).ExecuteReader
                    adaPol.Read()
                    Try
                        If adaPol.HasRows = True Then
                            contador = adaPol.Item(0)
                            contador = contador + 1
                        End If
                    Catch ex As Exception
                    End Try
                    ConPol.Close()
                End Using
            Loop
            conexFacte.Close()
        End Using
    End Sub
End Class
