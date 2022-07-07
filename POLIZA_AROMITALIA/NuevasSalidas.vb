Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Public Class NuevasSalidas
    Public Shared Sub Salidas()
        Using ConnEntrada As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaEntrada As New SqlDataAdapter("SELECT CVE_FOLIO FROM INSOFTEC_MINVE_OUT WHERE CAMPLIB1 IS NULL AND CAMPLIB3 IS NULL", ConnEntrada)
            Dim dtEntrada As New DataTable
            adaEntrada.Fill(dtEntrada)
            ConnEntrada.Open()
            Dim i As Integer = (dtEntrada.Rows.Count - 1)
            Dim j As Integer = 1
            Do While (j <= i)
                FOLIOMINVE01 = dtEntrada.AsEnumerable().ElementAtOrDefault(j).Item(0)

                Using ConnEntrada1 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaEntrada1 As New SqlDataAdapter("SELECT CVE_ART, NUM_MOV, CVE_CPTO, FECHA_DOCU, TIPO_DOC, REFER, CLAVE_CLPV, CANT, CANT_COST, PRECIO, COSTO, UNI_VENTA, FECHAELAB, CVE_FOLIO, SIGNO FROM MINVE01 WHERE CVE_FOLIO ='" & FOLIOMINVE01 & "' AND SIGNO=-1", ConnEntrada1)
                    Dim dtEntrada1 As New DataTable
                    adaEntrada1.Fill(dtEntrada1)
                    ConnEntrada1.Open()
                    Dim x As Integer = (dtEntrada1.Rows.Count - 1)
                    Dim y As Integer = 0

                    Do While (y <= x)
                        CVEART = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(0)
                        NUMMOV = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(1)
                        CVECPTO = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(2)
                        FECHADOCU = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(3)
                        TIPODOC = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(4)
                        REFERMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(5)
                        CLAVECLPV = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(6).ToString
                        CANTMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(7)
                        CANTCOST = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(8)
                        PRECIOMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(9)
                        COSTOMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(10)
                        UNIVENTA = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(11)
                        FECHAELABMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(12)
                        CVEFOLIO = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(13)
                        SIGNO = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(14)
                        Dim extrae_per As String = FECHAELABMIN
                        Dim Testarray() As String = extrae_per.Split("/")
                        'EJERCICIO = Testarray(2).Trim
                        EJERCICIO = FECHAELABMIN.Year
                        PERIODO = Testarray(1).Trim
                        PERIODO1 = Testarray(1).Trim
                        Dim recortar As String = EJERCICIO
                        Dim EJER_FISCAL As String = Microsoft.VisualBasic.Right(recortar, 2)
                        AUXILIAR = AUXI & EJER_FISCAL
                        POLIZA = POLIZ & EJER_FISCAL
                        CUENTA = CUENT & EJER_FISCAL
                        FECH = FECHAELABMIN.ToString("yyyy-dd-MM")
                        y += 1
                        Using ConPol As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConPol.Open()
                            Dim adaPol As SqlDataReader = New SqlCommand("SELECT MAX(NUM_POLIZ) FROM  " & POLIZA & " WHERE PERIODO ='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "' AND TIPO_POLI='Dr'", ConPol).ExecuteReader
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
                        If contador = 0 Then
                            contador = 1
                        End If

                        Using ConFol As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConFol.Open()
                            Dim cmdFol As New SqlCommand(("UPDATE " & FOLIOS & " SET FOLIO" & PERIODO1 & " ='" & contador + 1 & "' WHERE EJERCICIO ='" & EJERCICIO & "' AND TIPPOL='Dr'"), ConFol)
                            cmdFol.ExecuteNonQuery()
                            ConFol.Close()
                        End Using
                        'ITEMS
                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVEART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using
                        Using ConnAuxCon As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConnAuxCon.Open()
                            Dim adaConnAuxCon As SqlDataReader = New SqlCommand("SELECT COUNT(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='Dr'  AND NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", ConnAuxCon).ExecuteReader
                            adaConnAuxCon.Read()
                            con1 = adaConnAuxCon.Item(0)
                            ConnAuxCon.Close()
                        End Using
                        CAMPLIB8 = Replace(CAMPLIB8, "-", "")
                        CAMPLIB9 = Replace(CAMPLIB9, "-", "")
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            Dim cmdAux As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Dr',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & con1 + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CAMPLIB9 & "000000000002" & "','" & FECH & "','" & REFERMIN & " -" & CVEART & "','D',' " & Math.Round(CANTMIN * COSTOMIN, 2) & "','0','1','0','" & con1 + 1 & "','0','0');"), ConAux)
                            cmdAux.ExecuteNonQuery()
                            ConAux.Close()
                        End Using

                        Using ConnAuxCon As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConnAuxCon.Open()
                            Dim adaConnAuxCon As SqlDataReader = New SqlCommand("SELECT COUNT(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='Dr'  AND NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", ConnAuxCon).ExecuteReader
                            adaConnAuxCon.Read()
                            con2 = adaConnAuxCon.Item(0)
                            ConnAuxCon.Close()
                        End Using
                        CAMPLIB8 = Replace(CAMPLIB8, "-", "")
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            Dim cmdAux As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Dr',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & con2 + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CAMPLIB8 & "000000000003" & "','" & FECH & "','COSTO -" & CVEART & "','H','" & Math.Round(CANTMIN * COSTOMIN, 2) & "','0','1','0','" & con2 + 1 & "','0','0');"), ConAux)
                            cmdAux.ExecuteNonQuery()
                            ConAux.Close()
                        End Using
                        totpartida = 0
                        CVEART = ""
                        NUMMOV = ""
                        CVECPTO = ""
                        TIPODOC = ""
                        'REFERMIN = ""
                        CANTMIN = 0
                        CANTCOST = 0
                        PRECIOMIN = 0
                        COSTOMIN = 0
                        UNIVENTA = ""
                        SIGNO = 0
                        CAMPLIB8 = ""
                        CAMPLIB9 = ""

                        partidasNo += 1
                    Loop
                    ConnEntrada1.Close()
                End Using
                'HEADER
                Using ConnHeader As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnHeader.Open()
                    Dim cmdHeader As New SqlCommand(("INSERT INTO " & POLIZA & "  (TIPO_POLI,NUM_POLIZ,PERIODO,EJERCICIO,FECHA_POL,CONCEP_PO,NUM_PART,LOGAUDITA,CONTABILIZ,NUMPARCUA,TIENEDOCUMENTOS,ORIGEN) VALUES ('Dr',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & PERIODO & "','" & EJERCICIO & "','" & FECH & "','" & "SALIDA " & REFERMIN + "  " & FECHAELABMIN.ToString("dd-MMMM-yyyy").ToUpper & "','" & partidasNo & "','N','N','0','0','SAE-FAC-AUT');"), ConnHeader) 'totpartida
                    cmdHeader.ExecuteNonQuery()
                    ConnHeader.Close()
                End Using
                'UPDATE INSOFTEC_MINVE01
                Using ConnMinveInsof As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConnMinveInsof.Open()
                    Dim cmdConnMinveinsof As New SqlCommand(("UPDATE INSOFTEC_MINVE_OUT SET CAMPLIB1='Dr" & contador & "', CAMPLIB2='PROCESO 1', FECHA_DOCU='" & FECHADOCU.ToString("yyyy-dd-MM") & "' WHERE CVE_FOLIO='" & CVEFOLIO & "'"), ConnMinveInsof)
                    cmdConnMinveinsof.ExecuteNonQuery()
                    ConnMinveInsof.Close()
                End Using
                contador = 0
                FOLIOMINVE01 = ""

                j += 1
            Loop
            ConnEntrada.Close()
        End Using
    End Sub
End Class
