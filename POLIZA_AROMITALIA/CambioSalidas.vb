Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos

Public Class CambioSalidas
    Public Shared FOLIODR1 As String = ""
    Public Shared CERRADOSALIDAS As String = ""
    Public Shared Sub NuevoCambioSalida()
        Using ConnEntrada As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaEntrada As New SqlDataAdapter("SELECT CVE_FOLIO, CAMPLIB1 FROM INSOFTEC_MINVE_OUT WHERE CAMPLIB10=-1 AND CAMPLIB1<>'NULL'", ConnEntrada)
            Dim dtEntrada As New DataTable
            adaEntrada.Fill(dtEntrada)
            ConnEntrada.Open()
            Dim i As Integer = (dtEntrada.Rows.Count - 1)
            Dim j As Integer = 1
            Do While (j <= i)
                FOLIOMINVE01 = dtEntrada.AsEnumerable().ElementAtOrDefault(j).Item(0)
                FOLIODR1 = dtEntrada.AsEnumerable().ElementAtOrDefault(j).Item(1)
                Dim cut As String = FOLIODR1.Substring(2)
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
                        FECH = FECHAELABMIN.ToString("yyyy-MM-dd")

                        Using ConnPeriod As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConnPeriod.Open()
                            Dim AdaPeriod As SqlDataReader = New SqlCommand("SELECT * FROM ADMPER WHERE PERIODO='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "'", ConnPeriod).ExecuteReader
                            AdaPeriod.Read()
                            CERRADOSALIDAS = AdaPeriod.Item(2)
                            ConnPeriod.Close()
                        End Using

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

                        'ITEMS
                        COSTOMIN = Math.Round(COSTOMIN, 2)
                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVEART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using
                        contarPartidas2 = contarPartidas2 + 1
                        CAMPLIB9 = Replace(CAMPLIB9, "-", "")
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            If CERRADOSALIDAS = 0 Then '**************************PARA VALIDAR QUE EL PERIODO ESTE ABIERTO
                                Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTMIN * COSTOMIN & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='Dr' AND CONCEP_PO LIKE '%" & REFERMIN & " -" & CVEART & "%' AND NUM_CTA ='" & CAMPLIB9 & "000000000002" & "' AND NUM_PART='" & contarPartidas2 & "'", ConAux)
                                cmdAux.ExecuteNonQuery()
                            End If
                            ConAux.Close()
                        End Using


                        CAMPLIB8 = Replace(CAMPLIB8, "-", "")
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            If CERRADOSALIDAS = 0 Then '**************************PARA VALIDAR QUE EL PERIODO ESTE ABIERTO
                                Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTMIN * COSTOMIN & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='Dr' AND CONCEP_PO LIKE '%" & CVEART & "%' AND NUM_CTA ='" & CAMPLIB8 & "000000000003" & "' AND NUM_PART='" & contarPartidas2 + 1 & "'", ConAux)
                                cmdAux.ExecuteNonQuery()
                            End If
                            ConAux.Close()
                        End Using
                        totpartida = 0
                        CVEART = ""
                        NUMMOV = ""
                        CVECPTO = ""
                        TIPODOC = ""
                        REFERMIN = ""
                        CANTMIN = 0
                        CANTCOST = 0
                        PRECIOMIN = 0
                        COSTOMIN = 0
                        UNIVENTA = ""
                        SIGNO = 0
                        CAMPLIB8 = ""
                        CAMPLIB9 = ""


                        contarPartidas2 += 1
                    Loop
                    ConnEntrada1.Close()
                End Using
                'HEADER
                
                'UPDATE INSOFTEC_MINVE01
                Using ConnMinveInsof As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConnMinveInsof.Open()
                    Dim cmdConnMinveinsof As New SqlCommand("UPDATE INSOFTEC_MINVE_OUT SET CAMPLIB10=NULL WHERE CVE_FOLIO='" & CVEFOLIO & "'", ConnMinveInsof)
                    cmdConnMinveinsof.ExecuteNonQuery()
                    ConnMinveInsof.Close()
                End Using

                Using ConnPoliza As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnPoliza.Open()
                    If CERRADOSALIDAS = 0 Then '**************PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        Dim cmdPol As New SqlCommand("UPDATE " & POLIZA & " SET CONTABILIZ='N' WHERE CONCEP_PO ='" & "SALIDA " & REFERMIN + "  " & FECHAELABMIN.ToString("dd-MMMM-yyyy").ToUpper & "' AND NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO=' " & PERIODO & "' AND TIPO_POLI='Dr'", ConnPoliza)
                        cmdPol.ExecuteNonQuery()
                    End If
                    ConnPoliza.Close()
                End Using

                Using ConnMinveInsof As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConnMinveInsof.Open()
                    Dim cmdConnMinveinsof As New SqlCommand("DELETE INSOFTEC_CAMBIOS_IN_OUT WHERE CVE_FOLIO='" & CVEFOLIO & "' AND SIGNO=-1", ConnMinveInsof)
                    cmdConnMinveinsof.ExecuteNonQuery()
                    ConnMinveInsof.Close()
                End Using
                contador = 0
                FOLIOMINVE01 = ""
                FOLIODR1 = ""
                CERRADOSALIDAS = ""
                contarPartidas2 = 0
                j += 1
            Loop
            ConnEntrada.Close()
        End Using
    End Sub
End Class
