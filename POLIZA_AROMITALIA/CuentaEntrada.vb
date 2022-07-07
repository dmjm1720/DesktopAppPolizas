Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Imports POLIZA_AROMITALIA.EnvioMailEntrada
Public Class CuentaEntrada
    Public Shared CAMPLIB3E As String = ""
    Public Shared CAMPLIB4E As String = ""
    Public Shared Sub AccountIn()
        Using ConnEntrada As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaEntrada As New SqlDataAdapter("SELECT CVE_FOLIO FROM INSOFTEC_MINVE_IN WHERE CAMPLIB1 IS NULL", ConnEntrada)
            Dim dtEntrada As New DataTable
            adaEntrada.Fill(dtEntrada)
            ConnEntrada.Open()
            Dim i As Integer = (dtEntrada.Rows.Count - 1)
            Dim j As Integer = 1
            Do While (j <= i)
                FOLIOMINVE01 = dtEntrada.AsEnumerable().ElementAtOrDefault(j).Item(0)

                Using ConnEntrada1 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaEntrada1 As New SqlDataAdapter("SELECT CVE_ART, CVE_FOLIO, FECHAELAB, REFER FROM MINVE01 WHERE CVE_FOLIO ='" & FOLIOMINVE01 & "' AND SIGNO=1", ConnEntrada1)
                    Dim dtEntrada1 As New DataTable
                    adaEntrada1.Fill(dtEntrada1)
                    ConnEntrada1.Open()
                    Dim x As Integer = (dtEntrada1.Rows.Count - 1)
                    Dim y As Integer = 0

                    Do While (y <= x)
                        CVEART = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(0)
                        CVEFOLIO = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(1)
                        FECHAELABMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(2)
                        REFERMIN = dtEntrada1.AsEnumerable().ElementAtOrDefault(y).Item(3)
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

                        'ITEMS
                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVEART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using
                        If CAMPLIB8 = "" Or CAMPLIB9 = "" Then
                            CE = CE + 1
                        End If




                        y += 1
                    Loop
                    ConnEntrada1.Close()
                End Using
                Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConInveClib.Open()
                    Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB3, CAMPLIB4 FROM INSOFTEC_MINVE_IN WHERE CVE_FOLIO = '" & CVEFOLIO & "'", ConInveClib).ExecuteReader
                    adaInveClib.Read()
                    CAMPLIB3E = adaInveClib.Item(0).ToString
                    CAMPLIB4E = adaInveClib.Item(1).ToString
                    ConInveClib.Close()
                End Using

                Using ConnMinveInsof As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConnMinveInsof.Open()
                    If CE > 0 AndAlso CAMPLIB3E = "" AndAlso CAMPLIB4E = "" Then
                        EME()
                        Dim cmdConnMinveinsof As New SqlCommand(("UPDATE INSOFTEC_MINVE_IN SET CAMPLIB3='S/CTA', CAMPLIB4='ENVIADO' WHERE CVE_FOLIO='" & CVEFOLIO & "'"), ConnMinveInsof)
                        cmdConnMinveinsof.ExecuteNonQuery()
                    End If
                    If CE = 0 Then
                        Dim cmdConnMinveinsofC As New SqlCommand(("UPDATE INSOFTEC_MINVE_IN SET CAMPLIB3=NULL WHERE CVE_FOLIO='" & CVEFOLIO & "'"), ConnMinveInsof)
                        cmdConnMinveinsofC.ExecuteNonQuery()
                    End If
                    ConnMinveInsof.Close()
                End Using


                contador = 0
                CE = 0
                FOLIOMINVE01 = ""
                CVEFOLIO = ""
                CVEART = ""
                CAMPLIB8 = ""
                CAMPLIB9 = ""
                CAMPLIB3E = ""
                CAMPLIB4E = ""
                j += 1
            Loop
            ConnEntrada.Close()
        End Using
    End Sub
End Class
