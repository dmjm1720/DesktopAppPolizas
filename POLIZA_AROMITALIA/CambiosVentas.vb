Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Imports POLIZA_AROMITALIA.ProcesarNewFactura

Public Class CambiosVentas
    Public Shared CAMPOPOLIZA As String = ""
    Public Shared CERRADO As String = ""
    Public Shared Sub ActualizarCambiosVentas()
        Using ConFact As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaFact As New SqlDataAdapter("SELECT " & FACTF & ".TIP_DOC, " & FACTF & ".CVE_DOC, " & FACTF & ".CVE_CLPV, " & FACTF & ".STATUS, " & FACTF & ".FECHA_DOC, " & FACTF & ".FECHA_ENT, " & FACTF & ".FECHA_VEN, " & FACTF & ".CAN_TOT, " & FACTF & ".IMP_TOT1, " & FACTF & ".IMP_TOT4, " & FACTF & ".CVE_OBS, " & FACTF & ".NUM_ALMA, " & FACTF & ".ACT_CXC, " & FACTF & ".ACT_COI, " & FACTF & ".ENLAZADO, " & FACTF & ".TIP_DOC_E, " & FACTF & ".NUM_MONED, " & FACTF & ".TIPCAMB, " & FACTF & ".NUM_PAGOS, " & FACTF & ".PRIMERPAGO, " & FACTF & ".RFC, " & FACTF & ".CTLPOL, " & FACTF & ".ESCFD, " & FACTF & ".AUTORIZA, " & FACTF & ".SERIE, " & FACTF & ".FOLIO, " & FACTF & ".DAT_ENVIO, " & FACTF & ".CONTADO, " & FACTF & ".CVE_BITA, " & FACTF & ".BLOQ, " & FACTF & ".FORMAENVIO, " & FACTF & ".DES_FIN_PORC, " & FACTF & ".DES_TOT_PORC, " & FACTF & ".IMPORTE, " & FACTF & ".COM_TOT_PORC, " & FACTF & ".TIP_DOC_ANT," & FACTF & ".DOC_ANT,INSOFTEC_FACTF_CLIB01.CAMPLIB1 FROM  " & FACTF & "  INNER JOIN INSOFTEC_FACTF_CLIB01 ON " & FACTF & ".CVE_DOC = INSOFTEC_FACTF_CLIB01.CLAVE_DOC WHERE (INSOFTEC_FACTF_CLIB01.CAMPLIB10=0 AND INSOFTEC_FACTF_CLIB01.CAMPLIB1<>'NULL') AND (" & FACTF & ".STATUS <> 'C')", ConFact)
            Dim dtFact As New DataTable
            adaFact.Fill(dtFact)
            ConFact.Open()
            Dim i As Integer = (dtFact.Rows.Count - 1)
            Dim j As Integer = 0
            Do While (j <= i)
                TIP_DOC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(0).ToString
                CVE_DOC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(1).ToString
                CVE_CLPV = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(2).ToString
                STATUS = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(3).ToString
                FECHA_DOC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(4)
                FECHA_ENT = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(5)
                FECHA_VEN = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(6)
                CAN_TOT = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(7)
                IMP_TOT1 = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(8)
                IMP_TOT4 = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(9)
                CVE_OBS = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(10)
                NUM_ALMA = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(11)
                ACT_CXC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(12).ToString
                ACT_COI = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(13).ToString
                ENLAZADO = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(14).ToString
                TIP_DOC_E = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(15).ToString
                NUM_MONED = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(16)
                TIPCAMB = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(17)
                NUM_PAGOS = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(18)
                PRIMERPAGO = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(19)
                RFC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(20).ToString
                CTLPOL = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(21)
                ESCFD = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(22).ToString
                AUTORIZA = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(23)
                SERIE = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(24).ToString
                FOLIO = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(25)
                DAT_ENVIO = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(26)
                CONTADO = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(27).ToString
                CVE_BITA = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(28)
                BLOQ = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(29).ToString
                FORMAENVIO = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(30).ToString
                DES_FIN_PORC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(31)
                DES_TOT_PORC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(32)
                IMPORTE = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(33)
                COM_TOT_PORC = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(34)
                TIP_DOC_ANT = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(35).ToString
                DOC_ANT = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(36).ToString
                CAMPOPOLIZA = dtFact.AsEnumerable().ElementAtOrDefault(j).Item(37).ToString

                Dim extrae_per As String = FECHA_DOC
                Dim Testarray() As String = extrae_per.Split("/")
                EJERCICIO = Testarray(2).Trim
                PERIODO = Testarray(1).Trim
                PERIODO1 = Testarray(1).Trim
                Dim recortar As String = EJERCICIO
                Dim EJER_FISCAL As String = Microsoft.VisualBasic.Right(recortar, 2)
                Dim cut As String = CAMPOPOLIZA.Substring(2)

                AUXILIAR = AUXI & EJER_FISCAL
                POLIZA = POLIZ & EJER_FISCAL
                CUENTA = CUENT & EJER_FISCAL
                Dim FECH As String = FECHA_DOC.ToString("yyyy-MM-dd")

                Using ConnPeriod As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnPeriod.Open()
                    Dim AdaPeriod As SqlDataReader = New SqlCommand("SELECT * FROM ADMPER WHERE PERIODO='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "'", ConnPeriod).ExecuteReader
                    AdaPeriod.Read()
                    CERRADO = AdaPeriod.Item(2)
                    ConnPeriod.Close()
                End Using

                Using ConPol As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConPol.Open()
                    Dim adaPol As SqlDataReader = New SqlCommand("Select max(NUM_POLIZ) from " & POLIZA & " WHERE PERIODO ='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "' AND TIPO_POLI='" & TIPOFACTF & "'", ConPol).ExecuteReader
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


                Using ConItems As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConItems.Open()
                    Dim adaItems As SqlDataReader = New SqlCommand("SELECT COUNT(CVE_DOC) FROM " & PAR_FACTF & " WHERE (CVE_DOC='" & CVE_DOC & "')", ConItems).ExecuteReader
                    adaItems.Read()
                    totpartida = adaItems.Item(0)
                    ConItems.Close()
                End Using

                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTF & " WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
                    Dim dtItemsSearch As New DataTable
                    adaItemsSearch.Fill(dtItemsSearch)
                    ConItemsSearch.Open()
                    Dim x As Integer = (dtItemsSearch.Rows.Count - 1)
                    Dim y As Integer = 0
                    Do While (y <= x)
                        NUM_PAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(0)
                        CVE_ART = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(1).ToString
                        CANT = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(2)
                        PXS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(3)
                        PREC = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(4)
                        COST = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(5)
                        IMPU4 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(6)
                        IMP1APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(7).ToString
                        IMP2APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(8).ToString
                        IMP3APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(9).ToString
                        DESC1 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(10)
                        DESC2 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(11)
                        APAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(12)
                        ACT_INV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(13).ToString
                        NUM_ALM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(14)
                        TIP_CAM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(15)
                        UNI_VENTA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(16).ToString
                        TIPO_PROD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(17)
                        CVE_OBS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(18)
                        REG_SERIE = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(19)
                        E_LTPD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(20)
                        TIPO_ELEM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(21).ToString
                        NUM_MOV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(22)
                        CANTIDADTOTAL = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(23)
                        IMPRIMIR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(24).ToString


                        Using ConInve As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInve.Open()
                            Dim adaInve As SqlDataReader = New SqlCommand("SELECT CUENT_CONT, CTRL_ALM, DESCR FROM " & INVE & " WHERE CVE_ART ='" & CVE_ART & "'", ConInve).ExecuteReader
                            adaInve.Read()
                            CUENTA_COI = adaInve.Item(0).ToString
                            CTRL = adaInve.Item(1).ToString
                            MIDESCR = adaInve.Item(2).ToString
                            CUENTA_COI = Replace(CUENTA_COI, "-", "")
                            ConInve.Close()
                        End Using

                        Using ConCliente As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConCliente.Open()
                            Dim adaCliente As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE FROM " & CLIE & " WHERE CLAVE ='" & CVE_CLPV & "'", ConCliente).ExecuteReader
                            adaCliente.Read()
                            CUENTA_CONTABLE = adaCliente.Item(0).ToString
                            CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                            ConCliente.Close()
                        End Using

                        Using ConCuenta2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConCuenta2.Open()
                            Dim adaCuenta2 As SqlDataReader = New SqlCommand("SELECT CAPTURAUUID  FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE '" + CUENTA_CONTABLE + "%')", ConCuenta2).ExecuteReader
                            adaCuenta2.Read()
                            If adaCuenta2.HasRows = True Then
                                uuid = adaCuenta2.Item(0)
                            End If
                            ConCuenta2.Close()
                        End Using

                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9, CAMPLIB10, CAMPLIB11, CAMPLIB12, CAMPLIB13 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVE_ART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            CAMPLIB10 = adaInveClib.Item(2).ToString 'SERIE F, IVA 0
                            CAMPLIB11 = adaInveClib.Item(3).ToString 'SERIE F, IVA 16
                            CAMPLIB12 = adaInveClib.Item(4).ToString 'SERIE FD, IVA 0
                            CAMPLIB13 = adaInveClib.Item(5).ToString 'SERIE FD, IVA 16
                            ConInveClib.Close()
                        End Using

                        cuentacon = "0"
                        ccostos = "0"
                        CAMPLIB10 = Replace(CAMPLIB10, "-", "")
                        CAMPLIB11 = Replace(CAMPLIB11, "-", "")
                        CAMPLIB12 = Replace(CAMPLIB12, "-", "")
                        CAMPLIB13 = Replace(CAMPLIB13, "-", "")
                        'VALIDACIONES DE CUENTAS CONTABLES

                        If SERIE = "F" And IMPU4 = 0 Then
                            NUEVOIVA = CAMPLIB10 & "000000000003"
                        ElseIf SERIE = "F" And IMPU4 = 16 Then
                            NUEVOIVA = CAMPLIB11 & "000000000003"
                        ElseIf SERIE = "FD" And IMPU4 = 0 Then
                            NUEVOIVA = CAMPLIB12 & "000000000003"
                        ElseIf SERIE = "FD" And IMPU4 = 16 Then
                            NUEVOIVA = CAMPLIB13 & "000000000003"
                        End If
                        CANTIDADTOTAL = Math.Round(CANTIDADTOTAL, 2)
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                                Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTIDADTOTAL & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "'AND NUM_PART='" & y + 2 & "' AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_ART & "%' AND NUM_CTA ='" & NUEVOIVA & "'", ConAux)
                                cmdAux.ExecuteNonQuery()
                            End If
                            ConAux.Close()
                        End Using

                        NUM_PAR = 0
                        CVE_DOCNUM_PARCVE_ART = 0
                        CANT = 0
                        PXS = 0
                        PREC = 0
                        COST = 0
                        IMPU4 = 0
                        IMP1APLA = ""
                        IMP2APLA = ""
                        IMP3APLA = ""
                        IMP4APLA = 0
                        TOTIMP4 = 0
                        COMI = 0
                        APAR = 0
                        ACT_INV = ""
                        NUM_ALM = 0
                        POLIT_APLI = 0
                        TIP_CAM = 0
                        UNI_VENTA = ""
                        TIPO_PROD = ""
                        CVE_OBS = 0
                        REG_SERIE = 0
                        E_LTPD = 0
                        TIPO_ELEM = ""
                        NUM_MOV = 0
                        TOT_PARTIDAIMPRIMIR = 0
                        DESC1 = 0
                        DESC2 = 0
                        CVE_ART = ""
                        CANTIDADTOTAL = 0
                        y += 1
                    Loop
                    ConItemsSearch.Close()
                End Using



                'START CICLO H
                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTF & " WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
                    Dim dtItemsSearch As New DataTable
                    adaItemsSearch.Fill(dtItemsSearch)
                    ConItemsSearch.Open()
                    Dim w As Integer = (dtItemsSearch.Rows.Count - 1)
                    Dim z As Integer = 0
                    Do While (z <= w)
                        NUM_PAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(0)
                        CVE_ART = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(1).ToString
                        CANTIDAD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(2)
                        PXS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(3)
                        PREC = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(4)
                        PRECIO = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(5)
                        IMPU4 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(6)
                        IMP1APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(7).ToString
                        IMP2APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(8).ToString
                        IMP3APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(9).ToString
                        DESC1 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(10)
                        DESC2 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(11)
                        APAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(12)
                        ACT_INV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(13).ToString
                        NUM_ALM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(14)
                        TIP_CAM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(15)
                        UNI_VENTA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(16).ToString
                        TIPO_PROD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(17)
                        CVE_OBS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(18)
                        REG_SERIE = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(19)
                        E_LTPD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(20)
                        TIPO_ELEM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(21).ToString
                        NUM_MOV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(22)
                        TOT_PARTIDA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(23)
                        IMPRIMIR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(24).ToString


                        Using ConInve As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInve.Open()
                            Dim adaInve As SqlDataReader = New SqlCommand("SELECT CUENT_CONT, CTRL_ALM, DESCR FROM " & INVE & " WHERE CVE_ART ='" & CVE_ART & "'", ConInve).ExecuteReader
                            adaInve.Read()
                            CUENTA_COI = adaInve.Item(0).ToString
                            CTRL = adaInve.Item(1).ToString
                            MIDESCR = adaInve.Item(2).ToString
                            CUENTA_COI = Replace(CUENTA_COI, "-", "")
                            ConInve.Close()
                        End Using

                        Using ConCliente As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConCliente.Open()
                            Dim adaCliente As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE FROM " & CLIE & " WHERE CLAVE ='" & CVE_CLPV & "'", ConCliente).ExecuteReader
                            adaCliente.Read()
                            CUENTA_CONTABLE = adaCliente.Item(0).ToString
                            CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                            ConCliente.Close()
                        End Using

                        Using ConCuenta2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConCuenta2.Open()
                            Dim adaCuenta2 As SqlDataReader = New SqlCommand("SELECT CAPTURAUUID  FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE '" + CUENTA_CONTABLE + "%')", ConCuenta2).ExecuteReader
                            adaCuenta2.Read()
                            If adaCuenta2.HasRows = True Then
                                uuid = adaCuenta2.Item(0)
                            End If
                            ConCuenta2.Close()
                        End Using

                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVE_ART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using

                        Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            INSERT_TOTPART.Open()
                            Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUMDEPTO)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ LIKE '%" & cut & "%' AND NUMDEPTO <>0 AND CONCEP_PO LIKE '%" & CVE_DOC & "%'  AND PERIODO='" & PERIODO & "'", INSERT_TOTPART).ExecuteReader
                            adapter50.Read()
                            H = adapter50.Item(0) + 2
                            INSERT_TOTPART.Close()
                        End Using
                        PRECIO = Math.Round(PRECIO, 2)
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            CAMPLIB8 = Replace(CAMPLIB8, "-", "")
                            If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                                Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTIDAD * PRECIO & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "' AND NUM_PART='" & H & "' AND CONCEP_PO LIKE '%" & CVE_ART & "%' AND NUM_CTA ='" & CAMPLIB8 + "000000000003" & "' ", ConAux)
                                cmdAux.ExecuteNonQuery()
                            End If
                            ConAux.Close()
                        End Using

                        NUM_PAR = 0
                        CVE_DOCNUM_PARCVE_ART = 0
                        CANT = 0
                        PXS = 0
                        PREC = 0
                        COST = 0
                        IMPU4 = 0
                        IMP1APLA = ""
                        IMP2APLA = ""
                        IMP3APLA = ""
                        IMP4APLA = 0
                        TOTIMP4 = 0
                        COMI = 0
                        APAR = 0
                        ACT_INV = ""
                        NUM_ALM = 0
                        POLIT_APLI = 0
                        TIP_CAM = 0
                        UNI_VENTA = ""
                        TIPO_PROD = ""
                        CVE_OBS = 0
                        REG_SERIE = 0
                        E_LTPD = 0
                        TIPO_ELEM = ""
                        NUM_MOV = 0
                        TOT_PARTIDAIMPRIMIR = 0
                        DESC1 = 0
                        DESC2 = 0
                        CVE_ART = ""
                        CANTIDAD = 0
                        PRECIO = 0
                        z += 1
                        H += 1
                    Loop
                    ConItemsSearch.Close()
                    H = 0
                End Using
                'END CICLO D

                'START CLICLO H

                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTF & " WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
                    Dim dtItemsSearch As New DataTable
                    adaItemsSearch.Fill(dtItemsSearch)
                    ConItemsSearch.Open()
                    Dim w As Integer = (dtItemsSearch.Rows.Count - 1)
                    Dim z As Integer = 0
                    Do While (z <= w)
                        NUM_PAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(0)
                        CVE_ART = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(1).ToString
                        CANTIDAD2 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(2)
                        PXS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(3)
                        PREC = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(4)
                        PRECIO2 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(5)
                        IMPU4 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(6)
                        IMP1APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(7).ToString
                        IMP2APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(8).ToString
                        IMP3APLA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(9).ToString
                        DESC1 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(10)
                        DESC2 = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(11)
                        APAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(12)
                        ACT_INV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(13).ToString
                        NUM_ALM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(14)
                        TIP_CAM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(15)
                        UNI_VENTA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(16).ToString
                        TIPO_PROD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(17)
                        CVE_OBS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(18)
                        REG_SERIE = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(19)
                        E_LTPD = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(20)
                        TIPO_ELEM = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(21).ToString
                        NUM_MOV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(22)
                        TOT_PARTIDA = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(23)
                        IMPRIMIR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(24).ToString


                        Using ConInve As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInve.Open()
                            Dim adaInve As SqlDataReader = New SqlCommand("SELECT CUENT_CONT, CTRL_ALM, DESCR FROM " & INVE & " WHERE CVE_ART ='" & CVE_ART & "'", ConInve).ExecuteReader
                            adaInve.Read()
                            CUENTA_COI = adaInve.Item(0).ToString
                            CTRL = adaInve.Item(1).ToString
                            MIDESCR = adaInve.Item(2).ToString
                            CUENTA_COI = Replace(CUENTA_COI, "-", "")
                            ConInve.Close()
                        End Using

                        Using ConCliente As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConCliente.Open()
                            Dim adaCliente As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE FROM " & CLIE & " WHERE CLAVE ='" & CVE_CLPV & "'", ConCliente).ExecuteReader
                            adaCliente.Read()
                            CUENTA_CONTABLE = adaCliente.Item(0).ToString
                            CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                            ConCliente.Close()
                        End Using

                        Using ConCuenta2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConCuenta2.Open()
                            Dim adaCuenta2 As SqlDataReader = New SqlCommand("SELECT CAPTURAUUID  FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE '" + CUENTA_CONTABLE + "%')", ConCuenta2).ExecuteReader
                            adaCuenta2.Read()
                            If adaCuenta2.HasRows = True Then
                                uuid = adaCuenta2.Item(0)
                            End If
                            ConCuenta2.Close()
                        End Using

                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVE_ART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using
                        Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            INSERT_TOTPART.Open()
                            Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUMDEPTO)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ LIKE '%" & cut & "%' AND NUMDEPTO <>0 AND CONCEP_PO LIKE '%" & CVE_DOC & "%'  AND PERIODO='" & PERIODO & "'", INSERT_TOTPART).ExecuteReader
                            adapter50.Read()
                            H = adapter50.Item(0) + 2
                            INSERT_TOTPART.Close()
                        End Using
                        PRECIO2 = Math.Round(PRECIO2, 2)
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            CAMPLIB9 = Replace(CAMPLIB9, "-", "")
                            If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                                Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTIDAD2 * PRECIO2 & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "' AND NUM_PART='" & H & "' AND CONCEP_PO LIKE '%" & CVE_ART & "%' AND NUM_CTA ='" & CAMPLIB9 + "000000000002" & "'", ConAux) 'LIN_PROD2                           
                                cmdAux.ExecuteNonQuery()
                            End If
                            ConAux.Close()
                        End Using

                        NUM_PAR = 0
                        CVE_DOCNUM_PARCVE_ART = 0
                        CANT = 0
                        PXS = 0
                        PREC = 0
                        COST = 0
                        IMPU4 = 0
                        IMP1APLA = ""
                        IMP2APLA = ""
                        IMP3APLA = ""
                        IMP4APLA = 0
                        TOTIMP4 = 0
                        COMI = 0
                        APAR = 0
                        ACT_INV = ""
                        NUM_ALM = 0
                        POLIT_APLI = 0
                        TIP_CAM = 0
                        UNI_VENTA = ""
                        TIPO_PROD = ""
                        CVE_OBS = 0
                        REG_SERIE = 0
                        E_LTPD = 0
                        TIPO_ELEM = ""
                        NUM_MOV = 0
                        TOT_PARTIDAIMPRIMIR = 0
                        DESC1 = 0
                        DESC2 = 0
                        CVE_ART = ""
                        CANTIDAD2 = 0
                        PRECIO2 = 0
                        z += 1

                    Loop
                    ConItemsSearch.Close()
                    H = 0
                End Using
                'END CICLO H

                Using ConClie2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConClie2.Open()
                    Dim adaCliente As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE, NOMBRE FROM " & CLIE & " WHERE CLAVE ='" & CVE_CLPV & "'", ConClie2).ExecuteReader
                    adaCliente.Read()
                    CUENTA_CONTABLE = adaCliente.Item(0).ToString
                    NOMBRE_CLIENTE = adaCliente.Item(1).ToString
                    CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                    ConClie2.Close()
                End Using

                Using ConCuen3 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConCuen3.Open()
                    Dim adaCuen3 As SqlDataReader = New SqlCommand("SELECT NUM_CTA, NOMBRE FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE '" + CUENTA_CONTABLE + "%')", ConCuen3).ExecuteReader
                    adaCuen3.Read()
                    If adaCuen3.HasRows = True Then
                        CUENTA_CONTABLE1 = adaCuen3.Item(0).ToString
                        CONCEPTO = adaCuen3.Item(1).ToString
                    Else
                        CONCEPTO = ""
                        CUENTA_CONTABLE1 = "000000000003"
                    End If
                    ConCuen3.Close()
                End Using
                DEF1 = "250001000000000000002"
                DEF2 = "270002000000000000002"
                '****************************ENCABEZADO PARTIDA NO. 1
                IMPORTE = Math.Round(IMPORTE, 2)
                Using ConAux2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConAux2.Open()
                    If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        Dim cmdAux2 As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & IMPORTE & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & "FACTURA " & CVE_DOC & "    " & CONCEPTO & " " & "%' AND NUM_CTA ='" & CUENTA_CONTABLE + "000000000003" & "'", ConAux2)
                        cmdAux2.ExecuteNonQuery()
                    End If
                    ConAux2.Close()
                End Using

                '**************************IVA
                IMP_TOT4 = Math.Round(IMP_TOT4, 2)
                Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        If IMP_TOT4 > 0 Then
                            Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                INSERT_TOTPART.Open()
                                Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUMDEPTO)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ LIKE '%" & cut & "%' AND NUMDEPTO <>0 AND CONCEP_PO LIKE '%" & CVE_DOC & "%'  AND PERIODO='" & PERIODO & "'", INSERT_TOTPART).ExecuteReader
                                adapter50.Read()
                                H = adapter50.Item(0) + 1
                                INSERT_TOTPART.Close()
                            End Using
                            INSERT_PARTIDAS.Open()
                            Dim commandD As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & IMP_TOT4 & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND NUM_PART='" & H & "' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "'  AND CONCEP_PO LIKE '%" & "FACTURA " & CVE_DOC & "%' AND NUM_CTA ='" & DEF1 & "'", INSERT_PARTIDAS) 'DEF1                           
                            commandD.ExecuteNonQuery()
                            INSERT_PARTIDAS.Close()
                        End If
                    End If
                End Using
                IMP_TOT1 = Math.Round(IMP_TOT1, 2)
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADO = 0 Then
                        If IMP_TOT1 > 0 Then
                            Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                INSERT_TOTPART.Open()
                                Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUMDEPTO)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ LIKE '%" & cut & "%' AND NUMDEPTO <>0 AND CONCEP_PO LIKE '%" & CVE_DOC & "%'  AND PERIODO='" & PERIODO & "'", INSERT_TOTPART).ExecuteReader
                                adapter50.Read()
                                H = adapter50.Item(0) + 1
                                INSERT_TOTPART.Close()
                            End Using
                            INSERT_PARTIDASs.Open()
                            Dim commandDs As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & IMP_TOT1 & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND NUM_PART='" & H & "' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & "FACTURA " & CVE_DOC & " & %' AND NUM_CTA ='" & DEF2 & "'", INSERT_PARTIDASs) 'DEF2                       
                            commandDs.ExecuteNonQuery()
                            INSERT_PARTIDASs.Close()
                        End If
                    End If
                End Using



                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01 SET CAMPLIB10=NULL WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using


                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM " & PAR_FACTF & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        UNO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) - (TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) * (DESC3 / 100)) FROM " & PAR_FACTF & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        TRES = adapter2231.Item(0)
                        DESCCERO = UNO - TRES 'DESCUENTOS
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try


                DESCCERO = Math.Round(DESCCERO, 2)
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        If DESCCERO > 0 Then
                            Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                INSERT_TOTPART.Open()
                                Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUMDEPTO)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ LIKE '%" & cut & "%' AND NUMDEPTO <>0 AND CONCEP_PO LIKE '%" & CVE_DOC & "%'  AND PERIODO='" & PERIODO & "'", INSERT_TOTPART).ExecuteReader
                                adapter50.Read()
                                H = adapter50.Item(0) + 1
                                INSERT_TOTPART.Close()
                            End Using
                            INSERT_PARTIDASs.Open()
                            Dim commandDs As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & DESCCERO & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND NUM_PART='" & H & "' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & "DESCUENTO " & CVE_DOC & " & %' AND NUM_CTA ='410002000000000000002'", INSERT_PARTIDASs)
                            commandDs.ExecuteNonQuery()
                            INSERT_PARTIDASs.Close()
                        End If
                    End If
                End Using


                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM " & PAR_FACTF & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        adapter2231.Read()
                        CINCO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) - (TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) * (DESC3 / 100)) FROM " & PAR_FACTF & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        adapter2231.Read()
                        SEIS = adapter2231.Item(0)
                        DESC16 = CINCO - SEIS 'DESCUENTOS
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try
                DESC16 = Math.Round(DESC16, 2)
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADO = 0 Then '********PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        If DESC16 > 0 Then
                            Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                INSERT_TOTPART.Open()
                                Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUMDEPTO)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ LIKE '%" & cut & "%' AND NUMDEPTO <>0 AND CONCEP_PO LIKE '%" & CVE_DOC & "%'  AND PERIODO='" & PERIODO & "'", INSERT_TOTPART).ExecuteReader
                                adapter50.Read()
                                H = adapter50.Item(0) + 1
                                INSERT_TOTPART.Close()
                            End Using
                            INSERT_PARTIDASs.Open()
                            Dim commandDs As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & DESC16 & "', NUMDEPTO=1 WHERE NUM_POLIZ LIKE '%" & cut & "%' AND NUM_PART='" & H & "' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "'  AND CONCEP_PO LIKE '% " & "DESCUENTO " & CVE_DOC & "%' AND NUM_CTA ='410001000000000000002'", INSERT_PARTIDASs)
                            commandDs.ExecuteNonQuery()
                            INSERT_PARTIDASs.Close()
                        End If
                    End If
                End Using

                Using ConnPoliza As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnPoliza.Open()
                    If CERRADO = 0 Then '**************PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        Dim cmdPol As New SqlCommand("UPDATE " & POLIZA & " SET CONTABILIZ='N' WHERE CONCEP_PO ='" & "FACTURA " & CVE_DOC + "  " & CONCEPTO & "' AND NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO=' " & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "'", ConnPoliza)
                        cmdPol.ExecuteNonQuery()
                    End If
                    ConnPoliza.Close()
                End Using



                Using ConnPoliza As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnPoliza.Open()
                    If CERRADO = 0 Then '**************PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        Dim cmdPol As New SqlCommand("UPDATE " & AUXILIAR & " SET NUMDEPTO=0 WHERE CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO=' " & PERIODO & "' AND TIPO_POLI='" & TIPOFACTF & "'", ConnPoliza)
                        cmdPol.ExecuteNonQuery()
                    End If
                    ConnPoliza.Close()
                End Using

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand("DELETE INSOFTEC_CAMBIOS_VENTAS WHERE CLAVE_DOC='" & CVE_DOC & "'", INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                NUM_PAR = 0

                TIP_DOC = ""
                CVE_DOC = ""
                CVE_CLPV = ""
                STATUS = ""

                CVE_VEND = ""
                CVE_PEDI = ""

                CAN_TOT = 0
                IMP_TOT1 = 0
                IMP_TOT4 = 0
                DES_TOT = 0
                DES_FIN = 0
                COM_TOT = 0

                CVE_OBS = 0
                TIPCAMB = 0
                NUM_PAGOS = 0

                PRIMERPAGO = 0
                RFC = ""
                CTLPOL = 0
                ESCFD = ""
                AUTORIZA = 0
                SERIE = ""
                FOLIO = 0
                AUTOANIO = ""
                DAT_ENVIO = 0
                CONTADO = ""
                CVE_BITA = 0
                BLOQ = ""

                IMPORTE = 0
                COM_TOT_PORC = 0
                TIP_DOC_ANT = ""
                DOC_ANT = ""
                uuid = 0
                uuid3 = ""
                NUMREG = 0
                SUMADESCUENTO16 = 0
                SUMADESCUENTOCERO = 0
                DESC16 = 0
                DESCCERO = 0
                MIPAR1 = 0
                MIPAR2 = 0
                MIPAR3 = 0
                MIPAR4 = 0
                CANTIDAD = 0
                CANTIDAD2 = 0
                CANTIDADTOTAL = 0
                PRECIO = 0
                PRECIO2 = 0
                CAMPOPOLIZA = ""
                CERRADO = ""

                j += 1
            Loop
            ConFact.Close()
        End Using
    End Sub

End Class
