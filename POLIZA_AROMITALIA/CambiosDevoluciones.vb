Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Imports POLIZA_AROMITALIA.ProcesarNewFactura

Public Class CambiosDevoluciones
    Public Shared CAMPOPOLIZADEV As String = ""
    Public Shared CANTIDADTOTAL2DEV As Double = 0
    Public Shared PRECIO3DEV As Double = 0
    Public Shared PRECIO4DEV As Double = 0
    Public Shared CANTIDAD3DEV As Double = 0
    Public Shared CANTIDAD4DEV As Double = 0
    Public Shared Q1DEV As Integer = 0
    Public Shared Q2DEV As Integer = 0
    Public Shared Q3DEV As Integer = 0
    Public Shared Q4DEV As Integer = 0
    Public Shared CERRADODEV As String = ""
    Public Shared Sub ActualizarCambiosDevoluciones()
        Using ConFactD As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaFactD As New SqlDataAdapter("SELECT " & FACTD & ".TIP_DOC, " & FACTD & ".CVE_DOC, " & FACTD & ".CVE_CLPV, " & FACTD & ".STATUS, " & FACTD & ".FECHA_DOC, " & FACTD & ".FECHA_ENT, " & FACTD & ".FECHA_VEN, " & FACTD & ".CAN_TOT, " & FACTD & ".IMP_TOT1, " & FACTD & ".IMP_TOT4, " & FACTD & ".CVE_OBS, " & FACTD & ".NUM_ALMA, " & FACTD & ".ACT_CXC, " & FACTD & ".ACT_COI, " & FACTD & ".ENLAZADO, " & FACTD & ".TIP_DOC_E, " & FACTD & ".NUM_MONED, " & FACTD & ".TIPCAMB, " & FACTD & ".NUM_PAGOS, " & FACTD & ".PRIMERPAGO, " & FACTD & ".RFC, " & FACTD & ".CTLPOL, " & FACTD & ".ESCFD, " & FACTD & ".AUTORIZA, " & FACTD & ".SERIE, " & FACTD & ".FOLIO, " & FACTD & ".DAT_ENVIO, " & FACTD & ".CONTADO, " & FACTD & ".CVE_BITA, " & FACTD & ".BLOQ, " & FACTD & ".FORMAENVIO, " & FACTD & ".DES_FIN_PORC, " & FACTD & ".DES_TOT_PORC, " & FACTD & ".IMPORTE, " & FACTD & ".COM_TOT_PORC, " & FACTD & ".TIP_DOC_ANT," & FACTD & ".DOC_ANT, INSOFTEC_FACTD_CLIB01.CAMPLIB1 FROM  " & FACTD & " INNER JOIN INSOFTEC_FACTD_CLIB01 ON " & FACTD & ".CVE_DOC = INSOFTEC_FACTD_CLIB01.CLAVE_DOC WHERE (INSOFTEC_FACTD_CLIB01.CAMPLIB10=0 AND INSOFTEC_FACTD_CLIB01.CAMPLIB1<>'NULL') AND (" & FACTD & ".STATUS <> 'C')", ConFactD)
            Dim dtFactD As New DataTable
            adaFactD.Fill(dtFactD)
            ConFactD.Open()
            Dim i As Integer = (dtFactD.Rows.Count - 1)
            Dim j As Integer = 0
            Do While (j <= i)
                TIP_DOC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(0).ToString
                CVE_DOC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(1).ToString
                CVE_CLPV = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(2).ToString
                STATUS = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(3).ToString
                FECHA_DOC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(4)
                FECHA_ENT = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(5)
                FECHA_VEN = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(6)
                CAN_TOT = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(7)
                IMP_TOT1 = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(8)
                IMP_TOT4 = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(9)
                CVE_OBS = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(10)
                NUM_ALMA = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(11)
                ACT_CXC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(12).ToString
                ACT_COI = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(13).ToString
                ENLAZADO = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(14).ToString
                TIP_DOC_E = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(15).ToString
                NUM_MONED = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(16)
                TIPCAMB = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(17)
                NUM_PAGOS = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(18)
                PRIMERPAGO = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(19)
                RFC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(20).ToString
                CTLPOL = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(21)
                ESCFD = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(22).ToString
                AUTORIZA = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(23)
                SERIE = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(24).ToString
                FOLIO = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(25)
                DAT_ENVIO = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(26)
                CONTADO = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(27).ToString
                CVE_BITA = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(28)
                BLOQ = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(29).ToString
                FORMAENVIO = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(30).ToString
                DES_FIN_PORC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(31)
                DES_TOT_PORC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(32)
                IMPORTE = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(33)
                COM_TOT_PORC = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(34)
                TIP_DOC_ANT = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(35).ToString
                DOC_ANT = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(36).ToString
                CAMPOPOLIZADEV = dtFactD.AsEnumerable().ElementAtOrDefault(j).Item(37).ToString
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
                Dim cut As String = CAMPOPOLIZADEV.Substring(2)

                Using ConnPeriod As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnPeriod.Open()
                    Dim AdaPeriod As SqlDataReader = New SqlCommand("SELECT * FROM ADMPER WHERE PERIODO='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "'", ConnPeriod).ExecuteReader
                    AdaPeriod.Read()
                    CERRADODEV = AdaPeriod.Item(2)
                    ConnPeriod.Close()
                End Using

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

                Using ConItems As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConItems.Open()
                    Dim adaItems As SqlDataReader = New SqlCommand("SELECT COUNT(CVE_DOC) FROM " & PAR_FACTD & " WHERE (CVE_DOC='" & CVE_DOC & "')", ConItems).ExecuteReader
                    adaItems.Read()
                    totpartida = adaItems.Item(0)
                    ConItems.Close()
                End Using

                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTD & " WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
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
                        CANTIDADTOTAL2DEV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(y).Item(23)
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

                        If CTRL = "420001000000000000002" And IMPU4 = 16 Then
                        Else
                            CTRL = "420002000000000000002"
                        End If
                        If contador = 0 Then
                            contador = 1
                        End If
                        CANTIDADTOTAL2DEV = Math.Round(CANTIDADTOTAL2DEV, 2)
                        '***************************************PARTIDA NO. 1
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                                Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTIDADTOTAL2DEV & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "  " & CVE_ART & "  " & MIDESCR & "%' AND NUM_CTA ='" & CTRL & "'", ConAux)
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
                        CANTIDADTOTAL2DEV = 0
                        y += 1
                    Loop
                    ConItemsSearch.Close()
                End Using


                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTD & " WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
                    Dim dtItemsSearch As New DataTable
                    adaItemsSearch.Fill(dtItemsSearch)
                    ConItemsSearch.Open()
                    Dim w As Integer = (dtItemsSearch.Rows.Count - 1)
                    Dim z As Integer = 0
                    Do While (z <= w)
                        NUM_PAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(0)
                        CVE_ART = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(1).ToString
                        CANTIDAD3DEV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(2)
                        PXS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(3)
                        PREC = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(4)
                        PRECIO3DEV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(5)
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

                        Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            INSERT_TOTPART.Open()
                            Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTD & "' AND  NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                            adapter50.Read()
                            H = adapter50.Item(0)
                            INSERT_TOTPART.Close()
                        End Using
                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVE_ART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using
                        '*********************************PARTIDA NO 2
                        PRECIO3DEV = Math.Round(PRECIO3DEV, 2)
                        If PRECIO3DEV <> 0 Then
                            Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                ConAux.Open()
                                CAMPLIB8 = Replace(CAMPLIB8, "-", "")
                                If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                                    Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTIDAD3DEV * PRECIO3DEV & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "Costo de Ventas" & " " & CVE_DOC & "  " & CVE_ART & " " & MIDESCR & "%' AND NUM_CTA ='" & CAMPLIB8 & "000000000003" & "'", ConAux)
                                    cmdAux.ExecuteNonQuery()
                                End If
                                ConAux.Close()
                            End Using
                        End If

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
                        CANTIDAD3DEV = 0
                        PRECIO3DEV = 0
                        z += 1
                        H += 1
                    Loop
                    ConItemsSearch.Close()
                End Using

                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTD & " WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
                    Dim dtItemsSearch As New DataTable
                    adaItemsSearch.Fill(dtItemsSearch)
                    ConItemsSearch.Open()
                    Dim w As Integer = (dtItemsSearch.Rows.Count - 1)
                    Dim z As Integer = 0
                    Do While (z <= w)
                        NUM_PAR = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(0)
                        CVE_ART = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(1).ToString
                        CANTIDAD4DEV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(2)
                        PXS = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(3)
                        PREC = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(4)
                        PRECIO4DEV = dtItemsSearch.AsEnumerable().ElementAtOrDefault(z).Item(5)
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

                        Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            INSERT_TOTPART.Open()
                            Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTD & "' AND  NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                            adapter50.Read()
                            H = adapter50.Item(0)
                            INSERT_TOTPART.Close()
                        End Using

                        Using ConInveClib As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            ConInveClib.Open()
                            Dim adaInveClib As SqlDataReader = New SqlCommand("SELECT CAMPLIB8, CAMPLIB9 FROM INVE_CLIB01 WHERE CVE_PROD = '" & CVE_ART & "'", ConInveClib).ExecuteReader
                            adaInveClib.Read()
                            CAMPLIB8 = adaInveClib.Item(0).ToString
                            CAMPLIB9 = adaInveClib.Item(1).ToString
                            ConInveClib.Close()
                        End Using
                        '*********************************PARTIDA NO 3
                        PRECIO4DEV = Math.Round(PRECIO4DEV, 2)
                        If PRECIO4DEV <> 0 Then
                            Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                CAMPLIB9 = Replace(CAMPLIB9, "-", "")
                                ConAux.Open()
                                If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                                    Dim cmdAux As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & CANTIDAD4DEV * PRECIO4DEV & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "Costo de Ventas" & " " + CVE_DOC & "  " & CVE_ART & " " & MIDESCR & "%' AND NUM_CTA ='" & CAMPLIB9 & "000000000002" & "'", ConAux)
                                    cmdAux.ExecuteNonQuery()
                                End If
                                ConAux.Close()
                            End Using
                        End If

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
                        CANTIDAD4DEV = 0
                        PRECIO4DEV = 0
                        z += 1
                    Loop
                    ConItemsSearch.Close()
                End Using

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

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT COUNT (NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    Q1DEV = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using
                '*****************************PARTIDA NO 4
                IMPORTE = Math.Round(IMPORTE, 2)
                Using ConAux2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConAux2.Open()
                    CONCEPTO = Replace(CONCEPTO, "'", "")
                    If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                        Dim cmdAux2 As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & IMPORTE & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "DEVOLUCIÓN " & CVE_DOC & "    " & CONCEPTO & "%' AND NUM_CTA ='" & CUENTA_CONTABLE + "000000000003" & "'", ConAux2)
                        cmdAux2.ExecuteNonQuery()
                    End If
                    ConAux2.Close()
                End Using
                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT COUNT (NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    Q2DEV = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using
                IMP_TOT4 = Math.Round(IMP_TOT4, 2)
                Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                        If IMP_TOT4 > 0 Then
                            INSERT_PARTIDAS.Open()
                            CONCEPTO3 = Replace(CONCEPTO3, "'", "")
                            Dim commandD As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & IMP_TOT4 & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "DEVOLUCIÓN " & CVE_DOC & "%' AND NUM_CTA ='" & DEF1 & "'", INSERT_PARTIDAS) 'DEF1                    
                            commandD.ExecuteNonQuery()
                        End If
                    End If
                    INSERT_PARTIDAS.Close()
                End Using
                IMP_TOT1 = Math.Round(IMP_TOT1, 2)
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                        If IMP_TOT1 > 0 Then
                            INSERT_PARTIDASs.Open()
                            CONCEPTO = Replace(CONCEPTO, "'", "")
                            Dim commandDs As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & IMP_TOT1 & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "DEVOLUCIÓN " & CVE_DOC & "%' AND NUM_CTA ='" & DEF2 & "'", INSERT_PARTIDASs)
                            commandDs.ExecuteNonQuery()
                        End If
                    End If
                    INSERT_PARTIDASs.Close()
                End Using





                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand("UPDATE INSOFTEC_FACTD_CLIB01 SET CAMPLIB10=NULL WHERE CLAVE_DOC='" & CVE_DOC & "'", INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using


                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM " & PAR_FACTD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        UNO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) FROM " & PAR_FACTD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        TRES = adapter2231.Item(0)
                        DESCCERO = UNO - TRES 'DESCUENTOS
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try
                DESCCERO = Math.Round(DESCCERO, 2)
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                        If DESCCERO > 0 Then
                            INSERT_PARTIDASs.Open()
                            Dim commandDs As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & DESCCERO & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "DESCUENTO " & CVE_DOC & "%' AND NUM_CTA ='410002000000000000002'", INSERT_PARTIDASs)
                            commandDs.ExecuteNonQuery()
                            INSERT_PARTIDASs.Close()
                        End If
                    End If
                End Using



                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM " & PAR_FACTD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        adapter2231.Read()
                        CINCO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) FROM " & PAR_FACTD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        adapter2231.Read()
                        SEIS = adapter2231.Item(0)
                        DESC16 = CINCO - SEIS 'DESCUENTOS
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try
                DESC16 = Math.Round(DESC16, 2)
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If CERRADODEV = 0 Then '**************************PARA VALIDAR EL PERIODO QUE ESTE ABIERTO
                        If DESC16 > 0 Then
                            INSERT_PARTIDASs.Open()
                            Dim commandDs As New SqlCommand("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & DESC16 & "' WHERE NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO='" & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "' AND CONCEP_PO LIKE '%" & "DESCUENTO " & CVE_DOC & "%' AND NUM_CTA ='410001000000000000002'", INSERT_PARTIDASs)
                            commandDs.ExecuteNonQuery()
                            INSERT_PARTIDASs.Close()
                        End If
                    End If
                End Using

                Using ConnPoliza As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConnPoliza.Open()
                    If CERRADODEV = 0 Then '**************PARA VALIDAR SI ESTA CERRADO EL PERIODO
                        Dim cmdPol As New SqlCommand("UPDATE " & POLIZA & " SET CONTABILIZ='N' WHERE CONCEP_PO ='" & "DEVOLUCION " & CVE_DOC + "  " & CONCEPTO & "' AND NUM_POLIZ LIKE '%" & cut & "%' AND PERIODO=' " & PERIODO & "' AND TIPO_POLI='" & TIPOFACTD & "'", ConnPoliza)
                        cmdPol.ExecuteNonQuery()
                    End If
                    ConnPoliza.Close()
                End Using

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand("DELETE INSOFTEC_CAMBIOS_DEV WHERE CLAVE_DOC='" & CVE_DOC & "'", INSERT_CAMP)
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
                uuid4 = ""
                NUMREG = 0
                'FECH1 = ""
                SUMADESCUENTO16 = 0
                SUMADESCUENTOCERO = 0
                DESC16 = 0
                DESCCERO = 0
                contador = 0
                CANTIDADTOTAL2DEV = 0
                PRECIO3DEV = 0
                PRECIO4DEV = 0
                CANTIDAD3DEV = 0
                CANTIDAD4DEV = 0
                CAMPOPOLIZADEV = ""
                CERRADODEV = ""
                j += 1
            Loop
            ConFactD.Close()
        End Using
    End Sub

End Class
