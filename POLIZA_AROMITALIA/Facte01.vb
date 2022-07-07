Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.CorreoSinCasilla
Imports POLIZA_AROMITALIA.ConsultaTablaDatos

Public Class Facte01
    Public Shared CANTIDADTOTAL As Double = 0
    Public Shared PRECIO As Double = 0
    Public Shared CANTIDAD As Double = 0
    Public Shared PRECIO2 As Double = 0
    Public Shared CANTIDAD2 As Double = 0
    Public Shared NIVEL1 As String = ""
    Public Shared NIVEL2 As String = ""
    Public Shared NIVEL3 As String = ""
    Public Shared NIVEL11 As String = ""
    Public Shared NIVEL22 As String = ""
    Public Shared NIVEL33 As String = ""
    Public Shared NIVEL111 As String = ""
    Public Shared NIVEL222 As String = ""
    Public Shared NIVEL333 As String = ""
    Public Shared NIVEL1111 As String = ""
    Public Shared NIVEL2222 As String = ""
    Public Shared NIVEL3333 As String = ""
    Public Shared CADENA10 As String = ""
    Public Shared CADENA11 As String = ""
    Public Shared CADENA12 As String = ""
    Public Shared CADENA13 As String = ""
    Public Shared CADENA14 As String = ""
    Public Shared NIVEL18 As String = ""
    Public Shared NIVEL28 As String = ""
    Public Shared NIVEL38 As String = ""
    Public Shared NIVEL19 As String = ""
    Public Shared NIVEL29 As String = ""
    Public Shared NIVEL39 As String = ""
    Public Shared CADENA8 As String = ""
    Public Shared CADENA9 As String = ""
    Public Shared Sub ProcesarFacte01()
        Dim TOTAL_UNO As Double = 0.0
        Dim RESTA As Double = 0.0
        Dim MIMONTO As Double = 0.0
        Using ConFact As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adaFact As New SqlDataAdapter("SELECT FACTE01.TIP_DOC, FACTE01.CVE_DOC, FACTE01.CVE_CLPV, FACTE01.STATUS, FACTE01.FECHA_DOC, FACTE01.FECHA_ENT, FACTE01.FECHA_VEN, FACTE01.CAN_TOT, FACTE01.IMP_TOT1, FACTE01.IMP_TOT4, FACTE01.CVE_OBS, FACTE01.NUM_ALMA, FACTE01.ACT_CXC, FACTE01.ACT_COI, FACTE01.ENLAZADO, FACTE01.TIP_DOC_E, FACTE01.NUM_MONED, FACTE01.TIPCAMB, FACTE01.NUM_PAGOS, FACTE01.PRIMERPAGO, FACTE01.RFC, FACTE01.CTLPOL, FACTE01.ESCFD, FACTE01.AUTORIZA, FACTE01.SERIE, FACTE01.FOLIO, FACTE01.DAT_ENVIO, FACTE01.CONTADO, FACTE01.CVE_BITA, FACTE01.BLOQ, FACTE01.FORMAENVIO, FACTE01.DES_FIN_PORC, FACTE01.DES_TOT_PORC, FACTE01.IMPORTE, FACTE01.COM_TOT_PORC, FACTE01.TIP_DOC_ANT,FACTE01.DOC_ANT FROM  FACTE01 INNER JOIN INSOFTEC_FACTE_CLIB01 ON FACTE01.CVE_DOC = INSOFTEC_FACTE_CLIB01.CLAVE_DOC WHERE (INSOFTEC_FACTE_CLIB01.CAMPLIB1 IS NULL AND INSOFTEC_FACTE_CLIB01.CAMPLIB4 IS NULL) AND (FACTE01.STATUS <> 'C')", ConFact)
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
                Dim FECH As String = FECHA_DOC.ToString("yyyy-dd-MM")

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


                Using ConFol As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConFol.Open()
                    Dim cmdFol As New SqlCommand(("UPDATE " & FOLIOS & " SET FOLIO" & PERIODO1 & " ='" & contador & "' WHERE EJERCICIO ='" & EJERCICIO & "' AND TIPPOL='" & TIPOFACTF & "'"), ConFol)
                    cmdFol.ExecuteNonQuery()
                    ConFol.Close()
                End Using


                Using ConItems As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    ConItems.Open()
                    Dim adaItems As SqlDataReader = New SqlCommand("SELECT COUNT(CVE_DOC) FROM PAR_FACTE01 WHERE (CVE_DOC='" & CVE_DOC & "')", ConItems).ExecuteReader
                    adaItems.Read()
                    totpartida = adaItems.Item(0)
                    ConItems.Close()
                End Using

                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM PAR_FACTE01 WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
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

                        If CAMPLIB10 <> "" Then
                            NIVEL1 = CAMPLIB10.Substring(0, 3)
                            NIVEL1 = Replace(NIVEL1, "0", "")
                            NIVEL2 = CAMPLIB10.Substring(3, 3)
                            NIVEL2 = Replace(NIVEL2, "0", "")
                            NIVEL3 = CAMPLIB10.Substring(6, 3)
                            NIVEL3 = Replace(NIVEL3, "0", "")
                        End If

                        If NIVEL1 <> "" And NIVEL2 = "" And NIVEL3 = "" Then
                            CADENA10 = "000000000001"
                        ElseIf NIVEL2 <> "" And NIVEL3 = "" Then
                            CADENA10 = "000000000002"
                        ElseIf NIVEL3 <> "" Then
                            CADENA10 = "000000000003"
                        End If

                        If CAMPLIB11 <> "" Then
                            NIVEL11 = CAMPLIB11.Substring(0, 3)
                            NIVEL11 = Replace(NIVEL11, "0", "")
                            NIVEL22 = CAMPLIB11.Substring(3, 3)
                            NIVEL22 = Replace(NIVEL22, "0", "")
                            NIVEL33 = CAMPLIB11.Substring(6, 3)
                            NIVEL33 = Replace(NIVEL33, "0", "")
                        End If

                        If NIVEL11 <> "" And NIVEL22 = "" And NIVEL33 = "" Then
                            CADENA11 = "000000000001"
                        ElseIf NIVEL22 <> "" And NIVEL33 = "" Then
                            CADENA11 = "000000000002"
                        ElseIf NIVEL33 <> "" Then
                            CADENA11 = "000000000003"
                        End If

                        If CAMPLIB12 <> "" Then
                            NIVEL111 = CAMPLIB12.Substring(0, 3)
                            NIVEL111 = Replace(NIVEL111, "0", "")
                            NIVEL222 = CAMPLIB12.Substring(3, 3)
                            NIVEL222 = Replace(NIVEL222, "0", "")
                            NIVEL333 = CAMPLIB12.Substring(6, 3)
                            NIVEL333 = Replace(NIVEL333, "0", "")
                        End If

                        If NIVEL111 <> "" And NIVEL222 = "" And NIVEL333 = "" Then
                            CADENA12 = "000000000001"
                        ElseIf NIVEL222 <> "" And NIVEL333 = "" Then
                            CADENA12 = "000000000002"
                        ElseIf NIVEL333 <> "" Then
                            CADENA12 = "000000000003"
                        End If

                        If CAMPLIB13 <> "" Then
                            NIVEL1111 = CAMPLIB13.Substring(0, 3)
                            NIVEL1111 = Replace(NIVEL1111, "0", "")
                            NIVEL2222 = CAMPLIB13.Substring(3, 3)
                            NIVEL2222 = Replace(NIVEL2222, "0", "")
                            NIVEL3333 = CAMPLIB13.Substring(6, 3)
                            NIVEL3333 = Replace(NIVEL3333, "0", "")
                        End If

                        If NIVEL1111 <> "" And NIVEL2222 = "" And NIVEL3333 = "" Then
                            CADENA13 = "000000000001"
                        ElseIf NIVEL2222 <> "" And NIVEL3333 = "" Then
                            CADENA13 = "000000000002"
                        ElseIf NIVEL3333 <> "" Then
                            CADENA13 = "000000000003"
                        End If

                        If SERIE = "F" And IMPU4 = 0 Then
                            NUEVOIVA = CAMPLIB10 & CADENA10
                        ElseIf SERIE = "F" And IMPU4 = 16 Then
                            NUEVOIVA = CAMPLIB11 & CADENA11
                        ElseIf SERIE = "FD" And IMPU4 = 0 Then
                            NUEVOIVA = CAMPLIB12 & CADENA12
                        ElseIf SERIE = "FD" And IMPU4 = 16 Then
                            NUEVOIVA = CAMPLIB13 & CADENA13
                        End If




                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            Dim v = IsNumeric(DEPTO)
                            'If v = True Then
                            Dim cmdAux As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & y + 2 & "','" & PERIODO & "','" & EJERCICIO & "','" & NUEVOIVA & "','" & FECH & "','" + CVE_DOC + "  " + CVE_ART + "','H','" & Math.Round(CANTIDADTOTAL, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & y + 2 & "','0','0');"), ConAux)
                            cmdAux.ExecuteNonQuery()
                            'End If
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
                        NIVEL1 = ""
                        NIVEL2 = ""
                        NIVEL3 = ""
                        NIVEL11 = ""
                        NIVEL22 = ""
                        NIVEL33 = ""
                        NIVEL111 = ""
                        NIVEL222 = ""
                        NIVEL333 = ""
                        NIVEL1111 = ""
                        NIVEL2222 = ""
                        NIVEL3333 = ""
                        CADENA10 = ""
                        CADENA11 = ""
                        CADENA12 = ""
                        CADENA13 = ""
                        CADENA14 = ""
                        y += 1
                    Loop
                    ConItemsSearch.Close()
                End Using



                'START CICLO H
                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM  PAR_FACTE01 WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
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
                            Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                            adapter50.Read()
                            H = adapter50.Item(0)
                            INSERT_TOTPART.Close()
                        End Using


                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            CAMPLIB8 = Replace(CAMPLIB8, "-", "")

                            If CAMPLIB8 <> "" Then
                                NIVEL18 = CAMPLIB8.Substring(0, 3)
                                NIVEL18 = Replace(NIVEL18, "0", "")
                                NIVEL28 = CAMPLIB8.Substring(3, 3)
                                NIVEL28 = Replace(NIVEL28, "0", "")
                                NIVEL38 = CAMPLIB8.Substring(6, 3)
                                NIVEL38 = Replace(NIVEL38, "0", "")
                            End If

                            If NIVEL18 <> "" And NIVEL28 = "" And NIVEL38 = "" Then
                                CADENA8 = "000000000001"
                            ElseIf NIVEL28 <> "" And NIVEL38 = "" Then
                                CADENA8 = "000000000002"
                            ElseIf NIVEL38 <> "" Then
                                CADENA8 = "000000000003"
                            End If

                            Dim v = IsNumeric(DEPTO)
                            'If v = True Then

                            Dim cmdAux As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & H + 2 & "','" & PERIODO & "','" & EJERCICIO & "','" & CAMPLIB8 + CADENA8 & "','" & FECH & "',' " + "Costo Ventas" + " " + CVE_DOC + "  " + CVE_ART + "','H','" & Math.Round(CANTIDAD * PRECIO, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & H + 2 & "','0','0');"), ConAux)
                            cmdAux.ExecuteNonQuery()
                            'End If
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
                        NIVEL1 = ""
                        NIVEL2 = ""
                        NIVEL3 = ""
                        NIVEL11 = ""
                        NIVEL22 = ""
                        NIVEL33 = ""
                        NIVEL111 = ""
                        NIVEL222 = ""
                        NIVEL333 = ""
                        NIVEL1111 = ""
                        NIVEL2222 = ""
                        NIVEL3333 = ""
                        CADENA10 = ""
                        CADENA11 = ""
                        CADENA12 = ""
                        CADENA13 = ""
                        CADENA14 = ""
                        NIVEL18 = ""
                        NIVEL28 = ""
                        NIVEL38 = ""
                        NIVEL19 = ""
                        NIVEL29 = ""
                        NIVEL39 = ""
                        CADENA8 = ""
                        CADENA9 = ""
                        z += 1
                        H += 1
                    Loop
                    ConItemsSearch.Close()
                    H = 0
                End Using
                'END CICLO D

                'START CLICLO H

                Using ConItemsSearch As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim adaItemsSearch As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM PAR_FACTE01 WHERE CVE_DOC='" & CVE_DOC & "'", ConItemsSearch)
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
                            Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                            adapter50.Read()
                            H = adapter50.Item(0)
                            INSERT_TOTPART.Close()
                        End Using
                        Using ConAux As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConAux.Open()
                            CAMPLIB9 = Replace(CAMPLIB9, "-", "")

                            If CAMPLIB9 <> "" Then
                                NIVEL19 = CAMPLIB9.Substring(0, 3)
                                NIVEL19 = Replace(NIVEL19, "0", "")
                                NIVEL29 = CAMPLIB9.Substring(3, 3)
                                NIVEL29 = Replace(NIVEL29, "0", "")
                                NIVEL39 = CAMPLIB9.Substring(6, 3)
                                NIVEL39 = Replace(NIVEL39, "0", "")
                            End If

                            If NIVEL19 <> "" And NIVEL29 = "" And NIVEL39 = "" Then
                                CADENA9 = "000000000001"
                            ElseIf NIVEL29 <> "" And NIVEL39 = "" Then
                                CADENA9 = "000000000002"
                            ElseIf NIVEL39 <> "" Then
                                CADENA9 = "000000000003"
                            End If
                            Dim v = IsNumeric(DEPTO)
                            'If v = True Then

                            Dim cmdAux As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & H + 2 & "','" & PERIODO & "','" & EJERCICIO & "','" & CAMPLIB9 + CADENA9 & "','" & FECH & "',' " + "Costo Ventas" + " " + CVE_DOC + "  " + CVE_ART + "','D','" & Math.Round(CANTIDAD2 * PRECIO2, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & H + 2 & "','0','0');"), ConAux) 'LIN_PROD2                           
                            cmdAux.ExecuteNonQuery()
                            'End If
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
                        NIVEL1 = ""
                        NIVEL2 = ""
                        NIVEL3 = ""
                        NIVEL11 = ""
                        NIVEL22 = ""
                        NIVEL33 = ""
                        NIVEL111 = ""
                        NIVEL222 = ""
                        NIVEL333 = ""
                        NIVEL1111 = ""
                        NIVEL2222 = ""
                        NIVEL3333 = ""
                        CADENA10 = ""
                        CADENA11 = ""
                        CADENA12 = ""
                        CADENA13 = ""
                        CADENA14 = ""
                        NIVEL18 = ""
                        NIVEL28 = ""
                        NIVEL38 = ""
                        NIVEL19 = ""
                        NIVEL29 = ""
                        NIVEL39 = ""
                        CADENA8 = ""
                        CADENA9 = ""
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

                Using ConAux2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ConAux2.Open()
                    Dim v = IsNumeric(DEPTO)
                    CONCEPTO = Replace(CONCEPTO, "'", "")
                    'If v = True Then
                    Dim cmdAux2 As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'1','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_CONTABLE + "000000000003" & "','" & FECH & "','" + "FACTURA " & CVE_DOC + "    " + CONCEPTO + " " & "','D','" & Math.Round(IMPORTE, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','1','" & ccostos & "'," & cuentacon & ");"), ConAux2)
                    cmdAux2.ExecuteNonQuery()
                    'End If
                    ConAux2.Close()
                End Using

                Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If IMP_TOT4 > 0 Then
                        INSERT_PARTIDAS.Open()
                        Dim v = IsNumeric(DEPTO)
                        CONCEPTO3 = Replace(CONCEPTO3, "'", "")
                        If v = True Then
                            Dim commandD As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'555','" & PERIODO & "','" & EJERCICIO & "','" & DEF1 & "','" & FECH & "','" + "FACTURA " & CVE_DOC & "','H','" & Math.Round(IMP_TOT4, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS) 'DEF1
                            'Dim commandD As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K & "','" & PERIODO & "','" & EJERCICIO & "','" & DEF1 & "','" & FECH & "','" + "FACTURA " & CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO3 + "','H','" & IMP_TOT4 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS) 'DEF1
                            commandD.ExecuteNonQuery()

                        Else
                            DEPTO = "0" Or DEPTO = ""
                            Dim commandD As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'555','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "FACTURA" & CVE_DOC & "','H','" & Math.Round(IMP_TOT4, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS)
                            'Dim commandD As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "FACTURA" & CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO3 + "','H','" & IMP_TOT4 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS)
                            commandD.ExecuteNonQuery()
                        End If

                        INSERT_PARTIDAS.Close()
                    End If
                End Using


                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND  NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR3 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR3 & "', ORDEN='" & MIPAR3 & "' WHERE NUM_PART =555 AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If IMP_TOT1 > 0 Then
                        INSERT_PARTIDASs.Open()
                        Dim x = IsNumeric(DEPTO)
                        If x = x Then
                            CONCEPTO = Replace(CONCEPTO, "'", "")

                            Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('    ' + LTRIM(RTRIM(" & contador & ")), 5),'666','" & PERIODO & "','" & EJERCICIO & "','" & DEF2 & "','" & FECH & "','" + "FACTURA " & CVE_DOC & "','H','" & Math.Round(IMP_TOT1, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs) 'DEF1
                            'Dim commandDs As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve','" & contador & "','" & K & "','" & PERIODO & "','" & EJERCICIO & "','" & DEF1 & "','" & FECH & "','" + "FACTURA " & CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO3 + "','H','" & IMP_TOT1 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs) 'DEF1
                            commandDs.ExecuteNonQuery()

                        Else
                            Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & j & ")), 5),'666','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "FACTURA" & CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO3 + "','H','" & Math.Round(IMP_TOT1, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs)
                            commandDs.ExecuteNonQuery()
                        End If

                        INSERT_PARTIDASs.Close()
                    End If
                End Using

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "'  AND NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR4 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR4 & "', ORDEN='" & MIPAR4 & "' WHERE NUM_PART =666 AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Using INSERT_FOLIO As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_FOLIO.Open()
                    Dim command661 As New SqlCommand(("UPDATE " & AUXILIAR & " set NUM_PART=1 WHERE NUM_PART=0 AND TIPO_POLI='" & TIPOFACTF & "'"), INSERT_FOLIO)
                    command661.ExecuteNonQuery()
                    INSERT_FOLIO.Close()
                End Using

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE INSOFTEC_FACTE_CLIB01 SET CAMPLIB1='" & TIPOFACTF & contador & "', CAMPLIB3='PROCESO 1' WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_POLIZ)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "'  AND NUM_POLIZ=" & contador & " AND PERIODO=" & PERIODO & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    totpartida = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using
                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM  PAR_FACTE01  WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        UNO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        'Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) FROM " & PAR_FACTE01 & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) - (TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) * (DESC3 / 100)) FROM  PAR_FACTE01  WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        TRES = adapter2231.Item(0)
                        DESCCERO = UNO - TRES 'DESCUENTOS
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try


                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If DESCCERO > 0 Then
                        INSERT_PARTIDASs.Open()
                        Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'333','" & PERIODO & "','" & EJERCICIO & "','410002000000000000002','" & FECH & "','" + "DESCUENTO " & CVE_DOC + " ','D','" & Math.Round(DESCCERO, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & MIPAR1 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs)
                        commandDs.ExecuteNonQuery()
                        INSERT_PARTIDASs.Close()
                    End If
                End Using

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR1 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using 'conteo1
                '
                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR1 & "', ORDEN='" & MIPAR1 & "' WHERE NUM_PART =333 AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM PAR_FACTE01  WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader

                        adapter2231.Read()
                        CINCO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        ' Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) FROM " & PAR_FACTE01 & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) - (TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) * (DESC3 / 100)) FROM PAR_FACTE01  WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        adapter2231.Read()
                        SEIS = adapter2231.Item(0)
                        DESC16 = CINCO - SEIS 'DESCUENTOS
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try



                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    If DESC16 > 0 Then
                        INSERT_PARTIDASs.Open()
                        Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'444','" & PERIODO & "','" & EJERCICIO & "','410001000000000000002','" & FECH & "','" + "DESCUENTO " & CVE_DOC + " ','D','" & Math.Round(DESC16, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & MIPAR2 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs)
                        commandDs.ExecuteNonQuery()
                        INSERT_PARTIDASs.Close()
                    End If
                End Using


                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR2 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using 'conteo2

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR2 & "', ORDEN='" & MIPAR2 & "' WHERE NUM_PART =444 AND TIPO_POLI='" & TIPOFACTF & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Using INSERT_ENCABEZADO As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_ENCABEZADO.Open()
                    Dim command5 As New SqlCommand(("INSERT INTO " & POLIZA & "  (TIPO_POLI,NUM_POLIZ,PERIODO,EJERCICIO,FECHA_POL,CONCEP_PO,NUM_PART,LOGAUDITA,CONTABILIZ,NUMPARCUA,TIENEDOCUMENTOS,ORIGEN) VALUES ('" & TIPOFACTF & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & PERIODO & "','" & EJERCICIO & "','" & FECH & "','" & "FACTURA " & CVE_DOC + "  " & CONCEPTO & "','" & totpartida & "','N','N','0','0','SAE-FAC-AUT');"), INSERT_ENCABEZADO) 'totpartida
                    command5.ExecuteNonQuery()
                    INSERT_ENCABEZADO.Close()
                End Using


                Try
                    Using connection222540 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection222540.Open()
                        Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT UUID, FECHA_CERT FROM " & CFDI & " WHERE CVE_DOC ='" & CVE_DOC & "'", connection222540).ExecuteReader
                        adapter222540.Read()
                        If adapter222540.HasRows = True Then
                            uuid3 = adapter222540.Item(0)
                            FECHA_CERT_FACTURA = adapter222540.Item(1).ToString
                        End If
                    End Using
                Catch ex As Exception

                End Try

                Using connection222540 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connection222540.Open()
                    Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT MAX(CTUUIDCOMP) FROM " & CONTROL & "", connection222540).ExecuteReader
                    adapter222540.Read()
                    If adapter222540.HasRows = True Then
                        NUMREG = adapter222540.Item(0)
                    End If
                    NUMREG = NUMREG + 1
                End Using
                If uuid = 0 Then
                    CapturaXMLCasilla() 'se manda correo cuando no está habilitada la casilla de captura de comprobantes
                End If

                Dim FECH1 As String = FECHA_CERT_FACTURA.ToString("yyyy-dd-MM HH:mm:ss")


                Using ALFA As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ALFA.Open()
                    'Dim command6661 As New SqlCommand(("INSERT INTO UUIDTIMBRES (NUMREG, UUIDTIMBRE, MONTO, SERIE, FOLIO, RFCEMISOR, RFCRECEPTOR, ORDEN, FECHA, TIPOCOMPROBANTE) VALUES ('" & NUMREG & "','" & uuid3 & "','" & IMPORTE & "','" & SERIE & "','" & FOLIO & "','NME031127CH3','" & RFC & "','1','" & FECH1 & "','1')"), ALFA)
                    Dim command6661 As New SqlCommand(("INSERT INTO UUIDTIMBRES (NUMREG, UUIDTIMBRE, MONTO, SERIE, FOLIO, RFCEMISOR, RFCRECEPTOR, ORDEN, FECHA, TIPOCOMPROBANTE) VALUES ('" & NUMREG & "','" & uuid3 & "','" & IMPORTE & "','" & SERIE & "','" & FOLIO & "','NME031127CH3','" & RFC & "','1',GETDATE(),'1')"), ALFA)
                    command6661.ExecuteNonQuery()
                    ALFA.Close()
                End Using


                Using ALFA As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    ALFA.Open()
                    Dim command6661 As New SqlCommand(("UPDATE " & CONTROL & " SET CTUUIDCOMP = " & NUMREG & ""), ALFA)
                    command6661.ExecuteNonQuery()
                    ALFA.Close()
                End Using



                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & POLIZA & " SET UUID = '" & uuid3 & "', TIENEDOCUMENTOS=1 WHERE CONCEP_PO LIKE '%" & CVE_DOC & "%' AND TIPO_POLI='" & TIPOFACTF & "'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using


                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET ORDEN=NUM_PART  WHERE CONCEP_PO LIKE '%" & CVE_DOC & "%' AND TIPO_POLI='" & TIPOFACTF & "'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using


                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET IDUUID = '" & NUMREG & "' WHERE CONCEP_PO LIKE '%" & CVE_DOC & "%' AND CONCEP_PO LIKE '%" & CONCEPTO & "%' AND DEBE_HABER='D' AND TIPO_POLI='" & TIPOFACTF & "'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using
                '***PARA REVISAR LOS IMPORTES Y CUADRARLOS***
                Using connT1 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connT1.Open()
                    Dim adaT1 As SqlDataReader = New SqlCommand("SELECT SUM(MONTOMOV) FROM " & AUXILIAR & " WHERE DEBE_HABER='H' AND NUM_POLIZ= RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT1).ExecuteReader
                    adaT1.Read()
                    If adaT1.HasRows = True Then
                        TOTAL_UNO = adaT1.Item(0).ToString
                    End If
                    If TOTAL_UNO <> IMPORTE Then
                        If TOTAL_UNO < IMPORTE Then
                            RESTA = Math.Round(IMPORTE - TOTAL_UNO, 2)
                            RESTA = Math.Abs(RESTA)
                        End If
                        If TOTAL_UNO > IMPORTE Then
                            RESTA = Math.Round(IMPORTE - TOTAL_UNO, 2)
                        End If


                        Using connT2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            connT2.Open()
                            Dim adaT2 As SqlDataReader = New SqlCommand("SELECT MONTOMOV FROM " & AUXILIAR & " WHERE NUM_PART=2 AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT2).ExecuteReader
                            adaT2.Read()
                            If adaT2.HasRows = True Then
                                MIMONTO = adaT2.Item(0).ToString
                            End If
                        End Using
                        Using ConnReg As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ConnReg.Open()
                            Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & MIMONTO + RESTA & "'  WHERE NUM_PART=2 AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), ConnReg)
                            command666.ExecuteNonQuery()
                            ConnReg.Close()
                        End Using
                    End If
                End Using

                '***PARA REVISAR LOS IMPORTES Y CUADRARLOS***
                Dim SUMAH As Double = 0.0
                Dim SUMAD As Double = 0.0
                Dim RESTAHD As Double = 0.0
                Using connT1 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connT1.Open()
                    Dim adaT1 As SqlDataReader = New SqlCommand("SELECT SUM(MONTOMOV) FROM " & AUXILIAR & " WHERE DEBE_HABER='H' AND NUM_POLIZ= RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT1).ExecuteReader
                    adaT1.Read()
                    If adaT1.HasRows = True Then
                        SUMAH = adaT1.Item(0).ToString
                    End If
                End Using
                Using connT2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connT2.Open()
                    Dim adaT2 As SqlDataReader = New SqlCommand("SELECT SUM(MONTOMOV) FROM " & AUXILIAR & " WHERE DEBE_HABER='D' AND NUM_POLIZ= RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT2).ExecuteReader
                    adaT2.Read()
                    If adaT2.HasRows = True Then
                        SUMAD = adaT2.Item(0).ToString
                    End If
                End Using

                If SUMAH <> SUMAD Then
                    If SUMAH > SUMAD Then
                        RESTAHD = (SUMAH - SUMAD)
                    End If

                    If SUMAD > SUMAH Then
                        RESTAHD = (SUMAD - SUMAH)
                    End If
                    RESTAHD = Math.Round(RESTAHD, 2)
                    Dim NUEVOMONTO As Double = 0.0

                    Dim ITEMNNUMERO As Integer
                    If SUMAH > SUMAD Then
                        Using connT3 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            connT3.Open()
                            Dim adaT3 As SqlDataReader = New SqlCommand("SELECT TOP(1) MONTOMOV, NUM_PART FROM " & AUXILIAR & " WHERE DEBE_HABER='D' AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT3).ExecuteReader
                            adaT3.Read()
                            If adaT3.HasRows = True Then
                                NUEVOMONTO = adaT3.Item(0).ToString
                                ITEMNNUMERO = adaT3.Item(1)
                            End If
                        End Using
                        NUEVOMONTO = (NUEVOMONTO + RESTAHD)
                    End If

                    If SUMAD > SUMAH Then
                        Using connT3 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            connT3.Open()
                            Dim adaT3 As SqlDataReader = New SqlCommand("SELECT TOP(1) MONTOMOV, NUM_PART FROM " & AUXILIAR & " WHERE DEBE_HABER='H' AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT3).ExecuteReader
                            adaT3.Read()
                            If adaT3.HasRows = True Then
                                NUEVOMONTO = adaT3.Item(0).ToString
                                ITEMNNUMERO = adaT3.Item(1)
                            End If
                        End Using
                        NUEVOMONTO = (NUEVOMONTO + RESTAHD)
                    End If
                    Using ConnReg As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        ConnReg.Open()
                        Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & NUEVOMONTO & "'  WHERE NUM_PART=" & ITEMNNUMERO & " AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Ve' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), ConnReg)
                        command666.ExecuteNonQuery()
                        ConnReg.Close()
                    End Using
                End If
                '***PARA REVISAR LOS IMPORTES Y CUADRARLOS***
                TOTAL_UNO = 0.0
                RESTA = 0.0
                MIMONTO = 0.0

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
                NIVEL1 = ""
                NIVEL2 = ""
                NIVEL3 = ""
                NIVEL11 = ""
                NIVEL22 = ""
                NIVEL33 = ""
                NIVEL111 = ""
                NIVEL222 = ""
                NIVEL333 = ""
                NIVEL1111 = ""
                NIVEL2222 = ""
                NIVEL3333 = ""
                CADENA10 = ""
                CADENA11 = ""
                CADENA12 = ""
                CADENA13 = ""
                CADENA14 = ""

                j += 1
            Loop
            ConFact.Close()
        End Using
    End Sub
End Class
