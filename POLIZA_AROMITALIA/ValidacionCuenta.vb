Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.CorreoSinCuentaContable
Imports POLIZA_AROMITALIA.ConsultaTablaDatos

Imports POLIZA_AROMITALIA.CorreoProducto



Public Class ValidacionCuenta
    Public Shared NO_VALOR As Integer = 0
    Public Shared NO_ACUM As Integer = 0
    Public Shared CUENTA_PARTIDA As String
    Public Shared letra As String
    Public Shared Sub ValidacionCuentaContable()
        Dim num As Integer = 0

        'CONSUTA A LA TABLA FACTF01

        Using connection As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adapter As New SqlDataAdapter("SELECT " & FACTF & ".TIP_DOC, " & FACTF & ".CVE_DOC, " & FACTF & ".CVE_CLPV, " & FACTF & ".STATUS, " & FACTF & ".FECHA_DOC, " & FACTF & ".FECHA_ENT, " & FACTF & ".FECHA_VEN, " & FACTF & ".CAN_TOT, " & FACTF & ".IMP_TOT1, " & FACTF & ".IMP_TOT4, " & FACTF & ".CVE_OBS, " & FACTF & ".NUM_ALMA, " & FACTF & ".ACT_CXC, " & FACTF & ".ACT_COI, " & FACTF & ".ENLAZADO, " & FACTF & ".TIP_DOC_E, " & FACTF & ".NUM_MONED, " & FACTF & ".TIPCAMB, " & FACTF & ".NUM_PAGOS, " & FACTF & ".PRIMERPAGO, " & FACTF & ".RFC, " & FACTF & ".CTLPOL, " & FACTF & ".ESCFD, " & FACTF & ".AUTORIZA, " & FACTF & ".SERIE, " & FACTF & ".FOLIO, " & FACTF & ".DAT_ENVIO, " & FACTF & ".CONTADO, " & FACTF & ".CVE_BITA, " & FACTF & ".BLOQ, " & FACTF & ".FORMAENVIO, " & FACTF & ".DES_FIN_PORC, " & FACTF & ".DES_TOT_PORC, " & FACTF & ".IMPORTE, " & FACTF & ".COM_TOT_PORC, " & FACTF & ".TIP_DOC_ANT," & FACTF & ".DOC_ANT FROM  " & FACTF & " INNER JOIN INSOFTEC_FACTF_CLIB01 On " & FACTF & ".CVE_DOC = INSOFTEC_FACTF_CLIB01.CLAVE_DOC WHERE (INSOFTEC_FACTF_CLIB01.CAMPLIB1 IS NULL) AND (" & FACTF & ".STATUS <> 'C')", connection)

            Dim dataTable As New DataTable
            adapter.Fill(dataTable)
            connection.Open()
            Dim num6 As Integer = (dataTable.Rows.Count - 1)
            Dim i As Integer = 0
            Do While (i <= num6)

                TIP_DOC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(0).ToString
                CVE_DOC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(1).ToString
                CVE_CLPV = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(2).ToString
                STATUS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(3).ToString


                FECHA_DOC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(4)
                FECHA_ENT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(5)
                FECHA_VEN = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(6)
                CAN_TOT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(7)

                IMP_TOT1 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(8)
                IMP_TOT4 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(9)

                CVE_OBS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(10)
                NUM_ALMA = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(11)
                ACT_CXC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(12).ToString
                ACT_COI = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(13).ToString
                ENLAZADO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(14).ToString
                TIP_DOC_E = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(15).ToString
                NUM_MONED = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(16)
                TIPCAMB = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(17)
                NUM_PAGOS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(18)

                PRIMERPAGO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(19)
                RFC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(20).ToString
                CTLPOL = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(21)
                ESCFD = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(22).ToString
                AUTORIZA = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(23)
                SERIE = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(24).ToString
                FOLIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(25)

                DAT_ENVIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(26)
                CONTADO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(27).ToString
                CVE_BITA = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(28)
                BLOQ = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(29).ToString
                FORMAENVIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(30).ToString
                DES_FIN_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(31)

                DES_TOT_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(32)
                IMPORTE = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(33)
                COM_TOT_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(34)
                TIP_DOC_ANT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(35).ToString
                DOC_ANT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(36).ToString

                Dim extrae_per As String = FECHA_DOC
                Dim Testarray() As String = extrae_per.Split("/")

                EJERCICIO = Testarray(2).Trim
                PERIODO = Testarray(1).Trim
                PERIODO1 = Testarray(1).Trim
                'PARA TOMAR EL AÑO DEL EJERCICIO FISCAL
                Dim recortar As String = EJERCICIO
                Dim EJER_FISCAL As String = Microsoft.VisualBasic.Right(recortar, 2)
                AUXILIAR = AUXI & EJER_FISCAL
                POLIZA = POLIZ & EJER_FISCAL
                CUENTA = CUENT & EJER_FISCAL

                Using connection2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connection2.Open()
                    Dim adapter2 As SqlDataReader = New SqlCommand("SELECT max(NUM_POLIZ) FROM " & POLIZA & " WHERE PERIODO ='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "' AND TIPO_POLI='" & TIPOFACTF & "'", connection2).ExecuteReader
                    adapter2.Read()
                    If adapter2.HasRows = True Then
                        If IsDBNull(adapter2.Item(0)) Then
                            contador = 0
                        Else
                            contador = adapter2.Item(0)

                        End If
                    Else
                        contador = 0
                    End If
                    contador = contador + 1
                    connection2.Close()
                End Using

                'FOLIOS
                Using INSERT_FOLIO As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    Dim UPDATEFOLIOS As String = Now.Date.Month
                    INSERT_FOLIO.Open()
                    'ACTUALIZA EL FOLIO
                    'Dim command661 As New SqlCommand(("UPDATE FOLIOS01 SET FOLIO" & PERIODO1 & " ='" & contador & "' WHERE EJERCICIO ='" & EJERCICIO & "' AND TIPPOL='Ve'"), INSERT_FOLIO)
                    'command661.ExecuteNonQuery()
                    INSERT_FOLIO.Close()
                End Using

                Dim FECH As String = FECHA_DOC.ToString("yyyy-dd-MM")
                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)

                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT COUNT(CVE_DOC) FROM " & PAR_FACTF & " WHERE (CVE_DOC='" & CVE_DOC & "')", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    totpartida = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using

                ''''''''CONSULTA PARTIDAS
                Using BUSCAPARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim ada1 As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PXS,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,DESC1,DESC2,APAR,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA,IMPRIMIR FROM " & PAR_FACTF & " where CVE_DOC='" & CVE_DOC & "'", BUSCAPARTIDAS)
                    Dim data1 As New DataTable
                    ada1.Fill(data1)
                    BUSCAPARTIDAS.Open()

                    Dim num07 As Integer = (data1.Rows.Count - 1)
                    Dim j As Integer = 0
                    Dim CUENTA_PARTIDAS As Integer = data1.Rows.Count
                    Dim K As Integer = 2
                    Do While (j <= num07)

                        NUM_PAR = data1.AsEnumerable().ElementAtOrDefault(j).Item(0)
                        CVE_ART = data1.AsEnumerable().ElementAtOrDefault(j).Item(1).ToString
                        CANT = data1.AsEnumerable().ElementAtOrDefault(j).Item(2)
                        PXS = data1.AsEnumerable().ElementAtOrDefault(j).Item(3)
                        PREC = data1.AsEnumerable().ElementAtOrDefault(j).Item(4)
                        COST = data1.AsEnumerable().ElementAtOrDefault(j).Item(5)
                        IMPU4 = data1.AsEnumerable().ElementAtOrDefault(j).Item(6)
                        IMP1APLA = data1.AsEnumerable().ElementAtOrDefault(j).Item(7).ToString
                        IMP2APLA = data1.AsEnumerable().ElementAtOrDefault(j).Item(8).ToString
                        IMP3APLA = data1.AsEnumerable().ElementAtOrDefault(j).Item(9).ToString
                        DESC1 = data1.AsEnumerable().ElementAtOrDefault(j).Item(10)
                        DESC2 = data1.AsEnumerable().ElementAtOrDefault(j).Item(11)
                        APAR = data1.AsEnumerable().ElementAtOrDefault(j).Item(12)
                        ACT_INV = data1.AsEnumerable().ElementAtOrDefault(j).Item(13).ToString
                        NUM_ALM = data1.AsEnumerable().ElementAtOrDefault(j).Item(14)
                        TIP_CAM = data1.AsEnumerable().ElementAtOrDefault(j).Item(15)
                        UNI_VENTA = data1.AsEnumerable().ElementAtOrDefault(j).Item(16).ToString
                        TIPO_PROD = data1.AsEnumerable().ElementAtOrDefault(j).Item(17)
                        CVE_OBS = data1.AsEnumerable().ElementAtOrDefault(j).Item(18)
                        REG_SERIE = data1.AsEnumerable().ElementAtOrDefault(j).Item(19)
                        E_LTPD = data1.AsEnumerable().ElementAtOrDefault(j).Item(20)
                        TIPO_ELEM = data1.AsEnumerable().ElementAtOrDefault(j).Item(21).ToString
                        NUM_MOV = data1.AsEnumerable().ElementAtOrDefault(j).Item(22)
                        TOT_PARTIDA = data1.AsEnumerable().ElementAtOrDefault(j).Item(23)
                        IMPRIMIR = data1.AsEnumerable().ElementAtOrDefault(j).Item(24).ToString


                        '''''sumatorias
                        tot_part = 0
                        TOTIMP4 = 0

                        tot_part1 = TOT_PARTIDA + TOT_PARTIDA

                        TOTIMP41 = TOTIMP41 + IMPU4

                        NUM_PAR1 = NUM_PAR + 2
                        NUM_PAR2 = NUM_PAR + 2
                        NUM_PAR3 = NUM_PAR + 3

                        Using connection2225581 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            connection2225581.Open()
                            Dim adapter22255801 As SqlDataReader = New SqlCommand("SELECT CUENT_CONT, CTRL_ALM, DESCR FROM " & INVE & " WHERE CVE_ART ='" & CVE_ART & "'", connection2225581).ExecuteReader

                            adapter22255801.Read()
                            CUENTA_COI = adapter22255801.Item(0).ToString
                            CTRL = adapter22255801.Item(1).ToString
                            MIDESCR = adapter22255801.Item(2).ToString
                            CUENTA_COI = Replace(CUENTA_COI, "-", "")

                            Using connection222540 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                connection222540.Open()
                                Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT NUM_CTA, NOMBRE FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) = '" + CUENTA_COI + "')", connection222540).ExecuteReader
                                adapter222540.Read()
                                If adapter222540.HasRows = True Then
                                    LIN_PROD2 = adapter222540.Item(0).ToString
                                    CONCEPTO2 = adapter222540.Item(1).ToString
                                    'Else
                                    '    LIN_PROD2 = "00000000000000000000"
                                    '    CONCEPTO2 = ""
                                End If

                            End Using

                            'If CTRL = "AROMITALIA" Then
                            '    CTRL = "400001001000000000003"
                            'End If

                            'If CTRL = "S&G" Then
                            '    CTRL = "400001002000000000003"
                            'End If

                            'If CTRL = "INSUMOS" Then
                            '    CTRL = "400001003000000000003"
                            'End If

                            'If CTRL = "KAMPAI" Then
                            '    CTRL = "400001005000000000003"
                            'End If

                            'If CTRL = "MAQ" Then
                            '    CTRL = "400001005000000000003"
                            'End If

                            'If CTRL = "MYD" Then
                            '    CTRL = "400002009000000000003"
                            'End If

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
                            If SERIE = "F" And IMPU4 = 0 Or IMPU4 = 0.0 Then
                                NUEVOIVA = CAMPLIB10
                            ElseIf SERIE = "F" And IMPU4 = 16 Then
                                NUEVOIVA = CAMPLIB11
                            ElseIf SERIE = "FD" And IMPU4 = 0 Or IMPU4 = 0.0 Then
                                NUEVOIVA = CAMPLIB12
                            ElseIf SERIE = "FD" And IMPU4 = 16 Then
                                NUEVOIVA = CAMPLIB13
                            End If


                            'PARA VALIDAR LAS PARTIDAS CON CUENTA CONTABLE
                            Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                                INSERT_TOTPART.Open()
                                Dim adapter50 As SqlDataReader = New SqlCommand("SELECT CAMPLIB7 FROM INSOFTEC_FACTF_CLIB01 WHERE CLAVE_DOC = '" & CVE_DOC & "'", INSERT_TOTPART).ExecuteReader
                                adapter50.Read()
                                CAML4 = adapter50.Item(0).ToString
                                INSERT_TOTPART.Close()
                            End Using

                            If NUEVOIVA = "" Then
                                NO_VALOR = NO_VALOR + 1
                                NO_ACUM = NO_VALOR
                                'CUENTA_PARTIDA = CVE_ART

                                If NO_VALOR > 0 And NUEVOIVA = "" And CAML4 = "" Then
                                    ProductoSinCuenta()
                                    letra = String.Concat(String.Concat(NUEVOIVA, CVE_ART))
                                    'Using cn4 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                                    '    cn4.Open()
                                    '    Dim command6611 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01  SET CAMPLIB8='" & letra + "-" & "', CAMPLIB7='ENVIADO', CAMPLIB9='ERROR CTA.' WHERE CLAVE_DOC = '" & CVE_DOC & "' "), cn4)
                                    '    command6611.ExecuteNonQuery()
                                    '    cn4.Close()
                                    'End Using
                                End If


                            End If


                            CTRL = ""
                            CUENTA_PARTIDA = ""

                            K = K + 1

                            NUM_PAR = 0
                            CVE_DOCNUM_PARCVE_ART = 0
                            CANT = 0
                            PXS = 0
                            PREC = 0
                            COST = 0

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
                            NUEVOIVA = ""
                            connection2225581.Close()
                            'connection22.Close()
                        End Using

                        j += 1

                    Loop

                    'CLIE01
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE, NOMBRE FROM " & CLIE & " WHERE CLAVE ='" & CVE_CLPV & "'", connection22).ExecuteReader

                        adapter2231.Read()
                        CUENTA_CONTABLE = adapter2231.Item(0).ToString
                        NOMBRE_CLIENTE = adapter2231.Item(1).ToString
                        CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                        'connection22.Close()
                    End Using
                    'CONSULTAS CUENTA CONTABLE

                    Using connection2225 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        connection2225.Open()
                        Dim adapter2225 As SqlDataReader = New SqlCommand("SELECT NUM_CTA, NOMBRE FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE'" & CUENTA_CONTABLE & "%')", connection2225).ExecuteReader
                        adapter2225.Read()
                        If adapter2225.HasRows = True Then
                            CUENTA_CONTABLE1 = adapter2225.Item(0).ToString
                            CONCEPTO = adapter2225.Item(1).ToString
                        Else
                            CONCEPTO = ""
                            CUENTA_CONTABLE1 = ""

                        End If

                        connection2225.Close()

                    End Using

                    DEF1 = "250001000000000000002"
                    DEF2 = "270002000000000000002"



                    Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)

                        INSERT_PARTIDAS.Open()
                        Dim v = IsNumeric(DEPTO)
                        CONCEPTO = Replace(CONCEPTO, "'", "")
                        If v = True Then
                            'Dim command6 As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & NUM_PAR & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_CONTABLE + "000000000003" & "','" & FECH & "','" + "FACTURA " & CVE_DOC + "    " + CONCEPTO + " " & "','D','" & IMPORTE & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS)
                            'command6.ExecuteNonQuery()
                        Else
                            DEPTO = "0"
                            'Dim command6 As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & NUM_PAR + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_CONTABLE + "000000000003" & "','" & FECH & "','" + "FACTURA " & CVE_DOC & "','D','" & IMPORTE & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS)
                            'command6.ExecuteNonQuery()
                        End If
                        INSERT_PARTIDAS.Close()

                    End Using

                    Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)

                        INSERT_PARTIDAS.Open()
                        Dim v = IsNumeric(DEPTO)
                        CONCEPTO3 = Replace(CONCEPTO3, "'", "")
                        If v = True Then
                            'Dim commandD As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K & "','" & PERIODO & "','" & EJERCICIO & "','" & DEF1 & "','" & FECH & "','" + "FACTURA " & CVE_DOC & "','H','" & IMP_TOT4 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS) 'DEF1                       
                            'commandD.ExecuteNonQuery()

                        Else
                            DEPTO = "0"
                            'Dim commandD As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "FACTURA" & CVE_DOC & "','H','" & IMP_TOT4 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDAS)
                            'commandD.ExecuteNonQuery()
                        End If

                        INSERT_PARTIDAS.Close()

                    End Using



                    Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)

                        INSERT_PARTIDASs.Open()
                        Dim x = IsNumeric(DEPTO)
                        If x = x Then
                            CONCEPTO = Replace(CONCEPTO, "'", "")
                            'Dim commandDs As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('    ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & DEF2 & "','" & FECH & "','" + "FACTURA " & CVE_DOC & "','H','" & IMP_TOT1 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs) 'DEF1
                            'commandDs.ExecuteNonQuery()

                        Else
                            'Dim commandDs As New SqlCommand(("INSERT INTO AUXILIAR1501 (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('Ve',RIGHT('     ' + LTRIM(RTRIM(" & j & ")), 5),'" & K + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "FACTURA" & CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO3 + "','H','" & IMP_TOT1 & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs)
                            'commandDs.ExecuteNonQuery()
                        End If

                        INSERT_PARTIDASs.Close()

                    End Using

                End Using


                ''PARA CONTAR LAS PARTIDAS E INSERTARLAS EN LA TABLA POLIZAS1501

                If CUENTA_CONTABLE1 = "100000000000000000001" Then
                    CUENTA_CONTABLE1 = ""
                End If

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT CAMPLIB6 FROM INSOFTEC_FACTF_CLIB01 WHERE CLAVE_DOC = '" & CVE_DOC & "'", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    C4 = adapter50.Item(0).ToString

                    INSERT_TOTPART.Close()
                End Using
                If CUENTA_CONTABLE1 = "" Or CUENTA_CONTABLE1 = "100000000000000000001" Then
                    'If NO_VALOR > 0 And NUEVOIVA = "" And CAML4 = "" Then
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01 SET CAMPLIB4='ERROR CTA.', CAMPLIB6='ENVIADO' WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        If C4 = "" Then
                            EnviarCorreoSinCuenta()
                        End If
                        INSERT_CAMP.Close()
                    End Using
                End If

                If CUENTA_CONTABLE1 <> "" Then
                    ' If NO_VALOR = 0 Then
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01 SET CAMPLIB4 = NULL, CAMPLIB6 = NULL WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using
                End If

                If NO_VALOR > 0 Then
                    Using cn4 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        cn4.Open()
                        Dim command6611 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01  SET CAMPLIB8='" & letra + "-" & "', CAMPLIB7='ENVIADO', CAMPLIB9='ERROR CTA.' WHERE CLAVE_DOC = '" & CVE_DOC & "' "), cn4)
                        command6611.ExecuteNonQuery()
                        cn4.Close()
                    End Using
                End If
                If NO_ACUM = 0 Then
                    Using cn4 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        cn4.Open()
                        'ACTUALIZA EL FOLIO
                        Dim command6611 As New SqlCommand(("UPDATE INSOFTEC_FACTF_CLIB01  SET CAMPLIB8=NULL, CAMPLIB7=NULL,CAMPLIB9=NULL WHERE CLAVE_DOC = '" & CVE_DOC & "' "), cn4)
                        command6611.ExecuteNonQuery()
                        cn4.Close()
                    End Using
                End If

                NUM_PAR = 0
                C3 = ""
                C4 = ""
                TIP_DOC = ""
                CVE_DOC = ""
                CVE_CLPV = ""
                STATUS = ""
                IMPU4 = 0
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
                NOMBRE_CLIENTE = ""
                NO_VALOR = 0
                NO_ACUM = 0
                NO_VALOR = 0
                ' NUEVOIVA = ""
                CAML4 = ""


                i += 1

            Loop

            connection.Close()
        End Using
        contador1 = contador
    End Sub
End Class
