Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports FirebirdSql.Data.FirebirdClient
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.CorreoProducto
Imports POLIZA_AROMITALIA.CorreoSinCasilla
Imports POLIZA_AROMITALIA.EnviarCorreoSinUUID
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Imports POLIZA_AROMITALIA.ProdSinCuentaCompra
Imports POLIZA_AROMITALIA.SinCasillaXML

Public Class ProcesarDevolucionCompra

    Public Shared Sub DevolucionCompra()
        Dim num As Integer = 0
        Dim TOTAL_UNO As Double = 0.0
        Dim RESTA As Double = 0.0
        Dim MIMONTO As Double = 0.0
        'CONSUTA A LA TABLA COMPC01
        Using connection As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adapter As New SqlDataAdapter("SELECT " & COMPD & ".TIP_DOC, " & COMPD & ".CVE_DOC, " & COMPD & ".CVE_CLPV, " & COMPD & ".STATUS, " & COMPD & ".SU_REFER, " & COMPD & ".FECHA_DOC, " & COMPD & ".FECHA_REC, " & COMPD & ".FECHA_PAG, " & COMPD & ".CAN_TOT, " & COMPD & ".IMP_TOT1, " & COMPD & ".IMP_TOT2, " & COMPD & ".IMP_TOT4, " & COMPD & ".CVE_OBS, " & COMPD & ".NUM_ALMA, " & COMPD & ".ACT_CXP, " & COMPD & ".ACT_COI, " & COMPD & ".ENLAZADO, " & COMPD & ".TIP_DOC_E, " & COMPD & ".NUM_MONED, " & COMPD & ".TIPCAMB, " & COMPD & ".NUM_PAGOS, " & COMPD & ".CTLPOL, " & COMPD & ".ESCFD,  " & COMPD & ".SERIE, " & COMPD & ".FOLIO, " & COMPD & ".CONTADO, " & COMPD & ".BLOQ, " & COMPD & ".FORMAENVIO, " & COMPD & ".DES_FIN_PORC, " & COMPD & ".DES_TOT_PORC, " & COMPD & ".IMPORTE, " & COMPD & ".TIP_DOC_ANT, " & COMPD & ".DOC_ANT FROM  " & COMPD & " INNER JOIN INSOFTEC_COMPD_CLIB01 On " & COMPD & ".CVE_DOC = INSOFTEC_COMPD_CLIB01.CLAVE_DOC WHERE (INSOFTEC_COMPD_CLIB01.CAMPLIB1 Is NULL AND INSOFTEC_COMPD_CLIB01.CAMPLIB4 Is NULL) AND (" & COMPD & ".STATUS <> 'C')", connection)
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
                SU_REFER = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(4).ToString
                FECHA_DOC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(5)
                FECHA_REC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(6)
                FECHA_PAG = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(7)
                CAN_TOT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(8)
                IMP_TOT1 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(9)
                IMP_TOT2 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(10)
                IMP_TOT4 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(11)
                CVE_OBS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(12)
                NUM_ALMA = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(13)
                ACT_CXP = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(14).ToString
                ACT_COI = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(15).ToString
                ENLAZADO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(16).ToString
                TIP_DOC_E = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(17).ToString
                NUM_MONED = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(18)
                TIPCAMB = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(19)
                NUM_PAGOS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(20)
                CTLPOL = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(21)
                ESCFD = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(22).ToString
                SERIE = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(23).ToString
                FOLIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(24)
                CONTADO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(25).ToString
                BLOQ = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(26).ToString
                FORMAENVIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(27).ToString
                DES_FIN_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(28)
                DES_TOT_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(29)
                IMPORTE = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(30)
                'COM_TOT_PORC1 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(29).ToString
                TIP_DOC_ANT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(31).ToString
                DOC_ANT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(32).ToString

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
                    Dim adapter2 As SqlDataReader = New SqlCommand("SELECT max(NUM_POLIZ) FROM " & POLIZA & " WHERE PERIODO ='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "' AND TIPO_POLI='" & TIPOCOMPD & "'", connection2).ExecuteReader
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
                    Dim command661 As New SqlCommand(("UPDATE " & FOLIOS & " SET FOLIO" & PERIODO1 & " ='" & contador & "' WHERE EJERCICIO ='" & EJERCICIO & "' AND TIPPOL='" & TIPOCOMPD & "'"), INSERT_FOLIO)
                    command661.ExecuteNonQuery()
                    INSERT_FOLIO.Close()
                End Using

                Dim FECH As String = FECHA_DOC.ToString("yyyy-MM-dd")
                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)

                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT COUNT(CVE_DOC) FROM " & PAR_COMPD & " WHERE (CVE_DOC='" & CVE_DOC & "')", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    totpartida = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using

                ''''''''CONSULTA PARTIDAS
                Using BUSCAPARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim ada1 As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA FROM " & PAR_COMPD & " WHERE CVE_DOC='" & CVE_DOC & "'", BUSCAPARTIDAS)
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
                        'PXS = data1.AsEnumerable().ElementAtOrDefault(j).Item(3)
                        PREC = data1.AsEnumerable().ElementAtOrDefault(j).Item(3)
                        COST = data1.AsEnumerable().ElementAtOrDefault(j).Item(4)
                        IMPU4 = data1.AsEnumerable().ElementAtOrDefault(j).Item(5)
                        IMP1APLA = data1.AsEnumerable().ElementAtOrDefault(j).Item(6).ToString
                        IMP2APLA = data1.AsEnumerable().ElementAtOrDefault(j).Item(7).ToString
                        IMP3APLA = data1.AsEnumerable().ElementAtOrDefault(j).Item(8).ToString
                        'DESC1 = data1.AsEnumerable().ElementAtOrDefault(j).Item(10)
                        'DESC2 = data1.AsEnumerable().ElementAtOrDefault(j).Item(11)
                        'APAR = data1.AsEnumerable().ElementAtOrDefault(j).Item(12)
                        ACT_INV = data1.AsEnumerable().ElementAtOrDefault(j).Item(9).ToString
                        NUM_ALM = data1.AsEnumerable().ElementAtOrDefault(j).Item(10)
                        TIP_CAM = data1.AsEnumerable().ElementAtOrDefault(j).Item(11)
                        UNI_VENTA = data1.AsEnumerable().ElementAtOrDefault(j).Item(12).ToString
                        TIPO_PROD = data1.AsEnumerable().ElementAtOrDefault(j).Item(13)
                        CVE_OBS = data1.AsEnumerable().ElementAtOrDefault(j).Item(14)
                        REG_SERIE = data1.AsEnumerable().ElementAtOrDefault(j).Item(15)
                        E_LTPD = data1.AsEnumerable().ElementAtOrDefault(j).Item(16)
                        TIPO_ELEM = data1.AsEnumerable().ElementAtOrDefault(j).Item(17).ToString
                        NUM_MOV = data1.AsEnumerable().ElementAtOrDefault(j).Item(18)
                        TOT_PARTIDA = data1.AsEnumerable().ElementAtOrDefault(j).Item(19)
                        'IMPRIMIR = data1.AsEnumerable().ElementAtOrDefault(j).Item(24).ToString


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
                                Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT NUM_CTA, NOMBRE, CAPTURAUUID FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) = '" + CUENTA_COI + "')", connection222540).ExecuteReader
                                adapter222540.Read()
                                If adapter222540.HasRows = True Then
                                    LIN_PROD2 = adapter222540.Item(0).ToString
                                    CONCEPTO2 = adapter222540.Item(1).ToString
                                Else
                                    LIN_PROD2 = "000000000000000000000"
                                    CONCEPTO2 = ""
                                End If

                            End Using

                            Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                                connection22.Open()
                                Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT CAMPLIB8 FROM " & INVE_CLIB & " WHERE CVE_PROD ='" & CVE_ART & "'", connection22).ExecuteReader
                                adapter2231.Read()
                                MYACCOUNT = adapter2231.Item(0).ToString
                                MYACCOUNT = Replace(MYACCOUNT, "-", "")
                                Dim two As String = ""
                                Dim tree As String = ""
                                If MYACCOUNT >= "610000" And MYACCOUNT <= "620019" Then
                                    MYACCOUNT = MYACCOUNT + "000000000002"
                                Else
                                    MYACCOUNT = MYACCOUNT + "000000000003"
                                End If


                                'connection22.Close()
                            End Using

                            'PARA ENVIAR CORREO CUANDO EL CAMPO CTRL_ALM ESTÁ VACÍA
                            If MYACCOUNT = "" Then
                                'ProductoSinCuenta()
                                EnviarProdSCC()
                            End If

                            Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                                connection22.Open()
                                Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE FROM " & PROV & " WHERE CLAVE ='" & CVE_CLPV & "'", connection22).ExecuteReader

                                adapter2231.Read()
                                CUENTA_CONTABLE = adapter2231.Item(0).ToString
                                CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")

                                'connection22.Close()
                            End Using

                            Using connection222540 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                connection222540.Open()
                                Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT CAPTURAUUID  FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE '" + CUENTA_CONTABLE + "%')", connection222540).ExecuteReader
                                adapter222540.Read()
                                If adapter222540.HasRows = True Then
                                    uuid = adapter222540.Item(0)

                                End If

                            End Using

                            'CENTRO DE COSTOS

                            'Using connection222558 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                            '    connection222558.Open()
                            '    Dim adapter2225580 As SqlDataReader = New SqlCommand("SELECT CAMPLIB1,CAMPLIB2,CAMPLIB3 FROM INSOFTEC_FACTF_CLIB01 WHERE CLAVE_DOC='" & CVE_DOC & "'", connection222558).ExecuteReader
                            '    adapter2225580.Read()
                            '    DEPTO = adapter2225580.Item(0).ToString
                            '    ccostos = adapter2225580.Item(1).ToString
                            '    If adapter2225580.HasRows = True Then
                            '        cuentacon = adapter2225580.Item(2).ToString
                            '        Dim v = IsNumeric(cuentacon)
                            '        If v = False Then
                            '            cuentacon = "0"
                            '        End If
                            '    Else
                            '        cuentacon = "0"
                            '        ccostos = "0"
                            '    End If
                            '    connection222558.Close()
                            'End Using

                            'VALIDACIONES DE CUENTAS CONTABLES
                            ' CTRL = "117002000000000000002"
                            'If CTRL = "AROMITALIA" And IMPU4 = 0 Then
                            '    CTRL = "400001001000000000003"
                            'End If
                            'If CTRL = "AROMITALIA" And IMPU4 = 16 Then
                            '    CTRL = "400002001000000000003"
                            'End If
                            'If CTRL = "S&G" And IMPU4 = 0 Then
                            '    CTRL = "400001002000000000003"
                            'End If
                            'If CTRL = "S&G" And IMPU4 = 16 Then
                            '    CTRL = "400002002000000000003"
                            'End If
                            'If CTRL = "INSUMOS" And IMPU4 = 0 Then
                            '    CTRL = "400001003000000000003"
                            'End If
                            'If CTRL = "INSUMOS" And IMPU4 = 16 Then
                            '    CTRL = "400002003000000000003"
                            'End If
                            'If CTRL = "KAMPAI" And IMPU4 = 0 Then
                            '    CTRL = "400001005000000000003"
                            'End If
                            'If CTRL = "KAMPAI" And IMPU4 = 16 Then
                            '    CTRL = "400002005000000000003"
                            'End If
                            'If CTRL = "MAQ" And IMPU4 = 0 Then
                            '    CTRL = "400001005000000000003"
                            'End If
                            'If CTRL = "MAQ" And IMPU4 = 16 Then
                            '    CTRL = "400002008000000000003"
                            'End If
                            'If CTRL = "MYD" And IMPU4 = 16 Then
                            '    CTRL = "400002009000000000003"
                            'End If
                            'partidas

                            Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)

                                INSERT_PARTIDAS.Open()
                                Dim v = IsNumeric(DEPTO)
                                CONCEPTO2 = Replace(CONCEPTO2, "'", "")
                                If v = True Then
                                    Dim command6 As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K & "','" & PERIODO & "','" & EJERCICIO & "','" & MYACCOUNT & "','" & FECH & "','" + CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO2 + "','H','" & Math.Round(TOT_PARTIDA, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR1 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDAS) 'LIN_PROD2                           
                                    command6.ExecuteNonQuery()
                                Else
                                    DEPTO = "0"
                                    Dim command6 As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & K & "','" & PERIODO & "','" & EJERCICIO & "','" & MYACCOUNT & "','" & FECH & "','" + CVE_DOC + "  " + CVE_ART + "  " + MIDESCR + "','H','" & Math.Round(TOT_PARTIDA, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR1 - 1 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDAS) 'PARTIDAS SIN DESCUENTOS

                                    command6.ExecuteNonQuery()

                                End If
                                INSERT_PARTIDAS.Close()

                            End Using
                            K = K + 1

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

                            connection2225581.Close()
                            'connection22.Close()

                        End Using

                        j += 1

                    Loop


                    'CLIE01
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE, NOMBRE, RFC FROM " & PROV & " WHERE CLAVE ='" & CVE_CLPV & "'", connection22).ExecuteReader

                        adapter2231.Read()
                        CUENTA_CONTABLE = adapter2231.Item(0).ToString
                        NOMBRE_CLIENTE = adapter2231.Item(1).ToString
                        RFCPROVEEDOR = adapter2231.Item(2).ToString
                        CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                        'connection22.Close()
                    End Using


                    'CONSULTAS CUENTA CONTABLE

                    Using connection2225 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        connection2225.Open()
                        Dim adapter2225 As SqlDataReader = New SqlCommand("SELECT NUM_CTA, NOMBRE FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) LIKE '" + CUENTA_CONTABLE + "%')", connection2225).ExecuteReader
                        adapter2225.Read()
                        If adapter2225.HasRows = True Then
                            CUENTA_CONTABLE1 = adapter2225.Item(0).ToString
                            CONCEPTO = adapter2225.Item(1).ToString
                        Else
                            CONCEPTO = ""
                            CUENTA_CONTABLE1 = "000000000003"

                        End If

                        connection2225.Close()

                    End Using
                    'DEF1 = "250001000000000000002"
                    DEF1 = "117002000000000000002" 'cuenta del IVA
                    DEF2 = "270002000000000000002"



                    Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)

                        INSERT_PARTIDAS.Open()
                        Dim v = IsNumeric(DEPTO)
                        CONCEPTO = Replace(CONCEPTO, "'", "")
                        If v = True Then
                            Dim command6 As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & NUM_PAR & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_CONTABLE + "000000000003" & "','" & FECH & "','" + "COMPRA " & CVE_DOC + "    " + CONCEPTO + " " & "','D','" & Math.Round(IMPORTE, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR + 1 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDAS)
                            command6.ExecuteNonQuery()
                        Else
                            DEPTO = "0"
                            Dim command6 As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & NUM_PAR + 1 & "','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_CONTABLE + "000000000003" & "','" & FECH & "','" + "COMPRA " & CVE_DOC & "','D','" & Math.Round(IMPORTE, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDAS)
                            command6.ExecuteNonQuery()
                        End If
                        INSERT_PARTIDAS.Close()
                    End Using

                    Using INSERT_PARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        If IMP_TOT4 > 0 Then
                            INSERT_PARTIDAS.Open()
                            Dim v = IsNumeric(DEPTO)
                            CONCEPTO3 = Replace(CONCEPTO3, "'", "")
                            If v = True Then
                                Dim commandD As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'555','" & PERIODO & "','" & EJERCICIO & "','" & DEF1 & "','" & FECH & "','" + "COMPRA " & CVE_DOC & "','D','" & Math.Round(IMP_TOT4, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDAS) 'DEF1                           
                                commandD.ExecuteNonQuery()

                            Else
                                DEPTO = "0"
                                Dim commandD As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'555','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "COMPRA " & CVE_DOC & "','D','" & Math.Round(IMP_TOT4, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 + 1 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDAS)
                                commandD.ExecuteNonQuery()
                            End If

                            INSERT_PARTIDAS.Close()
                        End If
                    End Using


                    Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        INSERT_TOTPART.Open()
                        Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                        adapter50.Read()
                        MIPAR3 = adapter50.Item(0)
                        INSERT_TOTPART.Close()
                    End Using

                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR3 & "', ORDEN='" & MIPAR3 & "' WHERE NUM_PART =555 AND TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using

                    Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        If IMP_TOT1 > 0 Then
                            INSERT_PARTIDASs.Open()
                            Dim x = IsNumeric(DEPTO)
                            If x = x Then
                                CONCEPTO = Replace(CONCEPTO, "'", "")
                                Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('    ' + LTRIM(RTRIM(" & contador & ")), 5),'666','" & PERIODO & "','" & EJERCICIO & "','" & DEF2 & "','" & FECH & "','" + "COMPRA " & CVE_DOC & "','D','" & Math.Round(IMP_TOT1, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 + 1 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDASs) 'DEF1                         
                                commandDs.ExecuteNonQuery()
                            Else
                                Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & j & ")), 5),'666','" & PERIODO & "','" & EJERCICIO & "','" & CUENTA_COI & "','" & FECH & "','" + "COMPRA " & CVE_DOC + "  " + CVE_ART + "  " + CONCEPTO3 + "','D','" & Math.Round(IMP_TOT1, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & NUM_PAR2 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDASs)
                                commandDs.ExecuteNonQuery()
                            End If

                            INSERT_PARTIDASs.Close()
                        End If
                    End Using


                    Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        INSERT_TOTPART.Open()
                        Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                        adapter50.Read()
                        MIPAR4 = adapter50.Item(0)
                        INSERT_TOTPART.Close()
                    End Using
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR4 & "', ORDEN='" & MIPAR4 & "' WHERE NUM_PART =666 AND TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using


                End Using


                Using INSERT_FOLIO As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_FOLIO.Open()
                    Dim command661 As New SqlCommand(("UPDATE " & AUXILIAR & " set NUM_PART=1 WHERE NUM_PART=0 AND TIPO_POLI='" & TIPOCOMPD & "'"), INSERT_FOLIO)
                    command661.ExecuteNonQuery()
                    INSERT_FOLIO.Close()
                End Using


                'para las partidas de impuesto retenido
                Using INSERT_PARTIDASs As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)

                    IMP_RET = Math.Abs(IMP_TOT2)
                    If IMP_RET > 0 Then
                        INSERT_PARTIDASs.Open()
                        Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'333','" & PERIODO & "','" & EJERCICIO & "','200013000000000000002','" & FECH & "','" + "IMPUESTO RETENIDO " & CVE_DOC + " ','H','" & Math.Round(IMP_TOT2, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & MIPAR1 + 1 & "','" & ccostos & "','" & cuentacon & "');"), INSERT_PARTIDASs)
                        commandDs.ExecuteNonQuery()
                        INSERT_PARTIDASs.Close()
                    End If
                End Using

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR1 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using 'conteo1
                '
                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR1 & "', ORDEN='" & MIPAR1 & "' WHERE NUM_PART =333 AND TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using


                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)

                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE INSOFTEC_COMPC_CLIB01 SET CAMPLIB1='" & TIPOCOMPD & contador & "' WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    totpartida = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using


                ''PARA OTRA PARTIDA
                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM " & PAR_COMPD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
                        adapter2231.Read()
                        UNO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) FROM " & PAR_COMPD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=0", connection22).ExecuteReader
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
                        Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES (" & PAR_COMPD & ",RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'333','" & PERIODO & "','" & EJERCICIO & "','410002000000000000002','" & FECH & "','" + "DESCUENTO " & CVE_DOC + " ','H','" & Math.Round(DESCCERO, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & MIPAR1 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs)
                        commandDs.ExecuteNonQuery()
                        INSERT_PARTIDASs.Close()
                    End If
                End Using

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR1 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using 'conteo1
                '
                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR1 & "', ORDEN='" & MIPAR1 & "' WHERE NUM_PART =333 AND TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA) FROM " & PAR_COMPD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
                        adapter2231.Read()
                        CINCO = adapter2231.Item(0)
                        connection22.Close()
                    End Using
                Catch EX As Exception
                End Try

                Try
                    Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        connection22.Open()
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT SUM(TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100)) - ((TOT_PARTIDA - (TOT_PARTIDA * (DESC1 / 100))) * (DESC2 / 100))) FROM " & PAR_COMPD & " WHERE CVE_DOC ='" & CVE_DOC & "' AND IMPU4=16", connection22).ExecuteReader
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
                        Dim commandDs As New SqlCommand(("INSERT INTO " & AUXILIAR & " (TIPO_POLI, NUM_POLIZ, NUM_PART, PERIODO, EJERCICIO, NUM_CTA, FECHA_POL, CONCEP_PO, DEBE_HABER, MONTOMOV, NUMDEPTO,TIPCAMBIO,CONTRAPAR,ORDEN, CCOSTOS,CGRUPOS) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'444','" & PERIODO & "','" & EJERCICIO & "','410001000000000000002','" & FECH & "','" + "DESCUENTO " & CVE_DOC + " ','H','" & Math.Round(DESC16, 2) & "','" & DEPTO & "','" & TIPCAMB & "','0','" & MIPAR2 + 1 & "','" & ccostos & "'," & cuentacon & ");"), INSERT_PARTIDASs)
                        commandDs.ExecuteNonQuery()
                        INSERT_PARTIDASs.Close()
                    End If
                End Using


                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT count(NUM_PART)  FROM " & AUXILIAR & " WHERE TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' AND NUM_POLIZ=" & contador & "", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    MIPAR2 = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using 'conteo2

                Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_CAMP.Open()
                    Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART='" & MIPAR2 & "', ORDEN='" & MIPAR2 & "' WHERE NUM_PART =444 AND TIPO_POLI='" & TIPOCOMPD & "' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), INSERT_CAMP)
                    command666.ExecuteNonQuery()
                    INSERT_CAMP.Close()
                End Using

                Using INSERT_ENCABEZADO As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    INSERT_ENCABEZADO.Open()
                    Dim command5 As New SqlCommand(("INSERT INTO " & POLIZA & "  (TIPO_POLI,NUM_POLIZ,PERIODO,EJERCICIO,FECHA_POL,CONCEP_PO,NUM_PART,LOGAUDITA,CONTABILIZ,NUMPARCUA,TIENEDOCUMENTOS,ORIGEN) VALUES ('" & TIPOCOMPD & "',RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5),'" & PERIODO & "','" & EJERCICIO & "','" & FECH & "','" & "COMPRA " & CVE_DOC + "  " & CONCEPTO & "','" & totpartida & "','N','N','0','0','SAE-FAC-AUT');"), INSERT_ENCABEZADO) 'totpartida
                    command5.ExecuteNonQuery()
                    INSERT_ENCABEZADO.Close()
                End Using

                SU_REFER = Replace(SU_REFER, "F-", "")
                Try
                    Using connection222540 As FbConnection = New FbConnection(ConfigurationManager.ConnectionStrings.Item("FB").ToString)
                        connection222540.Open()
                        Dim adapter222540 As FbDataReader = New FbCommand("SELECT FECHA_EMISION, UUID FROM DATOSCFD WHERE FOLIO LIKE '" & SU_REFER & "%' AND RFC_EMISOR = '" & RFCPROVEEDOR & "'", connection222540).ExecuteReader
                        adapter222540.Read()
                        If adapter222540.HasRows = True Then
                            FECHA_CERT_DEV_COMPRA = adapter222540.Item(0).ToString
                            uuid2 = adapter222540.Item(1)
                        End If
                    End Using
                Catch ex As Exception

                End Try
                'Try
                '    Using connection222540 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                '        connection222540.Open()
                '        Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT UUID, FECHA_CERT FROM CFDI01 WHERE CVE_DOC ='" & CVE_DOC & "'", connection222540).ExecuteReader
                '        adapter222540.Read()
                '        If adapter222540.HasRows = True Then
                '            uuid2 = adapter222540.Item(0)
                '            FECHA_CERT = adapter222540.Item(1).ToString
                '        End If
                '    End Using
                'Catch ex As Exception

                'End Try

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
                    'CapturaXMLCasilla()
                    EnviarSinCasillaXML() 'se manda correo cuando no está habilitada la casilla de captura de comprobantes
                End If

                Dim FECH1 As String = FECHA_CERT_DEV_COMPRA.ToString("yyyy-MM-dd HH:mm:ss")
                Try
                    If uuid = 1 Or uuid = 0 Then
                        Using ALFA As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                            ALFA.Open()
                            'Dim command6661 As New SqlCommand(("INSERT INTO UUIDTIMBRES (NUMREG, UUIDTIMBRE, MONTO, SERIE, FOLIO, RFCEMISOR, RFCRECEPTOR, ORDEN, FECHA, TIPOCOMPROBANTE) VALUES ('" & NUMREG & "','" & uuid2 & "','" & IMPORTE & "','" & SERIE & "','" & FOLIO & "','" & RFCPROVEEDOR & "','NME031127CH3','1','" & FECH1 & "','1')"), ALFA)
                            'command6661.ExecuteNonQuery()
                            ALFA.Close()
                        End Using
                    End If
                Catch ex As Exception
                End Try

                If uuid = 1 Or uuid = 0 Then
                    Using ALFA As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        ALFA.Open()
                        'Dim command6661 As New SqlCommand(("UPDATE " & CONTROL & " SET CTUUIDCOMP = " & NUMREG & ""), ALFA)
                        'command6661.ExecuteNonQuery()
                        ALFA.Close()
                    End Using
                End If

                If uuid = 1 Or uuid = 0 Then
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        INSERT_CAMP.Open()
                        'Dim command666 As New SqlCommand(("UPDATE " & POLIZA & " SET UUID = '" & uuid2 & "', TIENEDOCUMENTOS=1 WHERE CONCEP_PO LIKE '%" & CVE_DOC & "%' AND TIPO_POLI='" & TIPOCOMPC & "'"), INSERT_CAMP)
                        'command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using
                End If
                If uuid2 = "" Then
                    EnviarUUID() 'Para mandar correo cuando no hay XML asociado al documento.
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE INSOFTEC_COMPC_CLIB01 SET CAMPLIB3='SIN XML'  WHERE CLAVE_DOC = '" & CVE_DOC & "'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using

                End If

                If uuid = 1 Or uuid = 0 Then
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                        INSERT_CAMP.Open()
                        'Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET IDUUID = '" & NUMREG & "' WHERE CONCEP_PO LIKE '%" & CVE_DOC & "%' AND CONCEP_PO LIKE '%" & CONCEPTO & "%' AND DEBE_HABER='H' AND TIPO_POLI='" & TIPOCOMPC & "' AND IDUUID=-1"), INSERT_CAMP)
                        'command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using
                End If
                '***PARA REVISAR LOS IMPORTES Y CUADRARLOS***
                Dim SUMAH As Double = 0.0
                Dim SUMAD As Double = 0.0
                Dim RESTAHD As Double = 0.0
                Using connT1 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connT1.Open()
                    Dim adaT1 As SqlDataReader = New SqlCommand("SELECT SUM(MONTOMOV) FROM " & AUXILIAR & " WHERE DEBE_HABER='H' AND NUM_POLIZ= RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Co' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT1).ExecuteReader
                    adaT1.Read()
                    If adaT1.HasRows = True Then
                        SUMAH = adaT1.Item(0).ToString
                    End If
                End Using
                Using connT2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                    connT2.Open()
                    Dim adaT2 As SqlDataReader = New SqlCommand("SELECT SUM(MONTOMOV) FROM " & AUXILIAR & " WHERE DEBE_HABER='D' AND NUM_POLIZ= RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Co' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT2).ExecuteReader
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
                            Dim adaT3 As SqlDataReader = New SqlCommand("SELECT TOP(1) MONTOMOV, NUM_PART FROM " & AUXILIAR & " WHERE DEBE_HABER='D' AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Co' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT3).ExecuteReader
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
                            Dim adaT3 As SqlDataReader = New SqlCommand("SELECT TOP(1) MONTOMOV, NUM_PART FROM " & AUXILIAR & " WHERE DEBE_HABER='H' AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Co' AND CONCEP_PO LIKE '%" & CVE_DOC & "%' ", connT3).ExecuteReader
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
                        Dim command666 As New SqlCommand(("UPDATE " & AUXILIAR & " SET MONTOMOV = '" & NUEVOMONTO & "'  WHERE NUM_PART=" & ITEMNNUMERO & " AND NUM_POLIZ=RIGHT('     ' + LTRIM(RTRIM(" & contador & ")), 5) AND TIPO_POLI='Co' AND CONCEP_PO LIKE '%" & CVE_DOC & "%'"), ConnReg)
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
                COM_TOT = 0 ''

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
                uuid2 = ""
                NUMREG = 0
                'FECH1 = ""
                SUMADESCUENTO16 = 0
                SUMADESCUENTOCERO = 0
                DESC16 = 0
                DESCCERO = 0
                MIPAR1 = 0
                MIPAR2 = 0
                MIPAR3 = 0
                MIPAR4 = 0

                i += 1

            Loop

            connection.Close()
        End Using
        contador1 = contador
    End Sub

End Class
