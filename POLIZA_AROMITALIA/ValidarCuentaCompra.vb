Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Imports POLIZA_AROMITALIA.CorreoSinCuentaContable
Imports POLIZA_AROMITALIA.ConsultaTablaDatos
Imports POLIZA_AROMITALIA.CorreoProducto
Imports POLIZA_AROMITALIA.CorreoSCCCompra
Imports POLIZA_AROMITALIA.RevisionXML
Imports POLIZA_AROMITALIA.EnviarCorreoSinUUID
Imports POLIZA_AROMITALIA.CorreoSinCasilla
Imports POLIZA_AROMITALIA.SinCasillaXML
Imports POLIZA_AROMITALIA.ProdSinCuentaCompra
Imports FirebirdSql.Data.FirebirdClient

Public Class ValidarCuentaCompra

    Public Shared Sub CuentaCompra()
        Dim num As Integer = 0
        Dim LEVEL1 As String = ""
        Dim LEVEL2 As String = ""
        Dim LEVEL3 As String = ""
        Dim CADENA As String = ""

        'CONSUTA A LA TABLA COMPC01

        Using connection As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            Dim adapter As New SqlDataAdapter("SELECT " & COMPC & ".TIP_DOC, " & COMPC & ".CVE_DOC, " & COMPC & ".CVE_CLPV, " & COMPC & ".STATUS, " & COMPC & ".SU_REFER, " & COMPC & ".FECHA_DOC, " & COMPC & ".FECHA_REC, " & COMPC & ".FECHA_PAG, " & COMPC & ".CAN_TOT, " & COMPC & ".IMP_TOT1, " & COMPC & ".IMP_TOT4, " & COMPC & ".CVE_OBS, " & COMPC & ".NUM_ALMA, " & COMPC & ".ACT_CXP, " & COMPC & ".ACT_COI, " & COMPC & ".ENLAZADO, " & COMPC & ".TIP_DOC_E, " & COMPC & ".NUM_MONED, " & COMPC & ".TIPCAMB, " & COMPC & ".NUM_PAGOS, " & COMPC & ".CTLPOL, " & COMPC & ".ESCFD,  " & COMPC & ".SERIE, " & COMPC & ".FOLIO, " & COMPC & ".CONTADO, " & COMPC & ".BLOQ, " & COMPC & ".FORMAENVIO, " & COMPC & ".DES_FIN_PORC, " & COMPC & ".DES_TOT_PORC, " & COMPC & ".IMPORTE, " & COMPC & ".TIP_DOC_ANT, " & COMPC & ".DOC_ANT FROM  " & COMPC & " INNER JOIN INSOFTEC_COMPC_CLIB01 On " & COMPC & ".CVE_DOC = INSOFTEC_COMPC_CLIB01.CLAVE_DOC WHERE (INSOFTEC_COMPC_CLIB01.CAMPLIB1 Is NULL) And (" & COMPC & ".STATUS <> 'C')", connection)

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
                IMP_TOT4 = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(10)
                CVE_OBS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(11)
                NUM_ALMA = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(12)
                ACT_CXP = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(13).ToString
                ACT_COI = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(14).ToString
                ENLAZADO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(15).ToString
                TIP_DOC_E = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(16).ToString
                NUM_MONED = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(17)
                TIPCAMB = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(18)
                NUM_PAGOS = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(19)
                CTLPOL = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(20)
                ESCFD = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(21).ToString
                SERIE = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(22).ToString
                FOLIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(23)
                CONTADO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(24).ToString
                BLOQ = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(25).ToString
                FORMAENVIO = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(26).ToString
                DES_FIN_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(27)
                DES_TOT_PORC = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(28)
                IMPORTE = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(29)
                TIP_DOC_ANT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(30).ToString
                DOC_ANT = dataTable.AsEnumerable().ElementAtOrDefault(i).Item(31).ToString

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
                    Dim adapter2 As SqlDataReader = New SqlCommand("Select max(NUM_POLIZ) from " & POLIZA & " WHERE PERIODO ='" & PERIODO & "' AND EJERCICIO='" & EJERCICIO & "' AND TIPO_POLI='Ve'", connection2).ExecuteReader
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

                Dim FECH As String = FECHA_DOC.ToString("yyyy-MM-dd")
                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)

                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT COUNT(CVE_DOC) FROM " & PAR_COMPC & " WHERE (CVE_DOC='" & CVE_DOC & "')", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    totpartida = adapter50.Item(0)
                    INSERT_TOTPART.Close()
                End Using

                ''''''''CONSULTA PARTIDAS
                Using BUSCAPARTIDAS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                    Dim ada1 As New SqlDataAdapter("SELECT NUM_PAR,CVE_ART,CANT,PREC,COST,IMPU4,IMP1APLA,IMP2APLA,IMP3APLA,ACT_INV,NUM_ALM,TIP_CAM,UNI_VENTA,TIPO_PROD,CVE_OBS,REG_SERIE,E_LTPD,TIPO_ELEM,NUM_MOV,TOT_PARTIDA FROM " & PAR_COMPC & " where CVE_DOC='" & CVE_DOC & "'", BUSCAPARTIDAS)
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
                                Dim adapter222540 As SqlDataReader = New SqlCommand("SELECT NUM_CTA, NOMBRE FROM " & CUENTA & " WHERE  (SUBSTRING(NUM_CTA, 1, 10) = '" + CUENTA_COI + "')", connection222540).ExecuteReader
                                adapter222540.Read()
                                If adapter222540.HasRows = True Then
                                    LIN_PROD2 = adapter222540.Item(0).ToString
                                    CONCEPTO2 = adapter222540.Item(1).ToString
                                Else
                                    LIN_PROD2 = "00000000000000000000"
                                    CONCEPTO2 = ""
                                End If

                            End Using


                            Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                                connection22.Open()
                                Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT CAMPLIB8 FROM " & INVE_CLIB & " WHERE CVE_PROD ='" & CVE_ART & "'", connection22).ExecuteReader
                                adapter2231.Read()
                                MYACCOUNT = adapter2231.Item(0).ToString
                                If MYACCOUNT <> "" Then
                                    LEVEL1 = MYACCOUNT.Substring(0, 3)
                                    LEVEL1 = Replace(LEVEL1, "0", "")
                                    LEVEL2 = MYACCOUNT.Substring(4, 3)
                                    LEVEL2 = Replace(LEVEL2, "0", "")
                                    LEVEL3 = MYACCOUNT.Substring(8, 3)
                                    LEVEL3 = Replace(LEVEL3, "0", "")
                                End If
                                MYACCOUNT = Replace(MYACCOUNT, "-", "")
                                connection22.Close()
                            End Using

                            If LEVEL1 <> "" And LEVEL2 = "" And LEVEL3 = "" Then
                                CADENA = "000000000001"
                            ElseIf LEVEL2 <> "" And LEVEL3 = "" Then
                                CADENA = "000000000002"
                            ElseIf LEVEL3 <> "" Then
                                CADENA = "000000000003"
                            End If

                            MYACCOUNT = MYACCOUNT + CADENA
                            'If MYACCOUNT > "610000" And MYACCOUNT < "620037" Then
                            '    MYACCOUNT = MYACCOUNT + "000000000002"
                            'Else
                            '    MYACCOUNT = MYACCOUNT + "000000000003"
                            'End If
                            'If MYACCOUNT = "620022002" Then
                            '    MYACCOUNT = MYACCOUNT + "000000000003"
                            'End If

                            Using connection22 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                                connection22.Open()
                                Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT NUM_CTA FROM " & CUENTA & " WHERE NUM_CTA LIKE '%" & MYACCOUNT & "%'", connection22).ExecuteReader
                                If adapter2231.HasRows = True Then
                                    adapter2231.Read()
                                    MYACCOUNT2 = adapter2231.Item(0).ToString
                                    'MYACCOUNT2 = Replace(MYACCOUNT, "-", "")
                                    connection22.Close()
                                End If
                            End Using

                            'If MYACCOUNT2 >= "610000" And MYACCOUNT <= "620019" Then
                            '    MYACCOUNT2 = MYACCOUNT2 + "000000000002"
                            'Else
                            '    MYACCOUNT2 = MYACCOUNT2 + "000000000003"
                            'End If

                            'CENTRO DE COSTOS

                            'If MYACCOUNT = "" Then
                            '    ProductoSinCuenta()
                            'End If

                            'partidaS

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
                        Dim adapter2231 As SqlDataReader = New SqlCommand("SELECT CUENTA_CONTABLE, NOMBRE FROM " & PROV & " WHERE CLAVE ='" & CVE_CLPV & "'", connection22).ExecuteReader

                        adapter2231.Read()
                        CUENTA_CONTABLE = adapter2231.Item(0).ToString
                        NOMBRE_CLIENTE = adapter2231.Item(1).ToString
                        CUENTA_CONTABLE = Replace(CUENTA_CONTABLE, "-", "")
                        connection22.Close()
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

                End Using


                ''PARA CONTAR LAS PARTIDAS E INSERTARLAS EN LA TABLA POLIZAS1501

                If CUENTA_CONTABLE1 = "" Then
                    CUENTA_CONTABLE1 = ""
                End If

                Using INSERT_TOTPART As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)

                    INSERT_TOTPART.Open()
                    Dim adapter50 As SqlDataReader = New SqlCommand("SELECT CAMPLIB5, CAMPLIB6  FROM INSOFTEC_COMPC_CLIB01 WHERE CLAVE_DOC = '" & CVE_DOC & "'", INSERT_TOTPART).ExecuteReader
                    adapter50.Read()
                    CAMPLIB5 = adapter50.Item(0).ToString
                    C4 = adapter50.Item(1).ToString
                    If CUENTA_CONTABLE1 = "" And C4 = "" Then
                        SCCCompra()      'Favor de validar o capturar la cuenta contable en el campo CUENTA CONTABLE
                    End If

                    If MYACCOUNT <> MYACCOUNT2 And CAMPLIB5 = "" Then 'And CAMPLIB5 = "" Then
                        EnviarProdSCC()    'Favor de validar o capturar la cuenta contable del producto indicado, en el campo CUENTA COMPRAS en la ficha de Datos Adicionales.
                    End If
                    INSERT_TOTPART.Close()
                End Using


                If CUENTA_CONTABLE1 = "" Or CUENTA_CONTABLE1 = "100000000000000000001" Or MYACCOUNT <> MYACCOUNT2 Then

                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE INSOFTEC_COMPC_CLIB01 SET CAMPLIB4='ERROR CTA.', CAMPLIB5 ='SIN CTA.', CAMPLIB6='ENVIADO' WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using
                End If

                If CUENTA_CONTABLE1 <> "" And CUENTA_CONTABLE1 <> "100000000000000000001" And MYACCOUNT = MYACCOUNT2 Then
                    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                        INSERT_CAMP.Open()
                        Dim command666 As New SqlCommand(("UPDATE INSOFTEC_COMPC_CLIB01 SET  CAMPLIB4 = NULL, CAMPLIB5 = NULL, CAMPLIB6 = NULL WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                        command666.ExecuteNonQuery()
                        INSERT_CAMP.Close()
                    End Using
                End If

                'If MYACCOUNT <> "" Then
                '    Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                '        INSERT_CAMP.Open()
                '        Dim command666 As New SqlCommand(("UPDATE INSOFTEC_COMPC_CLIB01 SET CAMPLIB5 = NULL WHERE CLAVE_DOC='" & CVE_DOC & "'"), INSERT_CAMP)
                '        command666.ExecuteNonQuery()
                '        INSERT_CAMP.Close()
                '    End Using

                'End If



                NUM_PAR = 0
                C3 = ""
                C4 = ""
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
                NOMBRE_CLIENTE = ""
                CAMPLIB5 = ""
                LEVEL1 = ""
                LEVEL2 = ""
                LEVEL3 = ""
                CADENA = ""

                i += 1
            Loop
            connection.Close()
        End Using
        'RevisarXML()
    End Sub

End Class
