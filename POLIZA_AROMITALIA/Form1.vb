Imports System.Data
Imports System.Configuration
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Data.Odbc
Imports System.Timers
Imports System
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.CancelarPoliza 'Para importar el método de cancelar pólizas ''Public Shared Sub FacturaPolizaCancelda()''
Imports POLIZA_AROMITALIA.ProcesarFacturas 'Para importar el método de procesar pólizas nuevas 'Public Shared Sub ProcesarPolizaNueva()
Imports POLIZA_AROMITALIA.CorreoProducto 'Para importar el método para envío de correos cuando el producto no lleva cuenta
Imports POLIZA_AROMITALIA.ValidacionCuenta 'Para importar el método para detectar si el cliente no tiene cuenta contable
Imports POLIZA_AROMITALIA.ProcesarDevolucion 'Para importar el método para procesar las devoluciones sobre venta
Imports POLIZA_AROMITALIA.CancelarPolizaDevolucion 'Para importar el método para procesar las devoluciones sobre venta canceladas
Imports POLIZA_AROMITALIA.ValidarCuentaCompra 'Para importar el método para procesar las cuentas de las compras
Imports POLIZA_AROMITALIA.ProcesarCompra 'Para importar el módulo para procesar las compras
Imports POLIZA_AROMITALIA.RevisionXML 'Para importar el módulo para revisar la captura de documentos en la base de datos de Firebird
Imports POLIZA_AROMITALIA.CancelarCompra 'Para importar el módulo de cancelar compras
Imports POLIZA_AROMITALIA.ConsultaTablaDatos 'para importar el módulo de las consultas a las tablas de consultas de datos
Imports POLIZA_AROMITALIA.ProcesarDevolucionCompra ' para importar el módulo de devoluciones sobre compra
Imports POLIZA_AROMITALIA.CancelarDevolucionCompra ' para importar el módulo de cancelar devs sobre compra
Imports POLIZA_AROMITALIA.ValidarCuentaPartidas '
Imports POLIZA_AROMITALIA.Delete_FACTD01
Imports POLIZA_AROMITALIA.ProcesarNewFactura
Imports POLIZA_AROMITALIA.ProcesarNewDevolucion
Imports POLIZA_AROMITALIA.NuevasEntradas
Imports POLIZA_AROMITALIA.NuevasSalidas
Imports POLIZA_AROMITALIA.CuentaEntrada
Imports POLIZA_AROMITALIA.InsertarSignoMas
Imports POLIZA_AROMITALIA.InsertarSignoMenos
Imports POLIZA_AROMITALIA.CuentaSalida
Imports POLIZA_AROMITALIA.CambiosVentas
Imports POLIZA_AROMITALIA.CambiosDevoluciones
Imports POLIZA_AROMITALIA.CambiosCompras
Imports POLIZA_AROMITALIA.CambioEntradas
Imports POLIZA_AROMITALIA.CambioSalidas
Imports POLIZA_AROMITALIA.Facte01NC
Imports POLIZA_AROMITALIA.Facte01NCCancelar



Public Class Form1
    Public Shared TIP_DOC As String = ""
    Public Shared CVE_DOC As String = ""
    Public Shared CVE_CLPV As String = ""
    Public Shared STATUS As String = ""
    Public Shared DAT_MOSTR As DateTime
    Public Shared CVE_VEND As String = ""
    Public Shared CVE_PEDI As String = ""
    Public Shared FECHA_DOC As DateTime
    Public Shared FECHA_ENT As DateTime
    Public Shared FECHA_VEN As DateTime
    Public Shared CAN_TOT As Double = 0
    Public Shared IMP_TOT1 As Double = 0
    Public Shared IMP_TOT2 As Double = 0
    Public Shared IMP_TOT3 As Double = 0
    Public Shared IMP_TOT4 As Double = 0
    Public Shared DES_TOT As Double = 0
    Public Shared DES_FIN As Double = 0
    Public Shared COM_TOT As Double = 0
    Public Shared CONDICION As String = ""
    Public Shared CVE_OBS As Integer = 0
    Public Shared NUM_ALMA As Integer = 0
    Public Shared ACT_CXC As String = ""
    Public Shared ACT_COI As String = ""
    Public Shared ENLAZADO As String = ""
    Public Shared TIP_DOC_E As String = ""
    Public Shared NUM_MONED As Integer = 0
    Public Shared TIPCAMB As Double = 0
    Public Shared NUM_PAGOS As Integer = 0
    Public Shared FECHAELAB As DateTime
    Public Shared PRIMERPAGO As Double = 0
    Public Shared RFC As String = ""
    Public Shared CTLPOL As Integer = 0
    Public Shared ESCFD As String = ""
    Public Shared AUTORIZA As Integer = 0
    Public Shared SERIE As String
    Public Shared FOLIO As Integer = 0
    Public Shared AUTOANIO As String = ""
    Public Shared DAT_ENVIO As Integer = 0
    Public Shared CONTADO As String = ""
    Public Shared CVE_BITA As Integer = 0
    Public Shared BLOQ As String = ""
    Public Shared FORMAENVIO As String = ""
    Public Shared DES_FIN_PORC As Double = 0
    Public Shared DES_TOT_PORC As Double = 0
    Public Shared IMPORTE As Double = 0
    Public Shared COM_TOT_PORC As Double = 0
    Public Shared COM_TOT_PORC1 As String = ""
    Public Shared TIP_DOC_ANT As String = ""
    Public Shared DOC_ANT As String = ""
    Public Shared contador As Integer = 0
    Public Shared contador1 As Integer = contador
    Public Shared cont As Integer = 0
    Public Shared menos As Integer = 0
    Public Shared totpartida As Integer
    Public Shared FECHA_REC As DateTime
    Public Shared FECHA_PAG As DateTime
    Public Shared ACT_CXP As String = ""


    'partidas variables
    Public Shared NUM_PAR As Integer = 1
    Public Shared CVE_ART As String = ""
    Public Shared CANT As Double = 0
    Public Shared PXS As Double = 0
    Public Shared PREC As Double = 0
    Public Shared COST As Double = 0
    Public Shared IMPU4 As Double = 0
    Public Shared IMP1APLA As String = ""
    Public Shared IMP2APLA As String = ""
    Public Shared IMP3APLA As String = ""
    Public Shared APAR As Double = 0
    Public Shared ACT_INV As String = ""
    Public Shared NUM_ALM As Integer = 0
    Public Shared TIP_CAM As Double = 0
    Public Shared UNI_VENTA As String = ""
    Public Shared TIPO_PROD As String = ""
    Public Shared CVE_OBSPAR As Integer = 0
    Public Shared REG_SERIE As Integer = 0
    Public Shared E_LTPD As Integer = 0
    Public Shared TIPO_ELEM As String = ""
    Public Shared NUM_MOV As Integer = 0
    Public Shared DESC1 As Double = 0
    Public Shared TOT_PARTIDA As Double = 0
    Public Shared IMPRIMIR As String = ""
    Public Shared CONCEPTO As String
    Public Shared CONCEPTO2 As String
    Public Shared CONCEPTO3 As String
    Public Shared CONCEPTIO As String
    Public Shared CUENTA_CONTABLE As String
    Public Shared CUENTA_CONTABLE1 As String
    Public Shared DEF1 As String
    Public Shared ccostos As String = ""
    Public Shared TOTIMP41 As Double
    Public Shared IMP_RET As Double = 0
    Public Shared IMP_RET2 As Double = 0
    Public Shared COMPRAFOLIO As String


    Public Shared DEPTO As String = ""
    Public Shared NUM_PAR1 As Integer = 0
    Public Shared NUM_PAR2 As Integer = 0
    Public Shared NUM_PAR3 As Integer = 0

    Public Shared LIN_PROD As String = ""
    Public Shared DESC_LIN As String = ""
    Public Shared LIN_PROD2 As String = ""
    Public Shared CUENTA_COI As String = ""
    Public Shared DEFAULT2 As String = ""

    Public Shared EJERCICIO As Integer
    Public Shared PERIODO As Integer
    Public Shared PERIODO1 As String

    Public Shared cuentacon As String = ""

    Public Shared CVE_DOCNUM_PARCVE_ART = ""
    Public Shared DESC2 As Double = 0
    Public Shared tot_part = 0
    Public Shared TOTIMP4 = 0
    Public Shared tot_part1 = 0
    Public Shared IMP4APLA = 0
    Public Shared COMI = 0
    Public Shared POLIT_APLI = 0
    Public Shared TOT_PARTIDAIMPRIMIR = 0
    Public Shared DEF2 = ""
    Public Shared CTRL = ""
    Public Shared MIDESCR = ""
    Public Shared DATO As String = ""
    Public Shared C3 As String = ""
    Public Shared C4 As String = ""
    Public Shared uuid As Integer = 0
    Public Shared uuid2 As String
    Public Shared uuid3 As String
    Public Shared uuid4 As String
    Public Shared NUMREG As Integer = 0
    Public Shared FECHA_CERT_FACTURA As DateTime
    Public Shared FECHA_CERT_DEVOLUCION As DateTime
    Public Shared FECHA_CERT_COMPRA As DateTime
    Public Shared FECHA_CERT_DEV_COMPRA As DateTime

    'datos_correo
    Public Shared CORREO_EMISOR As String = ""
    Public Shared PASSWORD_CORREO As String = ""
    Public Shared EMAIL As String = ""
    Public Shared TIPO As String = ""
    Public Shared COMENTARIO3 As String = ""
    Public Shared SALUDO As String = ""
    Public Shared NOMBRE_EMPRESA As String = ""
    Public Shared ASUNTO As String = ""
    Public Shared CORREO_RECEPTOR As String = ""
    Public Shared CORREO_RECEPTOR1 As String = ""
    Public Shared CORREO_RECEPTOR2 As String = ""
    Public Shared NOMBRE_HOST As String = ""
    Public Shared PUERTO As Integer = 0
    Public Shared MICUENTA As String = ""
    Public Shared SUMADESCUENTOCERO As Double = 0
    Public Shared SUMADESCUENTO16 As Double = 0
    Public Shared DESCCERO As Double = 0
    Public Shared DESC16 As Double = 0
    Public Shared UNO As Double = 0
    Public Shared DOS As Double = 0
    Public Shared TRES As Double = 0
    Public Shared CUATRO As Double = 0
    Public Shared M1 As Double = 0
    Public Shared CINCO As Double = 0
    Public Shared SEIS As Double = 0
    Public Shared SIETE As Double = 0
    Public Shared OCHO As Double = 0
    Public Shared M2 As Double = 0
    Public Shared MIPAR1 As Integer = 0
    Public Shared MIPAR2 As Integer = 0
    Public Shared MIPAR3 As Integer = 0
    Public Shared MIPAR4 As Integer = 0
    Public Shared MIPAR5 As Integer = 0
    Public Shared NOMBRE_CLIENTE As String = ""
    Public Shared correo As New MailMessage
    Public Shared envio As New SmtpClient
    Public Shared CONCATENA3 As String = ""
    Public Shared CONCATENA2 As String = ""
    Public Shared RFCPROVEEDOR As String = ""
    Public Shared SU_REFER As String = ""
    Public Shared MYACCOUNT As String = ""
    Public Shared MYACCOUNT2 As String = ""
    Public Shared CAMPLIB3 As String = ""
    Public Shared CAMPLIB5 As String
    Public Shared letra As String
    Public Shared CUENTA_PARTIDA As String
    Public Shared con1 As Integer
    Public Shared con2 As Integer

    Public Shared FOLIO_TIMBRADO As String = ""
    Public Shared SERIE_TIMBRADO As String = ""
    Public Shared CAMPLIB8 As String = ""
    Public Shared CAMPLIB9 As String = ""
    Public Shared CAMPLIB10 As String = ""
    Public Shared CAMPLIB11 As String = ""
    Public Shared CAMPLIB12 As String = ""
    Public Shared CAMPLIB13 As String = ""
    Public Shared NUEVOIVA As String = ""
    Public Shared NUEVOIVA2 As String = ""
    Public Shared NUEVOIVA3 As String = ""
    Public Shared NUEVOIVA4 As String = ""


    Public Shared CVEART As String = ""
    Public Shared NUMMOV As String = ""
    Public Shared CVECPTO As String = ""
    Public Shared FECHADOCU As DateTime
    Public Shared TIPODOC As String = ""
    Public Shared REFERMIN As String = ""
    Public Shared CLAVECLPV As String = ""
    Public Shared CANTMIN As Double = 0
    Public Shared CANTCOST As Double = 0
    Public Shared PRECIOMIN As Double = 0
    Public Shared COSTOMIN As Double = 0
    Public Shared UNIVENTA As String = ""
    Public Shared FECHAELABMIN As DateTime
    Public Shared CVEFOLIO As String = ""
    Public Shared SIGNO As Integer = 0
    Public Shared FOLIOMINVE01 As String = ""
    Public Shared FECH As String
    Public Shared partidasNo As Integer = 0
    Public Shared H As Integer = 0
    Public Shared D As Integer = 0

    Public Shared CE As Integer = 0
    Public Shared CS As Integer = 0
    Public Shared CAML4 As String = ""
    Public Shared MINDIRECTO As Double = 0
    Public Shared TOT_IND As Double = 0

    Public Shared contarPartidas As Integer = 0
    Public Shared contarPartidas2 As Integer = 0

    Private counter As Integer
    Private counter2 As Integer
    Private counter3 As Integer
    Private counter4 As Integer
    Private counter5 As Integer
    Private counter6 As Integer
    Private counter7 As Integer
    Private counter8 As Integer
    Private counter9 As Integer
    Private counter10 As Integer
    Private counter11 As Integer
    Private counter12 As Integer


    Public Sub InitializeTimer()
        counter = 0
        Timer1.Interval = 30000
        Timer1.Enabled = True
        Label30.Text = "Espere por favor"
    End Sub

    Public Sub StopTimer1()
        Timer1.Enabled = False
        Label13.Show()
        Label14.Hide()
    End Sub
    Public Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If counter < 0 Then
            Timer1.Enabled = False
            counter = 0
        Else
            ValidacionCuentaContable()
            NewFactura()
            ActualizarCambiosVentas() 'Procesar los cambios en los precios
            'ProcesarPolizaNueva() 'facturas
            counter = counter + 1
            Label30.Text = "" & counter.ToString
            Label13.Hide()
            Label14.Show()
            If counter = 1000 Then
                counter = 0
            End If
        End If

    End Sub
    Public Sub InitializeTimer2()
        counter2 = 0
        Timer2.Interval = 30000
        Timer2.Enabled = True
        Label31.Text = "Espere por favor"
    End Sub

    Public Sub StopTimer2()
        Timer2.Enabled = False
        Label16.Show()
        Label17.Hide()
    End Sub
    Public Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If counter2 < 0 Then
            Timer2.Enabled = False
            counter2 = 0
        Else
            FacturaPolizaCancelda()
            counter2 = counter2 + 1
            Label31.Text = "" & counter2.ToString
            Label16.Hide()
            Label17.Show()
            If counter2 = 1000 Then
                counter2 = 0
        End If
        End If

    End Sub

    Public Sub InitializeTimer3()
        counter3 = 0
        Timer3.Interval = 30000
        Timer3.Enabled = True
        Label32.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer3()
        Timer3.Enabled = False
        Label18.Show()
        Label19.Hide()
    End Sub
    Public Sub Timer3_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If counter3 < 0 Then
            Timer3.Enabled = False
            counter3 = 0
        Else
            Delete_CAMPLIB3()
            NewDevolucion()
            ActualizarCambiosDevoluciones() 'Procesar los cambios de precio en las devoluciones
            'DevolucionPoliza()
            counter3 = counter3 + 1
            Label32.Text = "" & counter3.ToString
            Label18.Hide()
            Label19.Show()
            If counter3 = 1000 Then
                counter3 = 0
                'ProgressBar3.Value = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer4()
        counter4 = 0
        Timer4.Interval = 30000
        Timer4.Enabled = True
        Label33.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer4()
        Timer4.Enabled = False
        Label20.Show()
        Label21.Hide()
    End Sub
    Public Sub Timer4_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        If counter4 < 0 Then
            Timer4.Enabled = False
            counter4 = 0
        Else
            FacturaPolizaCanceldaDevolucion()
            counter4 = counter4 + 1
            Label33.Text = "" & counter4.ToString
            Label20.Hide()
            Label21.Show()
            If counter4 = 1000 Then
                counter4 = 0
                ' ProgressBar4.Value = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer5()
        counter5 = 0
        Timer5.Interval = 30000
        Timer5.Enabled = True
        Label34.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer5()
        Timer5.Enabled = False
        Label22.Show()
        Label23.Hide()
    End Sub
    Public Sub Timer5_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        If counter5 < 0 Then
            Timer5.Enabled = False
            counter5 = 0
        Else
            CuentaCompra()
            ProcesarNuevaCompra()
            RevisarXML()
            NuevosCambiosCompras() 'Procesar los cambios de precios en las compras
            counter5 = counter5 + 1
            Label34.Text = "" & counter5.ToString
            Label22.Hide()
            Label23.Show()
            If counter5 = 1000 Then
                counter5 = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer6()
        counter6 = 0
        Timer6.Interval = 30000
        Timer6.Enabled = True
        Label35.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer6()
        Timer6.Enabled = False
        Label24.Show()
        Label25.Hide()
    End Sub
    Public Sub Timer6_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer6.Tick
        If counter6 < 0 Then
            Timer6.Enabled = False
            counter6 = 0
        Else
            CompraCancelada()
            counter6 = counter6 + 1
            Label35.Text = "" & counter6.ToString
            Label24.Hide()
            Label25.Show()
            If counter6 = 1000 Then
                counter6 = 0
            End If
        End If
    End Sub

    Public Sub InitializeTimer7()
        counter7 = 0
        Timer7.Interval = 30000
        Timer7.Enabled = True
        Label36.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer7()
        Timer7.Enabled = False
        Label26.Show()
        Label27.Hide()
    End Sub
    Public Sub Timer7_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer7.Tick
        If counter7 < 0 Then
            Timer7.Enabled = False
            counter7 = 0
        Else
            DevolucionCompra()
            counter7 = counter7 + 1
            Label36.Text = "" & counter7.ToString
            Label26.Hide()
            Label27.Show()
            If counter7 = 1000 Then
                counter7 = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer8()
        counter8 = 0
        Timer8.Interval = 30000
        Timer8.Enabled = True
        Label37.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer8()
        Timer8.Enabled = False
        Label28.Show()
        Label29.Hide()
    End Sub
    Public Sub Timer8_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer8.Tick
        If counter8 < 0 Then
            Timer8.Enabled = False
            counter8 = 0
        Else
            CancelarPolizaSV()
            counter8 = counter8 + 1
            Label37.Text = "" & counter8.ToString
            Label28.Hide()
            Label29.Show()
            If counter8 = 1000 Then
                counter8 = 0
            End If
        End If
    End Sub

    Public Sub InitializeTimer9()
        counter9 = 0
        Timer9.Interval = 30000
        Timer9.Enabled = True
        Label38.Text = "Espere por favor"
    End Sub
    Public Sub StopTimer9()
        Timer9.Enabled = False
        Label39.Show()
        Label40.Hide()
    End Sub
    Public Sub Timer9_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer9.Tick
        If counter9 < 0 Then
            Timer9.Enabled = False
            counter9 = 0
        Else
            InsertMas()
            AccountIn()
            Entradas()
            NuevoCambioEntrada()
            counter9 = counter9 + 1
            Label38.Text = "" & counter9.ToString
            Label39.Hide()
            Label40.Show()
            If counter9 = 1000 Then
                counter9 = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer10()
        counter10 = 0
        Timer10.Interval = 30000
        Timer10.Enabled = True
        Label41.Text = "Espera por favor"
    End Sub

    Public Sub StopTimer10()
        Timer10.Enabled = False
        Label42.Show()
        Label43.Hide()
    End Sub
    Public Sub Timer10_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer10.Tick
        If counter10 < 0 Then
            Timer10.Enabled = False
            counter10 = 0
        Else
            InsertMenos()
            AccountOut()
            Salidas()
            NuevoCambioSalida()
            counter10 = counter10 + 1
            Label41.Text = "" & counter10.ToString
            Label42.Hide()
            Label43.Show()
            If counter10 = 1000 Then
                counter10 = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer11()
        counter11 = 0
        Timer11.Interval = 30000
        Timer11.Enabled = True
        Label48.Text = "Espera por favor"
    End Sub
    Public Sub StopTimer11()
        Timer11.Enabled = False
        Label49.Show()
        Label50.Hide()
    End Sub
    Private Sub Timer11_Tick(sender As Object, e As EventArgs) Handles Timer11.Tick
        If counter11 < 0 Then
            Timer11.Enabled = False
            counter11 = 0

        Else
            'ProcesarFacte01()
            ProcesarFacte01NC()
            counter11 = counter11 + 1
            Label48.Text = "" & counter11.ToString
            Label49.Hide()
            Label50.Show()
            If counter11 = 1000 Then
                counter11 = 0
            End If
        End If
    End Sub
    Public Sub InitializeTimer12()
        counter12 = 0
        Timer12.Interval = 30000
        Timer12.Enabled = True
        Label51.Text = "Espera por favor"
    End Sub
    Public Sub StopTimer12()
        Timer12.Enabled = False
        Label52.Show()
        Label53.Hide()
    End Sub
    Private Sub Timer12_Tick(sender As Object, e As EventArgs) Handles Timer12.Tick
        If counter12 < 0 Then
            Timer12.Enabled = False
            counter12 = 0
        Else
            cancelarNC()
            counter12 = counter12 + 1
            Label51.Text = "" & counter12.ToString
            Label52.Hide()
            Label53.Show()
            If counter12 = 1000 Then
                counter12 = 0
            End If
        End If
    End Sub

    Public Sub cero()
        Using INSERT_FOLIO As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            INSERT_FOLIO.Open()
            Dim command661 As New SqlCommand(("UPDATE " & AUXILIAR & " SET NUM_PART=1 WHERE NUM_PART=0 AND TIPO_POLI='" & TIPOFACTD & "'"), INSERT_FOLIO)
            command661.ExecuteNonQuery()
            INSERT_FOLIO.Close()
        End Using
    End Sub

    Public Sub correo_prueba()
        Using cn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            Dim ada As New SqlDataAdapter("SELECT * FROM CORREOS_INSOFTEC", cn)
            Dim dt As New DataTable
            ada.Fill(dt)
            cn.Open()
            Dim w As Integer = (dt.Rows.Count)
            Dim i As Integer = 0
            Try
                Do While (i <= w)
                    TIPO = dt.AsEnumerable().ElementAtOrDefault(i).Item(1).ToString
                    EMAIL = dt.AsEnumerable().ElementAtOrDefault(i).Item(2).ToString
                    ProductoSinCuenta()
                    MsgBox("Se envío correo a: " + EMAIL + "    departamento: " + TIPO, vbInformation, "Correos de prueba")
                    i += 1
                Loop
            Catch ex As Exception
                Exit Try
            End Try
            cn.Close()
        End Using
    End Sub


    Sub crear_tabla_correos()

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("DROP TABLE [dbo].[CORREOS_INSOFTEC]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using

        Catch exp As Exception

        End Try

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("CREATE TABLE [dbo].[CORREOS_INSOFTEC]([ID] [int] IDENTITY(1,1) NOT NULL, [TIPO] [varchar](20) NOT NULL, [EMAIL] [varchar](150) NOT NULL, CONSTRAINT [PK_MAIL] PRIMARY KEY CLUSTERED ([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Sub crear_tabla_datos()
        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("DROP TABLE [dbo].[INSOFTEC_DATOS_TABLAS]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using

        Catch exp As Exception

        End Try

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("CREATE TABLE [dbo].[INSOFTEC_DATOS_TABLAS]([ID] [int] NOT NULL, [NOMBRE] [varchar](50) NOT NULL, [DESCRIPCION] [varchar](50) NOT NULL) ON [PRIMARY]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub crear_tablas_validaciones()
        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("DROP TABLE [dbo].[INSOFTEC_COMPC_CLIB01]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using

        Catch exp As Exception

        End Try

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("CREATE TABLE [dbo].[INSOFTEC_COMPC_CLIB01] ([CLAVE_DOC] [varchar](20) NOT NULL, [CAMPLIB1] [varchar](10) NULL, [CAMPLIB2] [varchar](10) NULL, [CAMPLIB3] [varchar](10) NULL, [CAMPLIB4] [varchar](10) NULL, [CAMPLIB5] [varchar](10) NULL, [CAMPLIB6] [varchar](10) NULL, CONSTRAINT [PK_INSOFTEC_COMPC_CLIB01] PRIMARY KEY CLUSTERED ([CLAVE_DOC] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("DROP TABLE [dbo].[INSOFTEC_FACTD_CLIB01]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using

        Catch exp As Exception

        End Try

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("CREATE TABLE [dbo].[INSOFTEC_FACTD_CLIB01]([CLAVE_DOC] [varchar](20) NOT NULL, [CAMPLIB1] [varchar](10) NULL, [CAMPLIB2] [varchar](1) NULL, [CAMPLIB3] [varchar](10) NULL, [CAMPLIB4] [varchar](10) NULL, [CAMPLIB5] [varchar](10) NULL, [CAMPLIB7] [varchar](10) NULL, [CAMPLIB6] [varchar](10) NULL, CONSTRAINT [PK_INSOFTEC_FACTD_CLIB01] PRIMARY KEY CLUSTERED ([CLAVE_DOC] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("DROP TABLE [dbo].[INSOFTEC_FACTF_CLIB01]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using

        Catch exp As Exception

        End Try

        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("CREATE TABLE [dbo].[INSOFTEC_FACTF_CLIB01] ([CLAVE_DOC] [varchar](20) NOT NULL, [CAMPLIB1] [varchar](10) NULL, [CAMPLIB2] [varchar](1) NULL, [CAMPLIB3] [varchar](10) NULL, [CAMPLIB4] [varchar](10) NULL, [CAMPLIB5] [varchar](10) NULL, [CAMPLIB6] [varchar](10) NULL, [CAMPLIB7] [varchar](10) NULL, CONSTRAINT [PK_INSOFTEC_FACTF_CLIB01] PRIMARY KEY CLUSTERED ([CLAVE_DOC] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Sub crear_triggers()
        Try
            Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
                INSERT_CAMP.Open()
                Dim command666 As New SqlCommand(("DROP TRIGGER [dbo].[UPDATE_COMPC_CLIB011]"), INSERT_CAMP)
                command666.ExecuteNonQuery()
                INSERT_CAMP.Close()
            End Using

        Catch exp As Exception

        End Try


    End Sub
    Private Sub CrearTablaParaCorreosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrearTablaParaCorreosToolStripMenuItem.Click
        If MsgBox("Se creará tabla de datos para consultas, deseas continuar...", MsgBoxStyle.YesNo, "Aviso importante") = MsgBoxResult.Yes Then 'if you select yes in the msgbox then it will close the window
            crear_tabla_datos()
            MsgBox("Se ha creado la tabla de datos para consultas", vbInformation, "¡Aviso importante!")
        Else
            MsgBox("Has cancelado la operación", vbCritical, "¡Aviso importante!")
        End If
    End Sub

    Private Sub IniciarSistemaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IniciarSistemaToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub CorreoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CorreoToolStripMenuItem.Click
        Form3.Show()
    End Sub

    Private Sub MandarCorreoDePruebaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MandarCorreoDePruebaToolStripMenuItem.Click
        correo_prueba()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConsultaTablas()

    End Sub

    Private Sub CrearTablaParaConsultasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrearTablaParaConsultasToolStripMenuItem.Click
        If MsgBox("Se creará tabla de correos, deseas continuar...", MsgBoxStyle.YesNo, "Aviso importante") = MsgBoxResult.Yes Then 'if you select yes in the msgbox then it will close the window
            crear_tabla_correos()
            MsgBox("Se ha creado la tabla de correos", vbInformation, "¡Aviso importante!")
        Else
            MsgBox("Has cancelado la operación", vbCritical, "¡Aviso importante!")
        End If
    End Sub

    Private Sub CrearTablaINSOFTECCOMPCCLIB01ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrearTablaINSOFTECCOMPCCLIB01ToolStripMenuItem.Click
        If MsgBox("Se creará tabla de validaciones, deseas continuar...", MsgBoxStyle.YesNo, "Aviso importante") = MsgBoxResult.Yes Then 'if you select yes in the msgbox then it will close the window
            'crear_tabla_correos()
            'crear_triggers()
            MsgBox("Se ha creado la tabla de validaciones", vbInformation, "¡Aviso importante!")
        Else
            MsgBox("Has cancelado la operación", vbCritical, "¡Aviso importante!")
        End If
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        Form2.Show()
    End Sub



End Class
