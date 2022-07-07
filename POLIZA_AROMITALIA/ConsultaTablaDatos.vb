Imports System.Data.SqlClient
Imports System.Configuration
Imports POLIZA_AROMITALIA.Form1
Public Class ConsultaTablaDatos
    'TABLAS PARA CONSULTAS
    Public Shared TCOMPRA As String = ""
    Public Shared COMPC As String = ""
    Public Shared TIPOCOMPC As String = ""
    Public Shared PAR_COMPC As String = ""
    Public Shared PARCOMPCNOMBRE As String = ""
    Public Shared EJERC As String = ""
    Public Shared EMPRESA As String = ""
    Public Shared POLIZA As String = ""
    Public Shared POLIZ As String = ""
    Public Shared AUXILIAR As String = ""
    Public Shared AUXI As String = ""
    Public Shared CUENTA As String = ""
    Public Shared CUENT As String = ""
    Public Shared FOLIOS As String = ""
    Public Shared NOMBREFOLIO As String = ""
    Public Shared NOMBREINVEN As String = ""
    Public Shared INVE As String = ""
    Public Shared NOMBREPROV As String = ""
    Public Shared PROV As String = ""
    Public Shared NOMBRECONTROL As String = ""
    Public Shared CONTROL As String = ""
    Public Shared NOMBREFACTF As String = ""
    Public Shared FACTF As String = ""
    Public Shared TIPOFACTF As String = ""
    Public Shared NOMBREPARTFACTF As String = ""
    Public Shared PAR_FACTF As String = ""
    Public Shared NOMBRECLIENTE As String = ""
    Public Shared CLIE As String = ""
    Public Shared NOMBREFACTDEV As String = ""
    Public Shared FACTD As String = ""
    Public Shared NOMBREPARFACTDEV As String = ""
    Public Shared PAR_FACTD As String = ""
    Public Shared TIPOFACTD As String = ""
    Public Shared NOMBRECDFI As String = ""
    Public Shared CFDI As String = ""
    Public Shared CFDII As String = ""
    Public Shared NOMBRECFDII As String = ""
    Public Shared NOMBREINVECLIB As String = ""
    Public Shared INVE_CLIB As String = ""
    Public Shared NOMBRECOMPRADEV As String = ""
    Public Shared COMPD As String = ""
    Public Shared TIPOCOMPD As String = ""
    Public Shared NOMBREPARCOMPDEV As String = ""
    Public Shared PAR_COMPD As String = ""


    Public Shared Sub ConsultaTablas()
        Using AUX As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            AUX.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EJERCICIO, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=1", AUX).ExecuteReader
            ada.Read()
            AUXI = ada.Item(0)
            EJERC = ada.Item(1)
            EMPRESA = ada.Item(2).ToString
            'AUXILIAR = AUXI & EJERC & EMPRESA
            AUX.Close()
        End Using

        Using POL As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            POL.Open()
            Dim ada As SqlDataReader = New SqlCommand("Select NOMBRE, EJERCICIO, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=2", POL).ExecuteReader
            ada.Read()
            POLIZ = ada.Item(0)
            EJERC = ada.Item(1)
            EMPRESA = ada.Item(2).ToString
            'POLIZA = POLIZ & EJERC & EMPRESA
            POL.Close()
        End Using

        Using CUE As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            CUE.Open()
            Dim ada As SqlDataReader = New SqlCommand("Select NOMBRE, EJERCICIO, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=3", CUE).ExecuteReader
            ada.Read()
            CUENT = ada.Item(0)
            EJERC = ada.Item(1)
            EMPRESA = ada.Item(2).ToString
            'CUENTA = CUENT & EJERC & EMPRESA
            CUE.Close()
        End Using

        Using COMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            COMP.Open()
            Dim ada As SqlDataReader = New SqlCommand("Select NOMBRE, EMPRESA, TIPO FROM INSOFTEC_DATOS_TABLAS WHERE ID=4", COMP).ExecuteReader
            ada.Read()
            TCOMPRA = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            TIPOCOMPC = ada.Item(2)
            COMPC = TCOMPRA & EMPRESA
            COMP.Close()
        End Using

        Using PAR_COMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            PAR_COMP.Open()
            Dim ada As SqlDataReader = New SqlCommand("Select NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=5", PAR_COMP).ExecuteReader
            ada.Read()
            PARCOMPCNOMBRE = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            PAR_COMPC = PARCOMPCNOMBRE & EMPRESA
        End Using

        Using FOLI As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            FOLI.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=6", FOLI).ExecuteReader
            ada.Read()
            NOMBREFOLIO = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            FOLIOS = NOMBREFOLIO & EMPRESA
            FOLI.Close()
        End Using

        Using INV As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            INV.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=7", INV).ExecuteReader
            ada.Read()
            NOMBREINVEN = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            INVE = NOMBREINVEN & EMPRESA
            INV.Close()
        End Using

        Using PROVE As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            PROVE.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=8", PROVE).ExecuteReader
            ada.Read()
            NOMBREPROV = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            PROV = NOMBREPROV & EMPRESA
            PROVE.Close()
        End Using

        Using CONTR As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            CONTR.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=9", CONTR).ExecuteReader
            ada.Read()
            NOMBRECONTROL = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            CONTROL = NOMBRECONTROL & EMPRESA
            CONTR.Close()
        End Using

        Using FACTR As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            FACTR.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA, TIPO FROM INSOFTEC_DATOS_TABLAS WHERE ID=10", FACTR).ExecuteReader
            ada.Read()
            NOMBREFACTF = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            TIPOFACTF = ada.Item(2)
            FACTF = NOMBREFACTF & EMPRESA
            FACTR.Close()
        End Using
        Using PARFACTF As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            PARFACTF.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=11", PARFACTF).ExecuteReader
            ada.Read()
            NOMBREFACTF = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            PAR_FACTF = NOMBREFACTF & EMPRESA
            PARFACTF.Close()
        End Using

        Using CLIENTES As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            CLIENTES.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=12", CLIENTES).ExecuteReader
            ada.Read()
            NOMBRECLIENTE = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            CLIE = NOMBRECLIENTE & EMPRESA
            CLIENTES.Close()
        End Using



        Using FACDEV As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            FACDEV.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA, TIPO FROM INSOFTEC_DATOS_TABLAS WHERE ID=13", FACDEV).ExecuteReader
            ada.Read()
            NOMBREFACTDEV = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            TIPOFACTD = ada.Item(2)
            FACTD = NOMBREFACTDEV & EMPRESA
            FACDEV.Close()
        End Using


        Using PARTFACDEV As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            PARTFACDEV.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=14", PARTFACDEV).ExecuteReader
            ada.Read()
            NOMBREPARFACTDEV = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            PAR_FACTD = NOMBREPARFACTDEV & EMPRESA
            PARTFACDEV.Close()
        End Using

        Using MICFDI As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            MICFDI.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=15", MICFDI).ExecuteReader
            ada.Read()
            NOMBRECDFI = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            CFDI = NOMBRECDFI & EMPRESA
            MICFDI.Close()
        End Using



        Using MICFDII As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            MICFDII.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=19", MICFDII).ExecuteReader
            ada.Read()
            NOMBRECFDII = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            CFDII = NOMBRECFDII & EMPRESA
            MICFDII.Close()
        End Using

        Using INVECLI As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            INVECLI.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=16", INVECLI).ExecuteReader
            ada.Read()
            NOMBREINVECLIB = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            INVE_CLIB = NOMBREINVECLIB & EMPRESA
            INVECLI.Close()
        End Using

        Using COMPDEV As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            COMPDEV.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA, TIPO FROM INSOFTEC_DATOS_TABLAS WHERE ID=17", COMPDEV).ExecuteReader
            ada.Read()
            NOMBRECOMPRADEV = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            TIPOCOMPD = ada.Item(2)
            COMPD = NOMBRECOMPRADEV & EMPRESA
            COMPDEV.Close()
        End Using

        Using PARCOMDEV As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("COI").ToString)
            PARCOMDEV.Open()
            Dim ada As SqlDataReader = New SqlCommand("SELECT NOMBRE, EMPRESA FROM INSOFTEC_DATOS_TABLAS WHERE ID=18", PARCOMDEV).ExecuteReader
            ada.Read()
            NOMBREPARCOMPDEV = ada.Item(0)
            EMPRESA = ada.Item(1).ToString
            PAR_COMPD = NOMBREPARCOMPDEV & EMPRESA
            PARCOMDEV.Close()
        End Using
    End Sub
End Class
