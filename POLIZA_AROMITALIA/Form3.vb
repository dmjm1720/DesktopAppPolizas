Imports System.Data.SqlClient
Imports System.Configuration
Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.CORREOS_INSOFTECTableAdapter.Fill(Me.COI7DataSet.CORREOS_INSOFTEC)
        Catch ex As Exception
            MsgBox("No existe la base de datos, favor de crearla", vbInformation, "!Atención¡")
            Me.Close()
        End Try
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            INSERT_CAMP.Open()
            Dim command666 As New SqlCommand(("DELETE FROM [dbo].[CORREOS_INSOFTEC]  WHERE ID='" + TxtDelete.Text + "' "), INSERT_CAMP)
            command666.ExecuteNonQuery()
            INSERT_CAMP.Close()
        End Using
        Me.CORREOS_INSOFTECTableAdapter.Fill(Me.COI7DataSet.CORREOS_INSOFTEC)
        TxtDelete.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            INSERT_CAMP.Open()
            Dim command666 As New SqlCommand(("UPDATE [dbo].[CORREOS_INSOFTEC] SET TIPO='" + CboUpdate.Text + "', EMAIL='" + TxtMailUpdate.Text + "' WHERE ID='" + TxtIdUpdate.Text + "' "), INSERT_CAMP)
            command666.ExecuteNonQuery()
            INSERT_CAMP.Close()
        End Using

        Me.CORREOS_INSOFTECTableAdapter.Fill(Me.COI7DataSet.CORREOS_INSOFTEC)
    End Sub

    Private Sub BtnInsert_Click(sender As Object, e As EventArgs) Handles BtnInsert.Click
        Using INSERT_CAMP As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            INSERT_CAMP.Open()
            Dim command666 As New SqlCommand(("INSERT INTO [dbo].[CORREOS_INSOFTEC] (TIPO, EMAIL) VALUES('" + CboInsert.Text + "', '" + TxtMailInsert.Text + "')"), INSERT_CAMP)
            command666.ExecuteNonQuery()
            INSERT_CAMP.Close()
        End Using
        Me.CORREOS_INSOFTECTableAdapter.Fill(Me.COI7DataSet.CORREOS_INSOFTEC)
    End Sub
End Class