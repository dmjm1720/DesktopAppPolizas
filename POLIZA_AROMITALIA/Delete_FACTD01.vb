Imports System.Data.SqlClient
Imports System.Configuration


Public Class Delete_FACTD01
    Public Shared Sub Delete_CAMPLIB3()
        Using update_table As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("SAE").ToString)
            update_table.Open()
            Dim command6661 As New SqlCommand(("UPDATE INSOFTEC_FACTD_CLIB01 SET CAMPLIB1=NULL WHERE CAMPLIB3 IS NULL"), update_table)
            command6661.ExecuteNonQuery()
            update_table.Close()
        End Using
    End Sub

End Class
