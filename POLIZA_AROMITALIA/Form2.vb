Imports System.Windows.Forms
Public Class Form2
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Call Form1.InitializeTimer()
            CheckBox1.Text = "Activado"
            MsgBox("Has activado el servicio de pólizas nuevas", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer1()
            CheckBox1.Text = "Desactivado"
            MsgBox("Servicio desactivado", vbInformation, "¡Aviso importante!")

        End If

    End Sub



    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Call Form1.InitializeTimer2()
            CheckBox2.Text = "Activado"
            MsgBox("Has activado el servicio de pólizas canceladas", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer2()
            CheckBox2.Text = "Desactivado"
            MsgBox("Servicio desactivado", vbInformation, "¡Aviso importante!")

        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            Call Form1.InitializeTimer3()
            CheckBox6.Text = "Activado"
            MsgBox("Has activado el servicio de devoluciones sobre venta", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer3()
            CheckBox6.Text = "Desactivado"
            MsgBox("Servicio desactivado", vbInformation, "¡Aviso importante!")

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked Then
            Call Form1.InitializeTimer4()
            CheckBox8.Text = "Activado"
            MsgBox("Has activado el servidio de devoluciones sobre venta canceladas", vbInformation, "!Aviso importante¡")
        Else
            Call Form1.StopTimer4()
            CheckBox8.Text = "Desactivado"
            MsgBox("Servicio desactivado", vbInformation, "!Aviso importante¡")

        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            Call Form1.InitializeTimer5()
            CheckBox3.Text = "Activado"
            MsgBox("Has activado el servicio de Compras", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer5()
            CheckBox3.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            Call Form1.InitializeTimer6()
            CheckBox4.Text = "Activado"
            MsgBox("Has activado el servicio de Compras", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer6()
            CheckBox4.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked Then
            Call Form1.InitializeTimer7()
            CheckBox7.Text = "Activado"
            MsgBox("Has activado el servicio de Devoluciones sobre Compras", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer7()
            CheckBox7.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            Call Form1.InitializeTimer8()
            CheckBox5.Text = "Activado"
            MsgBox("Has activado el servicio de Devoluciones sobre Compras canceladas", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer8()
            CheckBox5.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked Then
            Call Form1.InitializeTimer9()
            CheckBox9.Text = "Activado"
            MsgBox("Has activado el servicio de Movimientos al inventario", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer9()
            CheckBox9.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked Then
            Call Form1.InitializeTimer10()
            CheckBox10.Text = "Activado"
            MsgBox("Has activado el servicio de Movimientos al inventario", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer10()
            CheckBox10.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked Then
            Call Form1.InitializeTimer11()
            CheckBox11.Text = "Activado"
            MsgBox("Has activado el servicio de Notas de Crédito", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer11()
            CheckBox11.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked Then
            Call Form1.InitializeTimer12()
            CheckBox12.Text = "Activado"
            MsgBox("Has activado el servicio de Notas de Crédito Canceladas", vbInformation, "¡Aviso importante!")
        Else
            Call Form1.StopTimer12()
            CheckBox12.Text = "Desactivado"
            MsgBox("Servicio desactivado")
        End If
    End Sub
End Class