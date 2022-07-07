Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.Form1

Public Class EnviarCorreoSinUUID
    Public Shared Sub EnviarUUID()
        Using cn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            Dim adapter As New SqlDataAdapter("SELECT * FROM CORREOS_INSOFTEC WHERE TIPO='Compras'", cn)
            Dim dt As New DataTable
            adapter.Fill(dt)
            cn.Open()
            Dim y As Integer = (dt.Rows.Count)
            Dim i As Integer = 0
            Try
                Do While (i <= y)
                    EMAIL = dt.AsEnumerable().ElementAtOrDefault(i).Item(2).ToString
                    Dim emisor As String = "alertas@grupocomercialtria.com.mx"
                    Dim password As String = "QiVDZ90chqmp"
                    Dim mensaje As String = "<html><head><style type='text/css'><!--body,td,th {color: #000000;}--></style>" &
      "</head><body>" &
"<center><table width='985' border='0' align='left'><tr>" + "<br>" &
 " <p align='center' style='font-family:Calibri Light; font-size:20px'>[-Compra en SAE sin XML-] <br><br>Proveedor: (" & CVE_CLPV & " ) " & NOMBRE_CLIENTE & " <br><br>&rarr; La compra : " & CVE_DOC & " no tiene documento XML asociado. <br><br>Favor de asignar.<br><br>Gracias.<br><br></p>" &
"<th width='979' align='center' bgcolor='#E0F8E6' scope='col' style='font-family:Calibri Light; font-size:20px; color:#000000'>" &
"<p></p>" &
  "<p></p>" &
   "<p></p>" &
  "<p></p>" &
  "<p></p>" &
  "<p></p>" &
 "<p></p>" &
 "<p></p>" &
 " <p align='left' style='font-size:16px'></p>" &
"</p></th>" &
"</tr></table>" &
"<p></p>" &
"</center></body></html>"
                    Dim asunto As String = "SISTEMA AUTOMÁTICO DE PÓLIZAS"
                    Dim destinatarioTo As String = EMAIL
                    Try
                        correo.To.Clear()
                        correo.Body = ""
                        correo.Subject = ""
                        correo.Body = mensaje
                        correo.Subject = asunto
                        correo.IsBodyHtml = True
                        correo.To.Add(Trim(destinatarioTo))
                        'correo.CC.Add(Trim(destinatarioCC))
                        'correo.Bcc.Add(Trim(destinatarioBCC))
                        correo.From = New MailAddress(emisor)
                        envio.Credentials = New NetworkCredential(emisor, password)

                        envio.Host = "mail.grupocomercialtria.com.mx"
                        envio.Port = 587
                        envio.EnableSsl = True

                        envio.Send(correo)
                        'MsgBox("El mensaje fue enviado correctamente. ", MsgBoxStyle.Information, "Mensaje")

                    Catch ex As Exception
                        ' MessageBox.Show(ex.Message, "Mensajeria 1.0 vb.net ®", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    i += 1
                Loop
            Catch ex As Exception
                Exit Try
            End Try
            cn.Close()
        End Using

    End Sub

End Class
