Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports POLIZA_AROMITALIA.Form1

Public Class CorreoSCCCompra
    Public Shared Sub SCCCompra()
        Using cn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings.Item("COI").ToString)
            Dim adapter As New SqlDataAdapter("SELECT * FROM CORREOS_INSOFTEC WHERE TIPO='Compras'", cn)
            Dim dt As New DataTable
            adapter.Fill(dt)
            cn.Open()
            Dim z As Integer = (dt.Rows.Count)
            Dim i As Integer = 0
            Try
                Do While (i <= z)
                    EMAIL = dt.AsEnumerable().ElementAtOrDefault(i).Item(2).ToString
                    Dim emisor = "alertas@grupocomercialtria.com.mx"
                    Dim password As String = "QiVDZ90chqmp"
                    Dim mensaje As String = "<html><head><style type='text/css'><!--body,td,th {color: #000000;}--></style>" &
      "</head><body>" &
"<center><table width='985' border='0' align='left'><tr>" + "<br>" &
" <p align='center' style='font-family:Calibri Light; font-size:20px'>[-Error de Generaci&oacute;n de P&oacute;liza-] <br><br>/ Compra: " + CVE_DOC + " / Fecha: " & FECHA_DOC & " <br><br>Prveedor: (" & CVE_CLPV & " ) " & NOMBRE_CLIENTE & " <br><br>&rarr; Favor de validar o capturar la cuenta contable en el campo CUENTA CONTABLE del proveedor.<br><br>Gracias.<br><br></p>" &
"<th width='979' align='center' bgcolor='#E0F8E6' scope='col' style='font-family:Calibri ligth; font-size:12px, Calibri Ligth; color:#000000'><p></p>" &
  "<p></p>" &
  "<p></p>" &
  "<p></p>" &
  "<p></p>" &
 "<p></p>" &
 "<p>&nbsp;</p>" &
 " <p align='left' style='font-size:16px'>  </p><br>" &
"</p></th>" &
"</tr></table>" &
"<p><strong></strong><strong> </strong></p>" &
"</center></body></html>"
                    Dim asunto As String = "SISTEMA AUTOMTÁTICO DE PÓLIZAS"
                    Dim destinatarioTo As String = EMAIL
                    'Dim destinatarioCC As String = ""
                    'Dim destinatarioBCC As String = "marias@insoftec.com"
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
                        'MessageBox.Show(ex.Message, "Mensajeria 1.0 vb.net ®", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    i += 1

                Loop
                CVE_DOC = ""
                CVE_CLPV = ""
                NOMBRE_CLIENTE = ""
                CUENTA_CONTABLE = ""
                contador = 0
            Catch ex As Exception
                Exit Try
            End Try
        End Using
    End Sub

End Class
