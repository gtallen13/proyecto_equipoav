Public Class Login
    Dim password As String = "UNICAH_SI"
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If txtPassword.Text = password Then
            If Decision.valor = 0 Then
                Me.Dispose()
                DirectionXml.ShowDialog()

            ElseIf Decision.valor = 1 Then
                Me.Dispose()
                XmlVersiculo.ShowDialog()
            End If
        Else
            MessageBox.Show("La contrase�a es incorrecta", "Datos Incorrectos", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Dispose()
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtPassword.ResetText()
    End Sub
End Class
