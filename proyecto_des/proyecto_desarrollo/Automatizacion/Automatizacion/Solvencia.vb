Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.xml
Imports System.IO
Imports sigplusnet_vbnet_lcd15_demo
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Text.RegularExpressions




Public Class Form1
    Dim firmar As New sigplusnet_vbnet_lcd15_demo.Firma
    Dim Correlativo As String = "-PRES"
    Dim blnGuardado As Boolean = False
    Dim intCorrelativo As Integer = 0
    Dim checked As Integer = 0
    Dim cont_Microfono As Integer = 0
    Dim cont_Sonido As Integer = 0
    Dim cont_DataShow As Integer = 0
    Dim cont_Parlante As Integer = 0
    Dim cont_Otros As Integer = 0
    Dim cont_ControlAC As Integer = 0
    Dim cont_ControlP As Integer = 0
    Dim cont_ControlDS As Integer = 0
    Dim cont_Grabadora As Integer = 0
    Dim equipos As Integer = 0

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtSolicitante.Text = ""
        cbLugar.Text = ""
        txtFechaReservacion.Text = Nothing
        txtHoraInicial.Text = Nothing
        txtHoraFinal.Text = Nothing
        txtOtros1.Text = ""
        txtOtros2.Text = ""
        txtOtros2.Enabled = False
        txtEvento1.Text = ""
        txtEvento2.Text = ""
        txtEvento2.Enabled = False
        chkMicrofono.Checked = False
        chkSonido.Checked = False
        chkDataShow.Checked = False
        chkParlante.Checked = False
        chkOtros.Checked = False
        chkControlAC.Checked = False
        chkControlPantalla.Checked = False
        chkControlDataShow.Checked = False
        chkGrabadora.Checked = False
        chkDocente.Checked = False
        chkPersonalAdmin.Checked = False
        chkEstudiante.Checked = False


        TabControl1.SelectedTab = TabPage1


        Dim fecha As Date
        fecha = Now
        lblFecha.Text = Format(fecha, "dd/MM/yyyy")
        Label20.Text = "Favor Firmar el Documento"
        Label20.ForeColor = Color.Orange
        btnGuardar.Enabled = True
        txtOtros1.Enabled = False
        txtOtros2.Enabled = False


        Dim Doc As New XmlDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        Doc.Load(Application.StartupPath & "\Versiculo.xml")
        xmlnode = Doc.GetElementsByTagName("Versiculo")
        lbBiblia.Text = xmlnode(0).ChildNodes.Item(0).InnerText.Trim()
        If txtFechaReservacion.Text = Nothing Then
            txtFechaReservacion.ForeColor = Color.Gray
            txtFechaReservacion.Text = "MM/DD/YYYY"
        End If
        If txtHoraInicial.Text = Nothing Then
            txtHoraInicial.ForeColor = Color.Gray
            txtHoraInicial.Text = "XX:XX AM/PM"
        End If
        If txtHoraFinal.Text = Nothing Then
            txtHoraFinal.ForeColor = Color.Gray
            txtHoraFinal.Text = "XX:XX AM/PM"
        End If
    End Sub


    Private Sub Regresar_Click(sender As Object, e As EventArgs) Handles Regresar.Click
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub BotonFirmar_Click(sender As Object, e As EventArgs) Handles BotonFirmar.Click
        Dim firmar As New Firma
        Try
            'firmar.ShowDialog()
            firmar.firmado = True
            If firmar.firmado = True Then

                Label20.Text = "El documento ha sido Firmado"
                Label20.ForeColor = Color.LimeGreen
                btnGuardar.Enabled = True
                blnGuardado = True

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If blnGuardado = False Then
            MessageBox.Show("Por favor firme el documento", "Informacion Incompleta", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim direccionPlantilla As String
        Dim direccionCarpeta As String
        Dim direccionFirmas As String = Application.StartupPath()

        Dim doc As New XmlDocument
        Dim xmlnode As XmlNodeList
        Dim i As Integer = 0
        doc.Load(Application.StartupPath & "\Direccion.xml")
        xmlnode = doc.GetElementsByTagName("Direccion")
        direccionCarpeta = xmlnode(0).ChildNodes.Item(0).InnerText.Trim()
        direccionPlantilla = xmlnode(0).ChildNodes.Item(1).InnerText.Trim()


        Dim pdfTemplate As String = direccionPlantilla
        Dim newFile As String = direccionCarpeta & txtSolicitante.Text & Correlativo & ".pdf"

        'el siguiente codigo iterara sobre los archivo creados para asi asignar el numero seguido del Correlativo
        Dim directorioCarpeta As New DirectoryInfo(direccionCarpeta) 'obtiene la direccion de la carppeta
        Dim manejadorArchivos As FileInfo() = directorioCarpeta.GetFiles() 'utilizado para poder manejar los archivos
        Dim infoArchivo As FileInfo 'obtiene informacion del archivo
        For Each infoArchivo In manejadorArchivos
            If infoArchivo.FullName = newFile Then
                intCorrelativo += 1
                newFile = direccionCarpeta & txtSolicitante.Text & Correlativo & intCorrelativo & ".pdf"
            End If
        Next


        Dim pdfReader As New PdfReader(pdfTemplate)
        Dim pdfStamper As New PdfStamper(pdfReader, New FileStream(newFile, FileMode.Create))
        Dim pdfFormFields As AcroFields = pdfStamper.AcroFields
        Dim pcbContent As PdfContentByte = Nothing
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(direccionFirmas & "\firma.bmp") 'aqui ira la direccion de las firmas
        Dim sap As PdfSignatureAppearance = pdfStamper.SignatureAppearance
        Dim rect As iTextSharp.text.Rectangle = Nothing
        Dim imagen As iTextSharp.text.Image
        Dim loc As String


        loc = direccionFirmas & "\firma.bmp" 'aqui ira la direccion de las firmas

        imagen = iTextSharp.text.Image.GetInstance(loc)
        imagen.SetAbsolutePosition(70, 300)
        imagen.ScaleToFit(130, 130)
        pcbContent = pdfStamper.GetUnderContent(1)
        pcbContent.AddImage(imagen)

        ' set form pdfFormFields
        pdfFormFields.SetField("Solicitante", txtSolicitante.Text)
        pdfFormFields.SetField("Lugar", cbLugar.Text)
        pdfFormFields.SetField("FechaReservacion", txtFechaReservacion.Text)
        pdfFormFields.SetField("HoraInicial", txtHoraInicial.Text)
        pdfFormFields.SetField("HoraFinal", txtHoraFinal.Text)
        pdfFormFields.SetField("Fecha", lblFecha.Text)
        pdfFormFields.SetField("Otros_1", txtOtros1.Text)
        pdfFormFields.SetField("Otros_2", txtOtros2.Text)
        pdfFormFields.SetField("Evento_1", txtEvento1.Text)
        pdfFormFields.SetField("Evento_2", txtEvento2.Text)




        'pdfFormFields.SetField("signature5", TextBox1.Text)

        ' The form's checkboxes
        If chkMicrofono.Checked = True Then
            pdfFormFields.SetField("Microfono", "On")
        End If

        If chkSonido.Checked = True Then
            pdfFormFields.SetField("Sonido", "On")
        End If

        If chkDataShow.Checked = True Then
            pdfFormFields.SetField("DataShow", "On")
        End If

        If chkParlante.Checked = True Then
            pdfFormFields.SetField("Parlante", "On")
        End If

        If chkControlAC.Checked = True Then
            pdfFormFields.SetField("ControlAC", "On")
        End If

        If chkControlPantalla.Checked = True Then
            pdfFormFields.SetField("ControlPantalla", "On")
        End If

        If chkControlDataShow.Checked = True Then
            pdfFormFields.SetField("ControlDataShow", "On")
        End If

        If chkGrabadora.Checked = True Then
            pdfFormFields.SetField("Grabadora", "On")
        End If
        If chkDocente.Checked = True Then
            pdfFormFields.SetField("Docente", "On")
        End If
        If chkPersonalAdmin.Checked = True Then
            pdfFormFields.SetField("PersonalAdmin", "On")
        End If
        If chkEstudiante.Checked = True Then
            pdfFormFields.SetField("Estudiante", "On")
        End If
        If chkOtros.Checked = True Then
            pdfFormFields.SetField("Otros", "On")
        End If






        MessageBox.Show("Datos Guardados Satisfactoriamente")

        ' flatten the form to remove editting options, set it to false
        ' to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = False

        ' close the pdf
        pdfStamper.Close()


        txtSolicitante.Text = ""
        txtFechaReservacion.Text = ""
        cbLugar.Text = ""
        txtHoraInicial.Text = ""
        txtHoraFinal.Text = ""
        txtEvento1.Text = ""
        txtOtros1.Text = ""
        lblEvento2.Text = ""
        txtOtros2.Text = ""
        chkDocente.CheckState = 0
        chkPersonalAdmin.CheckState = 0
        chkEstudiante.CheckState = 0
        chkMicrofono.CheckState = 0
        chkSonido.CheckState = 0
        chkDataShow.CheckState = 0
        chkParlante.CheckState = 0
        chkOtros.CheckState = 0
        chkControlAC.CheckState = 0
        chkControlPantalla.CheckState = 0
        chkControlDataShow.CheckState = 0
        chkGrabadora.CheckState = 0

        Me.Close()

    End Sub



    Private Sub CheckOtros_Razones_CheckedChanged(sender As Object, e As EventArgs) Handles chkOtros.CheckedChanged
        If chkOtros.Checked = True Then
            txtOtros1.Enabled = True

        Else
            txtOtros1.Enabled = False
            txtOtros1.Text = ""
            txtOtros2.Enabled = False
            txtOtros2.Text = ""
        End If
        If cont_Otros = 0 Then
            cont_Otros += 1
        Else
            cont_Otros = cont_Otros - 1
        End If
    End Sub

    Private Sub pbSiguiente_Click(sender As Object, e As EventArgs) Handles pbSiguiente.Click
        equipos = cont_ControlAC + cont_ControlDS + cont_ControlP + cont_DataShow + cont_Grabadora + cont_Microfono + cont_Otros + cont_Parlante + cont_Sonido
        If txtFechaReservacion.Text <> "MM/DD/YYYY" Or txtHoraFinal.Text <> Nothing Or txtHoraInicial.Text <> Nothing Or cbLugar.Text <> Nothing Or txtSolicitante.Text <> Nothing Or txtEvento1.Text <> Nothing Then
            If EsTiempoValido(txtHoraFinal.Text) = True And EsTiempoValido(txtHoraInicial.Text) Then
                If chkDocente.Checked = False And chkEstudiante.Checked = False And chkPersonalAdmin.Checked = False Then
                    MessageBox.Show("Debe especificar el Solicitante.")
                ElseIf IsDate(txtFechaReservacion.Text) = False Then
                    MessageBox.Show("La fecha introducida no es correcta.")
                ElseIf chkControlAC.Checked = True Or chkControlDataShow.Checked = True Or chkControlPantalla.Checked = True Or chkDataShow.Checked = True Or chkOtros.Checked = True Or chkMicrofono.Checked = True Or chkGrabadora.Checked = True Or chkParlante.Checked = True Or chkSonido.Checked = True Then
                    If equipos > 5 Then
                        MessageBox.Show("No se pueden seleccionar mas de 5 equipos.")
                    Else
                        If chkOtros.Checked = True And txtOtros1.Text = Nothing Then
                            MessageBox.Show("No ha especificado el equipo.")
                        Else
                            TabControl1.SelectedTab = TabPage2
                        End If
                    End If
                Else
                    MessageBox.Show("No a seleccionado ningun equipo.")
                End If
            Else
                MessageBox.Show("Las horas ingresadas son incorrectas.")
            End If
        Else
            MessageBox.Show("Algunos datos no han sido llenados.")
        End If
    End Sub

    Private Sub txtSolicitante_KeyPress(sender As Object, e As KeyPressEventArgs) _
                              Handles txtSolicitante.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = "abcdefghijklmnopqrstuvwxyz "
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtLugar_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = "abcdefghijklmnopqrstuvwxyz 1234567890"
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtHoraInicial_KeyPress(sender As Object, e As KeyPressEventArgs) _
         Handles txtHoraInicial.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = ":1234567890pma "
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtHoraFinal_KeyPress(sender As Object, e As KeyPressEventArgs) _
         Handles txtHoraFinal.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = ":1234567890pma "
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub txtFechaReservacion_KeyPress(sender As Object, e As KeyPressEventArgs) _
                              Handles txtFechaReservacion.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = "1234567890/"
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub rtxtOtros_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = "abcdefghijklmnopqrstuvwxyz 1234567890"
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub rtxtEvento_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Not (Asc(e.KeyChar) = 8) Then
            Dim permitido As String = "abcdefghijklmnopqrstuvwxyz 1234567890"
            If Not permitido.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtFechaReservacion_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFechaReservacion.LostFocus

        If txtFechaReservacion.Text = Nothing Then
            txtFechaReservacion.ForeColor = Color.FromArgb(254, 197, 50)
            txtFechaReservacion.Text = "MM/DD/YYYY"
        End If

    End Sub
    Private Sub txtFechaReservacion_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFechaReservacion.GotFocus

        If txtFechaReservacion.Text = "MM/DD/YYYY" Then
            txtFechaReservacion.ForeColor = Color.FromArgb(254, 197, 50)
            txtFechaReservacion.Text = ""
        End If

    End Sub

    Private Sub txtHoraInicial_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHoraInicial.LostFocus

        If txtHoraInicial.Text = Nothing Then
            txtHoraInicial.ForeColor = Color.FromArgb(254, 197, 50)
            txtHoraInicial.Text = "XX:XX AM/PM"
        End If

    End Sub
    Private Sub txtHoraInicial_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHoraInicial.GotFocus

        If txtHoraInicial.Text = "XX:XX AM/PM" Then
            txtHoraInicial.ForeColor = Color.FromArgb(254, 197, 50)
            txtHoraInicial.Text = ""
        End If

    End Sub

    Private Sub txtHoraFinal_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHoraFinal.LostFocus

        If txtHoraFinal.Text = Nothing Then
            txtHoraFinal.ForeColor = Color.FromArgb(254, 197, 50)
            txtHoraFinal.Text = "XX:XX AM/PM"
        End If

    End Sub
    Private Sub txtHoraFinal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHoraFinal.GotFocus

        If txtHoraFinal.Text = "XX:XX AM/PM" Then
            txtHoraFinal.ForeColor = Color.FromArgb(254, 197, 50)
            txtHoraFinal.Text = ""
        End If

    End Sub

    Private Sub chkDocente_CheckedChanged(sender As Object, e As EventArgs) Handles chkDocente.CheckedChanged
        checked += 1
        If checked Mod 2 = 0 Then
            chkEstudiante.Enabled = True
            chkPersonalAdmin.Enabled = True
        Else
            chkEstudiante.Enabled = False
            chkPersonalAdmin.Enabled = False
        End If
    End Sub

    Private Sub chkPersonalAdmin_CheckedChanged(sender As Object, e As EventArgs) Handles chkPersonalAdmin.CheckedChanged
        checked += 1
        If checked Mod 2 = 0 Then
            chkEstudiante.Enabled = True
            chkDocente.Enabled = True
        Else
            chkEstudiante.Enabled = False
            chkDocente.Enabled = False
        End If
    End Sub

    Private Sub chkEstudiante_CheckedChanged(sender As Object, e As EventArgs) Handles chkEstudiante.CheckedChanged
        checked += 1
        If checked Mod 2 = 0 Then
            chkPersonalAdmin.Enabled = True
            chkDocente.Enabled = True
        Else
            chkPersonalAdmin.Enabled = False
            chkDocente.Enabled = False
        End If
    End Sub

    Private Sub chkControlAC_CheckedChanged(sender As Object, e As EventArgs) Handles chkControlAC.CheckedChanged
        If cont_ControlAC = 0 Then
            cont_ControlAC += 1
        Else
            cont_ControlAC = cont_ControlAC - 1
        End If
    End Sub

    Private Sub chkControlDataShow_CheckedChanged(sender As Object, e As EventArgs) Handles chkControlDataShow.CheckedChanged
        If cont_ControlDS = 0 Then
            cont_ControlDS += 1
        Else
            cont_ControlDS = cont_ControlDS - 1
        End If
    End Sub

    Private Sub chkControlPantalla_CheckedChanged(sender As Object, e As EventArgs) Handles chkControlPantalla.CheckedChanged
        If cont_ControlP = 0 Then
            cont_ControlP += 1
        Else
            cont_ControlP = cont_ControlP - 1
        End If
    End Sub

    Private Sub chkDataShow_CheckedChanged(sender As Object, e As EventArgs) Handles chkDataShow.CheckedChanged
        If cont_DataShow = 0 Then
            cont_DataShow += 1
        Else
            cont_DataShow = cont_DataShow - 1
        End If
    End Sub

    Private Sub chkGrabadora_CheckedChanged(sender As Object, e As EventArgs) Handles chkGrabadora.CheckedChanged
        If cont_Grabadora = 0 Then
            cont_Grabadora += 1
        Else
            cont_Grabadora = cont_Grabadora - 1
        End If
    End Sub

    Private Sub chkMicrofono_CheckedChanged(sender As Object, e As EventArgs) Handles chkMicrofono.CheckedChanged
        If cont_Microfono = 0 Then
            cont_Microfono += 1
        Else
            cont_Microfono = cont_Microfono - 1
        End If
    End Sub

    Private Sub chkParlante_CheckedChanged(sender As Object, e As EventArgs) Handles chkParlante.CheckedChanged
        If cont_Parlante = 0 Then
            cont_Parlante += 1
        Else
            cont_Parlante = cont_Parlante - 1
        End If
    End Sub

    Private Sub chkSonido_CheckedChanged(sender As Object, e As EventArgs) Handles chkSonido.CheckedChanged
        If cont_Sonido = 0 Then
            cont_Sonido += 1
        Else
            cont_Sonido = cont_Sonido - 1
        End If
    End Sub



    Private Sub txtEvento1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEvento1.KeyPress
        If txtEvento1.Text.Trim().Length() = 35 Then
            txtEvento2.Enabled = True
            txtEvento2.Select()
        End If
    End Sub

    Private Sub txtOtros1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtOtros1.KeyPress
        If txtOtros1.Text.Trim().Length() = 23 Then
            txtOtros2.Enabled = True
            txtOtros2.Select()

        End If
    End Sub


    Public Function EsTiempoValido(ByVal Tiempo As String) As Boolean
        Dim Formato As String = "^ *(1[0-2]|[1-9]|0[1-9]):[0-5][0-9] *(a|p|A|P)(m|M) *$"
        Dim Validar_Tiempo As New System.Text.RegularExpressions.Regex(Formato)
        Return Validar_Tiempo.IsMatch(Tiempo)
    End Function

    'ING. GREGORY LO AMAMOS <3
End Class
