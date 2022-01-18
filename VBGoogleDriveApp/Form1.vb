Imports Google.Apis.Drive.v3

Public Class Form1
    Dim service As New DriveService

    Private Sub createService()

    End Sub

    Private Sub btnSubir_Click(sender As Object, e As EventArgs) Handles btnSubir.Click

        Dim class1 As New Class1
        class1.Configurar("")
        class1.Autenticar()

        class1.Subir(FilePath.Text, "Estudio1")

        'MessageBox.Show(class1.ObtenerIDArchivo("DetalledeComprobantesdeVenta.xlsx"))
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

        createService()
        Dim buscar As New OpenFileDialog
        If buscar.ShowDialog = DialogResult.OK Then
            FilePath.Text = buscar.FileName
        Else
            Exit Sub
        End If

    End Sub

End Class
