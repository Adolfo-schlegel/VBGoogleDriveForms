Imports System.Configuration
Imports Google.Apis.Auth
Imports Google.Apis.Download
Imports Google.Apis.Drive.v2
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Drive.v2.Data
Imports System.Threading
Imports System.Data.SqlClient

Public Class Form1
    Dim service As New DriveService

    Private Sub createService()
        Dim clienteId = "260028708618-mssuc80o0rdt0k318p5c49l7mntjki6b.apps.googleusercontent.com"
        Dim clienteSecret = "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"

        Dim credenciales As UserCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(New ClientSecrets() With {.ClientId = clienteId, .ClientSecret = clienteSecret}, {DriveService.Scope.Drive}, "user", CancellationToken.None).Result

        service = New DriveService(New BaseClientService.Initializer() With {.HttpClientInitializer = credenciales, .ApplicationName = "Google Drive VB Dot Net"})
    End Sub

    Private Sub btnSubir_Click(sender As Object, e As EventArgs) Handles btnSubir.Click


        If service.ApplicationName <> "MYGoogleDrive" Then createService()

        Dim myarchivo As New File
        Dim bytearry As Byte() = System.IO.File.ReadAllBytes(FilePath.Text)
        Dim stream As New System.IO.MemoryStream(bytearry)
        Dim uploadRequest As FilesResource.InsertMediaUpload = service.Files.Insert(myarchivo, stream, myarchivo.MimeType)
        uploadRequest.Upload()
        Dim file As File = uploadRequest.ResponseBody
        MessageBox.Show("¡El archivo se ha subido correctamente! ")

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
