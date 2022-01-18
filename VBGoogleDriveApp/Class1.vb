Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Drive.v3
Imports Google.Apis.Drive.v3.Data
Imports Google.Apis.Services
Imports Google.Apis.Util.Store

Public Class Class1
    'Dim lGDrive As New GoogleDrive
    'lGDrive = New GoogleDrive("GPA.CONTABLE1")
    'lstr_res = lGDrive.Subir(txt_archivo.text, "Estudio")
    'lGDrive.Cerrar

    Private mstr_cd_gd As String = ""
    Private mstr_id As String = ""
    Private mstr_secreto As String = ""

    Public mService As New DriveService

    Public Sub New(Optional ByVal pstr_cd_gd As String = "")

        If pstr_cd_gd <> "" Then
            Me.Configurar(pstr_cd_gd)
        End If

    End Sub

    Public Function Configurar(ByVal pstr_cd_gd As String) As String

        Dim lreg As DataRow = Nothing
        Dim lstr_res As String = "OK"

        'mstr_cd_gd = pstr_cd_gd'

        'lstr_res = CVM.CodDetalle.BuscarDR(lreg, "GDRIVE", mstr_cd_gd)

        'If lstr_res = "OK" AndAlso lreg IsNot Nothing Then
        '    mstr_id = lreg("CD_REF1")
        '    mstr_secreto = lreg("CD_REF2")
        'End If

        mstr_id = "260028708618-mssuc80o0rdt0k318p5c49l7mntjki6b.apps.googleusercontent.com"
        mstr_secreto = "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"
        mstr_cd_gd = "C:\Users\adolf\AppData\Roaming\MyAppsToken"

        lreg = Nothing

        Return lstr_res

    End Function

    Public Function Autenticar() As String

        Dim lstr_res As String = "OK"

        'basicamente scopes son los alcances que quiero que tenga esta autorizacion,se pueden agregar mas autorizaciones si se desea'
        Dim scopes As String() = New String() {DriveService.Scope.Drive, DriveService.Scope.DriveFile}

        'Creacion de acceso y refrezco del token de acceso (dura 1 hora)'
        Dim Credential As UserCredential

        Credential = GoogleWebAuthorizationBroker.AuthorizeAsync(New ClientSecrets() With {.ClientId = mstr_id, .ClientSecret = mstr_secreto}, scopes, Environment.UserName, CancellationToken.None, New FileDataStore(mstr_cd_gd, True)).Result

        Console.WriteLine("El Token se guardo en " & mstr_cd_gd)

        'Hasta ahora estuvimos autorizandonos con las credenciales en OAuth 2.0 la cual nos arrojara el token de acceso
        '----------------------Conectando con el servicio y pasandole las credenciales verificadas--------------------'
        Dim service As DriveService = New DriveService(New BaseClientService.Initializer() With {
        .HttpClientInitializer = Credential,
        .ApplicationName = "Google Drive VB Dot Net"
        })

        mService = service 'servicio autenticado'

        Return lstr_res

    End Function

    Public Function Cerrar() As String

        Dim lstr_res As String = "OK"

        'mstr_service.Channels.Stop().Execute'¿¿¿¿

        Return lstr_res

    End Function

    Private Function pf_PuedoTrabajar() As String

        Dim lstr_res As String = "OK"

        If mstr_cd_gd = "" Then
            lstr_res = "Falta Especificar el Código de GD"
        End If

        If lstr_res = "OK" AndAlso (mstr_id = "" Or mstr_secreto = "") Then
            lstr_res = "Falta Configurar la Conexión"
        End If

        Return lstr_res

    End Function

    Public Function ObtenerMimeType(ByVal filename) As String

        'Esta pequeña funcion es la encargada de saber cual es el tipo del archivo'
        Dim mimeType As String = "application/unknown" 'ejemplo "application/txt", es lo que deberia de retornarme si es un txt'
        Dim ext As String = System.IO.Path.GetExtension(filename).ToLower()

        Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext)

        If regKey IsNot Nothing AndAlso regKey.GetValue("Content Type") IsNot Nothing Then
            mimeType = regKey.GetValue("Content Type").ToString()
        End If

        System.Diagnostics.Debug.WriteLine(mimeType)
        Return mimeType
    End Function
    Public Function Subir(ByVal pstr_archivo As String, ByVal pstr_carpeta_drive As String) As String

        Dim lstr_id As String = ""
        Dim lstr_res As String = "OK"

        'Valido Si puedo Operar en GD
        lstr_res = pf_PuedoTrabajar()
        If lstr_res <> "OK" Then Return lstr_res

        'Autentico
        lstr_res = Me.Autenticar
        'Dim lservice As DriveService = mService

        'Obtengo el ID de la carpeta en GD
        lstr_id = Me.ObtenerIDCarpeta(pstr_carpeta_drive, True)

        'Subo el archivo
        Try
            If System.IO.File.Exists(pstr_archivo) Then
                '/puedo subir un archivo y pasarle mis propias propiedades a travez del objeto file
                ' o
                'puedo utilizar System.IO.Path y absorver las propiedades de un archivo existente en el sistema'
                Dim body As Google.Apis.Drive.v3.Data.File = New Google.Apis.Drive.v3.Data.File()
                body.Name = System.IO.Path.GetFileName(pstr_archivo)
                body.MimeType = ObtenerMimeType(pstr_archivo)
                body.Parents = New List(Of String) From {lstr_id} 'guarda el archivo en una carpeta, por defecto te arroja el archivo a "mi unidad"{comentar linea si se desea guardar en mi unidad}'

                Dim bytearray As Byte() = System.IO.File.ReadAllBytes(pstr_archivo)
                Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream(bytearray)

                Dim request = mService.Files.Create(body, stream, body.MimeType) 'creo el archivo
                Dim result = request.Upload()
                'MessageBox.Show("Espere un momento...")
                'MessageBox.Show(result.Status & " :)")

            End If
        Catch ex As Exception
            lstr_res = ex.Message
        End Try

        Return lstr_res

    End Function

    Public Function ObtenerListaArchivos() As List(Of Google.Apis.Drive.v3.Data.File)

        'Dim service As DriveService = mService
        Dim result As List(Of Google.Apis.Drive.v3.Data.File) = New List(Of Google.Apis.Drive.v3.Data.File)()
        Dim Request As Google.Apis.Drive.v3.FilesResource.ListRequest = mService.Files.List()

        Do

            Try
                Dim files As FileList = Request.Execute()
                result.AddRange(files.Files)
                Request.PageToken = files.NextPageToken
            Catch ex As Exception
                Console.WriteLine("error " & vbLf & ex.Message)
            End Try
        Loop While Not String.IsNullOrEmpty(Request.PageToken)

        Return result
    End Function

    Private Function FiltrarArchivos(
        ByVal pstr_carpeta_nombre
        )
        '--------Me filtra todos los archivos GoogleDrive por carpeta y me retorna el id--------'
        Dim folders = ObtenerListaArchivos()

        For i As Integer = 0 To folders.Count - 1 - 1

            If folders(i).MimeType = "application/vnd.google-apps.folder" Then

                If folders(i).Name = pstr_carpeta_nombre Then
                    Return folders(i).Id
                End If
            End If
        Next
        Return ""
    End Function
    Public Function ObtenerIDCarpeta('falta crear carpeta si no existe'
        ByVal pstr_carpeta_nombre As String,
        ByVal pbln_CrearSiNoExiste As Boolean
        ) As String

        Dim lstr_id As String = ""

        Try

            lstr_id = FiltrarArchivos(pstr_carpeta_nombre)

        Catch ex As Exception
            MessageBox.Show("Error " & vbLf & ex.Message)
        End Try

        If lstr_id = "" And pbln_CrearSiNoExiste Then
            lstr_id = Me.CrearCarpeta(pstr_carpeta_nombre) 'si el id no existe, creeame una carpeta con ese nombre y retorname el su id'
        End If

        Return lstr_id

    End Function

    Public Function ObtenerIDArchivo(ByVal pstr_archivo As String) As String

        Dim lstr_id As String = ""

        Try
            Dim files = ObtenerListaArchivos()

            For i As Integer = 0 To files.Count - 1 - 1

                If files(i).MimeType <> "application/vnd.google-apps.folder" Then
                    If files(i).Name = pstr_archivo Then Return files(i).Id
                End If
            Next

        Catch ex As Exception
            Return "error" & ex.Message
        End Try

        Return lstr_id

    End Function

    Public Function CrearCarpeta(ByVal pstr_carpeta_nombre As String) As String

        Dim lstr_id As String = ""

        ' Dim driveService As DriveService = mService
        Dim body As Google.Apis.Drive.v3.Data.File = New Google.Apis.Drive.v3.Data.File()
        body.Name = pstr_carpeta_nombre
        body.MimeType = "application/vnd.google-apps.folder"
        Dim command = mService.Files.Create(body)
        Dim file = command.Execute()
        MessageBox.Show("Carpeta creada con exito id ( " & file.Id & " )") '

        lstr_id = file.Id
        Return lstr_id

    End Function

    Public Function Descargar(ByVal pstr_archivo As String, ByVal pstr_carpeta_destino As String) As String

        Dim lstr_id_archivo As String = ""
        Dim lstr_res As String = "OK"

        lstr_id_archivo = Me.ObtenerIDArchivo(pstr_archivo)

        If lstr_id_archivo = "" Then
            lstr_res = "Archivo Inexistente"
        Else
            '?????
        End If

        Return lstr_res

    End Function
End Class
