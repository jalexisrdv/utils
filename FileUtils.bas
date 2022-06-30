Attribute VB_Name = "FileUtils"
'Libreria utilizada: Microsoft Scripting Runtime
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'Microsoft Scripting Runtime'
'Documentacion oficial del objeto utilizado, proporcionado por la libreria: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object

'@Overview FileUtils proporciona funciones que facilita la manipulacion de archivos, permitiendo:
'abrir archivos de texto, eliminar archivos y carpetas, verificar la existencia de archivos y carpetas,
'obtener archivos y subfolders de una carpeta, obtener un archivo, recuperar el ultimo archivo modificado
'o creado.

Option Explicit
Option Private Module

Private fso As New FileSystemObject

'@Description crea un nuevo archivo de texto
'@Param path ruta donde se guardara el archivo
'@Return objeto de tipo TextStream, el cual facilita el acceso secuencial o manipulacion del archivo de texto
Public Function CreateTextFile(ByVal path As String) As TextStream
    Set CreateTextFile = fso.CreateTextFile(path, True)
End Function

'@Description abre un archivo de texto
'@Param path ruta del archivo de texto a abrir
'@Return objeto de tipo TextStream, el cual facilita el acceso secuencial o manipulacion del archivo de texto
Public Function OpenTextFile(ByVal path As String) As TextStream
    Set OpenTextFile = fso.OpenTextFile(path, ForReading)
End Function

'@Description elimina un archivo mediante la ruta especificada
'@Param path ruta del archivo a eliminar
'@Return True si el archivo es eliminado, False si no es eliminado
Public Sub DeleteFile(ByVal path As String)
    fso.DeleteFile path, True
End Sub

'@Description elimina un folder mediante la ruta especificada
'@Param path ruta del folder a eliminar
'@Return True si el folder es eliminado, False si no es eliminado
Public Sub DeleteFolder(ByVal path As String)
    fso.DeleteFolder path, True
End Sub

'@Description elimina todos los archivos de una ruta especificada
'@Param path ruta del folder o carpeta donde se eliminaran los archivos
Public Sub DeleteAllFilesInFolder(ByVal path As String)
    Dim files_ As Files
    Dim file_ As file
    Set files_ = GetFilesInFolder(path)
    For Each file_ In files_
        On Error Resume Next
        file_.Delete True
    Next
End Sub

'@Description elimina todos los subfolders (folders) de una ruta especificada
'@Param path ruta donde se eliminaran todos los subfolders (folders)
Public Sub DeleteAllSubFoldersInFolder(ByVal path As String)
    Dim folders_ As Folders
    Dim folder_ As Folder
    Set folders_ = GetSubFoldersInFolder(path)
    For Each folder_ In folders_
        On Error Resume Next 'en caso de obtener error de permiso denegado
        folder_.Delete True
    Next
End Sub

'@Description elimina todos los archivos y subfolders (folders) de una ruta especificada
'@Param path ruta donde se eliminaran todos los archivos y subfolders (folders)
Public Sub DeleteAllInFolder(ByVal path As String)
    DeleteAllSubFoldersInFolder path
    DeleteAllFilesInFolder path
End Sub

'@Description comprueba si un archivo existe
'@Param path ruta del archivo a comprobar existencia
'@Return True si el archivo existe, False si el archivo no existe
Public Function FileExist(ByVal path As String) As Boolean
    FileExist = fso.FileExists(path)
End Function

'@Description comprueba si un folder existe
'@Param path ruta del folder a comprobar existencia
'@Return True si el folder existe, False si el folder no existe
Public Function FolderExist(ByVal path As String) As Boolean
    FolderExist = fso.FolderExists(path)
End Function

'@Description lee todo el archivo de texto y devuelve una cadena de texto
'@Param path ruta del archivo de texto a leer
'@Return cadena de texto de todo el archivo de texto leido
Public Function ReadAll(ByVal path As String) As String
    ReadAll = OpenTextFile(path).ReadAll
End Function

'@Description obtiene todos los archivos dentro de un folder
'@Param path ruta del folder del cual se recuperaran los archivos
'@Return objeto Files que contiene todos los archivos contenidos dentro del folder
Public Function GetFilesInFolder(ByVal path As String) As Files
    Set GetFilesInFolder = fso.GetFolder(path).Files
End Function

'@Description obtiene todos los subfolders (folders) dentro de un folder
'@Param path ruta del folder del cual se recuperaran los subfolders
'@Return objeto Folders que contiene todos los subfolders contenidos dentro del folder
Public Function GetSubFoldersInFolder(ByVal path As String) As Folders
    Set GetSubFoldersInFolder = fso.GetFolder(path).SubFolders
End Function

'@Description crea un folder en la ruta especificada
'@Param path ruta en la cual se creara el folder
'@Return objeto Folder que representa el folder creado
Public Function CreateFolder(ByVal path As String) As Folder
    Set CreateFolder = fso.CreateFolder(path)
End Function

'@Description obtiene el archivo indicado
'@Param path ruta del archivo a recuperar
'@Return objeto File que representa el archivo recuperado
Public Function GetFile(ByVal path As String) As file
    Set GetFile = fso.GetFile(path)
End Function

'@Description devuelve el ultimo archivo modificado o creado de acuerdo a la extension especificada
'@Param path ruta donde se buscara el ultmo archivo modificado o creado
'@Param extension formato (ejemplo: pdf) del archivo que se buscara, util para ignorar los demas archivos con extension diferente
'@Param extensionLong longitud o numero de caracteres de la extension (formato)
'@Return objeto File que representa el ultimo archivo modificado o creado
Public Function GetLastFileModifiedByExtension(ByVal path As String, ByVal extension As String, ByVal extensionLong As Integer) As file
    Dim lastFileName As String
    Dim lastFile, archivo As file
    Dim archivos As Files
    Dim isFirstFile As Boolean
    Set lastFile = Nothing
    Set archivos = FileUtils.GetFilesInFolder(path)
    For Each archivo In archivos
        If InStr(LCase(Right(archivo.path, extensionLong)), LCase(extension)) > 0 Then
            If Not isFirstFile Then
                Set lastFile = archivo
                isFirstFile = True
            End If
            If lastFile.DateLastModified < archivo.DateLastModified Then
                Set lastFile = archivo
            End If
        End If
    Next
    Set GetLastFileModifiedByExtension = lastFile
End Function

'@Description espera por la existencia del ultimo archivo modificado o creado, util cuando se visualiza algun documento en SAP con
'el visualizador de documentos en una nueva ventana (nuevo modo), ya que la ejecucion del script podria continuar su ejecucion
'sin que el visualizador de documentos de SAP aun visualice el documento, lo cual provoca que el archivo aun no sea creado.
'Por lo que, si solo se se ejecuta el metodo GetLastFileModifiedByExtension este no encontraria ningun archivo, debido
'a que el visualizador de documentos de SAP aun no termina de mostrar y crear el documento.
'@Param path ruta donde se buscara el ultmo archivo modificado o creado
'@Param extension formato del archivo que se buscara, util para ignorar los demas archivos con extension diferente
'@Param extensionLong longitud o numero de caracteres de la extension
'@Param seconds segundos que se pausa o espera el metodo para proceder con la ejecucion
'@Return objeto File que representa el ultimo archivo modificado o creado
Public Function WaitLastFileModifiedByExtension(ByVal path As String, ByVal extension As String, ByVal extensionLong As Integer, Optional ByVal seconds As String = "1") As file
    Dim archivo As file
    Do While archivo Is Nothing
        Application.Wait (Now + TimeValue("0:00:" & seconds))
        Set archivo = GetLastFileModifiedByExtension(path, extension, extensionLong)
    Loop
    Set WaitLastFileModifiedByExtension = archivo
End Function

'@Description espera a que el archivo sea creado
'@Param path ruta con el nombre del archivo donde este sera creado
'@Param seconds segundos que se pausa o espera el metodo para proceder con la ejecucion
'@Return objeto File que representa el archivo creado
Public Function WaitExistenceFile(ByVal path As String, Optional ByVal seconds As String = "1") As file
    Dim exist As Boolean
    Do While exist <> True
        Application.Wait (Now + TimeValue("0:00:" & seconds))
        exist = FileExist(path)
    Loop
    Set WaitExistenceFile = GetFile(path)
End Function
