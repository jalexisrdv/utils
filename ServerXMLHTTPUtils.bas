Attribute VB_Name = "ServerXMLHTTPUtils"
'Librerias utilizadas: Microsoft XML, v6.0, Microsoft ActiveX Data Objects 2.8 Library, Microsoft ActiveX Data Objects Recordset 2.8 Library
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'Microsoft XML, v6.0', 'Microsoft ActiveX Data Objects 2.8 Library', 'Microsoft ActiveX Data Objects Recordset 2.8 Library'
'Documentacion oficial del objeto utilizado, proporcionado por la libreria:
'
'codigos de estado: https://developer.mozilla.org/es/docs/Web/HTTP/Status

'@Overview ServerXMLHTTPUtils proporciona funciones que facilitan la descarga de archivos externos a traves de HTTP

Option Explicit
Option Private Module

'codigo de errores del lado del servidor
Private Const INTERNAL_SERVER_ERROR As Long = 500

'codigo de errores de cliente
Private Const UNAUTHORIZED As Long = 401
Private Const NOT_FOUND As Long = 404

Private myServerXMLHTTP As New ServerXMLHTTP60
Private myOStream As New ADODB.Stream

'@Description descarga y almacena el archivo en una ruta
'@Param url direccion HTTP del archivo a descargar
'@Param pathname ruta donde se guardara el archivo descargado
'@Param username nombre de usuario en caso de ser necesario loguearse
'@Param password contrasena en caso de ser necesario loguearse
'@Return True significa que el archivo se descargo y guardo con exito
Public Function DownloadFile(ByVal url As String, ByVal pathname As String, Optional ByVal username As String = "", Optional ByVal password As String = "") As Boolean
    myServerXMLHTTP.Open "GET", url, False, username, password
    myServerXMLHTTP.setRequestHeader "Cache-Control", "no-cache"
    myServerXMLHTTP.setRequestHeader "Pragma", "no-cache"
    
    myServerXMLHTTP.send
    
    If myServerXMLHTTP.Status = 200 Then
        myOStream.Open
        myOStream.Type = 1
        myOStream.Write myServerXMLHTTP.responseBody
        myOStream.SaveToFile pathname, adSaveCreateOverWrite
        myOStream.Close
    End If
    
    DownloadFile = (myServerXMLHTTP.Status = 200)
End Function

Public Function GetStatus() As Long
    GetStatus = myServerXMLHTTP.Status
End Function

'@Description aborta la peticion realizada, limpiando los datos del objeto myServerXMLHTTP.
'Si ocurre algun error, es necesario invocar este metodo para limpiar el objeto, lo cual
'permite obtener los codigos de estados sin incongruencias
Public Sub abort()
    myServerXMLHTTP.abort
End Sub

'@Description muestra el mensaje de error al realizar la peticion HTTP
Public Sub ShowErrorMessageBox()
    Select Case GetStatus
        Case INTERNAL_SERVER_ERROR
            MsgBox "Error interno del servidor"
        Case UNAUTHORIZED
            MsgBox "Es necesario autenticarse para obtener la respuesta solicitada"
        Case NOT_FOUND
            MsgBox "El servidor no pudo encontrar el contenido solicitado"
        Case Else
            MsgBox "Archivo no descargado, el origen del error es desconocido"
    End Select
End Sub
