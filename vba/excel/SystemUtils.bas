Attribute VB_Name = "SystemUtils"
'@Overview SystemUtils proporciona funciones que facilitan recuperar informacion del sistema o ejecutar procesos,
'permitiendo: recuperar el nombre de usuario del sistema operativo

Option Explicit
Option Private Module

'@Description obtiene el nombre de usuario con el cual inicia sesion en windows
'@Return cadena de texto que representa el nombre de usario del sistema operativo
Public Function GetUserName() As String
'    GetUserName = Environ("username") 'hace uso de la variable de entorno username (esta puede no existir)
    GetUserName = Split(GetPathMyDocuments(), "\")(2)
End Function

'@Description obtiene la ruta a la carpeta documents
'@Return la ruta de la carpeta documents
Public Function GetPathMyDocuments() As String
    Dim wsh As Object
    Dim path As String
    
    Set wsh = CreateObject("WScript.Shell")
    path = wsh.SpecialFolders("MyDocuments")

    GetPathMyDocuments = path
End Function

'@Description obtiene la ruta a la carpeta SAP GUI
'@Return la ruta de la carpeta SAP GUI
Public Function GetPathSAPGUI() As String
    Dim path As String
    path = GetPathMyDocuments() & "\SAP" & "\SAP GUI"
    GetPathSAPGUI = path
End Function
