Attribute VB_Name = "WScriptUtils"
'@Overview WScriptUtils proporciona funciones que facilitan la creacion de argumentos, que podrian pasarse al ejecutar un script.
'Ademas, permite ejecutar scripts externos.

Option Explicit
Option Private Module

'@Description crea un objeto WScript.shell para ejecutar scripts externos
'@Return valor Object que representa el objeto WScript.shell
Private Function CreateWScriptObject() As Object
    Set CreateWScriptObject = CreateObject("WScript.shell")
End Function

'@Description prepara los argumentos nombrados que seran pasados al script a ejecutar
'@Param argsDictionary objeto Dictionary, la clave del diccionario representa el nombre del
'argumento y el valor del diccionario representa el valor del argumento (la clave y valor deben ser de tipo String)
'@Return valor String que contiene el formato utilizado para pasar argumentos al script a ejecutar
Public Function PrepareArgs(ByVal argsDictionary As dictionary) As String
    Dim Keys As Variant
    Dim args, key, value As String
    Dim i, limit As Integer
    limit = argsDictionary.Count - 1
    Keys = argsDictionary.Keys
    For i = 0 To limit
        key = Keys(i)
        value = Chr(34) & argsDictionary.item(key) & Chr(34)
        args = args & " /" & key & ":" & value
    Next
    PrepareArgs = args
End Function

'@Description ejecuta un script externo
'@Param pathfile ruta del script a ejecutar
'@Param args string con el formato de argumentos que se pasaran al script
Public Sub RunScript(ByVal pathfile As String, Optional ByVal args As String = "")
    Dim wscript As Object
    Set wscript = CreateWScriptObject()
    wscript.Run Chr(34) & pathfile & Chr(34) & args
    Set wscript = Nothing
End Sub
