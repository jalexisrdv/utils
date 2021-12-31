Attribute VB_Name = "SystemUtils"
'@Overview SystemUtils proporciona funciones que facilitan recuperar informacion del sistema o ejecutar procesos,
'permitiendo: recuperar el nombre de usuario del sistema operativo, o matar algun proceso en ejecucion.

Option Explicit
Option Private Module

'@Description obtiene el nombre de usuario con el cual inicia sesion en windows
'@Return cadena de texto que representa el nombre de usario del sistema operativo
Public Function GetUserName() As String
    GetUserName = Environ("username")
End Function

'@Description mata una tarea o proceso del sistema operativo mediante el nombre,
'para conocer el nombre del proceso abra el administrador de tareas y pulse la pestaña detalles
'@Return True si el proceso se ha matado o finalizado, False de lo contrario
Public Function TaskKill(taskName) As Boolean
    If CreateObject("WScript.Shell").Run("taskkill /f /im " & taskName, 0, True) = 0 Then
        TaskKill = True
    Else
        TaskKill = False
    End If
End Function
