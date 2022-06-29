Attribute VB_Name = "RunTimeMeterUtils"
'@Overview RunTimeMeterUtils permite medir el tiempo de ejecucion de una macro
Option Private Module
Option Explicit

Private t As Single

'@Description obtiene el tiempo inicial, al ejecutar alguna porcion de codigo
'@Return Single, tiempo recuperado
Public Function START_TIMER() As Single
    t = Timer
    START_TIMER = t
End Function

'@Description obtiene el tiempo final, tiempo de ejecucion de alguna porcion de codigo
'@Return Single, tiempo recuperado
Public Function END_TIMER() As Single
    t = Timer - t
    END_TIMER = t
End Function

'@Description muestra un mensaje emergente indicando el tiempo de ejecucion
'@Param msg mensaje personalizado que se mostrara con el mensaje que indica el tiempo de ejecucion
Public Sub ShowTime(Optional ByVal msg As String = vbNullString)
    If msg = vbNullString Then
        MsgBox "Tiempo de ejecución: " & t & " segundos", vbInformation
    Else
        MsgBox msg & vbNewLine & "Tiempo de ejecución: " & t & " segundos", vbInformation
    End If
End Sub
