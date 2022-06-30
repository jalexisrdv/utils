Attribute VB_Name = "AsynchronousRunUtils"
'@Overview AsynchronousRun permite ejecutar el metodo Run del objeto Application de manera "Asincrona",
'esto solo funciona cuando se manda a ejecutar el metodo Run desde otra instancia de Excel, por lo que,
'el metodo Start debe ser invocado dentro de un metodo perteneciente al Workbook que pertenece a la nueva
'instancia de Excel que fue creada. Ejemplo:
'
'Sub Launcher(ByVal pathfile_ As String)
'    pathfile = pathfile_
'    AsynchronousRun.Start "HandleWindows"
'End Sub
'
'El metodo anterior, es invocado desde el Worbook que pertenece a la instancia original de Excel.
'Set myExcel = ExcelObjectCreatorUtils.OpenWorkbookInNewInstance(pathfile)
'myExcel.Run "Launcher", pathfile_
'
'Reference: https://social.msdn.microsoft.com/Forums/en-US/0546f8eb-d786-4037-906e-1ee5d42e7484/asynchronous-applicationrun-call?forum=isvvba

Option Private Module
Option Explicit

'@Description ejecuta un procedimiento simulando la ejecucion asincrona, por lo que, al invocar cierto procedimiento
'la ejecucion de la macro continua sin esperar respuesta del procedimiento invocado. Nota: este metodo debe ser invocado desde una nueva instancia de Excel con
'el Workbook que realiza las tareas de su interes.
'@Param procedure procedimiento o funcion que se deasea ejecutar
Public Sub Start(ByVal procedure As String)
    Application.OnTime Now + TimeSerial(0, 0, 1), procedure
End Sub
