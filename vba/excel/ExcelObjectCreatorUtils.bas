Attribute VB_Name = "ExcelObjectCreatorUtils"
'@Overview ExcelObjectCreatorUtils permite ejecutar la macro en una segunda instancia de Excel o en otro proceso.
'¿Que me permite ejecutar la macro en una segunda instancia o en otro proceso?
'   1.- Bajar reportes de SAP con la transaccion GR55 usando el formato de salida de Excel, sin que el Excel de SAP se rompa, permitiendole cargar la información
'   2.- Evitar que los documentos de Excel que se encuentran en la instancia o proceso original se bloqueen. Ya que la macro corre sobre otro proceso o instancia.
'Esta libreria de utilidad tiene el metodo OpenThisWorkbookInNewInstance, el cual debe ser ejecutado en el evento Workbook_Open invocado al iniciar el documento Excel
'(solo si desea que el documento Excel automaticamente se abra en una nueva instancia de Excel).
'
'   Private Sub Workbook_Open()
'       ExcelObjectCreatorUtils.OpenThisWorkbookInNewInstance
'   End Sub
'

Option Private Module
Option Explicit

'@Description abre el mismo Workbook o archivo de Excel que ejecuta este metodo en una nueva instancia
Public Sub OpenThisWorkbookInNewInstance()
    Dim myExcel As Application
    Dim pathfile As String
    Dim createdWorkbook As Workbook
    
    pathfile = ThisWorkbook.path & "\" & ThisWorkbook.Name
    
    Set myExcel = OpenWorkbookInNewInstance(pathfile, True, False)
    
    Set createdWorkbook = Application.Workbooks.Add
    With createdWorkbook.Sheets(1).Range("A1")
        .value = "No cerrar documento de Excel hasta que la ejecución de la macro finalice (no es necesario guardar este documento)"
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    myExcel.enableEvents = True
    ThisWorkbook.Close SaveChanges:=False
End Sub

'@Description abre un Workbook o archivo de Excel en nueva instancia de Excel
'@Param pathfile ruta del archivo de Excel o Workbook que desea abrir
'@Param visible_ True para hacer visible el Workbook que abrira, False para ocultarlo
'@Param enableEvents_ True para habilitar eventos al Workbook que abrira, False para deshabilitarlos
'@Return Application objeto de la nueva instancia de Excel creada
Public Function OpenWorkbookInNewInstance(ByVal pathfile As String, Optional ByVal visible_ As Boolean = False, Optional ByVal enableEvents_ As Boolean = True) As Application
    Dim myExcel As Application
    Set myExcel = CreateExcelNewInstance
    myExcel.visible = visible_
    myExcel.enableEvents = enableEvents_
    OpenWorkbookWithInstance myExcel, pathfile
    Set OpenWorkbookInNewInstance = myExcel
End Function

'@Description abre un Workbook o archivo de Excel en una instancia de Excel especificada
'@Param myExcel instance de objeto de Excel
'@Param pathfile ruta del archivo de Excel o Workbook que desea abrir
'@Return Workbook objeto del archivo de Excel o Workbook que fue abierto
Public Function OpenWorkbookWithInstance(ByVal myExcel As Application, ByVal pathfile As String) As Workbook
    Set OpenWorkbookWithInstance = myExcel.Workbooks.Open(Filename:=pathfile, ReadOnly:=False)
End Function

'@Description crea una nueva instance de un objeto de Excel
'@Return Application nueva instance de un objeto Excel
Public Function CreateExcelNewInstance() As Application
    Set CreateExcelNewInstance = New Excel.Application
End Function
