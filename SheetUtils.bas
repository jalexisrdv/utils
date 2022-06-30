Attribute VB_Name = "SheetUtils"
'@Overview SheetUtils proporciona funciones que facilitan copiar informacion de una hoja fuente a una hoja objetivo,
'verificar la existencia de hojas dentro de un Workbook, copiar esquemas de agrupacion, y mas.

Option Explicit
Option Private Module

'@Description copia los niveles de agrupacion de una hoja fuente a la hoja objetivo
'@Param sourceSheet hoja de la cual se copiaran los niveles de agrupacion
'@Param targetSheet hoja en la cual se copiaran los niveles de agrupacion
'@Param sourceStartRow fila inicial de hoja fuente a partir de la cual se copiaran los niveles
'@Param sourceEndRow fila final de la hoja fuente hasta donde se copiaran los niveles
'@Param targetStartRow fila inicial de hoja objetivo a partir de la cual se comenzaran a establecer los niveles copiados
Public Sub CopyGroupingLevelsFrom( _
    ByVal sourceSheet As Worksheet, _
    ByVal targetSheet As Worksheet, _
    ByVal sourceStartRow As Long, _
    ByVal sourceEndRow As Long, _
    ByVal targetStartRow As Long _
)
    Dim row As Long
    For row = sourceStartRow To sourceEndRow
        targetSheet.rows(targetStartRow).OutlineLevel = sourceSheet.rows(row).OutlineLevel
        targetStartRow = targetStartRow + 1
    Next
End Sub

'@Description obtiene el nivel de agrupacion mas alto existente de una fila
'@Param sheet hoja de la cual se obtendra el nivel de agrupacion
'@Return valor Integer que representa el nivel de agrupacion mas alto
Public Function GetMaxRowOutlineLevel(ByVal sheet As Worksheet) As Long
    Dim row As Range
    Dim maxOutlineLevel As Long
    For Each row In sheet.UsedRange.rows
        If maxOutlineLevel < row.OutlineLevel Then
            maxOutlineLevel = row.OutlineLevel
        End If
    Next
    GetMaxRowOutlineLevel = maxOutlineLevel
End Function

'@Description obtiene el nivel de agrupacion mas alto existente de una columna
'@Param sheet hoja de la cual se obtendra el nivel de agrupacion
'@Return valor Integer que representa el nivel de agrupacion mas alto
Public Function GetMaxColumnOutlineLevel(ByVal sheet As Worksheet) As Long
    Dim column As Range
    Dim maxOutlineLevel As Long
    For Each column In sheet.UsedRange.Columns
        If maxOutlineLevel < column.OutlineLevel Then
            maxOutlineLevel = column.OutlineLevel
        End If
    Next
    GetMaxColumnOutlineLevel = maxOutlineLevel
End Function

'@Description solo visualiza las filas del nivel establecido, los demas niveles de las filas son contraidos
'@Param sheet hoja en la cual se mostraran las filas del nivel de agrupacion indicado
'@Param maxOutlineLevel el nivel de grupacion mas alto existente en una fila
'@Param levels nivel hasta el cual se visualizaran las filas
Public Sub ShowRowOutlineLevel(ByVal sheet As Worksheet, ByVal maxOutlineLevel As Long, ByVal levels As Long)
    Do While maxOutlineLevel >= levels
        sheet.Outline.ShowLevels rowLevels:=maxOutlineLevel
        maxOutlineLevel = maxOutlineLevel - 1
    Loop
End Sub

'@Description solo visualiza las columnas del nivel establecido, los demas niveles de las columnas son contraidos
'@Param sheet hoja en la cual se mostraran las columnas del nivel de agrupacion indicado
'@Param maxOutlineLevel el nivel de grupacion mas alto existente en una columna
'@Param levels nivel hasta el cual se visualizaran las columnas
Public Sub ShowColumnOutlineLevel(ByVal sheet As Worksheet, ByVal maxOutlineLevel As Long, ByVal levels As Long)
    Do While maxOutlineLevel >= levels
        sheet.Outline.ShowLevels columnLevels:=maxOutlineLevel
        maxOutlineLevel = maxOutlineLevel - 1
    Loop
End Sub

'@Description limpia o elimina el esquema de agrupacion establecido
'@Param sheet hoja en la cual se eliminara el esquema de agrupacion
Public Sub ClearOutline(ByVal sheet As Worksheet)
    sheet.Cells.ClearOutline
End Sub

'@Description copia la direccion del esquema de agrupacion desde una hoja fuente a una hoja objetivo
'@Param sourceSheet hoja fuente desde la cual se copiara la direccion del esquema de agrupacion
'@Param targetSheet hoja objetivo en la cual se establecera la direccion del esquema de agrupacion copiado
Public Sub CopyOutlineFrom(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)
    targetSheet.Outline.SummaryRow = sourceSheet.Outline.SummaryRow
    targetSheet.Outline.SummaryColumn = sourceSheet.Outline.SummaryColumn
End Sub

'@Description copia la informacion del rango usado desde la hoja fuente al rango indicado de la hoja objetivo
'@Param sourceSheet hoja desde la cual se copiara la informacion del rango usado
'@Param targetRange rango de la hoja objetivo donde se pegara la informacion copiada
Public Sub CopyByUsedRangeFrom(ByVal sourceSheet As Worksheet, ByVal targetRange As Range)
    sourceSheet.UsedRange.Copy targetRange
End Sub

'@Description copia los valores del rango usado de la hoja fuete a traves del array de valores del rango usado,
'este tipo de copiado es util para no hacer uso del portapapeles al copiar la informacion.
'@Param sourceSheet hoja desde la cual se copiara la informacion del rango usado
'@Param targetRange rango de la hoja objetivo donde se pegara la informacion copiada
Public Sub CopyUsedRangeValuesArrayFrom(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet, ByVal targetRange As Range)
    Dim auxNumberFormat As Variant
'    auxNumberFormat = targetSheet.Cells.NumberFormat 'creando copia del formato de celdas original
'    targetSheet.Cells.NumberFormat = "@" 'estableciendo formato de texto, para evitar convertir datos
    targetRange.Resize(UBound(sourceSheet.UsedRange.Value, 1), UBound(sourceSheet.UsedRange.Value, 2)).Value = sourceSheet.UsedRange.Value
'    targetSheet.Cells.NumberFormat = auxNumberFormat 'restaurando el formato de celdas original
End Sub

'@Description auto-ajusta el tamaño de las columnas de acuerdo al texto contenido dentro
'@Param sheet hoja en la cual se ajustara el tamaño de las columnas automaticamente
Public Sub AutoFitColumns(ByVal sheet As Worksheet)
    sheet.UsedRange.Columns.AutoFit
End Sub

'@Description verifica la existencia de una hoja dentro del Workbook o documento excel
'@Param targetWorkbook documento excel o Workbook donde se comprobara la existencia de la hoja indicada
'@Param sheetName nombre de la hoja a comprobar existencia
'@Return True si la hoja existe, False de lo contrario
Public Function ContaintsSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Boolean
On Error GoTo ErrorHandler
    If LCase(targetWorkbook.Worksheets(sheetName).Name) = LCase(sheetName) Then
        ContaintsSheet = True
    End If
    Exit Function
ErrorHandler:
    ContaintsSheet = False
End Function

'@Description cambia la fuente de datos de la tabla dinamica y actualiza o refresca la tabla para mostrar la informacion nueva.
'@Param targetWorkbook libro objetivo en el cual se creara un nuevo cache con la nueva informacion (fuente de datos)
'@Param targetSheet hoja de la cual se recuperan las tablas dinamicas
'@Param pivotTableName nombre de la tabla dinamica que se desea recuperar
'@Param sourceData fuente de datos que se usara para cambiar y alimentar los datos de la tabla dinamica (dentro del rango es importante incluir las columnas que identifican cada dato de la fila)
Public Sub ChangePivotTableSourceData(ByVal targetWorkbook As Workbook, ByVal targetSheet As Worksheet, ByVal pivotTableName As String, ByVal sourceData As Range)
    Dim targetPivotTable As PivotTable
    Dim newPivotCache As PivotCache
    
    Set targetPivotTable = targetSheet.PivotTables(pivotTableName)
    Set newPivotCache = targetWorkbook.PivotCaches.Create(xlDatabase, sourceData:=sourceData)
    targetPivotTable.ChangePivotCache newPivotCache
    targetPivotTable.PivotCache.Refresh 'actualizando el cache o fuente de datos de la tabla recuperada
    targetPivotTable.Update 'actualizando solo la tabla recuperada
End Sub

'@Description revisa si el rango de celdas seleccionado, tiene celdas visibles. Las celdas visibles se revisan cuando algun filtro de columna es aplicado.
'@Param targetRange rango a revisar si existen celdas visibles.
'@Return True, si contiene celdas visibles, Falso de lo contrario
Public Function HasVisibleCells(ByVal targetRange As Range) As Boolean
On Error GoTo ErrorHandler
    targetRange.SpecialCells (xlCellTypeVisible)
    HasVisibleCells = True
    Exit Function
ErrorHandler:
    HasVisibleCells = False
End Function

'@Description convierte los valores de texto contenidos en una columna a valores de tipo numero
'@Param sheet hoja donde existe la columna
'@Param column numero de columna en la cual desea convertir los valores de texto a numero
Public Sub ConvertTextColumnToNumberColumn(ByVal sheet As Worksheet, ByVal column As Long)
    With sheet.Columns(column)
        .NumberFormat = "0"
        .Value = .Value
    End With
End Sub

'@Description selecciona todos los valores de todos los filtros, esto hace que toda la informacion se muestre. En pocas palabras limpia el filtro
'@Param sheet hoja de la cual se quiere limpiar el filtro
Public Sub SelectAllFilters(ByVal sheet As Worksheet)
    On Error Resume Next
    sheet.ShowAllData
    On Error GoTo 0
End Sub
