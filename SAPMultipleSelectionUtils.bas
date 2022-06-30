Attribute VB_Name = "SAPMultipleSelectionUtils"
Option Explicit
Option Private Module

Private tableControl As GuiTableControl
Private verticalScrollBar As GuiScrollbar
Private tableRow As GuiTableRow
Private cTextField As GuiCTextField
Public session As GuiSession

Private row, endRow, rowData, startRowData, endRowData, startColumnData, endColumnData, column, startColumn, endColumn As Integer

'@Description muestra informacion sobre los tipos de elementos hijos que tiene una fila dentro de la tabla (ejemplo: muestra los GuiCTextField que tiene un elemento de fila como sus hijos)
Public Sub ShowInfoSingleValuesRowChildrenType(ByVal id As String)
    Set tableControl = session.findById(id).Parent
    Set verticalScrollBar = tableControl.verticalScrollBar
    Set tableRow = tableControl.Rows(0)
    endColumn = tableRow.Count - 1
    For column = 0 To endColumn
        Debug.Print "Column => " & column & " Type => " & tableRow(CInt(column)).Type
    Next
End Sub

'@Description se encarga de cargar los datos desde un arreglo en cada fila mostrada en la seleccion multiple
'@Param data arreglo de los datos a cargar
Public Sub LoadSingleValues(ByVal data As Variant, ByVal id As String)
    Set tableControl = session.findById(id).Parent
    Set verticalScrollBar = tableControl.verticalScrollBar
    
    startRowData = LBound(data, 1)
    endRowData = UBound(data, 1)
    
    row = 0
    endRow = tableControl.VisibleRowCount
    
    For rowData = startRowData To endRowData
    
        If row = endRow Then
            row = 1
            
            verticalScrollBar.position = verticalScrollBar.position + verticalScrollBar.PageSize
            Set tableControl = session.findById(id).Parent
            Set verticalScrollBar = tableControl.verticalScrollBar
        
            endRow = tableControl.VisibleRowCount
        End If
            
        Set tableRow = tableControl.Rows(row)
        
        Set cTextField = tableRow(1)
        cTextField.text = CStr(data(rowData))
        
        row = row + 1
    Next
End Sub

'@Description carga la informacion en la venntana de seleccion multiples, en la pestaña de intervalos seleccionada
'@Param data matriz de 2 columnas, el primer elemento representa el limite inferior y el segundo elemento (columna) representa el limite superior.
Public Sub LoadRanges(ByVal data As Variant, ByVal id As String)
    Set tableControl = session.findById(id).Parent
    Set verticalScrollBar = tableControl.verticalScrollBar
    
    startRowData = LBound(data, 1)
    endRowData = UBound(data, 1)
    startColumnData = LBound(data, 2)
    endColumnData = UBound(data, 2)
    
    row = 0
    endRow = tableControl.VisibleRowCount
    
    For rowData = startRowData To endRowData
    
        If row = endRow Then
            row = 1
            
            verticalScrollBar.position = verticalScrollBar.position + verticalScrollBar.PageSize
            Set tableControl = session.findById(id).Parent
            Set verticalScrollBar = tableControl.verticalScrollBar
        
            endRow = tableControl.VisibleRowCount
        End If
            
        Set tableRow = tableControl.Rows(row)
        
        Set cTextField = tableRow(1)
        cTextField.text = CStr(data(rowData, startColumnData))
        Set cTextField = tableRow(2)
        cTextField.text = CStr(data(rowData, endColumnData))
        
        row = row + 1
    Next
End Sub
