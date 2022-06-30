Attribute VB_Name = "DictionaryUtils"
'Libreria utilizada: Microsoft Scripting Runtime
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'Microsoft Scripting Runtime'
'Documentacion oficial del objeto utilizado, proporcionado por la libreria: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object

'@Overview DictionaryUtils proporciona funciones que facilitan la manipulacion del objeto Dictionary,
'permitiendo agregar, eliminar, recuperar y verificar existencia de elementos agregados al objeto.
'Ademas, permite convertir una hoja de excel que representa un diccionario de datos para clasificar informacion a un
'objeto Dictionary, lo cual mejora la eficiencia al buscar informacion.
Option Private Module
Option Explicit

Private dictionary As New dictionary

'@Description crea un objeto Dictionary a partir del valor de un rango, es importante que la seleccion del rango cree una matriz
'el diccionario creado sirve para clasificar informacion, la primera columna de la hoja de excel representa el valor clave (key)
'key sirve para buscar la informacion de manera eficiente dentro del objeto Dictionary.
'El valor (value o item) del objeto Dictionary sera una fila de la hoja excel que representa el diccionario, dicho
'valor o item o fila se representa mediante un arreglo de una dimension
'@Param matrix valor de una rango seleccionado que representa una matriz de 2 dimensiones
Public Sub CreateDictionary(ByVal matrix As Variant)
    If IsEmpty(matrix) Then
        Exit Sub
    End If

    dictionary.RemoveAll

    Dim row, column, startRow, endRow, startColumn, endColumn As Long
    Dim key As String
    
    startRow = LBound(matrix, 1)
    endRow = UBound(matrix, 1)
    startColumn = LBound(matrix, 2)
    endColumn = UBound(matrix, 2)
    
    For row = startRow To endRow
        Dim item() As Variant
        ReDim item(startColumn To endColumn)
        For column = startColumn To endColumn
            item(column) = Trim(matrix(row, column))
        Next
        key = UCase(Trim(matrix(row, startColumn)))
        If Not Contains(key) Then
            Add key, item
        End If
    Next
End Sub

'@Description agrega un elemento al objeto Dictionary
'@Param key valor con el cual se indentificara el elemento a agregar
'@Param item valor del elemento a agregar
Public Sub Add(ByVal key As Variant, ByVal item As Variant)
    dictionary.Add key, item
End Sub

'@Description actualiza un elemento existente del objeto Dictionary
'@Param key valor del item que se desea actualizar
'@Param item nuevo valor del item a actualizar
Public Sub UpdateItem(ByVal key As Variant, ByVal item As Variant)
    If Contains(key) Then
        Remove key
    End If
    Add key, item
End Sub

'@Description verifica la existencia de un elemento (valor) dentro del objeto Dictionary de acuerdo a la clave con
'el cual fue indentificado.
'@Param key valor clave con el cual se identifica el elemento (valor) agregado
'@Return True si el elemento existe, False de lo contrario
Public Function Contains(ByVal key As Variant) As Boolean
    Contains = dictionary.Exists(key)
End Function

'@Description obtiene un elemento (valor) dentro del objeto Dictionary de acuerdo a la clave con el cual fue identificado
'@Param key valor clave con el cual se identifica el elemento agregado
'@Return valor Variant que representa el elemento dentro del objeto Dictionary
Public Function GetItem(ByVal key As Variant) As Variant
    GetItem = dictionary.item(key)
End Function

'@Description obtiene todos los elementos (valores) contenidos dentro del objeto Dictionary
'@Return valor Variant que representa todos los elementos dentro del objeto Dictionary
Public Function GetItems() As Variant()
    GetItems = dictionary.items
End Function

'@Description obtiene todos las claves (keys) contenidas dentro del objeto Dictionary
'@Return valor Variant que representa todas las claves dentro del objeto Dictionary
Public Function GetKeys() As Variant()
    GetKeys = dictionary.Keys
End Function

'@Description elimina un elemento dentro del objeto Dictionary de acuerdo a la clave con el cual se indentifica
'@Param key valor clave con el cual se identifica el elemento agregado
Public Sub Remove(ByVal key As Variant)
    dictionary.Remove key
End Sub

'@Description elimina todos los elementos contenidos dentro del objeto Dictionary
Public Sub RemoveAll()
    dictionary.RemoveAll
End Sub

'@Description obtiene el numero de elementos que contiene el objeto Dictionary
'@Return valor Integer que representa el numero de elementos que almacena el objeto Dictionary
Public Function GetSize() As Integer
    GetSize = dictionary.Count
End Function


