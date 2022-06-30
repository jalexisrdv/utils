Attribute VB_Name = "GenerateFormulasCodeUtils"
Option Explicit
Option Private Module

Sub GenerateCode()
    GenerateConstants 29, 30, 1, 2, "BSEG DZ-AC"
    GenerateFormulaMapping 29, 30, 1, "BSEG DZ-AC"
End Sub

'@Description genera el codigo necesario para declarar las constantes que van en el modulo llamado Formulas
'@Param startColumn columna inicial que continiene formulas
'@Param endColumn ultima columna donde finalizan las coolumnas formuladas
'@Param headerRow fila que tiene el nombre de cabecera de las filas
'@Param formulaRow fila que contiene la formula
'@Param sheetName nombre de la hoja a partir de la cual vamos a generar el codigo
Public Sub GenerateConstants(ByVal startColumn As Long, ByVal endColumn As Long, ByVal headerRow As Long, ByVal formulaRow As Long, ByVal sheetName As String)
    Dim column As Long
    Dim constantName, formula, quotes As String
    Dim sheetTarget As Worksheet
    
    Set sheetTarget = ThisWorkbook.Worksheets(sheetName)
    quotes = Chr(34) 'Chr(34) representa el simbolo "
    
    For column = startColumn To endColumn
        constantName = sheetTarget.Cells(headerRow, column).Value
        constantName = sheetName & "_" & constantName
        constantName = SanitizeConstantName(constantName)
        formula = SanitizeFormula(sheetTarget.Cells(formulaRow, column).formula)
        Debug.Print "Public Const " & constantName & " As String = " & quotes & formula & quotes
    Next
End Sub


'@Description genera el codigo necesario que vamos a usar dentro del modulo ControllerFormulas, es el modulo que se encarga de asignar las formulas a un rango de celdas
'@Param startColumn columna inicial que continiene formulas
'@Param endColumn ultima columna donde finalizan las coolumnas formuladas
'@Param headerRow fila que tiene el nombre de cabecera de las filas
'@Param sheetName nombre de la hoja a partir de la cual vamos a generar el codigo
Public Sub GenerateFormulaMapping(ByVal startColumn As Long, ByVal endColumn As Long, ByVal headerRow As Long, ByVal sheetName As String)
    Dim column As Long
    Dim constantName, formula, quotes As String
    Dim sheetTarget As Worksheet
    
    Set sheetTarget = ThisWorkbook.Worksheets(sheetName)
    quotes = Chr(34) 'Chr(34) representa el simbolo "
    
    Debug.Print "Public Sub " & SanitizeConstantName(sheetName) & "_ADD_FORMULAS(ByVal sheet As Worksheet)"
    Debug.Print "   With sheet"
    Debug.Print "       startRow = "
    Debug.Print "       endRow = .UsedRange.Rows(.UsedRange.Rows.Count).row"
    
    For column = startColumn To endColumn
        constantName = sheetTarget.Cells(headerRow, column).Value
        constantName = sheetName & "_" & constantName
        constantName = SanitizeConstantName(constantName)
        Debug.Print "       .Range(" & ".Cells(startRow, " & column & "), .Cells(endRow, " & column & ")).Formula=" & " Formulas." & constantName
    Next
    
    Debug.Print "   End With"
    Debug.Print "End Sub"
End Sub

'@Description limpia string de formula para que esta pueda ser almacenada en una variable de tipo String
'@Return String formula con la sintaxis requerida para poder ser asignada en una variable de tipo String
Private Function SanitizeFormula(ByVal formula As String) As String
    Dim quotes As String
    quotes = Chr(34) 'Chr(34) representa el simbolo "
    formula = Replace(formula, quotes, quotes & quotes) 'remplazamos " por ""
    SanitizeFormula = formula
End Function

Private Function SanitizeConstantName(ByVal constantName As String) As String
    SanitizeConstantName = SanitizeVariableName(constantName)
End Function

'@Description limpia el string, con la finalidad de cumplir las reglas necesarias para el nombramiento de variables declaradas
'@Param variableName nombre de la variable a crear
'@Return String nombre de la variable ya sanitizada, nombre que cumple con la sintaxis de declaracion de variables
Private Function SanitizeVariableName(ByVal variableName As String) As String
    variableName = Replace(variableName, " ", "_")
    variableName = Replace(variableName, "-", "_")
    variableName = Replace(variableName, "'", "")
    variableName = EliminarAsentos(variableName)
    SanitizeVariableName = UCase(variableName)
End Function

Private Function EliminarAsentos(ByVal word As String) As String
    'Reemplazamos acentos en vocales minúsculas
    word = Replace(word, "á", "a")
    word = Replace(word, "é", "e")
    word = Replace(word, "í", "i")
    word = Replace(word, "ó", "o")
    word = Replace(word, "ú", "u")
    'Reemplazamos acentos en vocales mayúsculas
    word = Replace(word, "Á", "A")
    word = Replace(word, "É", "E")
    word = Replace(word, "Í", "I")
    word = Replace(word, "Ó", "O")
    word = Replace(word, "Ú", "U")
    
    EliminarAsentos = word
End Function
