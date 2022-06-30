Attribute VB_Name = "RegexUtils"
'Libreria utilizada: Microsoft VBScript Regular Expressions 5.5
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'Microsoft VBScript Regular Expressions 5.5'
'Documentacion oficial: https://docs.microsoft.com/en-us/dotnet/standard/base-types/the-regular-expression-object-model

'@Overview RegexUtils proporciona funciones que facilitan la manipulacion de cadenas de texto,
'coincidencia de patrones lo cual permite encontrar o verificar la existecia de cadenas de texto (que coiciden con el patron regex)
'dentro de una cadena de texto.

Option Explicit
Option Private Module

Private regex As New RegExp

'@Description Si es True encuentra todas las coincidencias dentro de la cadena o texto.
'Si es False solo encuentra la primer coincidencia.
'@Param global_ recibe un valor booleano
Public Sub SetGlobal(ByVal global_ As Boolean)
    regex.Global = global_
End Sub

'@Description devuelve el valor de la propiedad Global del objeto de clase RegExp
'@Return True o False
Public Function GetGlobal() As Boolean
    GetGlobal = regex.Global
End Function

'@Description Si es True ignora minusculas y mayusculas AAA = aaa
'Si es False es sensible a minusculas y mayusculas AAA <> aaa
'@Param ignoreCase_ recibe un valor booleano
Public Sub SetIgnoreCase(ByVal ignoreCase_ As Boolean)
    regex.IgnoreCase = ignoreCase_
End Sub

'@Description devuelve el valor de la propiedad IgnoreCase del objeto de clase RegExp
'@Return True o False
Public Function GetIgnoreCase() As Boolean
    GetIgnoreCase = regex.IgnoreCase
End Function

'@Description Si es True la coincidencia del patron indicado continua incluso despues
'del salto de linea, util cuando se usan los operadores ^ y $. Ejemplo:
'Pattern: hola$
'Texto donde se buscaran coincidencias:
'
'texto por aqui hola
'texto por alla hola
'
'Si es True solo se encuentra la ultima coincidencia, el hola despues de la palabra alla.
'Si es False se encuentran todas las coincidencias del patron en este caso los dos hola.
'@Param multiline_ recibe un valor booleano
Public Sub SetMultiLine(ByVal multiline_ As Boolean)
    regex.MultiLine = multiline_
End Sub

'@Description devuelve el valor de la propiedad MultiLine del objeto de clase RegExp
'@Return True o False
Public Function GetMultiLine()
    GetMultiLine = regex.MultiLine
End Function

'@Description encuentra una coincidencia con el patron indicado
'@Param pattern_ patron regex mediante el cual se buscara una coincidencia
'@Param source texto en el cual se buscara la coincidencia
'@Return True si encontro coincidencia con el patron proporcionado
Public Function Match(ByVal pattern_ As String, ByVal source As String) As Boolean
    regex.pattern = pattern_
    Match = regex.Test(source)
End Function

'@Description remplaza todas las coincidencias con el patron indicado
'@Param pattern_ patron regex mediante el cual se buscaran coincidencias
'@Param replace_ texto que remplazara todas las coincidencias
'@Param source texto en el cual se buscaran las coincidencias
'@Return el nuevo texto al cual se le han remplazado las coincidencias
Public Function ReplaceByRegex(ByVal pattern_ As String, ByVal replace_ As String, ByVal source As String) As String
    regex.pattern = pattern_
    ReplaceByRegex = regex.replace(source, replace_)
End Function

'@Description encuentra todas las coincidencias con el patron indicado
'@Param pattern_ patron regex mediante el cual se buscaran coincidencias
'@Param source texto en el cual se buscaran las coincidencias
'@Return un objeto MatchCollection con todas las coincidencias encontradas
Public Function FindAllMatches(ByVal pattern_ As String, ByVal source As String) As MatchCollection
    regex.pattern = pattern_
    Set FindAllMatches = regex.Execute(source)
End Function
