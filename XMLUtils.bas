Attribute VB_Name = "XMLUtils"
'Libreria utilizada: Microsoft XML Core Services (MSXML)
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'Microsoft XML, v6.0'
'Documentacion oficial: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms763742(v=vs.85)

'@Overview XMLUtils proporciona funciones que facilitan obtener informacion de los nodos de un documento XML

Option Explicit
Option Private Module

Private msxml As New MSXML2.DOMDocument60

'@Description establece el nombre de espacio a utilizar para leer el documento XML
'el nombre de espacio es el texto antes del nombre del nodo de un XML, ejemplo:
'//cfdi:Comprobante donde cfdi es el espacio de nombre de los nodos del XML
'@Param namespaces nombre de espacios del documento XML
Public Sub SetSelectionNamespaces(ByVal namespaces As String)
    msxml.setProperty "SelectionNamespaces", namespaces
End Sub

'@Description carga un documento xml
'@Param path ruta del documento XML a cargar
Public Sub Load(ByVal path As String)
    msxml.Load path
End Sub

'@Description obtiene un nodo de acuerdo al query especificado
'@Param query consulta del nodo a obtener dentro del XML
'@Return objeto IXMLDOMElement que representa el nodo recuperado
Public Function SelectNode(ByVal query As String) As IXMLDOMElement
    Set SelectNode = msxml.SelectSingleNode(query)
End Function

'@Description obtiene todos los nodos de acuerdo al query especificado
'@Param query consulta de los nodos a obtener dentro del XML
'@Return objeto IXMLDOMNodeList que representa una coleccion con todos los nodos obtenidos.
Public Function SelectNodes(ByVal query As String) As IXMLDOMNodeList
    Set SelectNodes = msxml.SelectNodes(query)
End Function

'@Description obtiene todo el documento XML en una cadena de texto
'@Return valor String que representa todo el documento XML
Public Function GetXMLDocument() As String
    GetXMLDocument = msxml.xml
End Function

