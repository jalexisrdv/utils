Attribute VB_Name = "HandlerWindowSaveAsUtils"
Option Explicit
Option Private Module

Private Const FILE_NAME As String = "5cead848f38540b70c048428225c04440fa4c641"

'Es importante indicar la ruta: C:\Users\60079123\Documents\SAP\SAP GUI
'en pathfile, ya que SAP, por defecto le asigna permisos de escritura a esta ruta, lo cual evita que salgan ventanas para permitir la escritura de archivos (estas ventanas salen en rutas diferentes
'a la configurada por defecto en SAP)
Private pathfile As String
Private pathfileArray() As Variant 'rutas
Private sourceWorkbook As Workbook
Private fileExtension As String

'@Description permite manipular la ventana de "seleccionar calculo costes tabla" y "guardar como", en caso de exportar reportes en Excel que abran la ventana guardar como de Windows
'@Param hWnd manejador de ventana SAP seleccionar calculo costes tabla la cual se quiere manipular, a partir de esta ventana de SAP, se abre la ventana de guardar de windows
'@Param patternCaption patron regex, estableciendo el titulo de la ventana, se recomienda ponerlo en ingles y espa?ol
Public Sub HandleWindow(Optional ByVal hWnd As Variant = -1, Optional ByVal patternCaption As String = "guardar como|save as", Optional ByVal fileExtension_ As String = "XLSX")
    Dim hWndSaveAs As Variant

    SetFileExtension fileExtension_
    SetPathFile GetPathSAPGUI() & "\" & GetFileNameWithExtension

    If hWnd <> -1 Then
        WindowsAPIUtils.SendKeyToWindow hWnd, vbKeyReturn
    End If
    
    hWndSaveAs = WindowsAPIUtils.WaitFindWindowHandlerByCaption(patternCaption)
    hWnd = WindowsAPIUtils.FindWindowComponentHandlerByPosition(hWndSaveAs, 0, "Edit")
    WindowsAPIUtils.SendTextToWindowComponent hWnd, pathfile
    hWnd = WindowsAPIUtils.FindWindowComponentHandlerByPosition(hWndSaveAs, 1, "Button")
    WindowsAPIUtils.SendClickToWindowComponent hWnd
    WaitTransferringPackage
End Sub

Private Function GetPathMyDocuments() As String
    Dim wsh As Object
    Dim path As String
    
    Set wsh = CreateObject("WScript.Shell")
    path = wsh.SpecialFolders("MyDocuments")

    GetPathMyDocuments = path
End Function

Public Function GetPathSAPGUI() As String
    Dim path As String
    path = GetPathMyDocuments() & "\SAP\SAP GUI"
    GetPathSAPGUI = path
End Function

Public Function OpenSourceWorkbook(Optional ByVal pathfile_ As String = vbNullString) As Workbook
    If pathfile_ = vbNullString Then
        pathfile_ = pathfile
    End If
'    While sourceWorkbook Is Nothing
'        On Error Resume Next
'        Set sourceWorkbook = Application.Workbooks.Open(pathfile_)
'    Wend
    Set OpenSourceWorkbook = sourceWorkbook
End Function

Public Sub CloseSourceWorkbook()
    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing
End Sub

Public Sub CloseWindows()
    CloseSourceWorkbook
    DeleteFile pathfile
End Sub

'@Description espera a que todos los paquetes terminen de ser exportados de SAP
Public Sub WaitTransferringPackage()
    SAPGUIUtils.WaitStatusBarMessageType
End Sub

Public Function DeleteFile(Optional ByVal pathfile_ As String = vbNullString) As String
    If pathfile_ = vbNullString Then
        pathfile_ = pathfile
    End If
    Kill pathfile_
End Function

Public Sub DeleteFiles(Optional ByVal pathfileArray_ As Variant = vbEmpty)
    Dim i, min, max As Integer
    Dim pathfile As String
    
    If pathfileArray_ = vbEmpty Then
        pathfileArray_ = pathfileArray
    End If

    min = LBound(pathfileArray_)
    max = UBound(pathfileArray_)

    i = min
    
    For i = min To max
        pathfile = pathfileArray_(i)
        Kill pathfile
    Next
End Sub

Public Sub SetPathFile(ByVal pathfile_ As String)
    pathfile = pathfile_
End Sub

Public Function GetPathFile() As String
    GetPathFile = pathfile
End Function

Public Sub SetPathFileArray(ByVal pathfileArray_ As Variant)
    pathfileArray = pathfileArray_
End Sub

Public Function GetPathFileArray() As Variant()
    GetPathFileArray = pathfileArray
End Function

Public Sub SetFileExtension(ByVal fileExtension_ As String)
    fileExtension = fileExtension_
End Sub

Public Function GetFileExtension() As String
    GetFileExtension = fileExtension
End Function

Public Function GetFileName() As String
    GetFileName = FILE_NAME
End Function

Public Function GetFileNameWithExtension() As String
    GetFileNameWithExtension = FILE_NAME & "." & GetFileExtension
End Function
