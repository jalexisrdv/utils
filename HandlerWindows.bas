Attribute VB_Name = "HandlerWindows"
Option Explicit
Option Private Module

Private Const DOCUMENT_NAME As String = "5cead848f38540b70c048428225c04440fa4c641.XLSX"

Private pathfile As String
Private sourceWorkbook As Workbook

Private hWnd As Variant
Private hWndSaveAs As Variant

Public Sub HandleWindow(ByVal hWndParent As Variant)
    pathfile = ThisWorkbook.Path & "\" & DOCUMENT_NAME

    WindowsAPIUtils.SendKeyToWindow hWndParent, vbKeyReturn
    
    hWndSaveAs = WindowsAPIUtils.WaitFindWindowHandlerByCaption("guardar como|save as")
    hWnd = WindowsAPIUtils.FindWindowComponentHandlerByPosition(hWndSaveAs, 0, "Edit")
    WindowsAPIUtils.SendTextToWindowComponent hWnd, pathfile
    hWnd = WindowsAPIUtils.FindWindowComponentHandlerByPosition(hWndSaveAs, 1, "Button")
    WindowsAPIUtils.SendClickToWindowComponent hWnd
    
    WaitFileExists pathfile, hWndParent
    
    'si la ventana de SAP para abrir el archivo de Excel exportado no sale (La aplicacion se queda en un blucle)
    'es importante asegurarnos de que la ventana salga
    hWnd = WindowsAPIUtils.WaitFindChildWindowHandler(hWndParent)
    hWnd = WindowsAPIUtils.FindWindowComponentHandlerByPosition(hWnd, 2, "Button")
    WindowsAPIUtils.SendClickToWindowComponent hWnd
End Sub

Function OpenSourceWorkbook() As Workbook
    Set sourceWorkbook = Application.Workbooks.Open(pathfile)
    Set OpenSourceWorkbook = sourceWorkbook
End Function

Sub CloseSourceWorkbook()
    sourceWorkbook.Close SaveChanges:=False
    DeleteFile pathfile
    SAPGUIUtils.CloseCurrentModo
End Sub

Private Function FileExists(ByVal pathfile As String) As Boolean
    On Error Resume Next
    On Error GoTo 0
    If Dir(pathfile) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Private Function WaitFileExists(ByVal pathfile As String, ByVal hWndParent As Variant) As Boolean
    Dim exists As Boolean
    exists = FileExists(pathfile)
    hWnd = 0
    Do While exists <> True
        'si el archivo aun no existe eso significa que aun no se crea el archivo o se debe gestionar la ventana de seguridad permitiendo crear dicho archivo
        exists = FileExists(pathfile)
        If Not exists And hWnd = 0 Then
            hWnd = WindowsAPIUtils.FindChildWindowHandler(hWndParent)
            hWnd = WindowsAPIUtils.FindWindowComponentHandlerByPosition(hWnd, 1, "Button")
            WindowsAPIUtils.SendClickToWindowComponent hWnd
        End If
    Loop
    WaitFileExists = exists
End Function

Private Function DeleteFile(ByVal pathfile As String) As String
    Kill pathfile
End Function
