Attribute VB_Name = "SAPGUIUtils"
'Libreria utilizada: SAP GUI Scripting API
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'SAP GUI Scripting API'.
'Por defecto la libreria no aparece, debe ser agregada mediante el archivo OCX, para ello pulsar el boton "Examinar"
'y abrir la siguiente ruta: C:\Program Files (x86)\SAP\FrontEnd\SAPgui\
'dentro de la ruta seleccione el archivo sapfewse.ocx
'Documentacion oficial: https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.03/en-US

'@Overview SAPGUIScriptingAPIUtils proporciona funciones que facilitan la conexion con la aplicacion SAP Logon,
'creacion de nuevos modos, aplicacion de transacciones, y verificacion de la existencia de componentes gui
'durante la ejecucion de un Script en SAP Logon.

Option Explicit
Option Private Module

Private sapGuiAuto As Object
Private sapApplication As GuiApplication
Private sapConnection As GuiConnection
Private session As GuiSession

'@Description obtiene la conexion con la aplicacion SAP Logon
'@Return objeto GuiSession que representa la sesion actual del usuario en SAP Logon
Public Function GetConnection() As GuiSession
    Set sapGuiAuto = GetObject("SAPGUI")
    Set sapApplication = sapGuiAuto.GetScriptingEngine
    Set sapConnection = sapApplication.Children(0)
    Set session = sapConnection.Children(0)
    Set GetConnection = session
End Function

'@Description obtiene la conexion con la aplicacion SAP Logon, solo se obtiene la conexion una unica vez.
'@Return objeto GuiSession que representa la sesion actual del usuario en SAP Logon
Public Function GetSingletonConnection() As GuiSession
    If session Is Nothing Then
        Set session = GetConnection
    End If
    Set GetSingletonConnection = session
End Function

'@Description devuelve el numero de sesiones activas en la aplicacion SAP Logon,
'cada sesion representa un nuevo modo en SAP Logon (equivalente a aplicar la transaccion /on)
'@Return valor Integer que representa el numero de sesiones establecidas o creadas
Public Function GetSessionsCount() As Integer
    GetSessionsCount = sapConnection.Sessions.Count
End Function

'@Description espera que la session sea establecida, cuando se crea un nuevo modo o se aplica
'la transacciion /on, una nueva ventana de SAP Logon es abierta, cuando la ventana es abierta
'en su totalidad, un nuevo objeto GuiSession es creado y por lo tanto el numero de sesiones
'establecidas aumenta.
'@Param sessionsCountBeforeModoNew numero de sesiones activas antes de crear un modo nuevo o aplicar la transaccion /on
'@Return objeto GuiSessiion que representa la ultima sesion establecida
Public Function WaitNewSession(ByVal sessionsCountBeforeModoNew As Integer) As GuiSession
    Do While sessionsCountBeforeModoNew = GetSessionsCount()
        'Wait...
    Loop
    Set WaitNewSession = GetLastSession
End Function

'@Description crea un modo o una sesion abriendo una nueva ventana de SAP Logon en la cual trabajara el Script SAP
'@Return objeto GuiSession que representa la ultima sesion establecida o creada
Public Function OpenNewModo() As GuiSession
    Set session = GetConnection
    Dim sessionsCountBeforeModoNew As Integer
    sessionsCountBeforeModoNew = GetSessionsCount()
    session.SendCommand "/on"
    Set session = WaitNewSession(sessionsCountBeforeModoNew)
    Set OpenNewModo = session
End Function

'@Description crea un modo o una sesion abriendo una nueva ventana de SAP Logon en la cual trabajara el Script SAP, este modo solo es abierto una unica vez.
'@Return objeto GuiSession que representa la ultima sesion establecida o creada
Public Function OpenNewModoSingleton() As GuiSession
    If session Is Nothing Then
        Set session = OpenNewModo
    End If
    Set OpenNewModoSingleton = session
End Function

'@Description cierra el ultimo modo creado o la ultima ventana de SAP Logon abierta o creada.
'Dicha sesion sera sobre la cual este trabajando el Scrip SAP
Public Sub CloseCurrentModo()
    session.SendCommand "/i"
    Set session = Nothing
End Sub

'@Description obtiene la sesion actual sobre la cual trabaja el Script SAP
'@Return objeto GuiSession que representa la sesion sobre la cual trabaja el Script SAP
Public Function GetCurrentSession() As GuiSession
    Set GetSession = session
End Function

'@Description obtiene la ultima sesion creada o establecida
'@Return objeto GuiSession que representa la ultima sesion creada o establecida
Public Function GetLastSession() As GuiSession
    Set GetLastSession = sapConnection.Children(GetSessionsCount() - 1)
End Function

'@Description abre una transaccion en un modo nuevo o una ventana nueva en SAP Logon
'@Param transactionCode codigo de la transaccion que desea abrir en un modo nuevo
'@Return objeto GuiSession que representa la sesion sobre la cual trabaja el Sctipt SAP
Public Function OpenTransactionInNewModo(ByVal transactionCode As String) As GuiSession
    Dim sessionsCountBeforeModoNew As Integer
    Set session = OpenNewModo()
    sessionsCountBeforeModoNew = GetSessionsCount()
    session.SendCommand transactionCode
    Set OpenTransactionInNewModo = session
End Function

'@Description obtiene el handler de la ventanada actualmente activa en SAP en la session sobre la cual
'esta trabajando el Script
'@Return Long handler de la ventana
Public Function GetActiveWindowHandle() As Long
    Dim sessionActiveWindow As GuiFrameWindow
    Set sessionActiveWindow = session.activeWindow
    GetActiveWindowHandle = sessionActiveWindow.Handle
End Function

'@Description se espera hasta obtener un tipo de mensaje retornado por la barra de estado.
'Este metodo suele ser de utilidad al momento de exportar un documento con la ventana de guardar como, de windows.
'¿Por que?
'Por que, al momento de abrir la ventana de Guardar como, de windows, esta ventana se manipula usando Windows API, entonces, SAP
'le sede el control de ejecucion a Windows API, por lo que, ahora SAP no sabe en que momento se guardara el documento, se cerrara la ventana de
'Guardar como, y se cerrara la ventana de Seleccion calculo de costes tabla de SAP. Entonces, al esperar un tipo de mensaje de respuesta de la
'barra de estatus, nos aseguramos de saber que el documento ya ha sido exportado y asi recuperamos el control de ejecucion en SAP.
Public Function WaitStatusBarMessageType() As String
    Dim statusBar As GuiStatusbar
    Dim messageType As String
    messageType = vbNullString
    Set statusBar = session.findById("wnd[0]/sbar")
    On Error Resume Next
    While Not messageType Like "[a-zA-Z]"
        messageType = statusBar.messageType
    Wend
    WaitStatusBarMessageType = messageType
End Function

'@Description comprueba si un componente de interfaz grafica se encuentra presente
'@Param id identificador del componente de interfaz grafica
'@Return True si el componente existe, de lo contrario False
Public Function ContainsGuiComponent(ByVal id As String) As Boolean
    If Not session.findById(id, False) Is Nothing Then
        ContainsGuiComponent = True
    Else
        ContainsGuiComponent = False
    End If
End Function
