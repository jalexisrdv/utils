Attribute VB_Name = "SAPLogonUtils"
'Libreria utilizada: SAP Logon Control
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'SAP Logon Control'.
'Por defecto la libreria no aparece, debe ser agregada mediante el archivo OCX, para ello pulsar el boton "Examinar"
'y abrir la siguiente ruta: C:\Program Files (x86)\SAP\FrontEnd\SAPgui\
'dentro de la ruta seleccione el archivo wdtlog.ocx

'@Resume SAPLogonUtils ofrece funciones que facilitan establecer conexion con el sistema SAP, a traves de
'SAP Logon Control, componente ofrecido por SAP para establecer conexion mientras se trabaja con interfaces de comunicacion de RFC y BAPI.
'RFC y BAPI son protocolos que permiten ejecutar transacciones y funciones sin necesidad de manipular la interfaz grafica de usuario de SAP.

Option Explicit
Option Private Module

Private Const TLO_RFC_NOT_CONNECTED As String = "0"
Private Const TLO_RFC_CONNECTED As String = "1"
Private Const TLO_RFC_CONNECT_CANCEL As String = "2"
Private Const TLO_RFC_NOT_CONNECT_PARAMETER_MISSGING As String = "4"
Private Const TLO_RFC_CONNECT_FAILED As String = "8"

Private mySapLogonControl As New sapLogonCtrl.sapLogonControl
Private myConnection As sapLogonCtrl.connection

'@Description crea una nueva conexion con el sistema de SAP a traves de SAP Logon Control
Public Sub NewConnection()
    Set myConnection = mySapLogonControl.NewConnection
End Sub

'@Description obtiene la conexion establecida mediante SAP Logon Control
'@Return sapLogonCtrl.connection objeto que representa la conexion establecida
Public Function GetConnection() As sapLogonCtrl.connection
    Set GetConnection = myConnection
End Function

'@Description establece al tipo de sistema con el cual se establecera la conexion,
'ya que SAP maneja diferentes sistemas con diferente proposito, ejemplo: Tenaris - PROD
'@Param system_ nombre del sistema con el cual se establecera la conexion
Public Sub SetSystem(ByVal system_ As String)
    myConnection.System = system_
End Sub

'@Description obtiene el nombre del sistema con el cual se establecio la conexion
'@Return nombre del sistema al que se establecio conexion
Public Function GetSystem() As String
    GetSystem = myConnection.System
End Function

'@Description establece el lenguaje con el cual se inicia el sistema de SAP, ejemplo: ES, EN, ...
'@Param language_ abreviacion del lenguaje a iniciar con el sistema de SAP
Public Sub SetLanguage(ByVal language_ As String)
    myConnection.Language = language_
End Sub

'@Description obtiene el lenguaje con el cual se inicio el sistema de SAP
'@Return lenguaje con el cual se inicio el sistema SAP
Public Function GetLanguage() As String
    GetLanguage = myConnection.Language
End Function

'@Description establece numero de mando (cliente) con el cual se establece la conexion
'@Param client_ numero de mando
Public Sub SetClient(ByVal client_ As String)
    myConnection.Client = client_
End Sub

'@Description obtiene el numero de mando (cliente) con el cual se establecio la conexion
'@Return client_ numero de mando
Public Function GetClient() As String
    GetClient = myConnection.Client
End Function

'@Description establece usuario con el cual se establecera la conexion
'@Param user_ usuario con el cual iniciara sesion
Public Sub SetUser(ByVal user_ As String)
    myConnection.User = user_
End Sub

'@Description obtiene usuario con el cual se establecera la conexion
'@Return user_ usuario con el cual se inicio sesion
Public Function GetUser() As String
    GetUser = myConnection.User
End Function

'@Description establece la contraseña para establecer la conexion
'@Param password_ contraseña del usuario
Public Sub SetPassword(ByVal password_ As String)
    myConnection.Password = password_
End Sub

'@Description obtiene la contraseña con la cual el usuario inicio sesion
'Return obtiene la contraseña del usuario logueado
Public Function GetPassword() As String
    GetPassword = myConnection.Password
End Function

'@Description inicia sesion al sistema de SAP con los parametros establecidos
'@Param hWnd manejador de la ventana donde se activara la ventana modal para iniciar sesion (esto con la finalidad de centrar la ventana modal)
'@Param bSilent indica si se iniciara la ventana modal para que el usuario ingrese los datos o se establecera el inicio de sesion sin ventana modal
Public Sub Logon(Optional ByVal hWnd As Long = 0, Optional ByVal bSilent As Boolean = False)
    myConnection.Logon hWnd, bSilent
    Select Case myConnection.IsConnected
        Case TLO_RFC_NOT_CONNECTED
            MsgBox "La conexión no fue establecida"
        Case TLO_RFC_CONNECTED
            MsgBox "La conexión fue establecida"
        Case TLO_RFC_CONNECT_CANCEL
            MsgBox "La conexión fue cancelada"
        Case TLO_RFC_NOT_CONNECT_PARAMETER_MISSGING
            MsgBox "La conexión no fue establecida, datos olvidados"
        Case TLO_RFC_CONNECT_FAILED
            MsgBox "La conexión falló, datos incorrectos"
    End Select
End Sub

'@Description cierra la conexion o sesion establecida al sistema SAP
Public Sub LogoOff()
    myConnection.Logoff
End Sub
