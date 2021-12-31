Attribute VB_Name = "WindowsAPIUtils"
'Libreria utilizada: Winuser.h
'Documentacion oficial: https://docs.microsoft.com/en-us/windows/win32/api/winuser/

'@Overview WindowsAPIUtils proporciona funciones que facilitan la manipulacion de ventanas,
'permitiendo enviar pulsaciones de teclas, enviar texto, enviar pulsacion de click, cerrar ventanas, y mas.
'Sin embargo, todo esto se realiza solo para el manejador recuperado de la ventana especificada (evitando enviar mensajes a otras ventanas)
'por lo que no hace falta tener el foco de la ventana objetivo activado.

'Founded: https://www.vbforums.com/showthread.php?208430-Use-sendmessage-to-close-an-application
'Founded: https://www.vbforums.com/showthread.php?323180-sending-text-to-other-applications
'Founded: https://fracta.net/fracta/index.php/forum/2-excel-vba-forum/8-excel-vba-send-keys-to-another-application-using-vba-sendkeys-command-and-user32dll-postmessage
'Keycode constants: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keycode-constants
'Los enlaces anteriores sirven como guia, sin embargo, el codigo fue adaptado y mejorado. Ademas, se solucionaron
'algunos errores presentados durante la implementacion.
'El TextBoxClass = lpsz1, TextBoxCaption = lpsz2, y lpClassName se pueden visualizar con la herramienta Microsoft Spy++

Option Explicit
Option Private Module

Private Const WM_CLOSE = &H10
Private Const WM_QUIT = &H12
Private Const WM_DESTROY = &H2
Private Const WM_NCDESTROY = &H82
Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const BM_CLICK = &HF5

'@Description busca la ventana correspondiente, mediante el nombre de titulo o nombre de clase de la ventana.
'@Param lpClassName nombre de clase de la ventana
'@Param lpWindowName nombre de la ventana o su nombre de titulo
'@Return valor Long que representa el manejador de la ventana
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

'@Description similar a FindWindow, sin embargo, FindWindowEx hace uso del manejador de ventana devuelta por FindWindow,
'FindWindowEx se encarga de buscar el elemento hijo de interfaz grafica dentro del manejador de ventana proporcionado por FindWindow.
'@Param hWndParent manejador padre donde se buscara el manejador del componente hijo
'@Param hWndChildAfter busca el siguiente  manejador del componente hijo dentro del manejador padre (es decir, busca todos los componentes hijos que existan)
'@Param lpszClass nombre de clase del componente (elemento) de interfaz grafica (ejemplo de nombre de clase: Button)
'@Param lpszWindow nombre del texto que contiene el elemento de interfaz grafica (ejemplo de nombre de texto: Cancelar),
'el ejemplo anterior es texto contenido dentro del botón (elemento de interfaz grafica)
'@Return valor Long que representa el manejador del componente gui
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String _
) As Long

'@Description Coloca (publica) un mensaje en la cola de mensajes asociada con el proceso que creó la ventana especificada
'y devuelve un valor sin esperar a que el proceso o hilo procese el mensaje.
'@Param hwnd manejador de la ventana a enviar mensaje
'@Param wMsg tipo de mensaje a enviar (consulte: https://docs.microsoft.com/en-us/windows/win32/winmsg/about-messages-and-message-queues)
'@Param wParam informacion adicional especifica del mensaje
'@Param lParam informacion adicional especifica del mensaje
'@Return valor Long, diferente de 0 si la funcion sucede o se ejecuta, 0 de lo contrario
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Any _
) As Long

'@Description Envía el mensaje especificado a una ventana o ventanas. La función SendMessage llama al procedimiento de ventana
'para la ventana especificada y no devuelve un valor hasta que el procedimiento de ventana ha procesado el mensaje.
'@Param hwnd manejador de la ventana a enviar mensaje
'@Param wMsg tipo de mensaje a enviar (consulte: https://docs.microsoft.com/en-us/windows/win32/winmsg/about-messages-and-message-queues)
'@Param wParam informacion adicional especifica del mensaje
'@Param lParam informacion adicional especifica del mensaje
'@Return valor Long, dicho valor depende del tipo de mensaje enviado
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Any _
) As Long

'@Description recupera el texto/titulo del manejador
'@Param hWnd manejador que contiene el texto a recuperar
'@Param lpString recibe variable a la cual se le establecera el texto recuperado (se asigna mediante referencia)
'@Param nMaxCount el numero maximo de caracteres que seran asignados a la variable pasada como argumento en el parametro lpString
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal nMaxCount As Long _
) As Long

'@Description recupera la longitud en caracteres del manejador
'@Param hWnd manejador del cual se recuperara la longitud en caracteres del texto/titulo contenido
'@Return Long, longitud en caracteres del texto
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
    ByVal hWnd As Long _
) As Long

'@Description encuentra el manejador (hwnd) de la ventana mediante el nombre de titulo o clase de la ventana
'@Param lpClassName nombre de clase de la ventana
'@Param lpWindowName nombre de la ventana o su nombre de titulo
'@Return valor Long que representa el manejador de la ventana, si el manejador no es encontrado se devuelve 0
Public Function FindWindowHandler(Optional ByVal lpClassName As String = vbNullString, Optional ByVal lpWindowName As String = vbNullString) As Long
    Dim hWnd As Long
    hWnd = FindWindow(lpClassName, lpWindowName)
    FindWindowHandler = hWnd
End Function

'@Description encuentra el manejador de la ventana mediante el nombre de titulo o clase de la ventana, a diferencia de FindWindowHandler
'WaitUntilFindWindowHandler espera hasta que la ventana este visible para recuperar la direccion del manejador
'@Param lpClassName nombre de clase de la ventana
'@Param lpWindowName nombre de la ventana o su nombre de titulo
'@Return valor Long que representa el manejador de la ventana
Public Function WaitUntilFindWindowHandler(Optional ByVal lpClassName As String = vbNullString, Optional ByVal lpWindowName As String = vbNullString) As Long
    Dim hWnd As Long
    hWnd = 0
    Do While hWnd = 0
        Application.Wait (Now + TimeValue("0:00:1")) 'un segundo
        hWnd = FindWindow(lpClassName, lpWindowName)
    Loop
    WaitUntilFindWindowHandler = hWnd
End Function

'@Description encuentra el manejador de un componente gui o elemento hijo dentro del manejador de la ventana recuperada con FindWindowHandler
'@Param hwnd manejador de la ventana en la cual se buscara algun componente hijo de interfaz grafica de usuario
'@Param componentClassName nombre de clase del componente a buscar
'@Param componentCaptionName nombre de caption del componente a buscar
'Return valor Long que representa el manejador del componente hijo de la interfaz grafica de usuario, si el manejador no es encontrado se devuelve 0
Public Function FindWindowComponentHandler(Optional ByVal hWnd As Long = 0, Optional ByVal componentClassName As String = vbNullString, Optional ByVal componentCaptionName As String = vbNullString) As Long
    Dim hWndChild As Long
    hWndChild = FindWindowEx(hWnd, 0, componentClassName, componentCaptionName)
    FindWindowComponentHandler = hWndChild
End Function

'@Description cierra la ventana del manejador de ventana indicado
'@Param hwnd manejador de la ventana a cerrar
'@Return True si la funcion se ejecuta o sucede, False de lo contrario
Public Function CloseWindow(ByVal hWnd As Long) As Boolean
    CloseWindow = PostMessage(hWnd, WM_CLOSE, CLng(0), CLng(0)) <> 0
End Function

'@Description envia simulacion de pulsacion de una tecla al manejador de la ventana o al manejador del componente gui de la ventana
'@Param hwnd manejador de la ventana a enviar simulacion de pulsacion de una tecla
'@Param keyCodeConstant constante del codigo de tecla a simular pulsacion (ejemplo: vbKeyReturn)
'@Return True si la funcion se ejecuta o sucede, False de lo contrario
Public Function SendKeyToWindow(ByVal hWnd As Long, ByVal keyCodeConstant As Long) As Boolean
    SendKeyToWindow = PostMessage(hWnd, WM_KEYDOWN, keyCodeConstant, CLng(0)) <> 0
End Function

'@Description envia simulacion de pulsacion de un click al manejador de un componente gui de la ventana
'@Param hwndchild manejador del componente gui del manejador de una ventana
Public Sub SendClickToWindowComponent(ByVal hWndChild As Long)
    SendMessage hWndChild, BM_CLICK, CLng(0), CLng(0)
End Sub

'@Description envia texto al manejador de un componente gui de la ventana
'@Param hwndchild manejador del componente gui del manejador de una ventana
'@Param text texto a enviar al componente gui
Public Sub SendTextToWindowComponent(ByVal hWndChild As Long, ByVal text As String)
    SendMessage hWndChild, WM_SETTEXT, 0, text
End Sub

'@Description imprime todos los captions los cuales son utilizados para buscar componentes hijos (handlers hijos) dentro de un manejador padre
'@Param hWnd handler o manejador padre donde se buscaran los componentes hijos (handlers hijos)
'@Description componentClassName nombre de clase de los handlers a imprimir su caption, ejemplo de clase: Button
Public Sub PrintCaptionsOfContainedComponentHandlers(ByVal hWnd As Long, Optional ByVal componentClassName As String = vbNullString)
    Dim hWndChild As Long
    Dim sChildText As String
    Do
        hWndChild = FindWindowEx(hWnd, hWndChild, componentClassName, vbNullString)
        If hWndChild <> 0 Then
            sChildText = String(GetWindowTextLength(hWndChild) + 1, Chr(0))
            GetWindowText hWndChild, sChildText, Len(sChildText)
            Debug.Print sChildText
        End If
    Loop While hWndChild <> 0
End Sub
