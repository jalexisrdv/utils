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
Private Const PROCESS_TERMINATE = &H1

#If VBA7 Then
    Private parentHandler As LongPtr
    Private hWndChild As LongPtr
    Private hWnd As LongPtr
#Else
    Private parentHandler As Long
    Private hWndChild As Long
    Private hWnd As Long
#End If

Private lpClassName As String
Private position As Long
Private i As Long
Private regex As Object

'@Description busca la ventana correspondiente, mediante el nombre de titulo o nombre de clase de la ventana.
'@Param lpClassName nombre de clase de la ventana
'@Param lpWindowName nombre de la ventana o su nombre de titulo
'@Return valor Long que representa el manejador de la ventana
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

'@Description similar a FindWindow, sin embargo, FindWindowEx hace uso del manejador de ventana devuelta por FindWindow,
'FindWindowEx se encarga de buscar el elemento hijo de interfaz grafica dentro del manejador de ventana proporcionado por FindWindow.
'@Param hWndParent manejador padre donde se buscara el manejador del componente hijo
'@Param hWndChildAfter busca el siguiente  manejador del componente hijo dentro del manejador padre (es decir, busca todos los componentes hijos que existan)
'@Param lpszClass nombre de clase del componente (elemento) de interfaz grafica (ejemplo de nombre de clase: Button)
'@Param lpszWindow nombre del texto que contiene el elemento de interfaz grafica (ejemplo de nombre de texto: Cancelar),
'el ejemplo anterior es texto contenido dentro del botón (elemento de interfaz grafica)
'@Return valor Long que representa el manejador del componente gui
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
#Else
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
#End If

'@Description Coloca (publica) un mensaje en la cola de mensajes asociada con el proceso que creó la ventana especificada
'y devuelve un valor sin esperar a que el proceso o hilo procese el mensaje.
'@Param hwnd manejador de la ventana a enviar mensaje
'@Param wMsg tipo de mensaje a enviar (consulte: https://docs.microsoft.com/en-us/windows/win32/winmsg/about-messages-and-message-queues)
'@Param wParam informacion adicional especifica del mensaje
'@Param lParam informacion adicional especifica del mensaje
'@Return valor Long, diferente de 0 si la funcion sucede o se ejecuta, 0 de lo contrario
#If VBA7 Then
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
#Else
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
#End If

'@Description Envía el mensaje especificado a una ventana o ventanas. La función SendMessage llama al procedimiento de ventana
'para la ventana especificada y no devuelve un valor hasta que el procedimiento de ventana ha procesado el mensaje.
'@Param hwnd manejador de la ventana a enviar mensaje
'@Param wMsg tipo de mensaje a enviar (consulte: https://docs.microsoft.com/en-us/windows/win32/winmsg/about-messages-and-message-queues)
'@Param wParam informacion adicional especifica del mensaje
'@Param lParam informacion adicional especifica del mensaje
'@Return valor Long, dicho valor depende del tipo de mensaje enviado
#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
#End If

'@Description recupera el texto/titulo del manejador
'@Param hWnd manejador que contiene el texto a recuperar
'@Param lpString recibe variable a la cual se le establecera el texto recuperado (se asigna mediante referencia)
'@Param nMaxCount el numero maximo de caracteres que seran asignados a la variable pasada como argumento en el parametro lpString
'@Return 0 si el handler no tiene texto, si tiene texto devuelve cualquier numero <> 0
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#Else
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#End If

'@Description recupera la longitud en caracteres del manejador
'@Param hWnd manejador del cual se recuperara la longitud en caracteres del texto/titulo contenido
'@Return Long, longitud en caracteres del texto
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
#End If

'@Description enumera todas las ventanas activas
'@Param lpEnumFunc callback function la cual sera invocada por cada handler, los cuales fueron enumerados
'@Param lParam un valor definido por la aplicación que se pasará a la función de devolución de llamada.
'@Return el valor de retorno no es usado
#If VBA7 Then
    Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
#Else
    Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
#End If

'@Description enumera todos los manejadores o componendes hijos del handler especificado
'@Param hWndParent handler del cual se quiere conocer todos sus componentes hijos (handlers hijos), normalmente
'se pasa un handler de ventana
'@Param lpEnumFunc callback function la cual sera invocada por cada handler, los cuales fueron enumerados
'@Param lParam un valor definido por la aplicación que se pasará a la función de devolución de llamada.
'@Return el valor de retorno no es usado
#If VBA7 Then
    Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
#Else
    Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
#End If
                                
'@Description verifica si un handler se encuentra visible
'@Param hWnd handler del cual requiere saber si esta visible
'@Return Long, 0 si la ventana no esta visble, de lo contrario cualquier numero diferente a 0
#If VBA7 Then
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
#Else
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
#End If
                                                  
'@Description recupera el nombre de clase de un handler
'@Param hWnd handler del cual se quiere recuperar el nombre de clase
'@Param lpClassName variable de tipo string en la cual se guardara el nombre de clase
'@Param nMaxCount longitud de caracteres de lpClassName
'@Return Long, 0 si no existe nombre de clase, de lo contrario devuelve la longitud del nombre de clase
#If VBA7 Then
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
#Else
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
#End If

'@Description obtiene el handler padre o propietario del handler pasado como argumento
'@Param hWnd handler del cual se quiere recuperar el handler padre
'@Return Long handler padre
#If VBA7 Then
    Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
#Else
    Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
#End If

'@Description recupera el ID del Thread que creo la ventana, opcionalmente
'si es pasada una variable en el segundo argumento, en esta se almacenara el ID
'del proceso que creo la ventana
'@Param hWnd handler o manejador de la ventana
'@Param lpdwProcessId variable como referencia donde se guardara el ID del proceso que creo la ventana
'@Return Long ID del Thread que creo la ventana
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
#Else
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
#End If

'@Description abre un proceso local existente
'@Param dwDesiredAccessas permite indicar el tipo de acceso al proceso
'@Param bInheritHandle Si este valor es VERDADERO, los procesos creados por este proceso heredarán el identificador.
'De lo contrario, los procesos no heredan este identificador.
'@Param dwProcId ID del proceso
'@Return Long handler o manejador del proceso con el ID establecido, 0 si no encuentra un manejador
#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As LongPtr
#Else
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
#End If

'@Description abre un Thread local existente
'@Param dwDesiredAccess permite indicar el tipo de acceso al Thread
'@Param bInheritHandle Si este valor es VERDADERO, los procesos creados por este proceso heredarán el identificador.
'De lo contrario, los procesos no heredan este identificador.
'@Param dwThreadId ID del proceso
'@Return Long handler o manejador del Thread con el ID establecido, 0 si no encuentra un manejador
#If VBA7 Then
    Private Declare PtrSafe Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As LongPtr
#Else
    Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
#End If

'@Description termina el proceso espeficificado
'@Param hProcess handler del proceso que se desea terminar
'@Param uExitCode se puede establecer cualquier valor, el valor de uExitCode puede ser alterado por el proceso
'dicho valor puede ser recuperado con GetExitCodeProcess
'@Return LongPtr 0 si la no se termina ningun proceso, de lo contrario devuelve un numero diferente
#If VBA7 Then
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
#Else
    Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
#End If

'@Description termina el Thread(hilo) espeficificado
'@Param hThread handler del Thread(hilo) que se desea terminar
'@Param dwExitCode se puede establecer cualquier valor, el valor de uExitCode puede ser alterado por el manejador del hThread
'dicho valor puede ser recuperado con GetExitCodeThread
'@Return LongPtr 0 si la no se termina ningun proceso, de lo contrario devuelve un numero diferente
#If VBA7 Then
    Private Declare PtrSafe Function TerminateThread Lib "kernel32" (ByVal hThread As LongPtr, ByVal dwExitCode As Long) As Long
#Else
    Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
#End If

'@Description cierra un handler de objeto abierto
'@Param un handler valido de un objeto abierto, por ejemplo: el handler de un proceso (hProcess)
'@Return 0 si la funcion no se ejecuta correctamente, de lo contrario un numero <> 0
#If VBA7 Then
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

'@Description encuentra el manejador (hwnd) de la ventana mediante el nombre de titulo o clase de la ventana
'@Param lpClassName nombre de clase de la ventana
'@Param lpWindowName nombre de la ventana o su nombre de titulo
'@Return valor Long que representa el manejador de la ventana, si el manejador no es encontrado se devuelve 0
#If VBA7 Then
    Public Function FindWindowHandler(Optional ByVal lpClassName As String = vbNullString, Optional ByVal lpWindowName As String = vbNullString) As LongPtr
#Else
    Public Function FindWindowHandler(Optional ByVal lpClassName As String = vbNullString, Optional ByVal lpWindowName As String = vbNullString) As Long
#End If
    hWnd = FindWindow(lpClassName, lpWindowName)
    FindWindowHandler = hWnd
End Function

'@Description encuentra el manejador de la ventana mediante el nombre de titulo o clase de la ventana, a diferencia de FindWindowHandler
'WaitUntilFindWindowHandler espera hasta que la ventana este visible para recuperar la direccion del manejador
'@Param lpClassName nombre de clase de la ventana
'@Param lpWindowName nombre de la ventana o su nombre de titulo
'@Return valor Long que representa el manejador de la ventana
#If VBA7 Then
    Public Function WaitUntilFindWindowHandler(Optional ByVal lpClassName As String = vbNullString, Optional ByVal lpWindowName As String = vbNullString) As LongPtr
#Else
    Public Function WaitUntilFindWindowHandler(Optional ByVal lpClassName As String = vbNullString, Optional ByVal lpWindowName As String = vbNullString) As Long
#End If
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
#If VBA7 Then
    Public Function FindWindowComponentHandler(Optional ByVal hWnd As LongPtr = 0, Optional ByVal componentClassName As String = vbNullString, Optional ByVal componentCaptionName As String = vbNullString) As LongPtr
#Else
    Public Function FindWindowComponentHandler(Optional ByVal hWnd As Long = 0, Optional ByVal componentClassName As String = vbNullString, Optional ByVal componentCaptionName As String = vbNullString) As Long
#End If
    hWndChild = FindWindowEx(hWnd, 0, componentClassName, componentCaptionName)
    FindWindowComponentHandler = hWndChild
End Function

'@Description cierra la ventana del manejador de ventana indicado
'@Param hwnd manejador de la ventana a cerrar
'@Return True si la funcion se ejecuta o sucede, False de lo contrario
#If VBA7 Then
    Public Function CloseWindow(ByVal hWnd As LongPtr) As Boolean
#Else
    Public Function CloseWindow(ByVal hWnd As Long) As Boolean
#End If
    CloseWindow = PostMessage(hWnd, WM_CLOSE, CLng(0), CLng(0)) <> 0
End Function

'@Description envia simulacion de pulsacion de una tecla al manejador de la ventana o al manejador del componente gui de la ventana
'@Param hwnd manejador de la ventana a enviar simulacion de pulsacion de una tecla
'@Param keyCodeConstant constante del codigo de tecla a simular pulsacion (ejemplo: vbKeyReturn)
'@Return True si la funcion se ejecuta o sucede, False de lo contrario
#If VBA7 Then
    Public Function SendKeyToWindow(ByVal hWnd As LongPtr, ByVal keyCodeConstant As Long) As Boolean
#Else
    Public Function SendKeyToWindow(ByVal hWnd As Long, ByVal keyCodeConstant As Long) As Boolean
#End If
    SendKeyToWindow = PostMessage(hWnd, WM_KEYDOWN, keyCodeConstant, CLng(0)) <> 0
End Function

'@Description envia simulacion de pulsacion de un click al manejador de un componente gui de la ventana
'@Param hwndchild manejador del componente gui del manejador de una ventana
#If VBA7 Then
    Public Sub SendClickToWindowComponent(ByVal hWndChild As LongPtr)
#Else
    Public Sub SendClickToWindowComponent(ByVal hWndChild As Long)
#End If
    SendMessage hWndChild, BM_CLICK, CLng(0), CLng(0)
End Sub

'@Description envia texto al manejador de un componente gui de la ventana
'@Param hwndchild manejador del componente gui del manejador de una ventana
'@Param text texto a enviar al componente gui
#If VBA7 Then
    Public Sub SendTextToWindowComponent(ByVal hWndChild As LongPtr, ByVal text As String)
#Else
    Public Sub SendTextToWindowComponent(ByVal hWndChild As Long, ByVal text As String)
#End If
    'se esperan 3 o X segundos necesearios, para que el componente se visualice y el texto enviado sea tomado
    'adecuadamente (ejemplo: en la ventana de Guardar como, es util esperar 3 segundos para que el texto enviado
    'al componente con nombre de clase Edit sea tomado correctamente al pulsar el boton Guardar mediante SendClickToWindowComponent)
    Application.Wait (Now + TimeValue("0:00:3")) '3 segundos
    SendMessage hWndChild, WM_SETTEXT, 0, text
End Sub

'@Description obtiene el texto o caption del handler
'@Param hWnd handler del cual se quiere recuperar el texto o caption
'@Return String texto o caption del handler
#If VBA7 Then
    Public Function GetHandlerText(ByVal hWnd As LongPtr) As String
#Else
    Public Function GetHandlerText(ByVal hWnd As Long) As String
#End If
    Dim caption As String
    Dim sanitizedCaption As String
    caption = Space$(256)
    If GetWindowText(hWnd, caption, Len(caption)) <> 0 Then
        sanitizedCaption = Left$(caption, InStr(caption, vbNullChar) - 1) 'se elimina caracter adicional y solo se deja el caption
        GetHandlerText = sanitizedCaption
    Else
        GetHandlerText = vbNullString
    End If
End Function

'@Description obtiene el nombre de clase del handler
'@Param hWnd handler del cual se quiere recuperar el nombre de clase
'@Return String nombre de clase del handler, ejemplo de clase: Button
#If VBA7 Then
    Public Function GetHandlerClassName(ByVal hWnd As LongPtr) As String
#Else
    Public Function GetHandlerClassName(ByVal hWnd As Long) As String
#End If
    Dim className As String
    Dim sanitizedClassName As String
    className = Space$(256)
    If GetClassName(hWnd, className, Len(className)) <> 0 Then
        sanitizedClassName = Left$(className, InStr(className, vbNullChar) - 1) 'se elimina caracter adicional y solo se deja el nombre de la clase
        GetHandlerClassName = sanitizedClassName
    Else
        GetHandlerClassName = vbNullString
    End If
End Function

'@Description obtiene el manejador de la ventana hijo, a partir del manejador de la ventana padre. Es decir, obtiene el manejador de la ventana hijo (segunda ventana)
'que fue creada (abierta) desde una ventana padre
'@Param parentHandler ventana padre a partir de la cual se recupera el manejador o ventanada hijo
'@Return Long manejador de la ventana hijo encontrada
#If VBA7 Then
    Public Function FindChildWindowHandler(ByVal parentHandler_ As LongPtr) As LongPtr
#Else
    Public Function FindChildWindowHandler(ByVal parentHandler_ As Long) As Long
#End If
    parentHandler = parentHandler_
    EnumWindows AddressOf FindChildWindowHandlerCallback, &H0
    FindChildWindowHandler = hWndChild
End Function

'@Description espera hasta obtener el manejador de la ventana hijo, a partir del manejador de la ventana padre. Es decir, obtiene el manejador de la ventana hijo (segunda ventana)
'que fue creada (abierta) desde una ventana padre
'@Param parentHandler ventana padre a partir de la cual se recupera el manejador o ventanada hijo
'@Return Long manejador de la ventana hijo encontrada
#If VBA7 Then
    Public Function WaitFindChildWindowHandler(ByVal parentHandler_ As LongPtr) As LongPtr
#Else
    Public Function WaitFindChildWindowHandler(ByVal parentHandler_ As Long) As Long
#End If
    hWnd = 0
    Do While hWnd = 0
        Application.Wait (Now + TimeValue("0:00:1")) 'un segundo
        hWnd = FindChildWindowHandler(parentHandler_)
    Loop
    WaitFindChildWindowHandler = hWnd
End Function

#If VBA7 Then
    Private Function FindChildWindowHandlerCallback(ByVal hWnd As LongPtr, ByVal lParam As Long) As Long
#Else
    Private Function FindChildWindowHandlerCallback(ByVal hWnd As Long, ByVal lParam As Long) As Long
#End If
    If IsWindowVisible(hWnd) Then
        If GetHandlerText(hWnd) <> vbNullString Then 'si tiene titulo la ventana o existe texto
            If parentHandler = GetParent(hWnd) Then
                hWndChild = hWnd
                Exit Function
            End If
        End If
    End If
    hWndChild = 0 'no se encontro nada, devuelve 0
    FindChildWindowHandlerCallback = 1
End Function

'@Description busca una ventana mediante su titulo, si el titulo de la ventana coincide con el patron regex significa que fue encontrada
'@Param pattern_ patron mediante regex, donse se indican las palabras con las cuales debe coincidir con el titulo de la ventana
'@Return Long manejador de la ventana encontrada
#If VBA7 Then
    Public Function FindWindowHandlerByCaption(ByVal pattern_ As String) As LongPtr
#Else
    Public Function FindWindowHandlerByCaption(ByVal pattern_ As String) As Long
#End If
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern_
    regex.IgnoreCase = True
    EnumWindows AddressOf FindWindowHandlerByCaptionCallback, &H0
    FindWindowHandlerByCaption = hWndChild
End Function


'@Description espera hasta encontrar una ventana mediante su titulo, si el titulo de la ventana coincide con el patron regex significa que fue encontrada
'@Param pattern_ patron mediante regex, donse se indican las palabras con las cuales debe coincidir con el titulo de la ventana
'@Return Long manejador de la ventana encontrada
#If VBA7 Then
    Public Function WaitFindWindowHandlerByCaption(ByVal pattern_ As String) As LongPtr
#Else
    Public Function WaitFindWindowHandlerByCaption(ByVal pattern_ As String) As Long
#End If
    hWnd = 0
    Do While hWnd = 0
        Application.Wait (Now + TimeValue("0:00:1")) 'un segundo
        hWnd = FindWindowHandlerByCaption(pattern_)
    Loop
    WaitFindWindowHandlerByCaption = hWnd
End Function

#If VBA7 Then
    Private Function FindWindowHandlerByCaptionCallback(ByVal hWnd As LongPtr, ByVal lParam As Long) As Long
#Else
    Private Function FindWindowHandlerByCaptionCallback(ByVal hWnd As Long, ByVal lParam As Long) As Long
#End If
    Dim caption As String
    If IsWindowVisible(hWnd) Then
        caption = GetHandlerText(hWnd)
        If caption <> vbNullString Then 'si tiene titulo la ventana o existe texto
            If regex.Test(caption) Then
                hWndChild = hWnd
                Exit Function
            End If
        End If
    End If
    hWndChild = 0 'no se encontro nada, devuelve 0
    FindWindowHandlerByCaptionCallback = 1
End Function

'@Description Encuentra un componente de acuerdo al indice de posicion
'@Param hWnd handler de la ventana padre a partir de la cual se recupera el handler del componente
'@Param position_ indice de posicion del componente a recuperar
'@Param lpClassName_ nombre de clase del handler de componente a recuperar
'@Return Long valor de handler del componente
#If VBA7 Then
    Public Function FindWindowComponentHandlerByPosition(ByVal hWnd As LongPtr, ByVal position_ As Long, Optional ByVal lpClassName_ As String = vbNullString) As LongPtr
#Else
    Public Function FindWindowComponentHandlerByPosition(ByVal hWnd As Long, ByVal position_ As Long, Optional ByVal lpClassName_ As String = vbNullString) As Long
#End If
    position = position_
    lpClassName = lpClassName_
    i = 0
    EnumChildWindows hWnd, AddressOf FindWindowComponentHandlerByPositionCallback, &H0
    FindWindowComponentHandlerByPosition = hWndChild
End Function

#If VBA7 Then
    Private Function FindWindowComponentHandlerByPositionCallback(ByVal hWnd As LongPtr, ByVal lParam As Long) As Long
#Else
    Private Function FindWindowComponentHandlerByPositionCallback(ByVal hWnd As Long, ByVal lParam As Long) As Long
#End If
    Dim className As String
    className = GetHandlerClassName(hWnd)
    Select Case lpClassName
        Case className
            If i = position Then
                hWndChild = hWnd
                Exit Function
            End If
            i = i + 1
        Case vbNullString
            If i = position Then
                hWndChild = hWnd
                Exit Function
            End If
            i = i + 1
    End Select
    hWndChild = 0 'no se encontro nada, devuelve 0
    FindWindowComponentHandlerByPositionCallback = 1
End Function

'@Description imprime informacion la cual te ayudara a buscar algun componente hijo, (ejemplo: la info ayuda a usar la funcion FindWindowComponentHandlerByPosition).
'El dato Position, indica el numero de indice. Por lo que, si se indica un nombre de clase (lpClassName_) el indice se ajusta solo a los elementos encontrados con dicha clase
'ejemplo: si existen 3 componentes de clase Button, las posiciones serian: 0, 1, 2. El numero de posicion es como si estuvieran almacenados en un arreglo los componentes encontrados.
'@Param hWnd handler o manejador padre donde se buscaran los componentes hijos (handlers hijos) (ejemplo: handler de la ventana)
'@Param lpClassName_ nombre de clase de los handlers a obtener informacion, ejemplo de clase: Button
#If VBA7 Then
    Public Sub PrintInfoComponentHandlers(ByVal hWnd As LongPtr, Optional ByVal lpClassName_ As String = vbNullString)
#Else
    Public Sub PrintInfoComponentHandlers(ByVal hWnd As Long, Optional ByVal lpClassName_ As String = vbNullString)
#End If
    i = 0
    lpClassName = lpClassName_
    EnumChildWindows hWnd, AddressOf PrintInfoComponentHandlersCallback, &H0
End Sub

#If VBA7 Then
    Private Function PrintInfoComponentHandlersCallback(ByVal hWnd As LongPtr, ByVal lParam As Long) As Long
#Else
    Private Function PrintInfoComponentHandlersCallback(ByVal hWnd As Long, ByVal lParam As Long) As Long
#End If
    Dim className As String
    className = GetHandlerClassName(hWnd)
    Select Case lpClassName
        Case className
            PrintInfoComponentHandler hWnd, className
            i = i + 1
        Case vbNullString
            PrintInfoComponentHandler hWnd, className
            i = i + 1
    End Select
    PrintInfoComponentHandlersCallback = 1
End Function

#If VBA7 Then
    Private Sub PrintInfoComponentHandler(ByVal hWnd As LongPtr, ByVal className As String)
#Else
    Private Sub PrintInfoComponentHandler(ByVal hWnd As Long, ByVal className As String)
#End If
    Debug.Print "========================= " & "Handler Info => " & hWnd & " ========================="
    Debug.Print "Parent Handler => " & GetParent(hWnd)
    Debug.Print "Caption => " & GetHandlerText(hWnd)
    Debug.Print "Class Name => " & className
    Debug.Print "Position => " & i
End Sub

'@Description termina la ejecucion del proceso que pertenece al manejador de ventana (hWnd)
'@Param hWnd handler o manejador de la venatana
'@Return True si el proceso fue terminado, False de lo contrario
#If VBA7 Then
    Function KillProcessByHwnd(ByVal hWnd As LongPtr) As Boolean
    Dim processID As LongPtr
    Dim processHandle As LongPtr
#Else
    Function KillProcessByHwnd(ByVal hWnd As Long)
    Dim processID As Long
    Dim processHandle As Long
#End If
    Dim succeeds As Long
    GetWindowThreadProcessId hWnd, processID
    processHandle = OpenProcess(PROCESS_TERMINATE, False, processID)
    succeeds = TerminateProcess(processHandle, 0)
    CloseHandle processHandle
    KillProcessByHwnd = succeeds <> 0
End Function

'@Description termina la ejecucion del Thread que pertenece al manejador de ventana (hWnd)
'@Param hWnd handler o manejador de la venatana
'@Return True si el Thread fue terminado, False de lo contrario
#If VBA7 Then
    Function KillThreadByHwnd(ByVal hWnd As LongPtr) As Boolean
    Dim threadID As LongPtr
    Dim threadHandle As LongPtr
#Else
    Function KillThreadByHwnd(ByVal hWnd As Long)
    Dim threadID As Long
    Dim threadHandle As Long
#End If
    Dim succeeds As Long
    threadID = GetWindowThreadProcessId(hWnd, 0)
    threadHandle = OpenThread(PROCESS_TERMINATE, False, threadID)
    succeeds = TerminateThread(threadHandle, 0)
    CloseHandle threadHandle
    KillThreadByHwnd = succeeds <> 0
End Function
