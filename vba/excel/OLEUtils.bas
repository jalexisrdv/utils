Attribute VB_Name = "OLEUtils"
'@Overview OLEUtils proporciona funciones que permiten desabilitar el mensaje emergente:
'"Microsoft Excel is waiting for another application to complete an OLE action"
'el mensaje anterior sale cuando una aplicacion externa tarda en dar respuesta, lo cual provoca que la ejecucion de la macro
'se detenga, y el usuario tenga que cerrar la ventana emergente para continuar la ejecucion de la macro.
'para evitar este problema, debe usar las 2 funciones proporcionadas para desabilitar el mensaje emergente y asi evitar que la
'ejecucion de la macro se detenga. El codigo debe estar encerrado entre las 2 funciones proporcionadas, ejemplo:
'
'START_DISABLE_POP_UP_FOR_OLE
'
'tu codigo aqui
'
'END_DISABLE_POP_UP_FOR_OLE
'

Option Private Module
Option Explicit

#If VBA7 Then

Private Declare PtrSafe Function CoRegisterMessageFilter Lib "OLE32.DLL" ( _
    ByVal lFilterIn As Long, _
    ByRef lPreviousFilter _
) As LongPtr
    
#Else

Private Declare Function CoRegisterMessageFilter Lib "OLE32.DLL" ( _
    ByVal lFilterIn As Long, _
    ByRef lPreviousFilter _
) As Long
    
#End If

Private lMsgFilter As Long

'@Description inicia el desbloqueo de la ventana emergente de OLE
'(desactiva el mensaje emergente: Microsoft Excel is waiting for another application to complete an OLE action)
Public Sub START_DISABLE_POP_UP_FOR_OLE()
    CoRegisterMessageFilter 0&, lMsgFilter
End Sub

'@Description finaliza el desbloqueo de la ventana emergente de OLE
Public Sub END_DISABLE_POP_UP_FOR_OLE()
    CoRegisterMessageFilter lMsgFilter, lMsgFilter
End Sub


