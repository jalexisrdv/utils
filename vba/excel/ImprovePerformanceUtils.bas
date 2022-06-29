Attribute VB_Name = "ImprovePerformanceUtils"
'@Overview ImprovePerformanceUtils establece las configuraciones necesarias para mejorar el rendimiento de la ejecucion de la macro,
'al terminar la configuracion se restablecen las configuraciones a su valor por defecto.
'Reference: https://techcommunity.microsoft.com/t5/excel/9-quick-tips-to-improve-your-vba-macro-performance/m-p/173687

Option Explicit
Option Private Module

Private saveCalculation As Long
Private saveScreenUpdating As Boolean
Private saveEnableAnimations  As Boolean
Private saveEnableEvents  As Boolean

'@Description establece las configuraciones recomendadas para mejorar el rendimiento de ejecucion de la macro
Public Sub START_SET_SETTINGS()
    With Application
        saveCalculation = .Calculation
        saveScreenUpdating = .ScreenUpdating
        saveEnableAnimations = .EnableAnimations
        saveEnableEvents = .EnableEvents
        
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableAnimations = False
        .EnableEvents = False
    End With
End Sub

'@Description restablece las configuraciones a su valor por defecto
Public Sub END_SET_SETTINGS()
    With Application
        .Calculation = saveCalculation
        .ScreenUpdating = saveScreenUpdating
        .EnableAnimations = saveEnableAnimations
        .EnableEvents = saveEnableEvents
    End With
End Sub

