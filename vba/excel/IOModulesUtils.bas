Attribute VB_Name = "IOModulesUtils"
'Libreria utilizada: Microsoft Visual Basic for Applications Extensibility 5.3
'La libreria puede ser agregada desde el menu 'Herramientas' - 'Referencias' - 'Microsoft Visual Basic for Applications Extensibility 5.3'

'@Overview IOModulesUtils permite exportar e importar el codigo del proyecto, el codigo se exporta en la carpeta de documentos
'creando un folder principal llamado VBAProjectFiles, dentro se crea un subfolder con el nombre del documento Excel, dentro de este subfolder
'se agrega el codigo del proyecto.
'
'Una vez tengas el codigo del proyecto en la carpeta correspondiente, recomiendo usar Git para que lleves un control de las modificaciones, y versiones del proyecto,
'para comenzar a usar git, primero debes crear el repositorio en la cuenta de github, una vez creado este te proporciona una URL.
'obtenida la url puedes abrir la consola Git Bach y aplicar el comando:
'
'git clone URL .
'
'El metodo ExportModules debe ser invocado antes de guardar el documento, con la finalidad de crear un respaldo automatico en su carpeta correspondiente,
'para ello debemos agregar el siguiente codigo en Thisworkbook:
'
'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    IOModulesUtils.ExportModules
'End Sub
'
'Entonces, cada que se guarde el documento se creara un respaldo del codigo y como ya hemos creado o clonado nuestro repositorio,
'git automaticamente detectara los cambios realizados al proyecto, esto se detecta mediante los archivos o modulos exportados.
'
'Recomiendo usar GIT ya que asi podremos llevar un historial del desarrollo, ademas a ti como desarrollador te brindara mayor experiencia ;)
'Git o un control de versiones es utilizado en la industria, asi que es un plus para ti, si o si debes aprenderlo.

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim ignoreModule As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.cls"
        Kill FolderWithVBAProjectFiles & "\*.frm"
        Kill FolderWithVBAProjectFiles & "\*.bas"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name
        
        ignoreModule = "IOModulesUtils"
        If InStr(1, szFileName, ignoreModule, vbTextCompare) <> 0 Then
            bExport = False
        End If

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

'    MsgBox "Export is ready"
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
        
    Next objFile
    
'    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath, projectName, path As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    projectName = fso.GetBaseName(ThisWorkbook.Name)

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    path = SpecialPath & "VBAProjectFiles"
    
    If fso.FolderExists(path) = False Then
        On Error Resume Next
        MkDir path
        On Error GoTo 0
    End If
    
    path = SpecialPath & "VBAProjectFiles\" & projectName
    
    If fso.FolderExists(path) = False Then
        On Error Resume Next
        MkDir path
        On Error GoTo 0
    End If
    
    If fso.FolderExists(path) = True Then
        FolderWithVBAProjectFiles = path
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
