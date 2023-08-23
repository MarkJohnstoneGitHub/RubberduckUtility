VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RubberduckUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Rubberduck utility to export all components according to @Folder annotation"
'@ModuleDescription("Rubberduck utility to export all components according to @Folder annotation")
'@Folder("Rubberduck Utilities")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/RubberduckUtility
'@Version v1.0 August 18, 2023
'@LastModified August 18, 2023

'@ReferenceAddin VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)

'@References
' https://codereview.stackexchange.com/questions/248109/add-default-rubberduck-vba-folder-annotation-to-vbproject-vbcomponents
' https://codereview.stackexchange.com/questions/256782/export-import-rubberduck-vba-virtual-folder-structure-with-library-references-to
' https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files/17362688#17362688
' https://stackoverflow.com/a/27385063/10759363
' https://en.wikipedia.org/wiki/Path_(computing)#:~:text=The%20delimiting%20character%20is%20most,may%20use%20a%20different%20delimiter.

'Useage:
'   Exports all components in the active project to a base file location, overwritting any existing files.
'   RubberduckUtility.ExportAllComponents "C:\VBA\Output\"

Option Explicit
Private Const VBAErrorFileNotFound As Long = 76   ' DirectoryNotFoundException
Private Const DirectorySeparatorChar As String = "\"

Private Type TProject
    Project As VBIDE.VBProject
End Type

Private this As TProject

Private Sub Class_Initialize()
    With this
        Set .Project = Application.VBE.ActiveVBProject
    End With
End Sub

''
'@Description("Export the active project modules to file according to Rubberduck @Folder annotation.")
'@param workingDirectory file path to export to.
'@Exceptions: VBAErrorFileNotFound
'@Remarks
' If Rubberduck folder annotation contains invalid characters for a folder name
' the file will be exported to the working directory.
'@TODO Validate inputs, implement optional filter for a RD Folder annotation?
Public Sub ExportAllComponents(ByVal workingDirectory As String)
Attribute ExportAllComponents.VB_Description = "Export the active project modules to file according to Rubberduck @Folder annotation."
    Dim trimmedWorkingDirectory As String
    trimmedWorkingDirectory = Trim$(workingDirectory)

    'If invalid throws error
    If Not DirectoryExist(trimmedWorkingDirectory) Then
        Err.Raise VBAErrorFileNotFound, "RubberduckUtility.ExportAllComponents", "Working directory not found: " & trimmedWorkingDirectory
    End If
    
    Dim component As VBComponent
    For Each component In this.Project.VBComponents
        Dim filename As String
        filename = component.Name
        Dim extension As String
        extension = ComponentExtension(component)

        Dim folderAnnotation As String
        folderAnnotation = ParseFolderAnnotation(component)
        
        Dim relativePath As String
        relativePath = GetFilePathFromFolderAnnotation(folderAnnotation)
        Dim fullpath As String
        fullpath = trimmedWorkingDirectory & relativePath
        
        On Error Resume Next
        Dim fileFullPath As String
        CreateDirectory fullpath    ' create directory structure according to RD @Folder annotation
        If Err.Number = 0 Then
            fileFullPath = fullpath & IIf(relativePath <> VBA.vbNullString, DirectorySeparatorChar, VBA.vbNullString) & filename & extension
        Else
            'catch error for invalid path due to RD @Folder annotation and export to default base path.
            fileFullPath = trimmedWorkingDirectory & filename & extension
            Debug.Print "Invalid rubberduck folder annotation, "; folderAnnotation; " "; filename & extension; " exported to working directory "; trimmedWorkingDirectory
        End If
        On Error GoTo 0 'reset error handling
        component.Export fileFullPath
    Next
End Sub

'@Description "Obtains the Rubberduck @Folder annotation."
Private Function ParseFolderAnnotation(ByVal component As VBComponent) As String
Attribute ParseFolderAnnotation.VB_Description = "Obtains the Rubberduck @Folder annotation."
    Const RubberduckFolderAnnotation As String = "'@Folder"
    Const annotationStartMatch As String = """"
    Const annotationEndMatch As String = """"

    Dim componentDeclareLines As String
    componentDeclareLines = component.CodeModule.Lines(1, component.CodeModule.CountOfDeclarationLines)
    Dim annotateStart As Long
    annotateStart = InStr(1, componentDeclareLines, RubberduckFolderAnnotation, vbTextCompare)
    If annotateStart > 0 Then
        annotateStart = InStr(annotateStart, componentDeclareLines, annotationStartMatch)
        Dim annotateEnd As Long
        annotateEnd = InStr(annotateStart + 1, componentDeclareLines, annotationEndMatch)
        ParseFolderAnnotation = Mid$(componentDeclareLines, annotateStart + 1, annotateEnd - annotateStart - 1)
    End If
End Function

'@Description "Obtains the path from the Rubberduck @Folder annotation"
Private Function GetFilePathFromFolderAnnotation(ByVal folderAnnotation As String) As String
Attribute GetFilePathFromFolderAnnotation.VB_Description = "Obtains the path from the Rubberduck @Folder annotation"
    Const RubberduckPathSeparator As String = "."
    
    Dim outputPath As String
    If folderAnnotation <> VBA.vbNullString Then
        outputPath = Replace$(folderAnnotation, RubberduckPathSeparator, DirectorySeparatorChar)
    End If
    GetFilePathFromFolderAnnotation = outputPath
End Function

'@Description "Obtains the componet file extension according to its type."
Private Function ComponentExtension(ByVal component As VBComponent) As String
Attribute ComponentExtension.VB_Description = "Obtains the componet file extension according to its type."
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_StdModule
            ComponentExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_ClassModule
            ComponentExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_MSForm
            ComponentExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_Document
            ComponentExtension = ".doccls"
    End Select
End Function

'@Description "Creates all directories and subdirectories in the specified path unless they already exist."
Private Sub CreateDirectory(ByVal fullpath As String)
Attribute CreateDirectory.VB_Description = "Creates all directories and subdirectories in the specified path unless they already exist."
    Dim folders As Variant
    folders = Split(fullpath, DirectorySeparatorChar)

    Dim path As String
    path = folders(0)
    If Dir$(path, VBA.vbDirectory) = vbNullString Then
        Err.Raise VBAErrorFileNotFound, "CreateDirectory", "Path not found"
    End If
    Dim folderIndex As Long
    For folderIndex = 1 To UBound(folders, 1)
        path = path & DirectorySeparatorChar & folders(folderIndex)
        If Dir$(path, vbDirectory) = vbNullString Then
            MkDir path
        End If
    Next folderIndex
End Sub

'@Description("Determines whether the given path refers to an existing directory on disk.")
Private Function DirectoryExist(ByVal path As String) As Boolean
Attribute DirectoryExist.VB_Description = "Determines whether the given path refers to an existing directory on disk."
    If Dir$(path, VBA.vbDirectory) <> vbNullString Then
        DirectoryExist = True
    End If
End Function
