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

'@Remarks PathSeparator is hardcoded possibily add conditional compiler for Mac and Windows or property/parameter?

'Useage:
'   Exports all components in the active project to a base file location, overwritting any existing files.
'   RubberduckUtility.ExportAllComponents "C:\VBA\Output\"

Option Explicit

Private Const PathSeparator As String = "\"

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
'@param destinationBasePath base file path to export to.
'@TODO Validate inputs, implement optional filter for a RD Folder annotation?
Public Sub ExportAllComponents(ByVal destinationBasePath As String)
Attribute ExportAllComponents.VB_Description = "Export the active project modules to file according to Rubberduck @Folder annotation."
    Dim component As VBComponent
    For Each component In this.Project.VBComponents
        Dim filename As String
        filename = component.Name
        Dim extension As String
        extension = ComponentExtension(component)

        Dim componentfolder As String
        componentfolder = ParseFolderAnnotation(component)
        
        Dim outputPath As String
        outputPath = GetFilePathFromAnnotation(componentfolder)
        'Debug.Print destinationBasePath; outputPath; IIf(outputPath <> VBA.vbNullString, PathSeparator, VBA.vbNullString); filename; extension; " : "; componentfolder
        Dim fullpath As String
        fullpath = destinationBasePath & outputPath
        CreateDirectory fullpath
        
        Dim fileFullPath As String
        fileFullPath = fullpath & IIf(outputPath <> VBA.vbNullString, PathSeparator, VBA.vbNullString) & filename & extension
        component.Export fileFullPath
    Next
End Sub

'@Description "Obtains the Rubberduck @Folder annotation."
Private Function ParseFolderAnnotation(ByVal component As VBComponent) As String
Attribute ParseFolderAnnotation.VB_Description = "Obtains the Rubberduck @Folder annotation."
    Const RubberduckFolderAnnotation As String = "'@Folder"
    
    Dim componentDeclareLines As String
    componentDeclareLines = component.CodeModule.Lines(1, component.CodeModule.CountOfDeclarationLines)
    Dim annotateStart As Long
    annotateStart = InStr(1, componentDeclareLines, RubberduckFolderAnnotation, vbTextCompare)
    If annotateStart > 0 Then
        annotateStart = InStr(annotateStart, componentDeclareLines, """")
        Dim annotateEnd As Long
        annotateEnd = InStr(annotateStart + 1, componentDeclareLines, """")
        ParseFolderAnnotation = Mid$(componentDeclareLines, annotateStart + 1, annotateEnd - annotateStart - 1)
    End If
End Function

'@Description "Obtains the path from the Rubberduck @Folder annotation"
Private Function GetFilePathFromAnnotation(ByVal folderAnnotation As String) As String
Attribute GetFilePathFromAnnotation.VB_Description = "Obtains the path from the Rubberduck @Folder annotation"
    Const RubberduckPathSeparator As String = "."
    
    Dim outputPath As String
    If folderAnnotation <> VBA.vbNullString Then
        outputPath = Replace$(folderAnnotation, RubberduckPathSeparator, PathSeparator)
    End If
    GetFilePathFromAnnotation = outputPath
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

'@Description "Creates path, including any non-existing subdirectories."
Private Sub CreateDirectory(ByVal fullpath As String)
Attribute CreateDirectory.VB_Description = "Creates path, including any non-existing subdirectories."
    Dim folders As Variant
    folders = Split(fullpath, PathSeparator)

    Dim path As String
    path = folders(0)
    If Dir$(path, VBA.vbDirectory) = vbNullString Then
        Err.Raise 76, "Rubberduck Export", "Path not found"
    End If
    Dim folderIndex As Long
    For folderIndex = 1 To UBound(folders, 1)
        path = path & PathSeparator & folders(folderIndex)
        If Dir$(path, vbDirectory) = vbNullString Then
            MkDir path
        End If
    Next folderIndex
End Sub
