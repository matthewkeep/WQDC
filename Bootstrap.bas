Option Explicit
' Bootstrap: Import all .bas modules from project folder.
' Usage: Import this file only, then run Bootstrap.LoadAll
'
' IMPORTANT: Enable "Trust access to VBA project object model"
'   Windows: File > Options > Trust Center > Macro Settings
'   Mac: May require different approach

Public Sub LoadAll()
    ' Imports all .bas files from the same folder as this workbook
    Dim folderPath As String, fileName As String
    Dim vbProj As Object, comp As Object
    Dim imported As Long, skipped As Long

    On Error GoTo ErrHandler

    ' Get folder path (where the .bas files are)
    folderPath = GetBasFolder()
    If Len(folderPath) = 0 Then
        MsgBox "Could not determine .bas folder. Place workbook in same folder as .bas files, or edit GetBasFolder().", vbExclamation
        Exit Sub
    End If

    Set vbProj = ThisWorkbook.VBProject

    ' Loop through .bas files
    fileName = Dir(folderPath & "*.bas")
    Do While Len(fileName) > 0
        ' Skip Bootstrap itself
        If LCase(fileName) <> "bootstrap.bas" Then
            ' Remove existing module if present
            On Error Resume Next
            Set comp = vbProj.VBComponents(Replace(fileName, ".bas", ""))
            If Not comp Is Nothing Then
                vbProj.VBComponents.Remove comp
            End If
            On Error GoTo ErrHandler

            ' Import the module
            vbProj.VBComponents.Import folderPath & fileName
            imported = imported + 1
        Else
            skipped = skipped + 1
        End If
        fileName = Dir()
    Loop

    ' Also import .cls files
    fileName = Dir(folderPath & "*.cls")
    Do While Len(fileName) > 0
        On Error Resume Next
        Set comp = vbProj.VBComponents(Replace(fileName, ".cls", ""))
        If Not comp Is Nothing Then
            vbProj.VBComponents.Remove comp
        End If
        On Error GoTo ErrHandler

        vbProj.VBComponents.Import folderPath & fileName
        imported = imported + 1
        fileName = Dir()
    Loop

    MsgBox "Imported " & imported & " modules." & vbNewLine & vbNewLine & _
           "Next steps:" & vbNewLine & _
           "1. Run Setup.BuildAll" & vbNewLine & _
           "2. Add sheet event code (see Events.bas header)", vbInformation, "Bootstrap"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description & vbNewLine & vbNewLine & _
           "Make sure 'Trust access to VBA project' is enabled.", vbExclamation, "Bootstrap"
End Sub

Private Function GetBasFolder() As String
    ' Returns the folder containing .bas files
    ' Edit this if your .bas files are in a different location

    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    ' If workbook is in the project folder, use that
    If Len(wbPath) > 0 Then
        #If Mac Then
            GetBasFolder = wbPath & "/"
        #Else
            GetBasFolder = wbPath & "\"
        #End If
        Exit Function
    End If

    ' Fallback: prompt user
    #If Mac Then
        GetBasFolder = InputBox("Enter path to .bas files folder:", "Bootstrap", "/Users/")
        If Len(GetBasFolder) > 0 And Right(GetBasFolder, 1) <> "/" Then GetBasFolder = GetBasFolder & "/"
    #Else
        GetBasFolder = InputBox("Enter path to .bas files folder:", "Bootstrap", "C:\")
        If Len(GetBasFolder) > 0 And Right(GetBasFolder, 1) <> "\" Then GetBasFolder = GetBasFolder & "\"
    #End If
End Function

Public Sub ListModules()
    ' Debug: List all modules in the project
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Debug.Print comp.Name, comp.Type
    Next comp
End Sub
