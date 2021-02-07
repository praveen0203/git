
call Unzip()

Function Unzip()
    'Dim FSO As Object
    'Dim oApp As Object
    'Dim Fname As Variant
    'Dim FileNameFolder As Variant
    'Dim DefPath As String
    'Dim strDate As String
    'Dim I As Long
    'Dim num As Long

    Fname = Application.GetOpenFilename(filefilter:="Zip Files (*.zip), *.zip", _MultiSelect:=True))
                                        
    If IsArray(Fname) = False Then
        'Do nothing
    Else
        'Root folder for the new folder.
        'You can also use DefPath = "C:\Users\Ron\test\"
        DefPath = Application.DefaultFilePath
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        'Create the folder name
        strDate = Format(Now, " dd-mm-yy h-mm-ss")
        FileNameFolder = DefPath & "MyUnzipFolder " & strDate & "\"

        'Make the normal folder in DefPath
        MkDir FileNameFolder

        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")

        For I = LBound(Fname) To UBound(Fname)
            num = oApp.Namespace(FileNameFolder).items.Count

            oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname(I)).items

        Next I

        MsgBox "You find the files here: " & FileNameFolder

        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.deletefolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Function
 

