
call m()

Sub m()

 HostFolder = "C:\Users\dell\Desktop\1"

Set FileSystem = CreateObject("Scripting.FileSystemObject")
DoFolder FileSystem.GetFolder(HostFolder)

Call DoFolder(HostFolder)

Msgbox "completed"
End Sub

Function DoFolder(Folder)
    'Dim SubFolder
    
    For Each SubFolder In Folder.SubFolders
    
    MsgBox SubFolder
        DoFolder SubFolder
      
    Next
   ' Dim File
    For Each File In Folder.Files
     
     filenm = File.Path
    'new File Name
   newfolder = "C:\Users\dell\Desktop\desti\" ' please add "\" as the end
   ' new path
   ' add \ at the end of folder
   'If VBA.Right(newfolder, 1) <> "\" Then newfolder = newfolder & "\"
     'new path of file
   'newpath = newfolder & VBA.Right(filenm, Len(filenm) - InStrRev(filenm, "\"))
 
    ' add some control check to avoid crashes
    
    If Right(File.Name, 3) = "zip" Then
  
    'move it finally
   Set fld = CreateObject("Scripting.FileSystemObject")
    fld.Movefile filenm, newfolder
    End If
   
   
    
        
    Next



End Function


