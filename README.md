# Download VBA Projects
Easy method to download a vba project from git into your excel project 

1. Download the desired project as zip from GitHub and unpack it. 

2. Add following code to a module in your VB Editor and change the string of the path variable with the path of the downloaded unpacked zip file.
```
Public Sub prcImport()
    Dim objVBComponents As Object, strFilename As String
    Dim path As String
    path = "C:\YOUR_PATH" 
    strFilename = Dir(path & "*.*")
    With ActiveWorkbook.VBProject
        Do While strFilename <> ""
            If UCase(right(strFilename, 4)) = ".BAS" Or _
                UCase(right(strFilename, 4)) = ".FRM" Or _
                UCase(right(strFilename, 4)) = ".CLS" Then
                .VBComponents.Import path & strFilename
            End If
            strFilename = Dir()
        Loop
    End With
End Sub
```
