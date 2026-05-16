' ActiveFolders.vbs
' Handles population and events for the ActiveFolders dropdown, listing paths from open Explorer windows.

Dim ActiveFoldersList(19) ' Fixed size array of 20 elements, indexed 0-19
Dim ActiveFoldersCount: ActiveFoldersCount = 0 ' Counter for valid entries
Dim PreviousFoldersState: PreviousFoldersState = "" ' State for detecting changes

Sub PopulateActiveFolders()
    ' Collect temporary paths first to detect changes
    Dim tempPaths(19)
    Dim tempCount: tempCount = 0
    Dim sh, win, i, url, path
    Set sh = CreateObject("Shell.Application")
    
    For i = 0 To sh.Windows.Count - 1
        Set win = sh.Windows.Item(i)
        On Error Resume Next
        url = win.LocationURL
        If Err.Number = 0 Then
            If url <> "" And Left(url, 8) = "file:///" Then
                path = Mid(url, 9)
                path = Replace(path, "%20", " ")  ' Spaces
				path = Replace(path, "%26", "&")  ' Ampersand (&)
				path = Replace(path, "%23", "#")  ' Hash (#) - common in folder namespath = Replace(path, "/", "\")
                path = Replace(path, "%20", " ")
                ' Only add if it's a folder (check if ends with \ or use FSO)
                If oFs.FolderExists(path) Then
                    tempPaths(tempCount) = path
                    tempCount = tempCount + 1
                    If tempCount >= 20 Then Exit For ' Prevent overflow
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next
    
    Set sh = Nothing
    
    ' Create sorted copy for state comparison
    Dim sortedPaths(19)
    For i = 0 To tempCount - 1
        sortedPaths(i) = tempPaths(i)
    Next
    Call BubbleSort(sortedPaths, tempCount)
    
    ' Build new state string (sorted, pipe-separated)
    Dim newState, j
    newState = ""
    For j = 0 To tempCount - 1
        If j > 0 Then newState = newState & "|"
        newState = newState & sortedPaths(j)
    Next
    
    ' Only update if changed or first run
    Dim doUpdate
    doUpdate = (newState <> PreviousFoldersState) Or (PreviousFoldersState = "")
    If doUpdate Then
        ' Clear the dropdown (loop backwards to avoid index shift errors)
        Dim activeSelect
        Set activeSelect = Document.GetElementById("ActiveFolders")
        For i = activeSelect.Options.Length - 1 To 0 Step -1
            activeSelect.RemoveChild activeSelect.Options(i)
        Next
        
        ' Add default option
        Set defaultOption = Document.CreateElement("OPTION")
        defaultOption.Text = "Select Folder"
        defaultOption.Value = ""
        activeSelect.Add defaultOption
        
        ' Initialize array to empty strings
        For i = 0 To 19
            ActiveFoldersList(i) = ""
        Next
        ActiveFoldersCount = 0
        
        ' Add options from tempPaths (original detection order)
        For i = 0 To tempCount - 1
            Set optionElem = Document.CreateElement("OPTION")
            optionElem.Text = oFs.GetBaseName(tempPaths(i)) & " (" & tempPaths(i) & ")"
            optionElem.Value = tempPaths(i)
            activeSelect.Add optionElem
            ActiveFoldersList(i) = tempPaths(i)
            ActiveFoldersCount = ActiveFoldersCount + 1
        Next
        
        ' Update state
        PreviousFoldersState = newState
        
        ' Optionally add a placeholder if no folders
        If tempCount = 0 Then
            Set noFolderOption = Document.CreateElement("OPTION")
            noFolderOption.Text = "No Active Folders"
            noFolderOption.Value = ""
            noFolderOption.Disabled = True
            activeSelect.Add noFolderOption
        End If
    End If
End Sub

Sub BubbleSort(arr, size)
    Dim i, j, temp
    For i = 0 To size - 2
        For j = 0 To size - 2 - i
            If arr(j) > arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next
    Next
End Sub

Sub ActiveFolders_OnChange()
    Dim selectedPath
    selectedPath = Document.GetElementById("ActiveFolders").Value
    If selectedPath <> "" Then
        document.getElementById("txtFolder").innerText = selectedPath
        ' Save to registry if desired, similar to Getfolder_Onclick
        On Error Resume Next
        oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\LastFolder", selectedPath, "REG_SZ"
        On Error GoTo 0
        ' Reset to default "Select Folder" option
        Document.GetElementById("ActiveFolders").Value = ""
        PopulateActiveFolders
    End If
End Sub

' Helper function to query ActiveFoldersList, skipping blanks
Function GetActiveFolderPath(index)
    Dim i: i = 0
    For k = 0 To 19
        If ActiveFoldersList(k) <> "" Then
            If i = index Then
                GetActiveFolderPath = ActiveFoldersList(k)
                Exit Function
            End If
            i = i + 1
        End If
    Next
    GetActiveFolderPath = ""
End Function

' Helper function to get count of valid paths
Function GetActiveFoldersCount()
    GetActiveFoldersCount = ActiveFoldersCount
End Function