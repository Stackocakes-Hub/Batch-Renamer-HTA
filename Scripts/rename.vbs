Const FOLDER_MISSING_ERR = "Folder not found: $FolderName"

Sub RenameFiles(psFolderPath, arrSearch, arrReplace)
    Dim oFolder,oFile,i,tnext,sNewName
    If Not oFs.FolderExists(psFolderPath) Then
        Log "<span style='color:#ef4444'>ERROR: " & Replace(FOLDER_MISSING_ERR,"$FolderName",psFolderPath) & "</span>"
        Exit Sub
    End If
    Set oFolder = oFs.GetFolder(psFolderPath)
    For Each oFile In oFolder.Files
        tnext = False
        For i = 0 To UBound(arrSearch)
            If arrSearch(i)<>"" And Not tnext Then
                If InStr(1,oFile.Name,arrSearch(i),vbTextCompare)>0 Then
                    If Trim(arrReplace(i)) = "" Then
                        sNewName = Replace(oFile.Name, arrSearch(i), "", 1, -1, vbTextCompare)
                    Else
                        sNewName = Replace(oFile.Name, arrSearch(i), arrReplace(i), 1, -1, vbTextCompare)
                    End If
                    RenameFile oFile.Path, sNewName
                    tnext = True
                End If
            End If
        Next
    Next
End Sub

Sub RenameFile(psFilePath, psTargetName)
    Dim sPath,sOldName,sNewFull
    sOldName = oFs.GetFileName(psFilePath)
    sPath    = Left(psFilePath,Len(psFilePath)-Len(sOldName))
    sNewFull = sPath & psTargetName

    If gPreview Then
        Log "<span style='color:#10b981'>[PREVIEW] Would rename <b>" & sOldName & "</b> to <b>" & psTargetName & "</b></span>"
    Else
        Log "Renaming <b>" & sOldName & "</b> to <b>" & psTargetName & "</b>"
        On Error Resume Next
        If oFs.FileExists(sNewFull) Then oFs.DeleteFile sNewFull, True
        oFs.MoveFile psFilePath, sNewFull
        If Err.Number<>0 Then Log "<span style='color:#ef4444'>ERROR: " & Err.Description & "</span>"
        On Error Goto 0
    End If
End Sub

Sub Log(s)
    Dim div : Set div = document.getElementById("log")
    div.innerHTML = div.innerHTML & s & "<br>"
    div.scrollTop = div.scrollHeight
End Sub

Sub ClearLog()
    document.getElementById("log").innerHTML = ""
End Sub

Sub Preview()
    ClearLog
    Log "<hr><b>=== PREVIEW ===</b>"
    gPreview = True
    RunRename
    Log "<b>Preview complete – no files changed.</b>"
End Sub

Sub RunRename()
    Dim folder, searchArr(), replaceArr(), i

    folder = Trim(document.getElementById("txtFolder").innerText)
	If folder = "" Then
        Log "<span style='color:#ef4444'>Please select a folder first.</span>"
        Exit Sub
    End If

    ReDim searchArr(visibleRows - 1)
    ReDim replaceArr(visibleRows - 1)
    For i = 0 To visibleRows - 1
        searchArr(i)  = Trim(rulesArray(i)(0))
        replaceArr(i) = Trim(rulesArray(i)(1))
    Next

    Dim cleanSearch(), cleanReplace(), cleanCount
    cleanCount = 0
    For i = 0 To visibleRows - 1
        If searchArr(i) <> "" Then
            ReDim Preserve cleanSearch(cleanCount)
            ReDim Preserve cleanReplace(cleanCount)
            cleanSearch(cleanCount) = searchArr(i)
            cleanReplace(cleanCount) = replaceArr(i)
            cleanCount = cleanCount + 1
        End If
    Next
    If cleanCount = 0 Then
        Log "<span style='color:#f59e0b'>No valid rules to apply (all Search fields empty).</span>"
        Exit Sub
    End If

    If Not gPreview Then
        ClearLog
        Log "<hr><b>=== RENAMING ===</b>"
    End If

    RenameFiles folder, cleanSearch, cleanReplace

    If Not gPreview Then Log "<b style='color:#16a34a'>Done.</b>"
    gPreview = False
End Sub