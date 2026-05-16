Const APP_WIDTH  = 700
Const APP_HEIGHT = 850

Dim oFs, oShellApp, oShell, gPreview, savedFolder, rulesArray, MAX_ROWS, visibleRows

Set oFs       = CreateObject("Scripting.FileSystemObject")
Set oShellApp = CreateObject("Shell.Application")
Set oShell    = CreateObject("WScript.Shell")

Function GetRegValue(key, def)
    On Error Resume Next
    GetRegValue = oShell.RegRead(key)
    If Err.Number <> 0 Then GetRegValue = def
    On Error Goto 0
End Function

Sub window_onload()
	RefreshFolders = window.setInterval("PopulateActiveFolders", 2000)
	MAX_ROWS = GetRegValue("HKEY_CURRENT_USER\Software\BatchRename\MaxRows", 20)
    ReDim rulesArray(MAX_ROWS - 1)
    Dim i
    For i = 0 To MAX_ROWS - 1
        rulesArray(i) = Array("", "")
    Next

    On Error Resume Next
    Dim autoData
    autoData = oShell.RegRead("HKEY_CURRENT_USER\Software\BatchRename\_AutoSave")
    If Err.Number = 0 And Not IsEmpty(autoData) And autoData <> "" Then
        LoadRulesetFromData autoData, "Last Session"
        Log "<span style='color:#6b7280'>Loaded: <b>Last Session</b></span>"
    Else
        LoadDefaultRules
        Log "<span style='color:#6b7280'>Loaded: <b>(Default)</b></span>"
    End If
    On Error Goto 0

    Dim appPath, folderPath
    appPath = Replace(Unescape(Document.Location.href), "file:///", "")
    appPath = Replace(appPath, "/", "\")
    folderPath = oFs.GetParentFolderName(appPath)
    window.document.title = "Batch Rename - " & folderPath

    On Error Resume Next
    savedFolder = oShell.RegRead("HKEY_CURRENT_USER\Software\BatchRename\LastFolder")
    On Error Goto 0
	If savedFolder <> "" And oFs.FolderExists(savedFolder) Then
		document.getElementById("txtFolder").innerText = savedFolder
	Else
		document.getElementById("txtFolder").innerText = oShellApp.Namespace(&H10&).Self.Path
	End If

    PopulateRulesetDropdown
    ResizeWindow APP_WIDTH, APP_HEIGHT
    BuildRows
End Sub

Sub AddRow()
    If visibleRows < MAX_ROWS Then
        visibleRows = visibleRows + 1
        document.getElementById("row" & (visibleRows - 1)).style.display = "flex"
    End If
End Sub

Sub DeleteRow(row)
    Dim i
    For i = row To visibleRows - 2
        rulesArray(i) = rulesArray(i + 1)
    Next
    rulesArray(visibleRows - 1) = Array("", "")
    visibleRows = visibleRows - 1
    BuildRows
End Sub

Sub BuildRows()
    Dim container, i, disp, searchInput, replaceInput
    
    Set container = document.getElementById("rules-container")
    If container Is Nothing Then
        Log "<span style='color:#ef4444'>ERROR: #rules-container not found.</span>"
        Exit Sub
    End If

    container.innerHTML = ""

    For i = 0 To MAX_ROWS - 1
        disp = "none"
        If i < visibleRows Then disp = "flex"

        container.innerHTML = container.innerHTML & _
            "<div id='row" & i & "' class='rule-row' style='display:" & disp & ";'>" & _
            "<input id='search" & i & "' type='text' onchange=\"UpdateRule " & i & ", 0\" placeholder='Find text'>" & _
            "&nbsp;&nbsp;<input id='replace" & i & "' type='text' onchange=\"UpdateRule " & i & ", 1\" placeholder='Replace with (blank = remove)'>" & _
            "&nbsp;&nbsp;<button class='btn-small' onclick=\"DeleteRow " & i & "\"'>X</button>" & _
            "</div>"
    Next

    On Error Resume Next
    For i = 0 To visibleRows - 1
        Set searchInput = document.getElementById("search" & i)
        Set replaceInput = document.getElementById("replace" & i)
        If Not searchInput Is Nothing Then searchInput.value = rulesArray(i)(0)
        If Not replaceInput Is Nothing Then replaceInput.value = rulesArray(i)(1)
    Next
    On Error Goto 0
End Sub

Sub ResizeWindow(w, h)
    Dim objWMIService, colItems, objItem
    Dim scrW, scrH
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
    For Each objItem in colItems
        scrW = objItem.ScreenWidth
        scrH = objItem.ScreenHeight
    Next
    Window.ResizeTo w, h
    Window.MoveTo (scrW - w) \ 2, (scrH - h) \ 2
End Sub

Function Pickfolder(Title)
    Dim F
    Set F = oShellApp.BrowseForFolder(0, Title, 16+32+64+512, 0)
    If (Not F Is Nothing) Then
        If F = "Desktop" Then
            Pickfolder = oShellApp.Namespace(&H10&).Self.Path
        Else
            Pickfolder = F.Items.Item.Path
        End If
    Else
        Pickfolder = ""
    End If
End Function

Sub Getfolder_Onclick()
    Dim sel
    sel = Pickfolder("Select the folder that contains the files to rename:")
    If sel <> "" Then
        oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\LastFolder", sel, "REG_SZ"
        document.getElementById("txtFolder").value = sel
    End If
End Sub

Sub window_onunload()
    On Error Resume Next
    SaveCurrentRuleset
    ' Force quit HTA process
    WScript.Quit
    On Error Goto 0
End Sub

Sub SaveCurrentRuleset()
    Dim data, i
    data = MAX_ROWS & vbCrLf & visibleRows & vbCrLf
    For i = 0 To visibleRows - 1
        data = data & rulesArray(i)(0) & "|" & rulesArray(i)(1) & vbCrLf
    Next
    oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\_AutoSave", data, "REG_SZ"
End Sub

Sub EditMaxRows()
    Dim inp, newMax
    inp = InputBox("Enter maximum number of rules (current: " & MAX_ROWS & ")", "Settings", MAX_ROWS)
    If IsNumeric(inp) Then
        newMax = CInt(inp)
        If newMax >= 5 And newMax <= 100 Then
            oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\MaxRows", newMax, "REG_DWORD"
            Log "<span style='color:#7c3aed'>Max rows set to " & newMax & ". Restart to apply.</span>"
        Else
            Log "<span style='color:#ef4444'>Enter a number between 5 and 100.</span>"
        End If
    End If
End Sub