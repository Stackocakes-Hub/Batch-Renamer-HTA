Function CreateOption(text, value)
    Dim opt
    Set opt = document.createElement("option")
    opt.text = text
    opt.value = value
    Set CreateOption = opt
End Function

Sub PopulateRulesetDropdown()
    Dim sel, known, names, i
    Set sel = document.getElementById("selRuleset")
    sel.innerHTML = ""

    sel.options.add CreateOption("Last Session", "_AutoSave")
    sel.options.add CreateOption("(Default)", "_DEFAULT_")

    known = GetRegValue("HKEY_CURRENT_USER\Software\BatchRename\knownRules", "")
    If known <> "" Then
        names = Split(known, ",")
        For i = 0 To UBound(names)
            If names(i) <> "" And names(i) <> "_AutoSave" Then
                sel.options.add CreateOption(names(i), names(i))
            End If
        Next
    End If

    sel.value = "_AutoSave"
End Sub

Sub LoadSelectedRuleset()
    Dim sel, value
    Set sel = document.getElementById("selRuleset")
    value = sel.value

    If value = "_AutoSave" Then
        On Error Resume Next
        Dim data
        data = oShell.RegRead("HKEY_CURRENT_USER\Software\BatchRename\_AutoSave")
        If Err.Number = 0 And data <> "" Then
            LoadRulesetFromData data, "Last Session"
            Log "<span style='color:#6b7280'>Switched to: <b>Last Session</b></span>"
        Else
            LoadDefaultRules
            Log "<span style='color:#f59e0b'>Last Session missing. Loaded <b>(Default)</b>.</span>"
        End If
        On Error Goto 0
    ElseIf value = "_DEFAULT_" Then
        LoadDefaultRules
        Log "<span style='color:#6b7280'>Switched to: <b>(Default)</b></span>"
    Else
        LoadRulesetByName value
        Log "<span style='color:#10b981'>Switched to: <b>" & value & "</b></span>"
    End If
    BuildRows
End Sub

Sub LoadRulesetByName(name)
    On Error Resume Next
    Dim data
    data = oShell.RegRead("HKEY_CURRENT_USER\Software\BatchRename\Rulesets\" & name)
    If Err.Number <> 0 Or data = "" Then
        Log "<span style='color:#ef4444'>Ruleset not found: " & name & "</span>"
        Err.Clear
        On Error Goto 0
        Exit Sub
    End If
    On Error Goto 0
    LoadRulesetFromData data, name
End Sub

Sub SaveCurrentAsRuleset()
    Dim name, data, i, known, names, found
    name = InputBox("Enter ruleset name:", "Save Ruleset", "")
    If name = "" Or name = "Last Session" Or name = "(Default)" Or name = "_AutoSave" Or name = "_DEFAULT_" Then
        Log "<span style='color:#f59e0b'>Invalid or reserved name.</span>"
        Exit Sub
    End If

    ' Build rules data
    data = MAX_ROWS & vbCrLf & visibleRows & vbCrLf
    For i = 0 To visibleRows - 1
        data = data & rulesArray(i)(0) & "|" & rulesArray(i)(1) & vbCrLf
    Next

    ' Save ruleset data
    On Error Resume Next
    oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\Rulesets\" & name, data, "REG_SZ"
    If Err.Number <> 0 Then
        Log "<span style='color:#ef4444'>Save failed: " & Err.Description & "</span>"
        Err.Clear
        On Error Goto 0
        Exit Sub
    End If
    On Error Goto 0

    ' Update knownRules (CSV)
    known = GetRegValue("HKEY_CURRENT_USER\Software\BatchRename\knownRules", "")
    names = Split(known, ",")
    found = False
    For i = 0 To UBound(names)
        If names(i) = name Then found = True
    Next
    If Not found Then
        If known = "" Then known = name Else known = known & "," & name
        oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\knownRules", known, "REG_SZ"
    End If

    PopulateRulesetDropdown
    document.getElementById("selRuleset").value = name
    Log "<span style='color:#10b981'>Ruleset saved: <b>" & name & "</b></span>"
End Sub

Sub RemoveSelectedRuleset()
    Dim sel, name, known, names, i, newList
    Set sel = document.getElementById("selRuleset")
    name = sel.value
    If name = "" Or name = "_AutoSave" Or name = "_DEFAULT_" Then
        Log "<span style='color:#f59e0b'>Cannot remove this ruleset.</span>"
        Exit Sub
    End If

    On Error Resume Next
    oShell.RegDelete "HKEY_CURRENT_USER\Software\BatchRename\Rulesets\" & name
    If Err.Number <> 0 Then
        Log "<span style='color:#ef4444'>Delete failed: " & Err.Description & "</span>"
        Err.Clear
    End If
    On Error Goto 0

    ' Update knownRules
    known = GetRegValue("HKEY_CURRENT_USER\Software\BatchRename\knownRules", "")
    names = Split(known, ",")
    newList = ""
    For i = 0 To UBound(names)
        If names(i) <> name And names(i) <> "" Then
            If newList = "" Then newList = names(i) Else newList = newList & "," & names(i)
        End If
    Next
    oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\knownRules", newList, "REG_SZ"

    PopulateRulesetDropdown
    sel.value = "_DEFAULT_"
    LoadDefaultRules
    BuildRows
    Log "<span style='color:#dc2626'>Removed: <b>" & name & "</b></span>"
End Sub

Sub ImportRuleset()
    Dim shell, folder, files, file, stream, line, parts, name, imported
    imported = 0

    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "Select folder containing .csv files", &H10&)
    If folder Is Nothing Then
        Log "<span style='color:#f59e0b'>Import cancelled.</span>"
        Exit Sub
    End If

    Set files = folder.Items
    For Each file In files
        If LCase(oFs.GetExtensionName(file.Name)) <> "csv" Then
        Else
            name = oFs.GetBaseName(file.Name)
            If name <> "" And name <> "_AutoSave" And name <> "_DEFAULT_" And name <> "Last Session" And name <> "(Default)" Then
                Set stream = oFs.OpenTextFile(file.Path, 1)
                Dim data, maxRows, visibleCount, hasHeader, i
                maxRows = 0: visibleCount = 0: hasHeader = False
                data = ""
                i = 0
                Do Until stream.AtEndOfStream
                    line = Trim(stream.ReadLine)
                    If line <> "" Then
                        If Not hasHeader Then
                            If Left(line, 9) = "MAX_ROWS," Then
                                maxRows = CLng(Mid(line, 10))
                            ElseIf Left(line, 13) = "VISIBLE_ROWS," Then
                                visibleCount = CLng(Mid(line, 14))
                            ElseIf InStr(line, "Search") > 0 And InStr(line, "Replace") > 0 Then
                                hasHeader = True
                            Else
                                parts = Split(line, ",")
                                If UBound(parts) >= 1 Then
                                    Dim searchVal, replaceVal
                                    searchVal = Replace(parts(0), """", "")
                                    replaceVal = Replace(parts(1), """", "")
                                    data = data & searchVal & "|" & replaceVal & vbCrLf
                                    i = i + 1
                                End If
                            End If
                        Else
                            parts = Split(line, ",")
                            If UBound(parts) >= 1 Then
                                searchVal = Replace(parts(0), """", "")
                                replaceVal = Replace(parts(1), """", "")
                                data = data & searchVal & "|" & replaceVal & vbCrLf
                                i = i + 1
                            End If
                        End If
                    End If
                Loop
                stream.Close

                If i > 0 And maxRows > 0 And visibleCount > 0 Then
                    data = maxRows & vbCrLf & visibleCount & vbCrLf & data
                    data = Left(data, Len(data) - Len(vbCrLf)) ' remove final CRLF
                    oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\Rulesets\" & name, data, "REG_SZ"

                    Dim known, names, found, j
                    known = GetRegValue("HKEY_CURRENT_USER\Software\BatchRename\knownRules", "")
                    names = Split(known, ",")
                    found = False
                    For j = 0 To UBound(names)
                        If names(j) = name Then found = True
                    Next
                    If Not found Then
                        If known = "" Then known = name Else known = known & "," & name
                        oShell.RegWrite "HKEY_CURRENT_USER\Software\BatchRename\knownRules", known, "REG_SZ"
                    End If

                    imported = imported + 1
                End If
            End If
        End If
    Next

    If imported > 0 Then
        PopulateRulesetDropdown
        Log "<span style='color:#10b981'>Imported " & imported & " ruleset(s) from CSV.</span>"
    Else
        Log "<span style='color:#f59e0b'>No valid .csv files found or imported.</span>"
    End If
End Sub

Sub ExportSelectedRuleset()
    Dim sel, name, data, folder, path, fso, file
    Set sel = document.getElementById("selRuleset")
    name = sel.options(sel.selectedIndex).text
    If name = "Last Session" Or name = "(Default)" Then
        Log "<span style='color:#f59e0b'>Cannot export this ruleset.</span>"
        Exit Sub
    End If

    On Error Resume Next
    data = oShell.RegRead("HKEY_CURRENT_USER\Software\BatchRename\Rulesets\" & sel.value)
    If Err.Number <> 0 Then
        Log "<span style='color:#ef4444'>Failed to read ruleset: " & name & "</span>"
        Exit Sub
    End If
    On Error Goto 0

    folder = Pickfolder("Select folder to export .csv file")
    If folder = "" Then
        Log "<span style='color:#f59e0b'>Export cancelled.</span>"
        Exit Sub
    End If

    path = folder & "\" & name & ".csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(path, True)
    
    Dim lines, maxRows, visibleCount, i, parts
    lines = Split(data, vbCrLf)
    maxRows = CInt(lines(0))
    visibleCount = CInt(lines(1))
    
    file.WriteLine "MAX_ROWS," & maxRows
    file.WriteLine "VISIBLE_ROWS," & visibleCount
    file.WriteLine "Search,Replace"
    
    For i = 2 To visibleCount + 1
        If i <= UBound(lines) Then
            parts = Split(lines(i), "|")
            If UBound(parts) >= 1 Then
                file.WriteLine """" & parts(0) & """","""" & parts(1) & """"
            End If
        End If
    Next
    file.Close

    Log "<span style='color:#7c3aed'>Exported: <b>" & name & ".csv</b> (" & visibleCount & " visible rows)</span>"
End Sub

Sub LoadDefaultRules()
    Dim defaultRules
    defaultRules = Array( _
        Array(" Model", ""), _
        Array("-Model", ""), _
        Array(" Layout1", ""), _
        Array("-Layout1", ""), _
        Array(" Color", "") _
    )
    visibleRows = UBound(defaultRules) + 1
    Dim i
    For i = 0 To UBound(defaultRules)
        rulesArray(i) = defaultRules(i)
    Next
    For i = visibleRows To MAX_ROWS - 1
        rulesArray(i) = Array("", "")
    Next
End Sub

Sub LoadRulesetFromData(data, sourceName)
    Dim lines, parts, j
    lines = Split(data, vbCrLf)
    If UBound(lines) < 1 Then Exit Sub

    On Error Resume Next
    MAX_ROWS = CInt(lines(0))
    visibleRows = CInt(lines(1))
    If Err.Number <> 0 Or MAX_ROWS < 1 Or visibleRows < 0 Or visibleRows > MAX_ROWS Then
        Err.Clear
        On Error Goto 0
        Exit Sub
    End If
    On Error Goto 0

    ReDim rulesArray(MAX_ROWS - 1)
    For j = 0 To visibleRows - 1
        If j + 2 <= UBound(lines) Then
            parts = Split(lines(j + 2), "|")
            If UBound(parts) >= 1 Then
                rulesArray(j) = Array(parts(0), parts(1))
            Else
                rulesArray(j) = Array(parts(0), "")
            End If
        Else
            rulesArray(j) = Array("", "")
        End If
    Next
    For j = visibleRows To MAX_ROWS - 1
        rulesArray(j) = Array("", "")
    Next
End Sub

Sub UpdateRule(row, col)
    Dim prefix, id, val
    prefix = "search"
    If col = 1 Then prefix = "replace"
    id = prefix & row
    val = document.getElementById(id).value
    rulesArray(row)(col) = val
    If row = visibleRows - 1 And Trim(val) <> "" And visibleRows < MAX_ROWS Then AddRow
End Sub