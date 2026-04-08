#SingleInstance Off
#Requires AutoHotkey v2.0

; ================================================================
; ELKHETA v3 - Google Sheets integration (Service Account JWT auth)
; When launched normally: shows the main GUI.
; When launched with /worker: runs as a background worker process.
; ================================================================

; ---- WORKER MODE: launched by the main instance with /worker ----
if A_Args.Length >= 4 && A_Args[1] = "/worker" {
    RunWorker(A_Args[2], A_Args[3], A_Args[4])
    ExitApp
}

; ---- SINGLE INSTANCE ENFORCEMENT FOR MAIN MODE ----
hMutex := DllCall("CreateMutex", "Ptr", 0, "Int", 1, "Str", "ElkhetaMainInstance", "Ptr")
if DllCall("GetLastError") = 183 {
    if WinExist("ahk_class AutoHotkeyGUI")
        WinActivate()
    ExitApp
}

; ================================================================
; MAIN INSTANCE BELOW
; ================================================================

; ---- GLOBAL VARIABLES ----
global FolderEdit := "", inputGui := ""
global folderOrder := [], folders := Map()
global searchRoots := []
global unreachablePaths := []
global sheetRowMap := Map()  ; maps folder name -> sheet row number from FetchFolderNamesFromSheet

; ---- PATHS ----
pathsFile      := A_AppData "\ElkhetaPaths.txt"
APP_VERSION    := "1.0.0"
MANIFEST_URL   := "https://raw.githubusercontent.com/WadeRK/raw/main/manifest.json"
; Embedded service account credentials
SA_CLIENT_EMAIL := "el-kheta-raw-mins@long-advice-488916-j7.iam.gserviceaccount.com"
SA_TOKEN_URI    := "https://oauth2.googleapis.com/token"
SA_PRIVATE_KEY  := "-----BEGIN PRIVATE KEY-----`nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDTHRRXQ6MuxDaD`nksVsuSjbscnSmlX3vjFlBjw+ld16SryMY0v7mnmOUeBE2hzXAR5EzFqHAWms0UJp`npZCI8s+zgBKKS+gyjziiGw9SqOpseBGSqpHXA/BUVlDMiZ9SjuahRQzLj4loP+e1`nBNz95Qqj4j3fRZqhG/XGlhjF6UDWZnbn3+jSScmROddR3ZXn5JyirAzwKzxlIA7a`nOc6IVrNksXKU+3Bj7SdNfoCbt8avauLT8ZZsqWnDz1J6A7LurMDFtfx8Do/F3yTr`nd++lyQohR6xXJMe9SoBgGA2x2yC4iI+v2f2DQUyfRQLIvXew6ePDI6/xEG/D9vha`n42VKuAclAgMBAAECggEADtBb0MrJfNreY2mtalVs2Uu2uI1FaorBTxA6sn6UKcRs`nSiKenpNmo7a5mMNAYlLtA5FeadpOxsZxVVX9mCqHqhYa1vMAImob5H4nG1sIhh9/`nxi2raa3YU6ax6URf9mxaISdmADO8heD0Gqbydet4TWsvAceR0+vRGYQHTX3yY4jw`n3mmkP99hq7kSd9lcWUQm6BrXj3zbGiwnD03DrWT0AWea7FiLz5zPl9RmoNf7QsS3`nkbmqt5kYimGwvRsLK9wyQQLYrylsHJVFvW0oaYvKa8RZIWx2u5Z4CsGeLBkHwrav`nORv/0FKtNxJ/KBgzZqWLKxbkdeKM/K5/MrcDz7v6sQKBgQD+BK/MOMcCwsRA7Kx7`n3H6tRPQV4bdcMp10NKqKkU3i74U11p6gC+Ociivtyp+kZTwg6G2DjxC9k5UfCNNA`nANKIPS98KjL/XsWcBci18mWuLKgVfwZmRhSp1EK7S55t3eJqiYb4TCxsZE1nKBIx`nFf8iDvW88bF8O9lQcTDZGInL0QKBgQDUwrSa0xBEO6awd0PvIcMyKp24/zRXKCDB`ngUKrF/8CLTBWobIS0NQamXYJMY9BkyZYy6WsD6plGj1lJoYQUeFIwz+yC0seutRv`nA8dBO3guPE7SWfGIMX1dAxBRu7eFw5BliliXZLqxWDj2ZuiClMv0gbpFEEPrA/2K`n5f9lFvwfFQKBgQCg9OlMF4IK/s0Kcq2MwpfGRRYBM4hTNINO2fxiV1YqASnAhqD2`nuvHcBDV3tNfZfIhQNdcG8MVjyrtH8fih6qN8zoBjRf4QkMXYalXW7KR/bC6JiHbO`noYOAZU5vjafy6BEK/t/2P5Y6jIf7YIm+bri+pQoTUnbrSKUX8tFIDuNpsQKBgEpy`nZufV+tclWEpfMTI3yil/p/jXs+TbcbrEFCPyHZURYtytb7YNxGoaUKce5FW5u61O`ndQYj5SfDasA+HqMPQ5lGWL9gHEUActz1oX894+upxprsRgu15XvqWod++9SefaRK`nKH2xDXKIwEXX9HvcvREtY5RPALT3jHxRxnAE/uuRAoGBAKIuXlNb3BSCTezQPu6k`nrZ2Z5ar/rXaPeT6fcf6a0WvZpTPforcGqTVTrqzE2ZsdmHS70Sl4HcIu0DY0P4i+`nHTRx+grbyz2NTayP+yiBnRV0gRX4mQfFfNQmT8CjOebjkEaoNqFO1pLE9vmC+0AZ`nuLOibhGYMY7Z9FI74dwvfRFw`n-----END PRIVATE KEY-----`n"

; Google Sheets config
SHEET_ID       := "1Hm7noXxv8ITMU3dNXQmqFEzfZY1mZlBJ4bQ9_ZIR0-M"
SHEET_NAME     := "OPERATIONS"
SHEET_GID      := 1476192399  ; Fixed sheetId for OPERATIONS tab
FILTER_COL_N   := 14   ; Column N = index 14 (1-based)
FILTER_COL_O   := 15   ; Column O = index 15
FILTER_COL_I   := 9    ; Column I = index 9
FILTER_COL_K   := 11   ; Column K = index 11
TARGET_COL_M   := 13   ; Column M = index 13

; ---- REMOTE UPDATE/DELETE POLICY CHECK (GITHUB MANIFEST) ----
ApplyRemotePolicy()

; ---- CLEAN UP ANY ORPHANED TEMP FILES FROM A PREVIOUS CRASHED RUN ----
Loop Files, A_Temp "\ElkhetaInput_*.txt"
    FileDelete(A_LoopFileFullPath)
Loop Files, A_Temp "\ElkhetaResult_*.txt"
    FileDelete(A_LoopFileFullPath)

; ---- CHECK CONFIG FILE ----
if !FileExist(pathsFile) {
    ShowConfigGUI()
} else {
    LoadSearchRoots()
    if searchRoots.Length = 0 {
        MsgBox("No valid search paths found in the config file.")
        Run(pathsFile)
        ExitApp
    }
    ShowFolderInputGUI()
}


; ================================================================
; GOOGLE SHEETS AUTH - entire JWT flow delegated to PowerShell
; AHK just passes the JSON key file path and gets back an access token.
; ================================================================

GetAccessToken() {
    suffix := A_TickCount
    tmpDir := EnvGet("TEMP")
    tmpPs  := tmpDir "\elkheta_auth_" suffix ".ps1"
    tmpOut := tmpDir "\elkheta_tok_"  suffix ".txt"
    tmpErr := tmpDir "\elkheta_err_"  suffix ".txt"

    psOut := StrReplace(tmpOut, "\", "/")
    psErr := StrReplace(tmpErr, "\", "/")
    ; Write embedded credentials to a temp JSON file for PowerShell
    tmpJson := tmpDir "\elkheta_sa_" suffix ".json"
    escapedKey := StrReplace(SA_PRIVATE_KEY, "`n", "\n")
    saJson := '{"client_email":"' SA_CLIENT_EMAIL '","token_uri":"' SA_TOKEN_URI '","private_key":"' escapedKey '"}'
    saBytes := Buffer(StrPut(saJson, "UTF-8") - 1)
    StrPut(saJson, saBytes, "UTF-8")
    fj := FileOpen(tmpJson, "w", "UTF-8-RAW")
    fj.RawWrite(saBytes)
    fj.Close()
    psJson := StrReplace(tmpJson, "\", "/")

    ; PowerShell script: reads JSON, builds JWT, signs it, exchanges for access token
    ps  := "try {" "`n"
    ps .= "  $json = Get-Content '" psJson "' -Raw | ConvertFrom-Json" "`n"
    ps .= "  $email = $json.client_email" "`n"
    ps .= "  $tokenUri = $json.token_uri" "`n"
    ps .= "  $pemRaw = $json.private_key" "`n"
    ps .= "  # Strip PEM headers and all whitespace to get raw base64" "`n"
    ps .= "  # Load PKCS8 key via RSACng (available in .NET 4.6+)" "`n"
    ps .= "  $pemClean = $pemRaw -replace '-----BEGIN [A-Z ]+-----','' -replace '-----END [A-Z ]+-----','' -replace '[\r\n\s]',''" "`n"
    ps .= "  $pkcs8bytes = [Convert]::FromBase64String($pemClean)" "`n"
    ps .= "  Add-Type -AssemblyName System.Security" "`n"
    ps .= "  $cng = [System.Security.Cryptography.CngKey]::Import($pkcs8bytes, [System.Security.Cryptography.CngKeyBlobFormat]::Pkcs8PrivateBlob)" "`n"
    ps .= "  $rsa = [System.Security.Cryptography.RSACng]::new($cng)" "`n"
    ps .= "  $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()" "`n"
    ps .= "  $exp = $now + 3600" "`n"
    ps .= "  function B64Url($bytes) {" "`n"
    ps .= "    [Convert]::ToBase64String($bytes) -replace '\+','-' -replace '/','_' -replace '=',''" "`n"
    ps .= "  }" "`n"
    ps .= "  $hdr = [ordered]@{ alg='RS256'; typ='JWT' } | ConvertTo-Json -Compress" "`n"
    ps .= "  $clm = [ordered]@{ iss=$email; scope='https://www.googleapis.com/auth/spreadsheets'; aud=$tokenUri; exp=$exp; iat=$now } | ConvertTo-Json -Compress" "`n"
    ps .= "  $headerB64   = B64Url ([System.Text.Encoding]::UTF8.GetBytes($hdr))" "`n"
    ps .= "  $claimB64    = B64Url ([System.Text.Encoding]::UTF8.GetBytes($clm))" "`n"
    ps .= "  $sigInput    = $headerB64 + '.' + $claimB64" "`n"
    ps .= "  $sigBytes = $rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($sigInput), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)" "`n"
    ps .= "  $jwt         = $sigInput + '.' + (B64Url $sigBytes)" "`n"
    ps .= "  # Exchange JWT for access token" "`n"
    ps .= "  $body = 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + $jwt" "`n"
    ps .= "  $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded'" "`n"
    ps .= "  [System.IO.File]::WriteAllText('" psOut "', $resp.access_token)" "`n"
    ps .= "} catch {" "`n"
    ps .= "  $errMsg = $_.Exception.Message" "`n"
    ps .= "  [System.IO.File]::WriteAllText('" psErr "', $errMsg)" "`n"
    ps .= "}" "`n"
    ; Write PS script without BOM
    psBytes := Buffer(StrPut(ps, "UTF-8") - 1)
    StrPut(ps, psBytes, "UTF-8")
    fp := FileOpen(tmpPs, "w", "UTF-8-RAW")
    fp.RawWrite(psBytes)
    fp.Close()

    RunWait('powershell.exe -ExecutionPolicy Bypass -File "' tmpPs '"',, "Hide")

    token := ""
    if FileExist(tmpOut) {
        token := Trim(FileRead(tmpOut, "UTF-8"))
        FileDelete(tmpOut)
    }

    if FileExist(tmpErr) {
        errText := Trim(FileRead(tmpErr, "UTF-8"))
        FileDelete(tmpErr)
        if errText != ""
            MsgBox("PowerShell error during auth:`n`n" errText, "Auth Error", 16)
    } else if token = "" {
        ; PS script ran but wrote neither token nor error - show the script for inspection
        psContent := FileExist(tmpPs) ? FileRead(tmpPs, "UTF-8") : "(ps file already deleted)"
        MsgBox("No token and no error file written.`n`nPS script preview:`n" SubStr(psContent, 1, 400), "Auth Debug", 16)
    }

    if FileExist(tmpPs)
        FileDelete(tmpPs)

    return token
}

; ---- FETCH SHEET DATA AND APPLY FILTERS ----
; Returns array of Column M values where Col N = "Smartboard" AND Col O is blank
FetchFolderNamesFromSheet() {
    global SA_CLIENT_EMAIL, SA_TOKEN_URI, SA_PRIVATE_KEY, SHEET_ID, SHEET_NAME, FILTER_COL_N, FILTER_COL_O, FILTER_COL_I, FILTER_COL_K, TARGET_COL_M

    accessToken := GetAccessToken()
    if !accessToken
        return []

    range := SHEET_NAME "!A:O"
    url   := "https://sheets.googleapis.com/v4/spreadsheets/" SHEET_ID "/values/" range

    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", url, false)
    whr.SetRequestHeader("Authorization", "Bearer " accessToken)
    whr.Send()

    resp := whr.ResponseText

    if whr.Status != 200 {
        MsgBox("Google Sheets API error " whr.Status ":`n" SubStr(resp, 1, 600), "API Error", 16)
        return []
    }

    names := []
    valStart := InStr(resp, '"values"')
    if !valStart {
        MsgBox("No data found in sheet response.", "Sheet Error", 16)
        return []
    }

    ; Build allowed teachers lookup map
    allowedTeachers := Map()
allowedTeachers["Nour Essam"] := true
allowedTeachers["Mohamed Hossam"] := true
allowedTeachers["Eslam Morsy"] := true
allowedTeachers["Ahmed Yehia"] := true
allowedTeachers["Eslam Abdelazeem"] := true
allowedTeachers["Mohamed Ebrahim"] := true
allowedTeachers["Ahmed Salah"] := true
allowedTeachers["Hossam Elashry"] := true
allowedTeachers["Atef Ramzy"] := true
allowedTeachers["Ekrami Eltaweel"] := true
allowedTeachers["Abanoub Hezkial"] := true
allowedTeachers["Ali Youssef"] := true
allowedTeachers["Mahitab Tarek"] := true
allowedTeachers["Hany Elrefaey"] := true
allowedTeachers["Karolen Samy"] := true
allowedTeachers["Omar Hussien"] := true
allowedTeachers["Saleh Abuzaid"] := true
allowedTeachers["Hoda Farouk"] := true
allowedTeachers["Mina Fayez"] := true
allowedTeachers["Ahmed bakr"] := true
allowedTeachers["Bosy Magdy"] := true
allowedTeachers["Fatma Ebrahim"] := true
allowedTeachers["Ahmed Salem"] := true

    firstRow := true
    sheetRow := 1
    searchFrom := valStart

    loop {
        rowStart := InStr(resp, "[", , searchFrom)
        if !rowStart
            break
        rowEnd := InStr(resp, "]", , rowStart)
        if !rowEnd
            break

        if SubStr(resp, rowStart + 1, 1) = "[" {
            searchFrom := rowStart + 1
            continue
        }

        rowJson := SubStr(resp, rowStart, rowEnd - rowStart + 1)

        if rowJson = "[" || StrLen(rowJson) < 3 {
            searchFrom := rowEnd + 1
            continue
        }

        if firstRow {
            firstRow := false
            searchFrom := rowEnd + 1
            continue
        }

        sheetRow++

        cells := []
        cellFrom := 2
        loop {
            qStart := InStr(rowJson, '"', , cellFrom)
            if !qStart
                break
            qEnd := qStart + 1
            loop {
                ch := SubStr(rowJson, qEnd, 1)
                if ch = ""
                    break
                if ch = '"' && SubStr(rowJson, qEnd - 1, 1) != Chr(92)
                    break
                qEnd++
            }
            cellVal := SubStr(rowJson, qStart + 1, qEnd - qStart - 1)
            cells.Push(cellVal)
            cellFrom := qEnd + 1
        }

        colN_val := cells.Length >= FILTER_COL_N ? cells[FILTER_COL_N] : ""
        colO_val := cells.Length >= FILTER_COL_O ? Trim(cells[FILTER_COL_O]) : ""
        colI_val := cells.Length >= FILTER_COL_I ? cells[FILTER_COL_I] : ""
        colK_val := cells.Length >= FILTER_COL_K ? Trim(cells[FILTER_COL_K]) : ""
        colM_val := cells.Length >= TARGET_COL_M ? Trim(cells[TARGET_COL_M]) : ""

        ; Extract teacher name from col K: find last hyphen and take everything after it
        lastHyphen := 0
        searchPos  := 1
        loop {
            found := InStr(colK_val, "-", , searchPos)
            if !found
                break
            lastHyphen := found
            searchPos  := found + 1
        }
        teacherName := lastHyphen ? Trim(SubStr(colK_val, lastHyphen + 1)) : colK_val

        colN_trimmed := Trim(colN_val)
        passNOI := (colN_trimmed = "SMARTBOARD" || colN_trimmed = "TO BE CHECKED LATER") && colO_val = "" && !InStr(colI_val, "(Q)") && !InStr(colI_val, "امتحان") && colM_val != ""

        if passNOI && allowedTeachers.Has(teacherName)
            names.Push({name: colM_val, row: sheetRow})

        searchFrom := rowEnd + 1
        if searchFrom > StrLen(resp)
            break
    }

    return names
}


; ================================================================
; LOAD SEARCH ROOTS
; ================================================================

LoadSearchRoots() {
    global searchRoots, pathsFile
    searchRoots := []
    if !FileExist(pathsFile)
        return
    raw := FileRead(pathsFile)
    for _, line in StrSplit(raw, "`n", "`r") {
        line := Trim(line)
        if line = "" || SubStr(line, 1, 1) = ";"
            continue
        if InStr(line, "=") {
            eqPos        := InStr(line, "=")
            friendlyName := Trim(SubStr(line, 1, eqPos - 1))
            path         := Trim(SubStr(line, eqPos + 1))
        } else {
            path := line
            if SubStr(path, 1, 2) = "\\" {
                parts        := StrSplit(LTrim(path, "\\"), "\\")
                friendlyName := parts[1]
            } else {
                SplitPath(path, , , , &friendlyName)
                if friendlyName = ""
                    friendlyName := path
            }
        }
        if path != ""
            searchRoots.Push({path: path, name: friendlyName})
    }
}

ApplyRemotePolicy() {
    global APP_VERSION, MANIFEST_URL

    if !MANIFEST_URL || InStr(MANIFEST_URL, "ORG/REPO")
        return

    manifest := HttpGetText(MANIFEST_URL)
    if manifest = ""
        return

    action := StrLower(JsonGetString(manifest, "action"))
    remoteVersion := JsonGetString(manifest, "version")
    remoteVersionLower := StrLower(remoteVersion)
    policyMessage := JsonGetString(manifest, "message")

    if (action = "disable" || remoteVersionLower = "disable") {
        if policyMessage != ""
            MsgBox(policyMessage, "Disabled", 48)
        ExitApp
    }

    if (action = "delete" || remoteVersionLower = "delete") {
        ScheduleSelfDelete(policyMessage)
        ExitApp
    }

    downloadUrl := JsonGetString(manifest, "download_url")
    if downloadUrl = ""
        return

    targetExt := A_IsCompiled ? ".exe" : ".ahk"
    remoteExt := GetUrlExtension(downloadUrl)
    if (remoteExt != "" && remoteExt != targetExt)
        return

    needsUpdate := false
    if (action = "force_update") {
        needsUpdate := true
    } else if (remoteVersion != "" && IsVersionNewer(remoteVersion, APP_VERSION)) {
        needsUpdate := true
    }

    if !needsUpdate
        return

    tempUpdateFile := A_Temp "\RawMinuteCounter_Update_" A_TickCount targetExt
    if !TryDownloadFile(downloadUrl, tempUpdateFile)
        return

    ScheduleSelfReplace(tempUpdateFile)
    ExitApp
}

HttpGetText(url) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    try {
        whr.Open("GET", url, false)
        whr.SetTimeouts(4000, 4000, 8000, 8000)
        whr.Send()
    } catch {
        return ""
    }
    if whr.Status != 200
        return ""
    return whr.ResponseText
}

TryDownloadFile(url, outFile) {
    try {
        Download(url, outFile)
        return FileExist(outFile)
    } catch {
        return false
    }
}

GetUrlExtension(url) {
    cut := url
    qPos := InStr(cut, "?")
    if qPos
        cut := SubStr(cut, 1, qPos - 1)
    hashPos := InStr(cut, "#")
    if hashPos
        cut := SubStr(cut, 1, hashPos - 1)
    SplitPath(cut, , , &ext)
    if ext = ""
        return ""
    return "." StrLower(ext)
}

JsonGetString(json, key) {
    pattern := '"' key '"\s*:\s*"((?:\\.|[^"\\])*)"'
    if !RegExMatch(json, pattern, &m)
        return ""
    return JsonUnescape(m[1])
}

JsonUnescape(text) {
    text := StrReplace(text, "\\/", "/")
    text := StrReplace(text, '\\"', '"')
    text := StrReplace(text, "\\n", "`n")
    text := StrReplace(text, "\\r", "`r")
    text := StrReplace(text, "\\t", "`t")
    text := StrReplace(text, "\\\\", Chr(92))
    return text
}

IsVersionNewer(remoteVersion, localVersion) {
    remoteParts := StrSplit(remoteVersion, ".")
    localParts := StrSplit(localVersion, ".")
    maxLen := (remoteParts.Length > localParts.Length) ? remoteParts.Length : localParts.Length
    Loop maxLen {
        idx := A_Index
        r := (idx <= remoteParts.Length) ? Integer(remoteParts[idx]) : 0
        l := (idx <= localParts.Length) ? Integer(localParts[idx]) : 0
        if r > l
            return true
        if r < l
            return false
    }
    return false
}

ScheduleSelfReplace(updateFilePath) {
    targetPath := A_ScriptFullPath
    launcherPath := A_Temp "\rmc_apply_update_" A_TickCount ".cmd"
    q := Chr(34)
    script := "@echo off`r`n"
    script .= "ping 127.0.0.1 -n 3 >nul`r`n"
    script .= ":retry`r`n"
    script .= "copy /Y " q updateFilePath q " " q targetPath q " >nul`r`n"
    script .= "if errorlevel 1 (`r`n"
    script .= "  ping 127.0.0.1 -n 2 >nul`r`n"
    script .= "  goto retry`r`n"
    script .= ")`r`n"
    script .= "start " q q " " q targetPath q "`r`n"
    script .= "del /q " q updateFilePath q " >nul 2>&1`r`n"
    script .= "del /q " q Chr(37) "~f0" q "`r`n"
    FileAppend(script, launcherPath, "CP0")
    Run('"' launcherPath '"',, "Hide")
}

ScheduleSelfDelete(reasonText := "") {
    if reasonText != ""
        MsgBox(reasonText, "Disabled", 48)
    targetPath := A_ScriptFullPath
    launcherPath := A_Temp "\rmc_self_delete_" A_TickCount ".cmd"
    q := Chr(34)
    script := "@echo off`r`n"
    script .= "ping 127.0.0.1 -n 3 >nul`r`n"
    script .= "del /f /q " q targetPath q " >nul 2>&1`r`n"
    script .= "del /q " q Chr(37) "~f0" q "`r`n"
    FileAppend(script, launcherPath, "CP0")
    Run('"' launcherPath '"',, "Hide")
}

SavePaths(cfgGui, inputEdit, *) {
    global pathsFile, searchRoots
    inputText := inputEdit.Value
    cfgGui.Destroy()
    if !inputText {
        MsgBox("No paths entered. Exiting.")
        ExitApp
    }
    if FileExist(pathsFile)
        FileDelete(pathsFile)
    FileAppend(inputText, pathsFile)
    LoadSearchRoots()
    if searchRoots.Length = 0 {
        MsgBox("No valid paths entered. Exiting.")
        ExitApp
    }
    ShowFolderInputGUI()
}

CancelScript(*) {
    ExitApp
}


; ================================================================
; CONFIG GUI
; ================================================================

ShowConfigGUI() {
    global pathsFile, APP_VERSION

    cfgGui := Gui("+AlwaysOnTop -Resize", "Configure Search Roots v" APP_VERSION)
    cfgGui.BackColor := "F0F4F8"
    cfgGui.SetFont("s10 c1A1A2E", "Segoe UI")

    cfgGui.Add("Text", "x0 y0 w500 h50 Background1A1A2E")
    cfgGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
    cfgGui.Add("Text", "x20 y14 w460 BackgroundTrans", "Configure Search Roots (v" APP_VERSION ")")

    cfgGui.SetFont("s9 c444444", "Segoe UI")
    cfgGui.Add("Text", "x20 y68 w460", "Enter one search root per line.")
    cfgGui.Add("Text", "x20 y84 w460", "Optionally give each a friendly name:")
    cfgGui.SetFont("s9 c888888 Italic", "Segoe UI")
    cfgGui.Add("Text", "x20 y100 w460", "Example:  Main Server = \\S1Storage\2026")

    cfgGui.SetFont("s10 c1A1A2E", "Segoe UI")
    inputEdit := cfgGui.Add("Edit", "x20 y124 w460 h160 Border", "")

    cfgGui.SetFont("s9 Bold cFFFFFF", "Segoe UI")
    okBtn := cfgGui.Add("Button", "x340 y358 w65 h28 Background1A6B3A", "Save")
    cfgGui.SetFont("s9 c555555", "Segoe UI")
    cancelBtn := cfgGui.Add("Button", "x414 y358 w66 h28", "Cancel")

    okBtn.OnEvent("Click", SavePaths.Bind(cfgGui, inputEdit))
    cancelBtn.OnEvent("Click", (*) => (cfgGui.Destroy(), ExitApp()))
    cfgGui.Show("w500 h400")
}

OpenConfig(*) {
    global pathsFile
    if !FileExist(pathsFile)
        MsgBox("Config file not found.")
    else
        Run(pathsFile)
}


; ================================================================
; MAIN FOLDER INPUT GUI  (now with "Load from Sheet" button)
; ================================================================

ShowFolderInputGUI() {
    global FolderEdit, inputGui, APP_VERSION

    inputGui := Gui("+AlwaysOnTop -Resize", "Raw Minute Counter v" APP_VERSION)
    inputGui.BackColor := "F0F4F8"
    inputGui.SetFont("s10 c1A1A2E", "Segoe UI")

    ; Header bar
    inputGui.Add("Text", "x0 y0 w500 h50 Background1A1A2E")
    inputGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
    inputGui.Add("Text", "x20 y14 w420 BackgroundTrans", "Footage Search (v" APP_VERSION ")")

    ; Gear button
    inputGui.SetFont("s11 cCCCCCC", "Segoe UI")
    configBtn := inputGui.Add("Button", "x441 y12 w40 h26 BackgroundTrans", "Cfg")

    ; Instruction label
    inputGui.SetFont("s9 c444444", "Segoe UI")
    inputGui.Add("Text", "x20 y64 w460", "Enter folder names to search, one per line:")

    ; Text area
    inputGui.SetFont("s10 c1A1A2E", "Segoe UI")
    FolderEdit := inputGui.Add("Edit", "x20 y82 w460 h240 Border", "")

    ; Sheet status label
    inputGui.SetFont("s9 c888888 Italic", "Segoe UI")
    sheetStatus := inputGui.Add("Text", "x20 y330 w460", "")

    ; Buttons row
    inputGui.SetFont("s9 Bold c1A1A2E", "Segoe UI")
    sheetBtn := inputGui.Add("Button", "x20 y358 w140 h28", "Load from Sheet")

    inputGui.SetFont("s9 Bold cFFFFFF", "Segoe UI")
    okBtn := inputGui.Add("Button", "x310 y358 w80 h28 Background1A6B3A", "Search")
    inputGui.SetFont("s9 c555555", "Segoe UI")
    cancelBtn := inputGui.Add("Button", "x400 y358 w80 h28", "Cancel")

    okBtn.OnEvent("Click", FolderOkClick)
    cancelBtn.OnEvent("Click", CancelScript)
    configBtn.OnEvent("Click", OpenConfig)

    sheetBtn.OnEvent("Click", (*) => LoadFromSheet(sheetStatus))

    inputGui.Show("w500 h404")
}

; ---- LOAD FROM SHEET BUTTON HANDLER ----
LoadFromSheet(statusLabel) {
    global FolderEdit, sheetRowMap
    statusLabel.Text := "Connecting to Google Sheets..."
    statusLabel.Opt("c888888")

    names := FetchFolderNamesFromSheet()

    if names.Length = 0 {
        statusLabel.Text := "No matching rows found (Col N = 'Smartboard', Col O blank)."
        statusLabel.Opt("cCC4400")
        return
    }

    ; Store row numbers for use in UpdateSheet
    sheetRowMap := Map()
    for _, entry in names {
        if !sheetRowMap.Has(entry.name)
            sheetRowMap[entry.name] := entry.row
    }

    ; Populate the text box so the user can review/edit before searching
    FolderEdit.Value := ""
    listText := ""
    for _, entry in names
        listText .= entry.name "`n"
    FolderEdit.Value := Trim(listText)

    statusLabel.Text := "Loaded " names.Length " folder(s) from sheet. Review above, then click Search."
    statusLabel.Opt("c1A6B3A")
}






; ================================================================
; SEARCH KICK-OFF
; ================================================================

FolderOkClick(*) {
    global folderOrder, FolderEdit, inputGui, folders, searchRoots
    LoadSearchRoots()
    if searchRoots.Length = 0 {
        MsgBox("No valid search paths found. Please configure them via Cfg.")
        return
    }
    folderList := FolderEdit.Value
    inputGui.Destroy()
    folderOrder := []
    folders := Map()
    for _, name in StrSplit(folderList, "`n", "`r") {
        name := Trim(name)
        if !name || folders.Has(name)
            continue
        folderOrder.Push(name)
        folders[name] := {total: 0, count: 0}
    }
    if folderOrder.Length = 0 {
        MsgBox("No folders entered. Try again.")
        ShowFolderInputGUI()
    } else {
        StartSearch()
    }
}


; ================================================================
; SEARCH FUNCTION  (unchanged from v2)
; ================================================================

StartSearch() {
    global folderOrder, folders, searchRoots, unreachablePaths, APP_VERSION
    unreachablePaths := []

    myGui := Gui("+AlwaysOnTop -Resize", "Searching... v" APP_VERSION)
    myGui.BackColor := "F0F4F8"

    myGui.Add("Text", "x0 y0 w560 h50 Background1A1A2E")
    myGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
    myGui.Add("Text", "x20 y14 w520 BackgroundTrans", "Searching... (v" APP_VERSION ")")

    myGui.SetFont("s9 c444444", "Segoe UI")
    myText := myGui.Add("Text", "x20 y62 w520", "Launching workers...")

    myGui.SetFont("s9 c1A1A2E", "Segoe UI")
    machineLabels := []
    yPos := 84
    for i, root in searchRoots {
        dot  := myGui.Add("Text", "x20 y" yPos " w14 h18 c888888", "o")
        lbl  := myGui.Add("Text", "x38 y" yPos " w500 h18", root.name " - Pinging...")
        machineLabels.Push({dot: dot, lbl: lbl, done: false})
        yPos += 22
    }

    barY := yPos + 10
    myGui.SetFont("s9 c888888", "Segoe UI")
    myGui.Add("Text", "x20 y" barY " w80", "Progress:")
    myProgress := myGui.Add("Progress", "x100 y" (barY+1) " w380 h16 Background1A1A2E", 0)

    timerY := barY + 28
    myGui.SetFont("s9 c888888", "Segoe UI")
    myTimer  := myGui.Add("Text", "x20 y" timerY " w300", "Elapsed: 0s")
    myGui.SetFont("s9 c555555", "Segoe UI")
    myCancel := myGui.Add("Button", "x450 y" (timerY - 3) " w90 h24", "Cancel")

    cancelledFlag := {value: false}
    myCancel.OnEvent("Click", (*) => cancelledFlag.value := true)
    myGui.OnEvent("Close",    (*) => cancelledFlag.value := true)

    totalH := timerY + 42
    myGui.Show("w560 h" totalH)

    startTime := A_TickCount
    UpdateTimer() {
        elapsed := Round((A_TickCount - startTime) / 1000)
        myTimer.Text := "Elapsed: " elapsed "s"
    }
    SetTimer(UpdateTimer, 1000)

    inputFile := A_Temp "\ElkhetaInput_" A_TickCount ".txt"
    folderListText := ""
    for _, name in folderOrder
        folderListText .= name "`n"
    FileAppend(folderListText, inputFile, "UTF-8")

    workerOutputFiles := []
    workerStatusFiles := []
    workerStartTimes  := []
    for i, root in searchRoots {
        outFile    := A_Temp "\ElkhetaResult_" A_TickCount "_" i ".txt"
        statusFile := A_Temp "\ElkhetaStatus_" A_TickCount "_" i ".txt"
        if FileExist(outFile)
            FileDelete(outFile)
        if FileExist(statusFile)
            FileDelete(statusFile)
        workerOutputFiles.Push(outFile)
        workerStatusFiles.Push(statusFile)
        workerStartTimes.Push(A_TickCount)
        Run('"' A_ScriptFullPath '" /worker "' inputFile '" "' outFile '" "' root.path '"',, "Hide")
    }

    workerTimeoutMs := 600000  ; 10 min per worker - accounts for WOL boot time + search
    timeoutMs       := 660000  ; 11 min overall hard timeout
    pollStart       := A_TickCount
    hungWorkers     := []

    loop {
        if cancelledFlag.value
            break
        if (A_TickCount - pollStart) > timeoutMs {
            MsgBox("Search timed out after 5 minutes.", "Timeout", 48)
            break
        }

        doneCount    := 0
        pendingNames := []
        for i, outFile in workerOutputFiles {
            if FileExist(outFile) {
                doneCount += 1
                if !machineLabels[i].done {
                    machineLabels[i].done := true
                    ; Read final status from status file
                    finalStatus := ""
                    if FileExist(workerStatusFiles[i])
                        finalStatus := Trim(FileRead(workerStatusFiles[i], "UTF-8"))
                    finalStatus := StrReplace(finalStatus, "STATUS:", "")
                    if finalStatus = ""
                        finalStatus := "Done."
                    machineLabels[i].dot.Opt("cFFFFFF")
                    machineLabels[i].dot.Text := "*"
                    machineLabels[i].lbl.Opt("c1A6B3A")
                    machineLabels[i].lbl.Text := searchRoots[i].name " - " finalStatus
                }
            } else {
                workerElapsed := A_TickCount - workerStartTimes[i]
                if workerElapsed > workerTimeoutMs {
                    alreadyFlagged := false
                    for _, hi in hungWorkers
                        if hi = i
                            alreadyFlagged := true
                    if !alreadyFlagged {
                        hungWorkers.Push(i)
                        machineLabels[i].dot.Opt("cFF4444")
                        machineLabels[i].dot.Text := "*"
                        machineLabels[i].lbl.Opt("cFF4444")
                        machineLabels[i].lbl.Text := searchRoots[i].name " - Not responding"
                        MsgBox("Warning: " searchRoots[i].name " is taking unusually long.`n`nResults from this location may be missing.", "Worker Warning", 48)
                    }
                } else {
                    ; Read live status from status file
                    liveStatus := ""
                    if FileExist(workerStatusFiles[i])
                        liveStatus := Trim(FileRead(workerStatusFiles[i], "UTF-8"))
                    liveStatus := StrReplace(liveStatus, "STATUS:", "")
                    if liveStatus = ""
                        liveStatus := "Pinging..."
                    machineLabels[i].dot.Opt("c1A6B3A")
                    machineLabels[i].dot.Text := "*"
                    machineLabels[i].lbl.Opt("c1A1A2E")
                    machineLabels[i].lbl.Text := searchRoots[i].name " - " liveStatus
                }
                pendingNames.Push(searchRoots[i].name)
            }
        }

        percent := Min(Round((doneCount / workerOutputFiles.Length) * 95), 95)
        myProgress.Value := percent

        if pendingNames.Length > 0 {
            waiting := ""
            for _, name in pendingNames
                waiting .= (waiting ? ", " : "") name
            myText.Text := "Searching " waiting "... (" doneCount " of " workerOutputFiles.Length " done)"
        } else {
            myText.Text := "All done! Finalizing..."
        }
        myGui.Title := percent "% - Footage Search (v" APP_VERSION ")"

        if doneCount = workerOutputFiles.Length
            break
        Sleep 500
    }

    SetTimer(UpdateTimer, 0)

    if cancelledFlag.value {
        if FileExist(inputFile)
            FileDelete(inputFile)
        for i, outFile in workerOutputFiles {
            if FileExist(outFile)
                FileDelete(outFile)
            if FileExist(workerStatusFiles[i])
                FileDelete(workerStatusFiles[i])
        }
        try myGui.Destroy()
        return
    }

    myText.Text := "Merging results..."
    myProgress.Value := 98

    for i, outFile in workerOutputFiles {
        if !FileExist(outFile)
            continue
        raw := FileRead(outFile, "UTF-8")
        if raw = "UNREACHABLE" {
            unreachablePaths.Push(searchRoots[i].name)
            FileDelete(outFile)
            continue
        }
        for _, line in StrSplit(raw, "`n", "`r") {
            line := Trim(line)
            if !line
                continue
            parts := StrSplit(line, "`t")
            if parts.Length < 3
                continue
            folderName := parts[1]
            count      := Integer(parts[2])
            totalSecs  := Integer(parts[3])
            if folders.Has(folderName) {
                folders[folderName].count += count
                folders[folderName].total += totalSecs
            }
        }
        FileDelete(outFile)
        if FileExist(workerStatusFiles[i])
            FileDelete(workerStatusFiles[i])
    }
    if FileExist(inputFile)
        FileDelete(inputFile)

    if unreachablePaths.Length > 0 {
        msg := "Warning: The script could not reach:`n`n"
        for _, name in unreachablePaths
            msg .= "  - " name "`n"
        MsgBox(msg, "Warning", 48)
    }

    myProgress.Value := 100
    Sleep 300
    myGui.Destroy()

    outputPath := A_ScriptDir "\RAWresults.xlsx"
    BuildXlsx(outputPath, folderOrder, folders)

    ; ---- PROMPT USER: view results or update sheet ----
    ShowResultsChoice(outputPath, folderOrder, folders)
}



; ================================================================
; RESULTS CHOICE GUI
; ================================================================

ShowResultsChoice(outputPath, folderOrder, folders) {
    global APP_VERSION
    rGui := Gui("+AlwaysOnTop -Resize", "Search Complete v" APP_VERSION)
    rGui.BackColor := "F0F4F8"

    rGui.Add("Text", "x0 y0 w420 h50 Background1A1A2E")
    rGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
    rGui.Add("Text", "x20 y14 w380 BackgroundTrans", "Search Complete (v" APP_VERSION ")")

    rGui.SetFont("s10 c1A1A2E", "Segoe UI")
    rGui.Add("Text", "x20 y66 w380", "What would you like to do with the results?")

    rGui.SetFont("s9 Bold cFFFFFF", "Segoe UI")
    sheetBtn := rGui.Add("Button", "x20  y104 w180 h32 Background1A6B3A", "Update Sheet (Col O)")
    rGui.SetFont("s9 Bold c1A1A2E", "Segoe UI")
    excelBtn := rGui.Add("Button", "x210 y104 w90  h32", "View Excel")
    rGui.SetFont("s9 c555555", "Segoe UI")
    skipBtn  := rGui.Add("Button", "x310 y104 w90  h32", "Close")

    sheetBtn.OnEvent("Click", (*) => (rGui.Destroy(), UpdateSheet(folderOrder, folders)))
    excelBtn.OnEvent("Click", (*) => (rGui.Destroy(), Run(outputPath), ExitApp()))
    skipBtn.OnEvent("Click",  (*) => (rGui.Destroy(), ExitApp()))
    rGui.Show("w420 h152")
}


; ================================================================
; WRITE RESULTS BACK TO GOOGLE SHEET (Column O)
; Uses the Sheets API batchUpdate to write only FOUND rows.
; Matches by Column M value. Skips NOT FOUND (count=0) rows.
; ================================================================

UpdateSheet(folderOrder, folders) {
    global SA_CLIENT_EMAIL, SA_TOKEN_URI, SA_PRIVATE_KEY, SHEET_ID, SHEET_NAME, SHEET_GID, sheetRowMap, APP_VERSION

    ; Get fresh access token
    accessToken := GetAccessToken()
    if !accessToken {
        MsgBox("Could not get access token. Sheet not updated.", "Error", 16)
        ExitApp
    }

    ; Build updates list using row numbers already captured by FetchFolderNamesFromSheet
    updates := []
    for _, folderName in folderOrder {
        if folders[folderName].count = 0
            continue
        if !sheetRowMap.Has(folderName)
            continue
        rowNum := sheetRowMap[folderName]
        minutes := Round(folders[folderName].total / 60)
        updates.Push({row: rowNum, minutes: minutes, folderName: folderName})
    }



    if updates.Length = 0 {
        MsgBox("No matching rows found in the sheet to update.", "Nothing to Update", 48)
        ExitApp
    }

    ; ---- PREVIEW: scrollable GUI showing every row that will be updated ----
    previewLines := ""
    for _, upd in updates
        previewLines .= "O" upd.row "`t" upd.minutes " min`t" upd.folderName "`n"

    pGui := Gui("+AlwaysOnTop -Resize", "Confirm Sheet Update v" APP_VERSION)
    pGui.BackColor := "F0F4F8"
    pGui.Add("Text", "x0 y0 w580 h50 Background1A1A2E")
    pGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
    pGui.Add("Text", "x20 y14 w540 BackgroundTrans", "Confirm Sheet Update (v" APP_VERSION ")")
    pGui.SetFont("s9 c444444", "Segoe UI")
    pGui.Add("Text", "x20 y60 w540", updates.Length " cell(s) in column O will be updated. No other columns touched.")
    pGui.SetFont("s9 c1A1A2E", "Segoe UI")
    pGui.Add("Text", "x20 y80 w80", "Cell")
    pGui.Add("Text", "x70 y80 w80", "Minutes")
    pGui.Add("Text", "x150 y80 w410", "Folder Name")
    previewEdit := pGui.Add("Edit", "x20 y98 w540 h320 ReadOnly -Wrap", previewLines)
    pGui.SetFont("s9 Bold cFFFFFF", "Segoe UI")
    confirmBtn := pGui.Add("Button", "x380 y430 w85 h28 Background1A6B3A", "Proceed")
    pGui.SetFont("s9 c555555", "Segoe UI")
    cancelBtn  := pGui.Add("Button", "x475 y430 w85 h28", "Cancel")

    confirmed := false
    confirmBtn.OnEvent("Click", (*) => (confirmed := true, pGui.Destroy()))
    cancelBtn.OnEvent("Click",  (*) => pGui.Destroy())
    pGui.OnEvent("Close",       (*) => pGui.Destroy())
    pGui.Show("w580 h474")
    WinWaitClose(pGui)

    if !confirmed {
        MsgBox("Sheet update cancelled. No changes were made.", "Cancelled", 64)
        ExitApp
    }

    ; ---- HIGHLIGHT TARGET CELLS PURPLE ----
    colorRequests := ""
    firstColor := true
    for _, upd in updates {
        if !firstColor
            colorRequests .= ","
        firstColor := false
        colorRequests .= '{"repeatCell":{"range":{"sheetId":' SHEET_GID ',"startRowIndex":' (upd.row - 1) ',"endRowIndex":' upd.row ',"startColumnIndex":14,"endColumnIndex":15},"cell":{"userEnteredFormat":{"backgroundColor":{"red":0.502,"green":0,"blue":0.502}}},"fields":"userEnteredFormat.backgroundColor"}}'
    }
    colorUrl := "https://sheets.googleapis.com/v4/spreadsheets/" SHEET_ID ":batchUpdate"
    whrColor := ComObject("WinHttp.WinHttpRequest.5.1")
    whrColor.Open("POST", colorUrl, false)
    whrColor.SetRequestHeader("Authorization", "Bearer " accessToken)
    whrColor.SetRequestHeader("Content-Type", "application/json")
    whrColor.Send('{"requests":[' colorRequests ']}')
    if whrColor.Status != 200 {
        MsgBox("Highlight step failed (HTTP " whrColor.Status ").`n`n" SubStr(whrColor.ResponseText, 1, 400) "`nNo values were written.", "Highlight Error", 16)
        ExitApp
    }

    resp := MsgBox("The cells that would be updated in column O have been highlighted in purple on the sheet.`n`n"
        . "Please review them in Google Sheets.`n`n"
        . "Click 'Yes' to proceed with writing the minutes to those highlighted cells,`n"
        . "or 'No' to cancel without making any changes.", "Confirm Write to Sheet", "YesNo 32")

    if resp = "No" {
        ; Clear the purple highlight before cancelling
        cancelClearReqs := ""
        firstCC := true
        for _, upd in updates {
            if !firstCC
                cancelClearReqs .= ","
            firstCC := false
            cancelClearReqs .= '{"repeatCell":{"range":{"sheetId":' SHEET_GID ',"startRowIndex":' (upd.row - 1) ',"endRowIndex":' upd.row ',"startColumnIndex":14,"endColumnIndex":15},"cell":{"userEnteredFormat":{"backgroundColor":{"red":1,"green":1,"blue":1}}},"fields":"userEnteredFormat.backgroundColor"}}'
        }
        whrCC := ComObject("WinHttp.WinHttpRequest.5.1")
        whrCC.Open("POST", "https://sheets.googleapis.com/v4/spreadsheets/" SHEET_ID ":batchUpdate", false)
        whrCC.SetRequestHeader("Authorization", "Bearer " accessToken)
        whrCC.SetRequestHeader("Content-Type", "application/json")
        whrCC.Send('{"requests":[' cancelClearReqs ']}')
        MsgBox("Sheet update cancelled. Highlight removed.", "Cancelled", 64)
        ExitApp
    }

    ; ---- WRITE MINUTES TO COLUMN O ----
    jsonData := '{"valueInputOption":"RAW","data":['
    first := true
    for _, upd in updates {
        if !first
            jsonData .= ","
        first := false
        jsonData .= '{"range":"' SHEET_NAME '!O' upd.row '","values":[[' upd.minutes ']]}'
    }
    jsonData .= "]}"
    updateUrl := "https://sheets.googleapis.com/v4/spreadsheets/" SHEET_ID "/values:batchUpdate"
    whr2 := ComObject("WinHttp.WinHttpRequest.5.1")
    whr2.Open("POST", updateUrl, false)
    whr2.SetRequestHeader("Authorization", "Bearer " accessToken)
    whr2.SetRequestHeader("Content-Type", "application/json")
    whr2.Send(jsonData)
    if whr2.Status != 200 {
        MsgBox("Sheet update failed (HTTP " whr2.Status ").`n`n" SubStr(whr2.ResponseText, 1, 400), "Error", 16)
        ExitApp
    }

    ; ---- CLEAR HIGHLIGHT ----
    clearRequests := ""
    firstClear := true
    for _, upd in updates {
        if !firstClear
            clearRequests .= ","
        firstClear := false
        clearRequests .= '{"repeatCell":{"range":{"sheetId":' SHEET_GID ',"startRowIndex":' (upd.row - 1) ',"endRowIndex":' upd.row ',"startColumnIndex":14,"endColumnIndex":15},"cell":{"userEnteredFormat":{"backgroundColor":{"red":1,"green":1,"blue":1}}},"fields":"userEnteredFormat.backgroundColor"}}'
    }
    whrClear := ComObject("WinHttp.WinHttpRequest.5.1")
    whrClear.Open("POST", "https://sheets.googleapis.com/v4/spreadsheets/" SHEET_ID ":batchUpdate", false)
    whrClear.SetRequestHeader("Authorization", "Bearer " accessToken)
    whrClear.SetRequestHeader("Content-Type", "application/json")
    whrClear.Send('{"requests":[' clearRequests ']}')

    ; ---- FETCH COL P FOR UPDATED ROWS AND COMPARE ----
    ; Build a range covering all updated rows in col P
    minRow := updates[1].row
    maxRow := updates[1].row
    for _, upd in updates {
        if upd.row < minRow
            minRow := upd.row
        if upd.row > maxRow
            maxRow := upd.row
    }

    pRange := SHEET_NAME "!P" minRow ":P" maxRow
    pUrl := "https://sheets.googleapis.com/v4/spreadsheets/" SHEET_ID "/values/" pRange
    whr3 := ComObject("WinHttp.WinHttpRequest.5.1")
    whr3.Open("GET", pUrl, false)
    whr3.SetRequestHeader("Authorization", "Bearer " accessToken)
    whr3.Send()

    ; Build a map of rowNum -> col P value from the response
    colPMap := Map()
    if whr3.Status = 200 {
        pResp := whr3.ResponseText
        inValues := false
        pRowNum := minRow - 1
        for _, ln in StrSplit(pResp, "`n", "`r") {
            trimmed := Trim(ln)
            if !inValues {
                if InStr(trimmed, '"values"')
                    inValues := true
                continue
            }
            if trimmed = "[" {
                pRowNum++
                continue
            }
            if StrLen(trimmed) >= 2 && SubStr(trimmed, 1, 1) = '"' && SubStr(trimmed, StrLen(trimmed), 1) = '"' {
                cellVal := SubStr(trimmed, 2, StrLen(trimmed) - 2)
                colPMap[pRowNum] := Integer(cellVal)
            }
        }
    }

    ; Compare O (written minutes) vs P and collect warnings
    warnings := ""
    for _, upd in updates {
        if !colPMap.Has(upd.row)
            continue
        pVal := colPMap[upd.row]
        diff := upd.minutes - pVal
        ; Warn if written value is more than 1 above or more than 10 below col P
        if diff > 1 || diff < -10
            warnings .= "Row " upd.row ": wrote " upd.minutes " min, col P = " pVal " min (diff " (diff > 0 ? "+" : "") diff ")`n" upd.folderName "`n`n"
    }

    if warnings != ""
        MsgBox("Done! Updated " updates.Length " row(s) in column O.`n`nWarnings - values outside expected range:`n`n" warnings, "Sheet Updated with Warnings", 48)
    else
        MsgBox("Done! Updated " updates.Length " row(s) in column O.", "Sheet Updated", 64)

    ExitApp
}

; ================================================================
; BUILD XLSX  (unchanged from v2)
; ================================================================

BuildXlsx(outputPath, folderOrder, folders) {
    try {
        xl := ComObject("Excel.Application")
    } catch {
        outputPath := StrReplace(outputPath, ".xlsx", ".csv")
        csv := "FolderName`tStatus`tFolderCount`tTotalMinutes`r`n"
        for _, folderName in folderOrder {
            safeName := StrReplace(folderName, '"', '""')
            if folders[folderName].count = 0 {
                csv .= '"' safeName '"' "`tNOT FOUND`t0`t`r`n"
            } else {
                minutes := Round(folders[folderName].total / 60)
                csv .= '"' safeName '"' "`tFOUND`t" folders[folderName].count "`t" minutes "`r`n"
            }
        }
        if FileExist(outputPath)
            FileDelete(outputPath)
        FileAppend(csv, outputPath, "UTF-16")
        return
    }

    xl.Visible := false
    xl.DisplayAlerts := false
    wb := xl.Workbooks.Add()
    ws := wb.Worksheets(1)
    ws.Name := "Results"

    lastRow := folderOrder.Length + 1

    headers := ["Folder Name", "Status", "Folder Count", "Total Minutes"]
    for i, h in headers
        ws.Cells(1, i).Value := h

    row := 2
    for _, folderName in folderOrder {
        ws.Cells(row, 1).Value := folderName
        if folders[folderName].count = 0 {
            ws.Cells(row, 2).Value := "NOT FOUND"
            ws.Cells(row, 3).Value := 0
        } else {
            ws.Cells(row, 2).Value := "FOUND"
            ws.Cells(row, 3).Value := folders[folderName].count
            ws.Cells(row, 4).Value := Round(folders[folderName].total / 60)
        }
        row++
    }

    tableRange := ws.Range("A1:D" lastRow)
    tbl := ws.ListObjects.Add(1, tableRange, , 1)
    tbl.Name := "ResultsTable"
    tbl.TableStyle := "TableStyleMedium2"

    headerRange := ws.Range("A1:D1")
    headerRange.Font.Bold := true
    headerRange.HorizontalAlignment := -4108

    ws.Columns(2).HorizontalAlignment := -4108
    ws.Columns(3).HorizontalAlignment := -4108
    ws.Columns(4).HorizontalAlignment := -4108

    ws.Columns(1).AutoFit()
    ws.Columns(2).ColumnWidth := 14
    ws.Columns(3).ColumnWidth := 14
    ws.Columns(4).ColumnWidth := 14

    ws.Rows(2).Select()
    ws.Application.ActiveWindow.FreezePanes := true

    if FileExist(outputPath)
        FileDelete(outputPath)
    wb.SaveAs(outputPath, 51)
    wb.Close(false)
    xl.Quit()
}


; ================================================================
; WORKER LOGIC  (unchanged from v2)
; ================================================================

RunWorker(inputFile, outputFile, rootPath) {
    WOL_EXE       := "C:\Program Files\Aquila Technology\WakeOnLAN\WakeOnLanC.exe"
    WOL_TIMEOUT   := 120000  ; 2 minutes max to wait for PC to boot
    WOL_PING_INTERVAL := 1000  ; re-ping every 1s after WOL (ping itself has -w 1000, so each cycle ~1s)

    ; ---- EXTRACT HOSTNAME FROM UNC PATH ----
    ; \\OP1\Share  ->  OP1
    ; Local paths get no WOL treatment
    hostname := ""
    if SubStr(rootPath, 1, 2) = "\\" {
        ; Skip the leading \\ then split on the first \ to get just the hostname
        stripped := SubStr(rootPath, 3)          ; "S1Storage\2026"
        parts    := StrSplit(stripped, "\")       ; ["S1Storage", "2026"]
        hostname := parts[1]                      ; "S1Storage"
    }

    ; ---- STATUS FILE (read by main instance to update the GUI label) ----
    ; outputFile is e.g. ElkhetaResult_xxx_1.txt - derive status path from it
    statusFile := StrReplace(outputFile, "ElkhetaResult_", "ElkhetaStatus_")

    WriteStatus(statusFile, msg) {
        try {
            if FileExist(statusFile)
                FileDelete(statusFile)
            FileAppend(msg, statusFile, "UTF-8")
        }
    }

    ; ---- LOAD FOLDER NAMES ----
    requested := Map()
    if FileExist(inputFile) {
        raw := FileRead(inputFile, "UTF-8")
        for _, line in StrSplit(raw, "`n", "`r") {
            line := Trim(line)
            if line != ""
                requested[line] := true
        }
    }
    if requested.Count = 0 {
        FileAppend("", outputFile)
        return
    }

    ; ---- DEBUG LOG ----
    dbgLog := A_Temp "\ElkhetaDebug_" A_TickCount "_" ProcessExist() ".txt"

    ; ---- CHECK REACHABILITY VIA PING (avoids DirExist hanging on offline UNC paths) ----
    weMadeItWake := false
    hostIsUp := false

    FileAppend(A_Now " Worker started. rootPath=" rootPath " hostname=" hostname "`n", dbgLog)
    FileAppend(A_Now " WOL_EXE exists: " FileExist(WOL_EXE) "`n", dbgLog)

    if hostname != "" {
        WriteStatus(statusFile, "STATUS:Pinging...")
        FileAppend(A_Now " Pinging " hostname "...`n", dbgLog)
        pingExit := RunWait('ping -n 2 -w 1000 ' hostname,, "Hide")
        FileAppend(A_Now " Ping exit code: " pingExit "`n", dbgLog)
        hostIsUp := (pingExit = 0)
    } else {
        WriteStatus(statusFile, "STATUS:Pinging...")
        hostIsUp := DirExist(rootPath) ? true : false
        FileAppend(A_Now " Local path check: " hostIsUp "`n", dbgLog)
    }

    FileAppend(A_Now " hostIsUp=" hostIsUp "`n", dbgLog)

    if !hostIsUp {
        if hostname = "" || !FileExist(WOL_EXE) {
            FileAppend(A_Now " No WOL possible. Marking UNREACHABLE.`n", dbgLog)
            WriteStatus(statusFile, "STATUS:Unreachable")
            FileAppend("UNREACHABLE", outputFile)
            return
        }
        ; Use the friendly hostname in the status message
        displayName := hostname
        WriteStatus(statusFile, "STATUS:Waking " displayName " up...")
        FileAppend(A_Now " Sending WOL wake: " WOL_EXE " -w -m " hostname "`n", dbgLog)
        Run('"' WOL_EXE '" -w -m ' hostname,, "Hide")
        weMadeItWake := true

        wakeStart := A_TickCount
        loop {
            pingExit := RunWait('ping -n 1 -w 1000 ' hostname,, "Hide")
            FileAppend(A_Now " Wake ping exit: " pingExit "`n", dbgLog)
            if pingExit = 0 {
                FileAppend(A_Now " Host responded to ping! Waiting for SMB share to become ready...`n", dbgLog)
                WriteStatus(statusFile, "STATUS:Waiting for share...")
                ; Ping success doesn't mean SMB is ready - poll the share path directly
                smbReady := false
                smbStart := A_TickCount
                smbTimeoutMs := 60000  ; up to 60s for SMB to come up after ping
                loop {
                    if DirExist(rootPath) {
                        FileAppend(A_Now " SMB share is accessible!`n", dbgLog)
                        smbReady := true
                        break
                    }
                    if (A_TickCount - smbStart) > smbTimeoutMs {
                        FileAppend(A_Now " SMB share never became accessible after ping.`n", dbgLog)
                        break
                    }
                    Sleep 500  ; check SMB every 500ms - react as soon as the share is up
                }
                hostIsUp := smbReady
                break
            }
            if (A_TickCount - wakeStart) > WOL_TIMEOUT {
                FileAppend(A_Now " WOL timeout exceeded.`n", dbgLog)
                break
            }
            Sleep WOL_PING_INTERVAL  ; brief pause between ping attempts while waiting for boot
        }

        if !hostIsUp {
            FileAppend(A_Now " PC never became fully accessible. Marking UNREACHABLE.`n", dbgLog)
            WriteStatus(statusFile, "STATUS:Unreachable")
            if weMadeItWake && hostname != "" {
                FileAppend(A_Now " Sending shutdown to " hostname " (woke it but SMB failed).`n", dbgLog)
                shutExit := RunWait(WOL_EXE ' -s -m ' hostname ' -t 0 -f',, "Hide")
                FileAppend(A_Now " shutdown.exe exit code: " shutExit (shutExit = 0 ? " (success)" : " (FAILED)") "`n", dbgLog)
            }
            FileAppend("UNREACHABLE", outputFile)
            return
        }
    }

    WriteStatus(statusFile, "STATUS:Searching...")
    FileAppend(A_Now " Proceeding to search.`n", dbgLog)

    ; ---- SEARCH ----
    shell := ComObject("Shell.Application")
    DURATION_COL := 27

    index := Map()
    Loop Files, rootPath "\*", "DR" {
        folderName := A_LoopFileName
        if !requested.Has(folderName)
            continue
        if !index.Has(folderName)
            index[folderName] := []
        index[folderName].Push(A_LoopFileFullPath)
    }

    results := Map()
    lastNamespacePath := ""
    cachedNamespace := ""

    for folderName, paths in index {
        results[folderName] := {total: 0, count: 0}
        for _, basePath in paths {
            results[folderName].count += 1
            Loop Files, basePath "\*MVI*", "FR" {
                SplitPath A_LoopFileFullPath, &fileName, &parentPath
                if parentPath != lastNamespacePath {
                    cachedNamespace := shell.Namespace(parentPath)
                    lastNamespacePath := parentPath
                }
                if !cachedNamespace
                    continue
                item := cachedNamespace.ParseName(fileName)
                if !item
                    continue
                duration := cachedNamespace.GetDetailsOf(item, DURATION_COL)
                if !duration
                    continue
                parts := StrSplit(duration, ":")
                if parts.Length = 3
                    results[folderName].total += (Integer(parts[1])*3600) + (Integer(parts[2])*60) + Integer(parts[3])
            }
        }
    }

    ; ---- WRITE OUTPUT ----
    out := ""
    for folderName, data in results
        out .= folderName "`t" data.count "`t" data.total "`n"

    if FileExist(outputFile)
        FileDelete(outputFile)
    FileAppend(out, outputFile, "UTF-8")

    ; ---- SHUT DOWN PC IF WE WOKE IT ----
    if weMadeItWake && hostname != "" {
        FileAppend(A_Now " Sending shutdown to " hostname " (we woke it, shutting it down now).`n", dbgLog)
        WriteStatus(statusFile, "STATUS:Done, turning " hostname " off...")
        shutExit := RunWait(WOL_EXE ' -s -m ' hostname ' -t 0 -f',, "Hide")
        FileAppend(A_Now " shutdown.exe exit code: " shutExit (shutExit = 0 ? " (success)" : " (FAILED - check admin rights, firewall, RPC on target)") "`n", dbgLog)
    } else {
        FileAppend(A_Now " No shutdown needed (weMadeItWake=" weMadeItWake " hostname=" hostname ").`n", dbgLog)
        WriteStatus(statusFile, "STATUS:Done.")
    }
}
