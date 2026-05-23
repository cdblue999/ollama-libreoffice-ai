Attribute VB_Name = "OllamaAI"
REM  ****  BASIC  ****

'====================================================================
' Ollama LibreOffice AI - Local AI Document Assistant for LibreOffice
' Version 1.0.0
'
' Single-file Basic module. No class modules, no external dependencies.
'
' Adapted from Ollama Outlook AI v1.1.1 by cdblue999
'
' Requirements:
'   - Ollama running at http://localhost:11434
'   - At least one model pulled (e.g., llama3.2:3b)
'   - LibreOffice 7.x or later
'   - Windows (uses COM for HTTP, registry, shell)
'
' All processing is fully local. No data leaves your machine.
'====================================================================

Option Explicit

'====================================================================
' Constants
'====================================================================
Private Const APP_NAME As String = "OllamaLibreOfficeAI"
Private Const APP_VERSION As String = "1.0.1"
Private Const OLLAMA_BASE_URL As String = "http://localhost:11434"
Private Const OLLAMA_DEFAULT_MODEL As String = "qwen2.5-coder:7b"
Private Const REQUEST_TIMEOUT_SECS As Long = 120
Private Const REG_PATH_ROOT As String = "HKEY_CURRENT_USER\Software\OllamaLibreOfficeAI"
Private Function CRLF() As String
    CRLF = Chr(13) + Chr(10)
End Function

'====================================================================
' Initialization - Run once after import to create default settings
'====================================================================
Public Sub Ollama_Initialize()
    On Error Resume Next

    Dim testVal As String
    testVal = GetRegSetting("Model", "")
    If testVal = "" Then
        SaveRegSetting "Model", OLLAMA_DEFAULT_MODEL
        SaveRegSetting "Prompt", "You are a helpful document assistant. Analyze the document content and respond professionally. Keep your response concise and well-structured."
        SaveRegSetting "Timeout", CStr(REQUEST_TIMEOUT_SECS)
        SaveRegSetting "ApiKey", ""
    End If

    MsgBox APP_NAME & " v" & APP_VERSION & " initialized." & CRLF & CRLF & "Run Tools > Macros > Ollama_ProcessDocument to use.", 64, APP_NAME
End Sub

'====================================================================
' Health Check - Is Ollama Running?
'====================================================================
Public Function Ollama_IsOllamaRunning() As Boolean
    On Error GoTo NotRunning
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", OLLAMA_BASE_URL & "/api/tags", False
    http.SetTimeouts 5000, 5000, 5000, 5000
    http.SetOption 0, "Ollama-LibreOffice-AI/1.0"
    http.Send
    Ollama_IsOllamaRunning = (http.Status = 200)
    Exit Function
NotRunning:
    Ollama_IsOllamaRunning = False
End Function

'====================================================================
' Refresh AI Models List (fetches from Ollama)
'====================================================================
Public Sub Ollama_RefreshModelsList()
    On Error GoTo RefreshError
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", OLLAMA_BASE_URL & "/api/tags", False
    http.SetTimeouts 5000, 5000, 5000, 5000
    http.SetOption 0, "Ollama-LibreOffice-AI/1.0"
    http.Send

    If http.Status = 200 Then
        Dim json As String
        json = http.ResponseText
        Dim models As String
        models = ExtractModelNames(json)
        If models <> "" Then
            SaveRegSetting "AvailableModels", models
        Else
            SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
        End If
    Else
        SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
    End If
    Exit Sub

RefreshError:
    SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
End Sub

'====================================================================
' Extract model names from Ollama API JSON response
'====================================================================
Private Function ExtractModelNames(ByVal json As String) As String
    On Error Resume Next
    Dim result As String
    result = ""

    Dim modelsStart As Long
    modelsStart = InStr(json, """models"":")
    If modelsStart = 0 Then
        ExtractModelNames = OLLAMA_DEFAULT_MODEL
        Exit Function
    End If

    Dim searchPos As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim modelName As String
    searchPos = modelsStart

    Do
        nameStart = InStr(searchPos, json, """name"":""")
        If nameStart = 0 Or nameStart > modelsStart + 10000 Then Exit Do
        nameStart = nameStart + 8
        nameEnd = InStr(nameStart, json, """")
        If nameEnd = 0 Or nameEnd <= nameStart Then Exit Do
        modelName = Mid(json, nameStart, nameEnd - nameStart)
        If modelName <> "" Then
            If result <> "" Then result = result & ","
            result = result & modelName
        End If
        searchPos = nameEnd + 1
    Loop

    If result = "" Then result = OLLAMA_DEFAULT_MODEL
    ExtractModelNames = result
End Function

'====================================================================
' Main AI Processing - Entry point for document processing
'====================================================================
Public Sub Ollama_ProcessDocument()
    On Error GoTo ProcessError

    If Not Ollama_IsOllamaRunning() Then
        MsgBox "Ollama is not running." & CRLF & CRLF & "Please start Ollama and try again.", 48, APP_NAME
        Exit Sub
    End If

    Dim doc As Object
    doc = ThisComponent

    If doc Is Nothing Then
        MsgBox "No document is open.", 48, APP_NAME
        Exit Sub
    End If

    Dim docText As String
    Dim docType As String

    If doc.supportsService("com.sun.star.text.TextDocument") Then
        docType = "Writer"
        docText = GetWriterDocumentText(doc)
    ElseIf doc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        docType = "Calc"
        docText = GetCalcDocumentText(doc)
    ElseIf doc.supportsService("com.sun.star.presentation.PresentationDocument") Then
        docType = "Impress"
        docText = GetImpressDocumentText(doc)
    Else
        MsgBox "Unsupported document type. Only Writer, Calc, and Impress are supported.", 48, APP_NAME
        Exit Sub
    End If

    If docText = "" Then
        MsgBox "No text content found in the document.", 48, APP_NAME
        Exit Sub
    End If

    If Len(docText) > 50000 Then
        docText = Left(docText, 50000)
    End If

    Dim modelName As String
    Dim systemPrompt As String
    Dim timeoutSecs As Long

    modelName = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    systemPrompt = GetRegSetting("Prompt", "You are a helpful document assistant. Analyze the document content and respond professionally. Keep your response concise and well-structured.")
    timeoutSecs = CLng(GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS)))

    Dim userContent As String
    userContent = "Document Type: " & docType & CRLF & CRLF & "Document Content:" & CRLF & docText

    Dim response As String
    response = Ollama_ProcessRequest(systemPrompt, userContent, modelName, timeoutSecs)

    If response <> "" Then
        If Left(response, 1) = "[" And InStr(response, "Ollama Error:") > 0 Then
            MsgBox response, 48, APP_NAME
        Else
            InsertResponseIntoDocument response, doc
        End If
    Else
        MsgBox "No response received from Ollama." & CRLF & CRLF & "Check that:" & CRLF & "  - Ollama is running (ollama list)" & CRLF & "  - The selected model is downloaded" & CRLF & "  - Your model name matches exactly", 48, APP_NAME
    End If

    Exit Sub

ProcessError:
    MsgBox "Error processing document: " & Err.Description, 16, APP_NAME
End Sub

'====================================================================
' Get text content from a Writer document
'====================================================================
Private Function GetWriterDocumentText(ByVal doc As Object) As String
    On Error Resume Next
    Dim text As Object
    Set text = doc.Text
    If text Is Nothing Then
        GetWriterDocumentText = ""
        Exit Function
    End If
    GetWriterDocumentText = text.String
End Function

'====================================================================
' Get text content from a Calc spreadsheet
'====================================================================
Private Function GetCalcDocumentText(ByVal doc As Object) As String
    On Error Resume Next
    Dim result As String
    result = ""
    Dim sheets As Object
    Set sheets = doc.Sheets
    If sheets Is Nothing Then
        GetCalcDocumentText = ""
        Exit Function
    End If
    Dim sheetCount As Long
    sheetCount = sheets.getCount()
    Dim i As Long
    For i = 0 To sheetCount - 1
        Dim sheet As Object
        Set sheet = sheets.getByIndex(i)
        If Not sheet Is Nothing Then
            If result <> "" Then result = result & CRLF & CRLF
            result = result & "=== Sheet: " & sheet.Name & " ===" & CRLF
            Dim cursor As Object
            Set cursor = sheet.createCursor()
            If Not cursor Is Nothing Then
                cursor.gotoStartOfUsedArea(False)
                cursor.gotoEndOfUsedArea(True)
                Dim rangeAddr As Object
                Set rangeAddr = cursor.getRangeAddress()
                Dim maxRow As Long
                maxRow = rangeAddr.EndRow
                Dim maxCol As Long
                maxCol = rangeAddr.EndColumn
                Dim row As Long
                Dim col As Long
                For row = 0 To maxRow
                    Dim rowText As String
                    rowText = ""
                    For col = 0 To maxCol
                        On Error Resume Next
                        Dim cellValue As String
                        cellValue = sheet.getCellByPosition(col, row).String
                        If Err.Number = 0 And cellValue <> "" Then
                            If rowText <> "" Then rowText = rowText & " | "
                            rowText = rowText & cellValue
                        End If
                        On Error GoTo 0
                    Next col
                    If rowText <> "" Then
                        result = result & rowText & CRLF
                    End If
                Next row
            End If
        End If
    Next i
    GetCalcDocumentText = result
End Function

'====================================================================
' Get text content from an Impress presentation
'====================================================================
Private Function GetImpressDocumentText(ByVal doc As Object) As String
    On Error Resume Next
    Dim result As String
    result = ""
    Dim slides As Object
    Set slides = doc.Slides
    If slides Is Nothing Then
        GetImpressDocumentText = ""
        Exit Function
    End If
    Dim slideCount As Long
    slideCount = slides.getCount()
    Dim i As Long
    For i = 0 To slideCount - 1
        Dim slide As Object
        Set slide = slides.getByIndex(i)
        If Not slide Is Nothing Then
            If result <> "" Then result = result & CRLF
            result = result & "=== Slide " & (i + 1) & " ===" & CRLF
            Dim shapeIdx As Long
            Dim shapeCount As Long
            shapeCount = slide.getCount()
            For shapeIdx = 0 To shapeCount - 1
                On Error Resume Next
                Dim shape As Object
                Set shape = slide.getByIndex(shapeIdx)
                If Err.Number = 0 And Not shape Is Nothing Then
                    If shape.supportsService("com.sun.star.drawing.TextShape") Then
                        Dim shapeText As String
                        shapeText = shape.String
                        If shapeText <> "" Then
                            result = result & shapeText & CRLF
                        End If
                    End If
                End If
                On Error GoTo 0
            Next shapeIdx
        End If
    Next i
    GetImpressDocumentText = result
End Function

'====================================================================
' Insert response text into the document
'====================================================================
Private Sub InsertResponseIntoDocument(ByVal responseText As String, ByVal doc As Object)
    On Error Resume Next

    If doc Is Nothing Then Exit Sub

    If doc.supportsService("com.sun.star.text.TextDocument") Then
        Dim cursor As Object
        Set cursor = doc.Text.createTextCursor()
        If Not cursor Is Nothing Then
            doc.Text.insertString(cursor, CRLF & CRLF & responseText, False)
        End If
    ElseIf doc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        Dim sheets As Object
        Set sheets = doc.Sheets
        If sheets.getCount() > 0 Then
            Dim sheet As Object
            Set sheet = sheets.getByIndex(0)
            Dim targetCell As Object
            On Error Resume Next
            Set targetCell = doc.CurrentSelection
            If targetCell Is Nothing Then
                Set targetCell = sheet.getCellByPosition(0, 0)
            End If
            If Not targetCell Is Nothing Then
                targetCell.String = responseText
            End If
            On Error GoTo 0
        End If
    ElseIf doc.supportsService("com.sun.star.presentation.PresentationDocument") Then
        Dim curSel As Object
        Set curSel = doc.CurrentSelection
        If Not curSel Is Nothing Then
            If curSel.supportsService("com.sun.star.drawing.TextShape") Then
                curSel.String = curSel.String & CRLF & responseText
            Else
                MsgBox "Select a text box on the slide first, then run the macro.", 64, APP_NAME
            End If
        Else
            MsgBox "Select a text box on the slide first, then run the macro.", 64, APP_NAME
        End If
    End If
End Sub

'====================================================================
' Settings Form (PowerShell-based Windows Forms)
'====================================================================
Public Sub Ollama_ShowConfigurationForm()
    On Error GoTo ConfigFormError

    Dim currentModel As String
    Dim currentPrompt As String
    Dim currentTimeout As String
    Dim currentApiKey As String

    currentModel = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    currentPrompt = GetRegSetting("Prompt", "")
    currentTimeout = GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS))
    currentApiKey = GetRegSetting("ApiKey", "")

    Dim psCode As String
    psCode = "$models = @(); "
    psCode = psCode & "try { "
    psCode = psCode & "$resp = Invoke-RestMethod -Uri 'http://localhost:11434/api/tags' -TimeoutSec 5 -ErrorAction Stop; "
    psCode = psCode & "$models = $resp.models | ForEach-Object { $_.name }; "
    psCode = psCode & "} catch { $models = @('" & EscapeForPS(currentModel) & "') }; "
    psCode = psCode & "if ($models.Count -eq 0) { $models = @('" & EscapeForPS(currentModel) & "') }; "
    psCode = psCode & "$dm = '" & EscapeForPS(currentModel) & "'; "
    psCode = psCode & "$dp = '" & EscapeForPS(currentPrompt) & "'; "
    psCode = psCode & "$dt = " & currentTimeout & "; "
    psCode = psCode & "$ak = '" & EscapeForPS(currentApiKey) & "'; "
    psCode = psCode & "Add-Type -AssemblyName System.Windows.Forms; "
    psCode = psCode & "Add-Type -AssemblyName System.Drawing; "
    psCode = psCode & "$form = New-Object System.Windows.Forms.Form; "
    psCode = psCode & "$form.Text = 'Ollama LibreOffice AI - Settings'; "
    psCode = psCode & "$form.Size = New-Object System.Drawing.Size(520,470); "
    psCode = psCode & "$form.StartPosition = 'CenterScreen'; "
    psCode = psCode & "$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $pid).MainModule.FileName); "
    psCode = psCode & "$form.FormBorderStyle = 'FixedDialog'; "
    psCode = psCode & "$form.MaximizeBox = $false; "
    psCode = psCode & "$lblModel = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblModel.Text = 'AI Model:'; "
    psCode = psCode & "$lblModel.Location = New-Object System.Drawing.Point(15,20); "
    psCode = psCode & "$lblModel.Size = New-Object System.Drawing.Size(100,25); "
    psCode = psCode & "$form.Controls.Add($lblModel); "
    psCode = psCode & "$cmbModel = New-Object System.Windows.Forms.ComboBox; "
    psCode = psCode & "$cmbModel.Location = New-Object System.Drawing.Point(120,18); "
    psCode = psCode & "$cmbModel.Size = New-Object System.Drawing.Size(370,25); "
    psCode = psCode & "$cmbModel.DropDownStyle = 'DropDownList'; "
    psCode = psCode & "foreach ($m in $models) { [void]$cmbModel.Items.Add($m); if ($m -eq $dm) { $cmbModel.SelectedItem = $m } }; "
    psCode = psCode & "if ($cmbModel.SelectedIndex -eq -1 -and $cmbModel.Items.Count -gt 0) { $cmbModel.SelectedIndex = 0 }; "
    psCode = psCode & "$form.Controls.Add($cmbModel); "
    psCode = psCode & "$lblPrompt = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblPrompt.Text = 'System Prompt:'; "
    psCode = psCode & "$lblPrompt.Location = New-Object System.Drawing.Point(15,55); "
    psCode = psCode & "$lblPrompt.Size = New-Object System.Drawing.Size(100,25); "
    psCode = psCode & "$form.Controls.Add($lblPrompt); "
    psCode = psCode & "$txtPrompt = New-Object System.Windows.Forms.TextBox; "
    psCode = psCode & "$txtPrompt.Location = New-Object System.Drawing.Point(15,80); "
    psCode = psCode & "$txtPrompt.Size = New-Object System.Drawing.Size(475,120); "
    psCode = psCode & "$txtPrompt.Multiline = $true; "
    psCode = psCode & "$txtPrompt.ScrollBars = 'Vertical'; "
    psCode = psCode & "$txtPrompt.Text = $dp; "
    psCode = psCode & "$form.Controls.Add($txtPrompt); "
    psCode = psCode & "$lblTimeout = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblTimeout.Text = 'Timeout (seconds):'; "
    psCode = psCode & "$lblTimeout.Location = New-Object System.Drawing.Point(15,215); "
    psCode = psCode & "$lblTimeout.Size = New-Object System.Drawing.Size(120,25); "
    psCode = psCode & "$form.Controls.Add($lblTimeout); "
    psCode = psCode & "$nudTimeout = New-Object System.Windows.Forms.NumericUpDown; "
    psCode = psCode & "$nudTimeout.Location = New-Object System.Drawing.Point(140,213); "
    psCode = psCode & "$nudTimeout.Size = New-Object System.Drawing.Size(80,25); "
    psCode = psCode & "$nudTimeout.Minimum = 10; "
    psCode = psCode & "$nudTimeout.Maximum = 600; "
    psCode = psCode & "$nudTimeout.Value = $dt; "
    psCode = psCode & "$form.Controls.Add($nudTimeout); "
    psCode = psCode & "$lblApiKey = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblApiKey.Text = 'API Key (optional):'; "
    psCode = psCode & "$lblApiKey.Location = New-Object System.Drawing.Point(15,250); "
    psCode = psCode & "$lblApiKey.Size = New-Object System.Drawing.Size(120,25); "
    psCode = psCode & "$form.Controls.Add($lblApiKey); "
    psCode = psCode & "$txtApiKey = New-Object System.Windows.Forms.TextBox; "
    psCode = psCode & "$txtApiKey.Location = New-Object System.Drawing.Point(140,250); "
    psCode = psCode & "$txtApiKey.Size = New-Object System.Drawing.Size(350,25); "
    psCode = psCode & "$txtApiKey.Text = $ak; "
    psCode = psCode & "$form.Controls.Add($txtApiKey); "
    psCode = psCode & "$lblInfo = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblInfo.Text = 'Settings are saved to Windows Registry (HKCU\Software\OllamaLibreOfficeAI).'; "
    psCode = psCode & "$lblInfo.Location = New-Object System.Drawing.Point(15,290); "
    psCode = psCode & "$lblInfo.Size = New-Object System.Drawing.Size(475,25); "
    psCode = psCode & "$lblInfo.ForeColor = 'Gray'; "
    psCode = psCode & "$form.Controls.Add($lblInfo); "
    psCode = psCode & "$btnOK = New-Object System.Windows.Forms.Button; "
    psCode = psCode & "$btnOK.Text = 'Save'; "
    psCode = psCode & "$btnOK.Location = New-Object System.Drawing.Point(160,335); "
    psCode = psCode & "$btnOK.Size = New-Object System.Drawing.Size(90,30); "
    psCode = psCode & "$btnOK.Add_Click({ $form.Tag = $cmbModel.SelectedItem + '|' + $txtPrompt.Text + '|' + $nudTimeout.Value + '|' + $txtApiKey.Text; $form.DialogResult = 'OK'; $form.Close() }); "
    psCode = psCode & "$form.Controls.Add($btnOK); "
    psCode = psCode & "$btnCancel = New-Object System.Windows.Forms.Button; "
    psCode = psCode & "$btnCancel.Text = 'Cancel'; "
    psCode = psCode & "$btnCancel.Location = New-Object System.Drawing.Point(270,335); "
    psCode = psCode & "$btnCancel.Size = New-Object System.Drawing.Size(90,30); "
    psCode = psCode & "$btnCancel.Add_Click({ $form.DialogResult = 'Cancel'; $form.Close() }); "
    psCode = psCode & "$form.Controls.Add($btnCancel); "
    psCode = psCode & "$result = $form.ShowDialog(); "
    psCode = psCode & "if ($result -eq 'OK') { "
    psCode = psCode & "$parts = $form.Tag -split '\|'; "
    psCode = psCode & "Write-Output $parts[0]; "
    psCode = psCode & "Write-Output $parts[1]; "
    psCode = psCode & "Write-Output ([int]$parts[2]); "
    psCode = psCode & "Write-Output $parts[3]; "
    psCode = psCode & "} else { Write-Output 'CANCELLED' }"

    Dim result As String
    result = RunPowerShell(psCode)

    If result <> "CANCELLED" And result <> "" Then
        Dim lines() As String
        lines = Split(result, CRLF)

        If UBound(lines) >= 0 Then
            Dim modelVal As String
            modelVal = Trim(lines(0))
            If modelVal <> "" Then SaveRegSetting "Model", modelVal
        End If

        If UBound(lines) >= 1 Then
            Dim promptVal As String
            promptVal = Trim(lines(1))
            If promptVal <> "" Then SaveRegSetting "Prompt", promptVal
        End If

        If UBound(lines) >= 2 Then
            Dim timeoutVal As String
            timeoutVal = Trim(lines(2))
            If timeoutVal <> "" Then SaveRegSetting "Timeout", timeoutVal
        End If

        If UBound(lines) >= 3 Then
            Dim apiKeyVal As String
            apiKeyVal = Trim(lines(3))
            SaveRegSetting "ApiKey", apiKeyVal
        End If

        MsgBox "Settings saved successfully.", 64, APP_NAME
    End If

    Exit Sub

ConfigFormError:
    MsgBox "Could not open configuration form: " & Err.Description, 48, APP_NAME
End Sub

'====================================================================
' About Form
'====================================================================
Public Sub Ollama_ShowAboutForm()
    Dim msg As String
    msg = APP_NAME & " v" & APP_VERSION & CRLF & CRLF
    msg = msg & "Local AI Document Assistant for LibreOffice" & CRLF
    msg = msg & "Powered by Ollama" & CRLF & CRLF
    msg = msg & "All processing is fully local." & CRLF
    msg = msg & "No data leaves your machine."
    MsgBox msg, 64, "About " & APP_NAME
End Sub

'====================================================================
' API Communication - Send request to Ollama
'====================================================================
Private Function Ollama_ProcessRequest(ByVal systemPrompt As String, ByVal userContent As String, ByVal modelName As String, ByVal timeoutSecs As Long) As String
    On Error GoTo RequestError

    Dim payload As String
    payload = "{"
    payload = payload & """model"":""" & EscapeJson(modelName) & ""","
    payload = payload & """messages"":["
    payload = payload & "{""role"":""system"",""content"":""" & EscapeJson(systemPrompt) & """},"
    payload = payload & "{""role"":""user"",""content"":""" & EscapeJson(userContent) & """}"
    payload = payload & "],"
    payload = payload & """stream"":false"
    payload = payload & "}"

    Dim url As String
    url = OLLAMA_BASE_URL & "/v1/chat/completions"

    Dim response As String
    response = Ollama_SendHttpRequest(url, "POST", payload, timeoutSecs)

    If response = "" Then
        Ollama_ProcessRequest = ""
        Exit Function
    End If

    Dim content As String
    content = ExtractResponseContent(response)
    If content = "" Then
        Dim errMsg As String
        errMsg = ExtractApiError(response)
        If errMsg <> "" Then
            Ollama_ProcessRequest = "[Ollama Error: " & errMsg & "]"
        Else
            Ollama_ProcessRequest = ""
        End If
    Else
        Ollama_ProcessRequest = content
    End If
    Exit Function

RequestError:
    Ollama_ProcessRequest = ""
End Function

'====================================================================
' HTTP Request using WinHttp
'====================================================================
Private Function Ollama_SendHttpRequest(ByVal url As String, ByVal method As String, ByVal payload As String, ByVal timeoutSecs As Long) As String
    On Error GoTo HttpError

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim timeoutMs As Long
    timeoutMs = timeoutSecs * 1000

    http.Open method, url, False
    http.SetTimeouts 10000, 15000, timeoutMs, timeoutMs
    http.SetOption 0, "Ollama-LibreOffice-AI/1.0"
    http.SetRequestHeader "Content-Type", "application/json"

    Dim apiKey As String
    apiKey = GetRegSetting("ApiKey", "")
    If apiKey <> "" Then
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
    End If

    If method = "POST" Then
        http.Send payload
    Else
        http.Send
    End If

    Dim status As Integer
    status = http.Status
    If status >= 200 And status < 300 Then
        Ollama_SendHttpRequest = http.ResponseText
    Else
        Ollama_SendHttpRequest = http.ResponseText
    End If

    Exit Function

HttpError:
    Ollama_SendHttpRequest = ""
End Function

'====================================================================
' Extract response content from Ollama API JSON
'====================================================================
Private Function ExtractResponseContent(ByVal json As String) As String
    On Error Resume Next

    Dim contentMarker As String
    contentMarker = """content"":"""

    Dim msgStart As Long
    msgStart = InStr(json, """message"":{")
    If msgStart = 0 Then
        msgStart = InStr(json, contentMarker)
        If msgStart = 0 Then
            ExtractResponseContent = ""
            Exit Function
        End If
    Else
        msgStart = InStr(msgStart, json, contentMarker)
        If msgStart = 0 Then
            ExtractResponseContent = ""
            Exit Function
        End If
    End If

    Dim contentStart As Long
    contentStart = msgStart + Len(contentMarker)

    Dim contentEnd As Long
    contentEnd = contentStart

    Do While contentEnd <= Len(json)
        Dim ch As String
        ch = Mid(json, contentEnd, 1)
        If ch = "\" Then
            contentEnd = contentEnd + 2
        ElseIf ch = """" Then
            Exit Do
        Else
            contentEnd = contentEnd + 1
        End If
    Loop

    If contentEnd > contentStart Then
        Dim raw As String
        raw = Mid(json, contentStart, contentEnd - contentStart)
        Dim result As String
        result = raw
        result = Replace(result, "\\", Chr(1))
        result = Replace(result, "\n", CRLF)
        result = Replace(result, "\t", Chr(9))
        result = Replace(result, "\r", Chr(13))
        result = Replace(result, "\""", """")
        result = Replace(result, Chr(1), "\")
        ExtractResponseContent = result
    Else
        ExtractResponseContent = ""
    End If
End Function

'====================================================================
' Extract error message from Ollama API error response
'====================================================================
Private Function ExtractApiError(ByVal json As String) As String
    On Error Resume Next
    Dim errStart As Long
    errStart = InStr(json, """error"":""")
    If errStart = 0 Then
        errStart = InStr(json, """message"":""")
    End If
    If errStart = 0 Then
        ExtractApiError = ""
        Exit Function
    End If
    errStart = errStart + 10
    Dim errEnd As Long
    errEnd = InStr(errStart, json, """")
    If errEnd = 0 Or errEnd <= errStart Then
        ExtractApiError = ""
        Exit Function
    End If
    ExtractApiError = Mid(json, errStart, errEnd - errStart)
End Function

'====================================================================
' Escape strings for JSON
'====================================================================
Private Function EscapeJson(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, CRLF, "\\n")
    result = Replace(result, Chr(13), "\\r")
    result = Replace(result, Chr(10), "\\n")
    result = Replace(result, Chr(9), "\\t")
    EscapeJson = result
End Function

'====================================================================
' Escape strings for PowerShell single-quoted strings
'====================================================================
Private Function EscapeForPS(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "'", "''")
    EscapeForPS = result
End Function

'====================================================================
' Registry Helpers (via WScript.Shell)
'====================================================================
Private Function GetRegSetting(ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim regPath As String
    regPath = REG_PATH_ROOT & "\" & key
    Dim value As Variant
    value = ws.RegRead(regPath)
    If Err.Number = 0 Then
        GetRegSetting = CStr(value)
    Else
        GetRegSetting = defaultValue
    End If
End Function

Private Sub SaveRegSetting(ByVal key As String, ByVal value As String)
    On Error Resume Next
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim regPath As String
    regPath = REG_PATH_ROOT & "\" & key
    ws.RegWrite regPath, value, "REG_SZ"
End Sub

'====================================================================
' Run PowerShell script and capture output
'====================================================================
Private Function RunPowerShell(ByVal script As String) As String
    On Error GoTo PSError

    Dim tempFile As String
    Dim tempDir As String
    Dim fso As Object
    Dim f As Object
    Dim ws As Object
    Dim exec As Object
    Dim shellCmd As String
    Dim shellOut As String
    Dim startTime As Double

    tempDir = Environ("TEMP")
    tempFile = tempDir & "\ollama_config.ps1"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(tempFile, True)
    f.Write script
    f.Close

    Set ws = CreateObject("WScript.Shell")

    shellCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & tempFile & """"
    Set exec = ws.Exec(shellCmd)

    startTime = Timer
    Do While exec.Status = 0
        If Timer > startTime + 60 Then
            exec.Terminate
            Exit Do
        End If
    Loop

    If exec.Status <> 0 Then
        shellOut = exec.StdOut.ReadAll()
    End If

    On Error Resume Next
    fso.DeleteFile tempFile
    On Error GoTo 0

    RunPowerShell = shellOut
    Exit Function

PSError:
    RunPowerShell = ""
End Function

'====================================================================
' Screenshot Analysis - Pick an image and analyze it with a vision model
'====================================================================
Public Sub Ollama_AnalyzeScreenshot()
    On Error GoTo ScreenshotError

    If Not Ollama_IsOllamaRunning() Then
        MsgBox "Ollama is not running." & CRLF & CRLF & "Please start Ollama and try again.", 48, APP_NAME
        Exit Sub
    End If

    Dim doc As Object
    doc = ThisComponent
    If doc Is Nothing Then
        MsgBox "No document is open.", 48, APP_NAME
        Exit Sub
    End If

    Dim imageDataUri As String
    imageDataUri = PickImageAndGetBase64()

    If imageDataUri = "CANCELLED" Or imageDataUri = "" Then
        Exit Sub
    End If

    Dim modelName As String
    Dim systemPrompt As String
    Dim timeoutSecs As Long

    modelName = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    systemPrompt = GetRegSetting("Prompt", "You are a helpful document assistant. Analyze the document content and respond professionally. Keep your response concise and well-structured.")
    timeoutSecs = CLng(GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS)))

    Dim prompt As String
    prompt = "Analyze this error screenshot or image. Explain what the error means, what caused it, and how to fix it."

    Dim response As String
    response = Ollama_ProcessVisionRequest(systemPrompt, prompt, imageDataUri, modelName, timeoutSecs)

    If response <> "" Then
        If Left(response, 1) = "[" And InStr(response, "Ollama Error:") > 0 Then
            MsgBox response, 48, APP_NAME
        Else
            InsertResponseIntoDocument response, doc
        End If
    Else
        MsgBox "No response received from Ollama." & CRLF & CRLF & "Check that:" & CRLF & "  - Ollama is running" & CRLF & "  - The selected model supports vision (e.g., llama3.2-vision, llava)", 48, APP_NAME
    End If

    Exit Sub

ScreenshotError:
    MsgBox "Error analyzing screenshot: " & Err.Description, 16, APP_NAME
End Sub

'====================================================================
' File picker + base64 via PowerShell
'====================================================================
Private Function PickImageAndGetBase64() As String
    On Error GoTo PickError

    Dim psCode As String
    psCode = "Add-Type -AssemblyName System.Windows.Forms; "
    psCode = psCode & "$ofd = New-Object System.Windows.Forms.OpenFileDialog; "
    psCode = psCode & "$ofd.Title = 'Select Screenshot or Error Image'; "
    psCode = psCode & "$ofd.Filter = 'Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif'; "
    psCode = psCode & "if ($ofd.ShowDialog() -eq 'OK') { "
    psCode = psCode & "$bytes = [System.IO.File]::ReadAllBytes($ofd.FileName); "
    psCode = psCode & "$b64 = [System.Convert]::ToBase64String($bytes); "
    psCode = psCode & "$ext = [System.IO.Path]::GetExtension($ofd.FileName).TrimStart('.').ToLower(); "
    psCode = psCode & "if ($ext -eq 'jpg') { $ext = 'jpeg' }; "
    psCode = psCode & "Write-Output ('data:image/' + $ext + ';base64,' + $b64) "
    psCode = psCode & "} else { Write-Output 'CANCELLED' }"

    PickImageAndGetBase64 = RunPowerShell(psCode)
    Exit Function

PickError:
    PickImageAndGetBase64 = ""
End Function

'====================================================================
' Vision API Request - Send image with text to vision model
'====================================================================
Private Function Ollama_ProcessVisionRequest(ByVal systemPrompt As String, ByVal userPrompt As String, ByVal imageDataUri As String, ByVal modelName As String, ByVal timeoutSecs As Long) As String
    On Error GoTo VisionError

    Dim payload As String
    payload = "{"
    payload = payload & """model"":""" & EscapeJson(modelName) & ""","
    payload = payload & """messages"":["
    payload = payload & "{""role"":""system"",""content"":""" & EscapeJson(systemPrompt) & """},"
    payload = payload & "{""role"":""user"",""content"":["
    payload = payload & "{""type"":""text"",""text"":""" & EscapeJson(userPrompt) & """},"
    payload = payload & "{""type"":""image_url"",""image_url"":{""url"":""" & EscapeJson(imageDataUri) & """}}"
    payload = payload & "]}"
    payload = payload & "],"
    payload = payload & """stream"":false"
    payload = payload & "}"

    Dim url As String
    url = OLLAMA_BASE_URL & "/v1/chat/completions"

    Dim response As String
    response = Ollama_SendHttpRequest(url, "POST", payload, timeoutSecs)

    If response = "" Then
        Ollama_ProcessVisionRequest = ""
        Exit Function
    End If

    Dim content As String
    content = ExtractResponseContent(response)
    If content = "" Then
        Dim errMsg As String
        errMsg = ExtractApiError(response)
        If errMsg <> "" Then
            Ollama_ProcessVisionRequest = "[Ollama Error: " & errMsg & "]"
        Else
            Ollama_ProcessVisionRequest = ""
        End If
    Else
        Ollama_ProcessVisionRequest = content
    End If
    Exit Function

VisionError:
    Ollama_ProcessVisionRequest = ""
End Function

'====================================================================
' End of Module
'====================================================================
