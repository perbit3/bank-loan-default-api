Attribute VB_Name = "LoanDefaultApi"
Option Explicit

' Calls the FastAPI /predict endpoint and writes the probability of default back to the workbook.
' Works on Excel for Mac (curl through libc popen) and Excel for Windows (WinHttp).

#If Mac Then
    Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
    Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
    Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
    Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr
#End If

Private Const TIMEOUT_SECONDS As Long = 120


' ---------------------------------------------------------------------------------------------
' Submit button on the "Predict" sheet
' ---------------------------------------------------------------------------------------------
Public Sub SubmitPrediction()
    Dim ws As Worksheet, values(1 To 6) As Double, payload As String, response As String
    Dim prob As Double, errMsg As String

    Set ws = ThisWorkbook.Worksheets("Predict")
    ws.Range("ResultProbability").ClearContents
    ws.Range("ResultClass").ClearContents
    ws.Range("ResultLabel").ClearContents
    ws.Range("ResultTime").ClearContents
    ws.Range("ResultStatus").Value = "Sending request..."

    If Not ReadInputs(ws, values, errMsg) Then
        ws.Range("ResultStatus").Value = errMsg
        MsgBox errMsg, vbExclamation, "Check the customer details"
        Exit Sub
    End If

    payload = BuildPayload(values)
    Application.StatusBar = "Calling the prediction API (the first call can take up to a minute while the free server wakes up)..."
    Application.Cursor = xlWait

    If Not PostJson(ApiUrl(), payload, response, errMsg) Then
        Application.StatusBar = False
        Application.Cursor = xlDefault
        ws.Range("ResultStatus").Value = "Error: " & errMsg
        MsgBox "The API call failed." & vbNewLine & vbNewLine & errMsg, vbCritical, "Prediction failed"
        Exit Sub
    End If

    prob = JsonNumber(response, "probability_of_default")
    ws.Range("ResultProbability").Value = prob
    ws.Range("ResultClass").Value = CLng(JsonNumber(response, "predicted_defaulter"))
    ws.Range("ResultLabel").Value = JsonString(response, "risk_label")
    ws.Range("ResultTime").Value = Now
    ws.Range("ResultStatus").Value = "OK - response received from " & ApiUrl()

    Application.StatusBar = False
    Application.Cursor = xlDefault
End Sub


' ---------------------------------------------------------------------------------------------
' "Score test dataset" button on the "Test Predictions" sheet: calls /predict for every row
' ---------------------------------------------------------------------------------------------
Public Sub ScoreTestDataset()
    Dim ws As Worksheet, r As Long, lastRow As Long, c As Long
    Dim values(1 To 6) As Double, response As String, errMsg As String, scored As Long

    Set ws = ThisWorkbook.Worksheets("Test Predictions")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range(ws.Cells(5, 9), ws.Cells(lastRow, 10)).ClearContents
    Application.Cursor = xlWait

    For r = 5 To lastRow
        Application.StatusBar = "Scoring customer " & ws.Cells(r, 1).Value & " of " & (lastRow - 4) & "..."
        For c = 1 To 6
            values(c) = CDbl(ws.Cells(r, c + 1).Value)   ' columns B:G = AGE..OTHDEBT
        Next c
        If Not PostJson(ApiUrl(), BuildPayload(values), response, errMsg) Then
            Application.StatusBar = False
            Application.Cursor = xlDefault
            MsgBox "The API call failed for SN " & ws.Cells(r, 1).Value & "." & vbNewLine & vbNewLine & errMsg, vbCritical, "Scoring failed"
            Exit Sub
        End If
        ws.Cells(r, 9).Value = JsonNumber(response, "probability_of_default")
        ws.Cells(r, 10).Value = CLng(JsonNumber(response, "predicted_defaulter"))
        scored = scored + 1
        DoEvents
    Next r

    ws.Range("TestScoredAt").Value = Now
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox scored & " customers scored by the API.", vbInformation, "Test dataset scored"
End Sub


' ---------------------------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------------------------
Private Function ApiUrl() As String
    ApiUrl = Trim$(ThisWorkbook.Worksheets("Predict").Range("ApiUrl").Value)
End Function

Private Function ReadInputs(ws As Worksheet, values() As Double, errMsg As String) As Boolean
    Dim names As Variant, i As Long, v As Variant
    names = Array("InAge", "InEmploy", "InAddress", "InDebtInc", "InCredDebt", "InOthDebt")
    For i = 0 To 5
        v = ws.Range(names(i)).Value
        If IsEmpty(v) Or Not IsNumeric(v) Then
            errMsg = ws.Range(names(i)).Offset(0, -1).Value & " must be a number."
            Exit Function
        End If
        If CDbl(v) < 0 Then
            errMsg = ws.Range(names(i)).Offset(0, -1).Value & " cannot be negative."
            Exit Function
        End If
        values(i + 1) = CDbl(v)
    Next i
    If values(1) <> 1 And values(1) <> 2 And values(1) <> 3 Then
        errMsg = "AGE must be 1, 2 or 3."
        Exit Function
    End If
    ReadInputs = True
End Function

Private Function BuildPayload(values() As Double) As String
    BuildPayload = "{""AGE"":" & Num(values(1)) & ",""EMPLOY"":" & Num(values(2)) & _
                   ",""ADDRESS"":" & Num(values(3)) & ",""DEBTINC"":" & Num(values(4)) & _
                   ",""CREDDEBT"":" & Num(values(5)) & ",""OTHDEBT"":" & Num(values(6)) & "}"
End Function

' Number as text with a full stop as the decimal separator, whatever the regional settings
Private Function Num(ByVal x As Double) As String
    Num = Trim$(Str$(x))
    If Left$(Num, 1) = "." Then Num = "0" & Num
    If Left$(Num, 2) = "-." Then Num = "-0" & Mid$(Num, 2)
End Function

' POST a JSON body and return the response text. Returns False with errMsg on failure.
Private Function PostJson(ByVal url As String, ByVal body As String, response As String, errMsg As String) As Boolean
#If Mac Then
    Dim cmd As String, output As String, statusCode As String, exitCode As Long, p As Long
    cmd = "curl -s -S -m " & TIMEOUT_SECONDS & " -X POST -H 'Content-Type: application/json' " & _
          "-d '" & body & "' -w '\n%{http_code}' '" & url & "' 2>&1"
    output = ExecShell(cmd, exitCode)
    If exitCode <> 0 Then
        errMsg = "curl error: " & Trim$(output)
        Exit Function
    End If
    output = Replace(output, vbCr, "")
    If Right$(output, 1) = vbLf Then output = Left$(output, Len(output) - 1)
    p = InStrRev(output, vbLf)
    statusCode = Mid$(output, p + 1)
    response = Left$(output, IIf(p > 0, p - 1, 0))
    If statusCode <> "200" Then
        errMsg = "HTTP " & statusCode & ": " & response
        Exit Function
    End If
    PostJson = IsPrediction(response, errMsg)
#Else
    Dim http As Object
    On Error GoTo Failed
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, TIMEOUT_SECONDS * 1000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body
    response = http.ResponseText
    If http.Status <> 200 Then
        errMsg = "HTTP " & http.Status & ": " & response
        Exit Function
    End If
    PostJson = IsPrediction(response, errMsg)
    Exit Function
Failed:
    errMsg = Err.Description
#End If
End Function

#If Mac Then
Private Function ExecShell(ByVal cmd As String, exitCode As Long) As String
    Dim f As LongPtr, chunk As String, n As Long, result As String
    f = popen(cmd, "r")
    If f = 0 Then
        exitCode = -1
        Exit Function
    End If
    Do While feof(f) = 0
        chunk = VBA.Space$(256)
        n = CLng(fread(chunk, 1, Len(chunk) - 1, f))
        If n > 0 Then result = result & VBA.Left$(chunk, n)
    Loop
    exitCode = pclose(f)
    ExecShell = result
End Function
#End If

Private Function IsPrediction(ByVal response As String, errMsg As String) As Boolean
    If InStr(response, """probability_of_default""") = 0 Then
        errMsg = "Unexpected response from the API (the server may still be starting up, please try again in a moment)."
    Else
        IsPrediction = True
    End If
End Function

' Minimal JSON readers for the flat response returned by /predict
Private Function JsonNumber(ByVal json As String, ByVal key As String) As Double
    Dim p As Long, q As Long, s As String
    p = InStr(json, """" & key & """")
    If p = 0 Then Err.Raise vbObjectError + 1, , "Key '" & key & "' not found in response: " & json
    p = InStr(p, json, ":") + 1
    q = p
    Do While q <= Len(json) And InStr(",}", Mid$(json, q, 1)) = 0
        q = q + 1
    Loop
    s = Trim$(Mid$(json, p, q - p))
    JsonNumber = Val(s)
End Function

Private Function JsonString(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p, json, ":")
    p = InStr(p, json, """") + 1
    q = InStr(p, json, """")
    JsonString = Mid$(json, p, q - p)
End Function
