' ============================================================
' Outlook VBA 모듈: Gemini 자동 답변 초안 생성 (UI 버전)
' ThisOutlookSession 또는 별도 모듈에 붙여넣기
' ============================================================

Sub GeminiAutoReply()
    Dim pythonExe As String
    Dim scriptPath As String

    pythonExe = "python"
    scriptPath = "C:\Users\82108\Downloads\outlook_gemini\ui.py"

    ' UI 창 표시 (1 = 기본 창)
    Shell pythonExe & " """ & scriptPath & """", 1
End Sub
