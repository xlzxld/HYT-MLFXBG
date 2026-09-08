'@ 列出当前方案的全部结果图名 → temp\available_plots.json (配置界面"从 Moldflow 刷新结果"用)
'@ 纯只读: 不显示/不创建/不修改任何结果。连不上 Moldflow 时写空清单+原因, 退出码恒为 0。
'@
Option Explicit
Dim FSO, WshShell, BaseDir, TempDir, OutPath
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)
TempDir = BaseDir & "\temp"
If Not FSO.FolderExists(TempDir) Then FSO.CreateFolder TempDir
OutPath = TempDir & "\available_plots.json"

Dim Synergy, saEnv, StudyDoc, PlotManager, P
Dim Names(), N
N = 0
ReDim Names(255)

' 连接 (与 AutoReport.vbs 同序: SAInstance → GetObject → CreateObject)
Set Synergy = Nothing
saEnv = WshShell.ExpandEnvironmentStrings("%SAInstance%")
On Error Resume Next
If saEnv <> "" And saEnv <> "%SAInstance%" Then
    Dim Getter
    Set Getter = GetObject(saEnv)
    If Not Getter Is Nothing Then Set Synergy = Getter.GetSASynergy
End If
If Synergy Is Nothing Then Set Synergy = GetObject(, "Synergy.Synergy")
If Synergy Is Nothing Then Set Synergy = CreateObject("Synergy.Synergy")
On Error GoTo 0
If Synergy Is Nothing Then
    WriteJson "", N, Array(), "未连接到 Moldflow (请先打开 Moldflow)"
    WScript.Quit 0
End If

Set StudyDoc = Nothing
On Error Resume Next
Set StudyDoc = Synergy.StudyDoc()
On Error GoTo 0
If StudyDoc Is Nothing Then
    WriteJson "", N, Array(), "方案未打开"
    WScript.Quit 0
End If

' 枚举结果名 (官方 GetFirstPlot/GetNextPlot, 只取名字)
Set PlotManager = Nothing
On Error Resume Next
Set PlotManager = Synergy.PlotManager()
On Error GoTo 0
If PlotManager Is Nothing Then
    WriteJson StudyDoc.StudyName, 0, Array(), "PlotManager 不可用"
    WScript.Quit 0
End If

Dim nm
On Error Resume Next
Set P = PlotManager.GetFirstPlot()
Do While Not P Is Nothing
    nm = ""
    nm = P.GetName()
    If Err.Number <> 0 Then
        Err.Clear
        nm = ""
    End If
    If Len(nm) > 0 Then
        If N > UBound(Names) Then ReDim Preserve Names(UBound(Names) * 2 + 1)
        Names(N) = nm
        N = N + 1
    End If
    Set P = PlotManager.GetNextPlot(P)
Loop
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

If N = 0 Then
    WriteJson StudyDoc.StudyName, 0, Array(), "方案中没有结果 (请先完成分析)"
Else
    WriteJson StudyDoc.StudyName, N, Names, ""
End If
WScript.Quit 0

Sub WriteJson(study, count, arr, errMsg)
    Dim i, s
    s = "{" & vbCrLf & _
        "  ""study"": """ & EscapeJson(CStr(study)) & """," & vbCrLf & _
        "  ""error"": """ & EscapeJson(CStr(errMsg)) & """," & vbCrLf & _
        "  ""plots"": ["
    For i = 0 To count - 1
        If i > 0 Then s = s & ","
        s = s & vbCrLf & "    """ & EscapeJson(CStr(arr(i))) & """"
    Next
    s = s & vbCrLf & "  ]" & vbCrLf & "}"
    WriteUtf8TextFile OutPath, s
End Sub

Function EscapeJson(strText)
    Dim s
    s = Replace(strText, "\", "\\")
    s = Replace(s, """", "\""")
    EscapeJson = s
End Function

Sub WriteUtf8TextFile(Path, Text)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText Text
    stream.SaveToFile Path, 2
    stream.Close
End Sub
