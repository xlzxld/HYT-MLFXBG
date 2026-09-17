'%RunPerInstance
'@ 说明: 生成冷却水路测试数据 —— 第一版独立存档 (用户实测可用, 2026-09-16)。
'@   与 MakeCoolingCircuit.vbs (带清理/重试/诊断) 并列保留, 两者互不影响。
'@   实测: 有网格的研究上 Generate 返回 True (直径 8, 距离 25, 3 通道)。
'@ 用法: 先打开研究 -> Tools -> Macro -> Play Macro -> 选本文件。
'@ 说明: 不修改模型 (纯生成), 不自动保存; 需要时在软件里手动保存。
'@ 提示: 研究若没有网格 (MeshStatus=New), 生成会半途失败, 请先生成网格。
Option Explicit

Dim FSO, WshShell, BaseDir, TempDir, LogPath
Dim SynergyGetter, Synergy, CircuitGen, saEnv, result

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)
TempDir = BaseDir & "\temp"
If Not FSO.FolderExists(TempDir) Then FSO.CreateFolder TempDir
LogPath = TempDir & "\cooling.log"

' ---- CONFIG: 冷却回路参数 (单位: mm) ----
Dim CfgDiameter, CfgDistance, CfgSpacing, CfgOverhang, CfgChannels, CfgUseHoses, CfgXAlign
CfgDiameter = 8.0
CfgDistance = 25.0
CfgSpacing = 30.0
CfgOverhang = 20.0
CfgChannels = 3
CfgUseHoses = False
CfgXAlign = False

Call LogMsg("=== 冷却水路生成开始 (v1 独立版) ===")

' 连接 Synergy (三级兜底, 与 AutoReport 同源)
Set Synergy = Nothing
Set SynergyGetter = Nothing
saEnv = WshShell.ExpandEnvironmentStrings("%SAInstance%")
On Error Resume Next
If saEnv <> "" And saEnv <> "%SAInstance%" Then
    Set SynergyGetter = GetObject(saEnv)
    If Not SynergyGetter Is Nothing Then Set Synergy = SynergyGetter.GetSASynergy
End If
On Error GoTo 0
If Synergy Is Nothing Then
    On Error Resume Next
    Set Synergy = GetObject(, "Synergy.Synergy")
    If Synergy Is Nothing Then Set Synergy = CreateObject("Synergy.Synergy")
    On Error GoTo 0
End If
If Synergy Is Nothing Then
    Call LogMsg("ERROR: 未连接上 Synergy")
    MsgBox "没有连上 Moldflow。请在 Moldflow 里打开研究后, 用 Tools -> Macro 运行本宏。", 16, "生成失败"
    WScript.Quit 1
End If

Set CircuitGen = Nothing
On Error Resume Next
Set CircuitGen = Synergy.CircuitGenerator()
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If CircuitGen Is Nothing Then
    Call LogMsg("ERROR: CircuitGenerator 不可用")
    MsgBox "当前版本没有 CircuitGenerator 接口, 无法自动生成冷却水路。" & vbCrLf & _
           "请在界面里用 冷却回路向导 手工创建。", 16, "生成失败"
    WScript.Quit 1
End If

On Error Resume Next
CircuitGen.DeleteOld = True
CircuitGen.Diameter = CfgDiameter
CircuitGen.Distance = CfgDistance
CircuitGen.Spacing = CfgSpacing
CircuitGen.Overhang = CfgOverhang
CircuitGen.NumChannels = CfgChannels
CircuitGen.UseHoses = CfgUseHoses
CircuitGen.XAlign = CfgXAlign
If Err.Number <> 0 Then
    Call LogMsg("WARN: 设置回路参数异常: " & Err.Description)
    Err.Clear
End If

result = CircuitGen.Generate()
If Err.Number <> 0 Then
    Call LogMsg("ERROR: Generate 异常: " & Err.Description)
    Err.Clear
    result = False
End If
On Error GoTo 0
Call LogMsg("Generate 返回: " & CStr(result) & _
            " (直径 " & CfgDiameter & ", 间距 " & CfgDistance & ", 通道数 " & CfgChannels & ")")

If result Then
    MsgBox "冷却水路已生成。" & vbCrLf & vbCrLf & _
           "参数: 直径 " & CfgDiameter & "mm, 间距 " & CfgDistance & "mm, " & CfgChannels & " 条通道" & vbCrLf & _
           "本宏不会自动保存; 需要时请在 Moldflow 里自行保存。", 64, "生成完成"
Else
    MsgBox "冷却水路生成失败。" & vbCrLf & "请看 temp\cooling.log 并发给维护者。", 16, "生成失败"
End If

Call LogMsg("=== 冷却水路生成结束 ===")
WScript.Quit 0

Sub LogMsg(msg)
    On Error Resume Next
    Dim f
    Set f = FSO.OpenTextFile(LogPath, 8, True)
    f.WriteLine Now & " - " & msg
    f.Close
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
