'%RunPerInstance
'@ 描述: 生成冷却水路测试数据 (v2: 带诊断/重试/计数)
'@   用官方 CircuitGenerator 在当前研究里生成一组简单冷却回路。
'@
'@ 用法: 先复制一份研究 -> 打开副本 -> Tools -> Macro -> Play Macro -> 本文件
'@ 说明: 会修改模型 (新增几何), 不自动保存; 参数见下方 CONFIG 区; 日志 temp/cooling.log
'@ 提示: 生成后用报告配置里的"冷却水(水路)"勾选框即可测试显示效果
Option Explicit

Dim FSO, WshShell, BaseDir, TempDir, LogPath
Dim SynergyGetter, Synergy, CircuitGen, saEnv, result, attempt
Dim CfgDiameter, CfgDistance, CfgSpacing, CfgOverhang, CfgChannels, CfgUseHoses, CfgXAlign

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)
TempDir = BaseDir & "\temp"
If Not FSO.FolderExists(TempDir) Then FSO.CreateFolder TempDir
LogPath = TempDir & "\cooling.log"

' ---- CONFIG: 冷却回路参数 (单位: mm) ----
CfgDiameter = 8.0
CfgDistance = 25.0
CfgSpacing = 30.0
CfgOverhang = 20.0
CfgChannels = 3
CfgUseHoses = False
CfgXAlign = False

Call LogMsg("=== 冷却水路生成开始 (v2) ===")

' 连接 Synergy (三级兜底)
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

Dim StudyDoc
Set StudyDoc = Synergy.StudyDoc()
Call LogMsg("研究: " & CStr(StudyDoc.StudyName))
Call LogMsg("网格状态: MeshStatus=" & CStr(StudyDoc.MeshStatus()) & ", MeshType=" & CStr(StudyDoc.MeshType))

Call LogCoolingCounts("生成前")
' 生成器是"追加"式: 残留的旧回路会让新回路生成失败 (实测 2026-09-16),
' 因此生成前先删除已有的冷却实体 (管道 40480 + 冷却液入口 40020)。
Call DeleteExistingCooling()

Set CircuitGen = Nothing
On Error Resume Next
Set CircuitGen = Synergy.CircuitGenerator()
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If CircuitGen Is Nothing Then
    Call LogMsg("ERROR: CircuitGenerator 不可用")
    MsgBox "当前版本没有 CircuitGenerator 接口, 请在界面里用冷却回路向导手工创建。", 16, "生成失败"
    WScript.Quit 1
End If

result = False
For attempt = 1 To 2
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
        Call LogMsg("WARN: 设置回路参数异常 #" & Err.Number & " [" & Err.Source & "] " & Err.Description)
        Err.Clear
    End If

    result = False
    result = CBool(CircuitGen.Generate())
    If Err.Number <> 0 Then
        Call LogMsg("ERROR: Generate 第" & attempt & "次异常 #" & Err.Number & " [" & Err.Source & "] " & Err.Description)
        Err.Clear
        result = False
    Else
        Call LogMsg("Generate 第" & attempt & "次返回: " & CStr(result))
    End If
    On Error GoTo 0
    Call LogCoolingCounts("第" & attempt & "次生成后")
    If result Then Exit For
    If attempt = 1 Then
        Call LogMsg("提示: 第一次失败, 2 秒后重试一次...")
        Call SleepSec(2)
    End If
Next

If result Then
    MsgBox "冷却水路已生成。" & vbCrLf & vbCrLf & _
           "参数: 直径 " & CfgDiameter & "mm, 间距 " & CfgDistance & "mm, " & CfgChannels & " 条通道" & vbCrLf & _
           "本宏不会自动保存; 需要时请在 Moldflow 里自行保存。", 64, "生成完成"
Else
    MsgBox "冷却水路生成失败。" & vbCrLf & _
           "日志: temp\cooling.log (含错误号与前后计数)" & vbCrLf & vbCrLf & _
           "可先在界面里手工用冷却回路向导建一条, 或换一份干净副本再试。", 16, "生成失败"
End If

Call LogMsg("=== 冷却水路生成结束 ===")
WScript.Quit 0

Sub LogCoolingCounts(tag)
    ' 统计冷却相关实体 (属性识别: 管道 40480 / 冷却液入口 40020)
    Dim pe, pm, ts, tsets, prop, name, lst, pred, p2, pOr, size, total, det
    total = 0
    det = ""
    Set pm = Synergy.PredicateManager()
    Set pe = Synergy.PropertyEditor()
    tsets = Array(40480, 40020)
    Dim i, j
    For j = 0 To UBound(tsets)
        ts = tsets(j)
        Set prop = Nothing
        On Error Resume Next
        Set prop = pe.GetFirstProperty(ts)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not prop Is Nothing Then
            name = ""
            On Error Resume Next
            name = CStr(prop.Name)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            Set lst = StudyDoc.CreateEntityList()
            Set p2 = pm.CreatePropTypePredicate(ts)
            lst.SelectFromPredicate p2
            size = 0
            On Error Resume Next
            size = CLng(lst.Size)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            total = total + size
            det = det & " | " & ts & "[" & name & "]=" & size
        End If
    Next
    Call LogMsg("冷却实体计数(" & tag & "): 合计 " & total & det)
End Sub

Sub DeleteExistingCooling()
    ' 删除残留冷却实体; 仅限冷却相关属性, 不动其它几何
    Dim pm, meshEd, tsets, j, ts, prop, p2, pred, pOr, lst, size, ok_del
    Set pm = Synergy.PredicateManager()
    Set meshEd = Synergy.MeshEditor()
    tsets = Array(40480, 40020)
    Set pred = Nothing
    For j = 0 To UBound(tsets)
        Set p2 = Nothing
        On Error Resume Next
        Set p2 = pm.CreatePropTypePredicate(tsets(j))
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not p2 Is Nothing Then
            If pred Is Nothing Then
                Set pred = p2
            Else
                Set pOr = Nothing
                On Error Resume Next
                Set pOr = pm.CreateBoolOrPredicate(pred, p2)
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                If Not pOr Is Nothing Then Set pred = pOr
            End If
        End If
    Next
    If pred Is Nothing Then Exit Sub
    Set lst = StudyDoc.CreateEntityList()
    lst.SelectFromPredicate pred
    size = 0
    On Error Resume Next
    size = CLng(lst.Size)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If size <= 0 Then
        Call LogMsg("无残留冷却实体, 跳过清理")
        Exit Sub
    End If
    ok_del = False
    On Error Resume Next
    meshEd.Delete lst
    If Err.Number <> 0 Then
        Call LogMsg("WARN: 删除残留冷却实体异常 #" & Err.Number & " [" & Err.Source & "] " & Err.Description)
        Err.Clear
    Else
        ok_del = True
    End If
    On Error GoTo 0
    Call LogMsg("已删除残留冷却实体 " & size & " 个 (ok=" & CStr(ok_del) & ")")
End Sub

Sub SleepSec(n)
    On Error Resume Next
    WshShell.Run "cmd.exe /c ping -n " & (CLng(n) + 1) & " 127.0.0.1 >nul 2>&1", 0, True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Sub LogMsg(msg)
    On Error Resume Next
    Dim f
    Set f = FSO.OpenTextFile(LogPath, 8, True)
    f.WriteLine Now & " - " & msg
    f.Close
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
