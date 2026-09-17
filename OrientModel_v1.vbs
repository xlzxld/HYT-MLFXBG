'%RunPerInstance
'@
'@ 描述: 【存档版】Moldflow 模型定向宏 第一版 (2026-09-16 原始版)
'@   选中面 -> 面心移到原点附近 + 面朝外方向对齐 Z (绕面心旋转) ; 会弹"是否保存"询问。
'@ 已知局限 (保留以便对照测试): 用多范围拼串选择实体 (实测只有第一个范围生效)、
'@   绕面心旋转、aMerge=True (会合并重合节点) —— 新版本已全部修正, 本文件仅供对照测试。
'@ 运行: Tools -> Macro -> Play Macro (需先打开研究, 并在网格上选中一个面)
Option Explicit

Dim FSO, WshShell, BaseDir, TempDir, ConfigPath, LogPath
Dim SynergyGetter, Synergy, StudyDoc, Modeler, PredicateManager, Viewer
Dim saEnv
Dim ConfigJsonStr, ConfigObj, HTML
Dim TargetAxis, FlipNormal
Dim Sel, SelStr, SelCount, selIdx, entOne
Dim MinX, MinY, MinZ, MaxX, MaxY, MaxZ, NodeCount, TriCount
Dim NX, NY, NZ, CX, CY, CZ, NLen
Dim AxisX, AxisY, AxisZ, AngleDeg, DotVal, TX, TY, TZ, CrossLen
Dim ListAll, PredAll, EntCount
Dim VecCenter, VecAxis, VecMove
Dim okRot, okMov, DoSave
Dim MsgText

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)
TempDir = BaseDir & "\temp"
If Not FSO.FolderExists(TempDir) Then FSO.CreateFolder TempDir
ConfigPath = BaseDir & "\report_config.json"
LogPath = TempDir & "\orient.log"

' 1. 读配置 (缺失/坏格式一律回落到默认值)
TargetAxis = "Z"
FlipNormal = False
If FSO.FileExists(ConfigPath) Then
    On Error Resume Next
    ConfigJsonStr = ReadUtf8TextFile(ConfigPath)
    Set HTML = CreateObject("htmlfile")
    HTML.parentWindow.execScript "function parseJSON(s) { if (typeof JSON !== 'undefined' && JSON.parse) { try { return JSON.parse(s); } catch (e) {} } return eval('(' + s + ')'); }", "JScript"
    Set ConfigObj = HTML.parentWindow.parseJSON(ConfigJsonStr)
    TargetAxis = UCase(Trim(CStr(ConfigObj.orient_settings.target_axis)))
    FlipNormal = CBool(ConfigObj.orient_settings.flip_normal)
    If Err.Number <> 0 Then
        Err.Clear
        TargetAxis = "Z"
        FlipNormal = False
    End If
    If Len(TargetAxis) = 0 Then TargetAxis = "Z"
    On Error GoTo 0
End If
If TargetAxis <> "X" And TargetAxis <> "Y" And TargetAxis <> "Z" Then
    Call LogMsg("WARN: target_axis 取值非法 (" & TargetAxis & "), 回落为 Z")
    TargetAxis = "Z"
End If

Call LogMsg("=== 模型定向开始 (v1 存档版) === target_axis=" & TargetAxis & ", flip_normal=" & CStr(FlipNormal))

' 2. 连接 Synergy (与 AutoReport 同源的三级兜底)
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
    MsgBox "没有连上 Moldflow。请在 Moldflow 里打开研究后, 用 Tools -> Macro 运行本宏。", 16, "定向失败"
    WScript.Quit 1
End If

Set StudyDoc = Nothing
On Error Resume Next
Set StudyDoc = Synergy.StudyDoc()
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If StudyDoc Is Nothing Then
    Call LogMsg("ERROR: 当前没有打开的研究")
    MsgBox "当前没有打开的研究。请先打开研究并划好网格, 再运行本宏。", 48, "定向失败"
    WScript.Quit 1
End If

Set Modeler = Nothing
Set Viewer = Nothing
Set PredicateManager = Nothing
On Error Resume Next
Set Modeler = Synergy.Modeler()
Set Viewer = Synergy.Viewer()
Set PredicateManager = Synergy.PredicateManager()
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If Modeler Is Nothing Then
    Call LogMsg("ERROR: Modeler 对象不可用")
    MsgBox "Modeler 对象不可用, 无法执行模型变换。", 16, "定向失败"
    WScript.Quit 1
End If

' 3. 读取当前选择 (网格显示下点选的面)
Set Sel = Nothing
SelCount = 0
SelStr = ""
On Error Resume Next
Set Sel = StudyDoc.Selection
If Not Sel Is Nothing Then
    SelCount = CLng(Sel.Size)
    SelStr = CStr(Sel.ConvertToString())
End If
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If SelCount <= 0 Then
    Call LogMsg("ERROR: 选择为空")
    MsgBox "没有检测到选中的面。" & vbCrLf & vbCrLf & "请先在网格显示下点选一个面 (三角形), 再运行本宏。", 48, "定向失败"
    WScript.Quit 1
End If
Call LogMsg("选中 " & SelCount & " 个实体: " & SelStr)

' 4. 统计选中面的节点范围 / 三角形法线
MinX = 1E+30 : MinY = 1E+30 : MinZ = 1E+30
MaxX = -1E+30 : MaxY = -1E+30 : MaxZ = -1E+30
NodeCount = 0 : TriCount = 0
NX = 0 : NY = 0 : NZ = 0
For selIdx = 0 To SelCount - 1
    Set entOne = Nothing
    On Error Resume Next
    Set entOne = Sel.Entity(selIdx)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If Not entOne Is Nothing Then Call CollectElem(StudyDoc, entOne)
Next
If Err.Number <> 0 Then
    Call LogMsg("WARN: 选择遍历异常: " & Err.Description)
    Err.Clear
End If

If NodeCount = 0 Or TriCount = 0 Then
    Call LogMsg("ERROR: 未取到三角形面 (节点数=" & NodeCount & ", 三角形数=" & TriCount & ")")
    MsgBox "没能从选择里取到三角形网格。" & vbCrLf & vbCrLf & _
           "请确认: 1) 模型已划网格; 2) 在网格显示下选中一个面。" & vbCrLf & _
           "当前选择: " & SelStr, 48, "定向失败"
    WScript.Quit 1
End If

CX = (MinX + MaxX) / 2
CY = (MinY + MaxY) / 2
CZ = (MinZ + MaxZ) / 2
NLen = Sqr(NX * NX + NY * NY + NZ * NZ)
If NLen < 1E-9 Then
    Call LogMsg("ERROR: 法线求和为零 (三角形可能正反抵消)")
    MsgBox "选中面的法线互相抵消, 无法确定朝向。" & vbCrLf & "请只选中同一个面 (不要选到正反面各一半)。", 48, "定向失败"
    WScript.Quit 1
End If
NX = NX / NLen : NY = NY / NLen : NZ = NZ / NLen
If FlipNormal Then
    NX = -NX : NY = -NY : NZ = -NZ
    Call LogMsg("按配置翻转法线")
End If

' 5. 计算"把法线对齐到目标轴"的旋转轴与角度
If TargetAxis = "X" Then
    TX = 1 : TY = 0 : TZ = 0
ElseIf TargetAxis = "Y" Then
    TX = 0 : TY = 1 : TZ = 0
Else
    TX = 0 : TY = 0 : TZ = 1
End If
DotVal = NX * TX + NY * TY + NZ * TZ
If DotVal > 1 Then DotVal = 1
If DotVal < -1 Then DotVal = -1
AxisX = NY * TZ - NZ * TY
AxisY = NZ * TX - NX * TZ
AxisZ = NX * TY - NY * TX
CrossLen = Sqr(AxisX * AxisX + AxisY * AxisY + AxisZ * AxisZ)
If CrossLen < 1E-9 Then
    If DotVal > 0 Then
        AxisX = 1 : AxisY = 0 : AxisZ = 0
        AngleDeg = 0
    Else
        If Abs(NX) > 0.9 Then
            AxisX = 0 : AxisY = 1 : AxisZ = 0
        Else
            AxisX = 1 : AxisY = 0 : AxisZ = 0
        End If
        AngleDeg = 180
    End If
Else
    AxisX = AxisX / CrossLen : AxisY = AxisY / CrossLen : AxisZ = AxisZ / CrossLen
    AngleDeg = Arccos(DotVal) * 180 / 3.14159265358979
End If

Call LogMsg("面心=(" & Round(CX, 3) & ", " & Round(CY, 3) & ", " & Round(CZ, 3) & _
            ") 法线=(" & Round(NX, 4) & ", " & Round(NY, 4) & ", " & Round(NZ, 4) & _
            ") 旋转角=" & Round(AngleDeg, 2) & " 度, 轴=(" & Round(AxisX, 4) & ", " & Round(AxisY, 4) & ", " & Round(AxisZ, 4) & ")")

MsgText = "将执行模型定向 (第一版):" & vbCrLf & vbCrLf & _
          "选中: " & SelCount & " 个实体 (" & TriCount & " 个三角形, " & NodeCount & " 个节点)" & vbCrLf & _
          "面心: (" & Round(CX, 3) & ", " & Round(CY, 3) & ", " & Round(CZ, 3) & ")" & vbCrLf & _
          "面法线: (" & Round(NX, 4) & ", " & Round(NY, 4) & ", " & Round(NZ, 4) & ")" & vbCrLf & _
          "旋转: " & Round(AngleDeg, 2) & " 度, 轴 (" & Round(AxisX, 4) & ", " & Round(AxisY, 4) & ", " & Round(AxisZ, 4) & ")" & vbCrLf & vbCrLf & _
          "效果: 面心 -> (0, 0, 0); 面朝向 -> " & TargetAxis & " 轴 (点""前""视图正对该面)" & vbCrLf & vbCrLf & _
          "注意: 本操作会修改模型 (网格/节点/CAD 一起变换)。" & vbCrLf & _
          "Moldflow 可能提示结果失效, 请按提示处理。" & vbCrLf & vbCrLf & _
          "继续执行吗?"
If MsgBox(MsgText, 4 + 48, "模型定向确认") <> 6 Then
    Call LogMsg("用户取消")
    WScript.Quit 0
End If

' 6. 组装"全部实体"列表并执行旋转 + 平移
Set ListAll = Nothing
Set PredAll = Nothing
EntCount = 0
On Error Resume Next
Set ListAll = StudyDoc.CreateEntityList()
Set PredAll = PredicateManager.CreateLabelPredicate("N1:N9999999 T1:T9999999 TE1:TE9999999 B1:B9999999 C1:C9999999 S1:S9999999 F1:F9999999 BD1:BD9999999 STL1:STL9999999")
ListAll.SelectFromPredicate PredAll
EntCount = CLng(ListAll.Size)
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
Call LogMsg("全部实体数: " & EntCount)
If EntCount <= 0 Then
    MsgBox "没能枚举出模型实体 (标签谓词返回 0)。" & vbCrLf & "请把 temp\orient.log 发给维护者排查。", 16, "定向失败"
    WScript.Quit 1
End If

Set VecCenter = Synergy.CreateVector()
Set VecAxis = Synergy.CreateVector()
Set VecMove = Synergy.CreateVector()
Call VecCenter.SetXYZ(CX, CY, CZ)
Call VecAxis.SetXYZ(AxisX, AxisY, AxisZ)
Call VecMove.SetXYZ(-CX, -CY, -CZ)

okRot = True
okMov = True
On Error Resume Next
If AngleDeg > 0.001 Then
    okRot = CBool(Modeler.Rotate(ListAll, VecCenter, VecAxis, AngleDeg, False, 1, True))
End If
If Err.Number <> 0 Then
    Call LogMsg("ERROR: Rotate 异常: " & Err.Description)
    Err.Clear
    okRot = False
End If
okMov = CBool(Modeler.Translate(ListAll, VecMove, False, 1, True))
If Err.Number <> 0 Then
    Call LogMsg("ERROR: Translate 异常: " & Err.Description)
    Err.Clear
    okMov = False
End If
On Error GoTo 0
Call LogMsg("Rotate=" & CStr(okRot) & ", Translate=" & CStr(okMov))

If Not okMov Then
    MsgBox "平移失败, 模型可能只转了一半。" & vbCrLf & "请看 temp\orient.log 并发给维护者。", 16, "定向失败"
    WScript.Quit 1
End If

' 7. 视图切前视并刷新
On Error Resume Next
Call Viewer.Rotate(0, 0, 0)
Viewer.Fit
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
Call SleepSec(1)

' 8. 询问是否保存研究
DoSave = (MsgBox("定向完成。" & vbCrLf & vbCrLf & "现在保存研究吗?" & vbCrLf & _
                 "选""否""则只改内存, 你可以在 Moldflow 里撤销/手动保存。", 4 + 64, "模型定向") = 6)
If DoSave Then
    Dim saveOK
    saveOK = False
    On Error Resume Next
    saveOK = CBool(StudyDoc.Save())
    If Err.Number <> 0 Then
        Call LogMsg("ERROR: Save 异常: " & Err.Description)
        Err.Clear
        saveOK = False
    End If
    On Error GoTo 0
    Call LogMsg("Save=" & CStr(saveOK))
Else
    Call LogMsg("用户选择不保存")
End If

Call LogMsg("=== 模型定向结束 ===")
MsgBox "定向完成。" & vbCrLf & vbCrLf & _
       "面心已到 (0, 0, 0), 面朝向 " & TargetAxis & " 轴。" & vbCrLf & _
       "点 Moldflow 的""前""视图即可正对该面。", 64, "模型定向"
WScript.Quit 0

' ============================================================
' 辅助过程
' ============================================================

Function Arccos(x)
    If x >= 1 Then
        Arccos = 0
    ElseIf x <= -1 Then
        Arccos = 3.14159265358979
    Else
        Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End If
End Function

Sub Accum(pX, pY, pZ)
    If pX < MinX Then MinX = pX
    If pY < MinY Then MinY = pY
    If pZ < MinZ Then MinZ = pZ
    If pX > MaxX Then MaxX = pX
    If pY > MaxY Then MaxY = pY
    If pZ > MaxZ Then MaxZ = pZ
    NodeCount = NodeCount + 1
End Sub

Function GetCoord(SD, pEnt, pOutX, pOutY, pOutZ)
    Dim pCoord
    GetCoord = False
    Set pCoord = Nothing
    On Error Resume Next
    Set pCoord = SD.GetNodeCoord(pEnt)
    If Err.Number <> 0 Then
        Err.Clear
        Set pCoord = Nothing
    End If
    On Error GoTo 0
    If pCoord Is Nothing Then Exit Function
    Dim pX, pY, pZ
    pX = 0 : pY = 0 : pZ = 0
    On Error Resume Next
    pX = CDbl(pCoord.X)
    pY = CDbl(pCoord.Y)
    pZ = CDbl(pCoord.Z)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    pOutX = pX : pOutY = pY : pOutZ = pZ
    GetCoord = True
End Function

Sub CollectElem(SD, pEnt)
    Dim pNodes, pCnt, pK, pNd, pX, pY, pZ, pGot
    Dim pGx(2), pGy(2), pGz(2)
    Dim pAx, pAy, pAz, pBx, pBy, pBz

    pGot = 0
    Set pNodes = Nothing
    On Error Resume Next
    Set pNodes = SD.GetElemNodes(pEnt)
    If Err.Number <> 0 Then
        Err.Clear
        Set pNodes = Nothing
    End If
    On Error GoTo 0

    If pNodes Is Nothing Then
        If GetCoord(SD, pEnt, pX, pY, pZ) Then Call Accum(pX, pY, pZ)
        Exit Sub
    End If

    pCnt = -1
    On Error Resume Next
    pCnt = CLng(pNodes.Size)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If pCnt >= 0 Then
        For pK = 0 To pCnt - 1
            Set pNd = Nothing
            On Error Resume Next
            Set pNd = pNodes.Entity(pK)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            If Not pNd Is Nothing Then
                If GetCoord(SD, pNd, pX, pY, pZ) Then
                    Call Accum(pX, pY, pZ)
                    If pGot < 3 Then
                        pGx(pGot) = pX : pGy(pGot) = pY : pGz(pGot) = pZ
                        pGot = pGot + 1
                    End If
                End If
            End If
        Next
    ElseIf IsArray(pNodes) Then
        Dim pU
        pU = UBound(pNodes)
        For pK = 0 To pU
            Set pNd = pNodes(pK)
            If GetCoord(SD, pNd, pX, pY, pZ) Then
                Call Accum(pX, pY, pZ)
                If pGot < 3 Then
                    pGx(pGot) = pX : pGy(pGot) = pY : pGz(pGot) = pZ
                    pGot = pGot + 1
                End If
            End If
        Next
    End If

    If pGot = 3 Then
        pAx = pGx(1) - pGx(0) : pAy = pGy(1) - pGy(0) : pAz = pGz(1) - pGz(0)
        pBx = pGx(2) - pGx(0) : pBy = pGy(2) - pGy(0) : pBz = pGz(2) - pGz(0)
        NX = NX + (pAy * pBz - pAz * pBy)
        NY = NY + (pAz * pBx - pAx * pBz)
        NZ = NZ + (pAx * pBy - pAy * pBx)
        TriCount = TriCount + 1
    End If
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

Function ReadUtf8TextFile(Path)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile Path
    ReadUtf8TextFile = stream.ReadText
    stream.Close
End Function
