'%RunPerInstance
'@
'@ 描述: Moldflow 模型定向宏 (需求 1, 2026-09-16)。网格显示下选中一个面, 把该面
'@   "转正 + 归零": 面心移动到世界原点 (0,0,0), 面朝外方向对齐 target_axis (默认 Z);
'@   之后点"前"视图即正对该面。复刻 GUI 手工流程: 全选所有实体 -> 先移到原点 -> 再绕原点旋转。
'@ 运行: Tools -> Macro -> Play Macro (需先打开研究, 并在网格上选中一个面)
'@ 配置: report_config.json -> orient_settings; 日志: temp/orient.log
Option Explicit

Dim FSO, WshShell, BaseDir, TempDir, ConfigPath, LogPath
Dim SynergyGetter, Synergy, StudyDoc, Modeler, Viewer, PredicateManager
Dim saEnv
Dim ConfigJsonStr, ConfigObj, HTML
Dim TargetAxis, FlipNormal
Dim Sel, SelStr, SelCount, selIdx, entOne
Dim MinX, MinY, MinZ, MaxX, MaxY, MaxZ, NodeCount, TriCount
Dim NX, NY, NZ, CX, CY, CZ, NLen
Dim AxisX, AxisY, AxisZ, AngleDeg, DotVal, TX, TY, TZ, CrossLen
Dim VecOrigin, VecAxis, VecMove
Dim ListAll, EntCount
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

Call LogMsg("=== 模型定向开始 === target_axis=" & TargetAxis & ", flip_normal=" & CStr(FlipNormal))

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

MsgText = "即将把选中的面『转正 + 对齐原点』:" & vbCrLf & vbCrLf & _
          "· 面心会移动到坐标 0,0,0" & vbCrLf & _
          "· 这个面会正对屏幕 (点 Moldflow 的『前』视图就是正对它看)" & vbCrLf & _
          "· 全部实体一起动 (网格/流道/水路/CAD), 与手工全选旋转等效" & vbCrLf & vbCrLf & _
          "选中: " & TriCount & " 个三角形, 面心 (" & Round(CX, 3) & ", " & Round(CY, 3) & ", " & Round(CZ, 3) & ")" & vbCrLf & _
          "旋转 " & Round(AngleDeg, 2) & " 度 (轴 " & Round(AxisX, 3) & ", " & Round(AxisY, 3) & ", " & Round(AxisZ, 3) & ")" & vbCrLf & vbCrLf & _
          "注意1: 会修改模型; 若研究里已有分析结果, Moldflow 会让结果失效," & vbCrLf & _
          "       首次可能等待几分钟 (这是结果失效的一次性处理, 不是卡死), 请勿中断。" & vbCrLf & _
          "注意2: 本宏不会自动保存, 需要保存时请你在 Moldflow 里自行保存。" & vbCrLf & vbCrLf & _
          "现在执行吗?"
If MsgBox(MsgText, 4 + 48, "模型定向确认") <> 6 Then
    Call LogMsg("用户取消")
    WScript.Quit 0
End If

' 6. 先拍一张"调整前"截图 (temp/orient_before.png), 便于对照
Call SnapView(Viewer, TempDir & "\orient_before.png")

' 7. 全选所有实体 (关键: 必须一次调用移动所有实体。实测曲线与梁共用节点,
'    分趟移动会把共用节点移动两次 -> 模型错乱), 先平移到原点, 再绕原点旋转。
Dim AllOK
AllOK = False
Set ListAll = Nothing
EntCount = 0
On Error Resume Next
Set ListAll = AllEntities(StudyDoc, PredicateManager)
If Not ListAll Is Nothing Then EntCount = CLng(ListAll.Size)
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
Call LogMsg("全选实体数: " & EntCount)

If EntCount > 0 Then
    Set VecOrigin = Synergy.CreateVector()
    Set VecAxis = Synergy.CreateVector()
    Set VecMove = Synergy.CreateVector()
    Call VecOrigin.SetXYZ(0, 0, 0)
    Call VecAxis.SetXYZ(AxisX, AxisY, AxisZ)
    Call VecMove.SetXYZ(-CX, -CY, -CZ)

    Dim okMov, okRot
    okMov = True
    okRot = True
    On Error Resume Next
    okMov = CBool(Modeler.Translate(ListAll, VecMove, False, 1, False))
    If Err.Number <> 0 Then
        Call LogMsg("ERROR: Translate 异常: " & Err.Description)
        Err.Clear
        okMov = False
    End If
    If AngleDeg > 0.001 Then
        okRot = CBool(Modeler.Rotate(ListAll, VecOrigin, VecAxis, AngleDeg, False, 1, False))
    End If
    If Err.Number <> 0 Then
        Call LogMsg("ERROR: Rotate 异常: " & Err.Description)
        Err.Clear
        okRot = False
    End If
    On Error GoTo 0
    Call LogMsg("Translate=" & CStr(okMov) & ", Rotate=" & CStr(okRot))
    AllOK = okMov And okRot
Else
    Call LogMsg("ERROR: 全选谓词返回 0 个实体, 兜底走分趟变换")
    AllOK = FallbackMultiPass(StudyDoc, Modeler, PredicateManager, AxisX, AxisY, AxisZ, AngleDeg, CX, CY, CZ)
End If

If Not AllOK Then
    MsgBox "变换执行失败或不完整。" & vbCrLf & "请看 temp\orient.log 并发给维护者。", 16, "定向失败"
    WScript.Quit 1
End If

' 7.2 CAD 本体 (BD*) 变换趟
'     实机取证 (2026-09-16 录制脚本): GUI 旋转 CAD 时用的是
'     EntList.SelectFromString " BD1 BD2 ... " + Modeler.Rotate;
'     实机确证 (2026-09-16, 宏上下文对照试验): 选择解析器对 N/FD/L 前缀正常,
'     但 BD* 在任何上下文都选不中 —— CAD 实体不在 API 实体空间内, 本趟永远探测为 0。
'     保留本趟意义: 万一未来版本放开, 无需改宏; 探测不到时在日志给出可照抄的手工旋转参数。
Dim CadSpec, CadProbe, cadK, CadN, cadOK
Dim cadNodeRef, cadNodeBefore, cadNodeAfter, cadCoord, cadCoord2, cadHint
CadSpec = ""
CadN = 0
Set CadProbe = Nothing
' 安全采样: 记一个网格节点坐标, CAD 趟结束后对比 (防止 CAD 与网格共用节点被二次位移)
Set cadNodeRef = Nothing
cadNodeBefore = ""
cadNodeAfter = ""
On Error Resume Next
Set cadNodeRef = StudyDoc.GetFirstNode()
If Not cadNodeRef Is Nothing Then
    Set cadCoord = StudyDoc.GetNodeCoord(cadNodeRef)
    If Not cadCoord Is Nothing Then cadNodeBefore = CStr(cadCoord.X) & "," & CStr(cadCoord.Y) & "," & CStr(cadCoord.Z)
End If
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
On Error Resume Next
For cadK = 1 To 200
    ' 与录制脚本对齐: 选择串前后加空格 (" BD1 "), 且每次用新列表, 避免解析/状态问题
    Set CadProbe = Modeler.CreateEntityList()
    If Not CadProbe Is Nothing Then
        CadProbe.SelectFromString " BD" & cadK & " "
        If CLng(CadProbe.Size) > 0 Then
            If Len(CadSpec) > 0 Then CadSpec = CadSpec & " "
            CadSpec = CadSpec & "BD" & cadK
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
Next
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
If Len(CadSpec) > 0 Then
    Call LogMsg("CAD 本体探测到: " & CadSpec)
    cadOK = False
    On Error Resume Next
    Set CadProbe = Modeler.CreateEntityList()
    CadProbe.SelectFromString " " & CadSpec & " "
    CadN = CLng(CadProbe.Size)
    If CadN > 0 Then
        Call LogMsg("CAD 选择串回读: " & CStr(CadProbe.ConvertToString()))
    End If
    If CadN > 0 Then
        Dim okMov2, okRot2
        okMov2 = CBool(Modeler.Translate(CadProbe, VecMove, False, 1, False))
        If AngleDeg > 0.001 Then
            okRot2 = CBool(Modeler.Rotate(CadProbe, VecOrigin, VecAxis, AngleDeg, False, 1, False))
        Else
            okRot2 = True
        End If
        cadOK = okMov2 And okRot2
        Call LogMsg("CAD 本体变换: " & CadN & " 个, Translate=" & CStr(okMov2) & ", Rotate=" & CStr(okRot2))
    End If
    If Err.Number <> 0 Then
        Call LogMsg("WARN: CAD 本体变换异常: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
    ' 安全校验: CAD 趟不应影响网格节点
    If Not cadNodeRef Is Nothing Then
        On Error Resume Next
        Set cadCoord2 = StudyDoc.GetNodeCoord(cadNodeRef)
        If Not cadCoord2 Is Nothing Then cadNodeAfter = CStr(cadCoord2.X) & "," & CStr(cadCoord2.Y) & "," & CStr(cadCoord2.Z)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Len(cadNodeBefore) > 0 And Len(cadNodeAfter) > 0 Then
            If cadNodeBefore <> cadNodeAfter Then
                Call LogMsg("WARN: CAD 趟影响了网格节点 (" & cadNodeBefore & " -> " & cadNodeAfter & "), 请检查模型!")
            Else
                Call LogMsg("安全校验: CAD 趟未影响网格节点 (正常)")
            End If
        End If
    End If
Else
    cadHint = "未探测到 CAD 本体 —— 已实机确证: CAD 实体不在 API 实体空间内 (BD* 两种连接路径均选不中)。"
    cadHint = cadHint & " 若你的研究含 CAD 几何 (CAD 几何文件夹), 请在 GUI 里手工旋转: 参考点 (0,0,0), 轴 ("
    cadHint = cadHint & Round(AxisX, 4) & ", " & Round(AxisY, 4) & ", " & Round(AxisZ, 4) & "), 角度 "
    cadHint = cadHint & Round(AngleDeg, 2) & " 度"
    Call LogMsg(cadHint)
End If

' 8. 校验: 重新读选中面的中心, 期望落在原点附近
Dim vx, vy, vz, dev
Call SelectionCenter(StudyDoc, Sel, vx, vy, vz)
dev = Sqr(vx * vx + vy * vy + vz * vz)
Call LogMsg("校验: 选中面中心=(" & Round(vx, 3) & ", " & Round(vy, 3) & ", " & Round(vz, 3) & "), 偏差=" & Round(dev, 3))

' 9. 视图切前视并刷新
On Error Resume Next
Call Viewer.Rotate(0, 0, 0)
Viewer.Fit
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0
Call SleepSec(1)
Call SnapView(Viewer, TempDir & "\orient_after.png")

' 10. 收尾: 不自动保存 (用户明确要求: 默认不保存, 需要时自己保存)
Call LogMsg("定向完成: 面心=(" & Round(vx, 3) & ", " & Round(vy, 3) & ", " & Round(vz, 3) & _
            "), 偏差=" & Round(dev, 3) & "; 未自动保存 (需要保存请在 Moldflow 里自行保存)")

Call LogMsg("=== 模型定向结束 ===")
WScript.Quit 0

' ============================================================
' 变换与辅助过程
' ============================================================

Function AllEntities(SD, pPredMgr)
    ' 全选所有实体: 官方谓词一次只能带一个范围, 但可以用
    ' "不存在的属性类型" 构造空谓词, 再取反 (NOT) 得到全部实体。
    Dim pPredNone, pPredAll, pLst
    Set AllEntities = Nothing
    Set pPredNone = Nothing
    Set pPredAll = Nothing
    On Error Resume Next
    Set pPredNone = pPredMgr.CreatePropTypePredicate(999999)
    If pPredNone Is Nothing Then
        Err.Clear
        Exit Function
    End If
    Set pPredAll = pPredMgr.CreateBoolNotPredicate(pPredNone)
    If pPredAll Is Nothing Then
        Err.Clear
        Exit Function
    End If
    Set pLst = SD.CreateEntityList()
    If pLst Is Nothing Then
        Err.Clear
        Exit Function
    End If
    pLst.SelectFromPredicate pPredAll
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    Set AllEntities = pLst
End Function

Function FallbackMultiPass(SD, pModeler, pPredMgr, pAx, pAy, pAz, pAngle, pCx, pCy, pCz)
    ' 兜底: 老版本没有 BoolNot 谓词时按类型分批 (注意: 曲线与梁共用节点,
    ' 该路径可能双重位移, 仅作最后手段; 日志会明确警告)
    Dim pLst, pPred, pIdx, pPrefixes, pOk, pMov, pRot
    Dim pVecO, pVecA, pVecM
    Call LogMsg("WARN: 进入兜底分趟路径 (可能双重位移, 建议改用支持谓词取反的版本)")
    Set pVecO = Synergy.CreateVector()
    Set pVecA = Synergy.CreateVector()
    Set pVecM = Synergy.CreateVector()
    Call pVecO.SetXYZ(0, 0, 0)
    Call pVecA.SetXYZ(pAx, pAy, pAz)
    Call pVecM.SetXYZ(-pCx, -pCy, -pCz)
    pPrefixes = Array("N")
    pOk = True
    For pIdx = 0 To UBound(pPrefixes)
        Set pLst = SD.CreateEntityList()
        Set pPred = pPredMgr.CreateLabelPredicate(pPrefixes(pIdx) & "1:" & pPrefixes(pIdx) & "9999999")
        pLst.SelectFromPredicate pPred
        If CLng(pLst.Size) > 0 Then
            pMov = pModeler.Translate(pLst, pVecM, False, 1, False)
            pRot = pModeler.Rotate(pLst, pVecO, pVecA, pAngle, False, 1, False)
            If Not pMov Then pOk = False
        End If
    Next
    FallbackMultiPass = pOk
End Function

Function CadIdString(SD, pPrefix, pCap)
    ' 备用: CAD 面/体不被标签谓词支持, 用 SelectFromString 指数探测 + 线性补齐
    Dim pLst, pK, pUpper, pStr
    CadIdString = ""
    Set pLst = Nothing
    On Error Resume Next
    Set pLst = SD.CreateEntityList()
    On Error GoTo 0
    If pLst Is Nothing Then Exit Function
    pUpper = 0
    pK = 1
    Do While pK <= pCap
        On Error Resume Next
        pLst.SelectFromString pPrefix & pK
        If CLng(pLst.Size) > 0 Then pUpper = pK
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        On Error GoTo 0
        If pUpper < pK Then Exit Do
        pK = pK * 2
    Loop
    If pUpper = 0 Then Exit Function
    If pUpper > pCap Then pUpper = pCap
    pStr = ""
    For pK = 1 To pUpper
        On Error Resume Next
        pLst.SelectFromString pPrefix & pK
        If CLng(pLst.Size) > 0 Then
            If Len(pStr) > 0 Then pStr = pStr & " "
            pStr = pStr & pPrefix & pK
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    Next
    CadIdString = pStr
End Function

Sub SelectionCenter(SD, pSel, pOutX, pOutY, pOutZ)
    Dim pCnt, pI, pEnt
    Call ResetAccum()
    pCnt = 0
    On Error Resume Next
    pCnt = CLng(pSel.Size)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    For pI = 0 To pCnt - 1
        Set pEnt = Nothing
        On Error Resume Next
        Set pEnt = pSel.Entity(pI)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not pEnt Is Nothing Then Call CollectElem(SD, pEnt)
    Next
    If NodeCount > 0 Then
        pOutX = (MinX + MaxX) / 2
        pOutY = (MinY + MaxY) / 2
        pOutZ = (MinZ + MaxZ) / 2
    Else
        pOutX = 0 : pOutY = 0 : pOutZ = 0
    End If
End Sub

Sub ResetAccum()
    MinX = 1E+30 : MinY = 1E+30 : MinZ = 1E+30
    MaxX = -1E+30 : MaxY = -1E+30 : MaxZ = -1E+30
    NodeCount = 0 : TriCount = 0
    NX = 0 : NY = 0 : NZ = 0
End Sub

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
    Dim pCoord, pX, pY, pZ
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
    ' 收集一个实体的节点; 若恰好 3 个节点则按右手法则累加三角形法线
    Dim pNodes, pCnt, pK, pNd, pX, pY, pZ, pGot, pU
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

Sub SnapView(pViewer, pPath)
    On Error Resume Next
    pViewer.SaveImage3 pPath, 1600, 900, True, True, True, True, True, True, True, True, True
    If Err.Number <> 0 Then
        Call LogMsg("WARN: 截图失败 (" & pPath & "): " & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
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
