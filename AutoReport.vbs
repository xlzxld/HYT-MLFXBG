'%RunPerInstance
'@
'@ DESCRIPTION
'@ Autodesk Moldflow 2023 模流分析报告 PPT 全自动生成 (宏)
'@ 支持: 中性面(Midplane) / 双层面(Dual Domain) / 3D 实体全自适应
'@
'@@
Option Explicit

Dim FSO, WshShell, ScriptDir, BaseDir, TempDir, ConfigPath, ModeADir, ModeBDir, ScreenshotMode
Dim ConfigJsonStr, HTML, ConfigObj
Dim ImageWidth, ImageHeight, KeepView, NFrames, DelayMs
Dim SynergyGetter, Synergy, StudyDoc, PlotManager, Viewer, DiagnosisManager
Dim StudyName, MeshTypeRaw, DetectedMeshType
Dim MeshSummary, MeshText
Dim MatID, MatSubID, MatPlot
Dim i, pObj, pKey, pName, pType, pEnabled, PlotObj, ImgPath, GifPath
Dim PlotsArray, ManifestText, TodayStr, PyCmd, ret, ExportCount, saEnv

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' 1. 确定基准目录与临时目录
' 可移植性: 基准目录一律取本脚本所在目录 (模板/输出/临时目录均相对它解析), 不再硬编码机器路径
BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)

TempDir = BaseDir & "\temp"
ConfigPath = BaseDir & "\report_config.json"

If Not FSO.FolderExists(TempDir) Then
    FSO.CreateFolder TempDir
End If
ModeADir = TempDir & "\mode_a"
ModeBDir = TempDir & "\mode_b"
If Not FSO.FolderExists(ModeADir) Then FSO.CreateFolder ModeADir
If Not FSO.FolderExists(ModeBDir) Then FSO.CreateFolder ModeBDir

Call LogMsg("=== 模流报告生成开始 ===")
Call LogMsg("BaseDir: " & BaseDir)

' 2. 读取配置文件
If Not FSO.FileExists(ConfigPath) Then
    MsgBox "找不到配置文件: " & ConfigPath, 16, "错误"
    Call LogMsg("ERROR: 配置文件不存在: " & ConfigPath)
    WScript.Quit 1
End If

ConfigJsonStr = ReadUtf8TextFile(ConfigPath)

Set HTML = CreateObject("htmlfile")
HTML.parentWindow.execScript "function parseJSON(s) { return eval('(' + s + ')'); } function getArrayItem(arr, i) { return arr[i]; }", "JScript"
Set ConfigObj = HTML.parentWindow.parseJSON(ConfigJsonStr)

' 弹出配置界面 (可选, 由 show_gui_before_run 控制)
' 由 GUI 拉起时跳过 (GUI 自己会拉起本脚本, 防双重运行)
Dim fromGui
fromGui = WshShell.ExpandEnvironmentStrings("%MLFXBG_FROM_GUI%")
If CBool(ConfigObj.show_gui_before_run) And fromGui <> "1" Then
    Call LogMsg("弹出配置界面...")
    ret = WshShell.Run("python """ & BaseDir & "\config_gui.py""", 1, True)
    ConfigJsonStr = ReadUtf8TextFile(ConfigPath)
    Set ConfigObj = HTML.parentWindow.parseJSON(ConfigJsonStr)

'  完成标记守卫: GUI 里点生成会自行拉起本脚本跑完整流程; 用户关窗返回到本实例后,
'  若标记存在说明本次已生成过, 直接退出 — 否则会把 Moldflow 导出 + PPT 再跑一遍。
    If FSO.FileExists(TempDir & "\gui_run_done.txt") Then
        Call LogMsg("检测到 gui_run_done.txt: 本次生成已由配置界面驱动完成, 退出避免重复运行")
        WScript.Quit 0
    End If
End If

ImageWidth = CLng(ConfigObj.image_settings.width)
ImageHeight = CLng(ConfigObj.image_settings.height)
KeepView = CBool(ConfigObj.image_settings.keep_view)
NFrames = CLng(ConfigObj.animation_settings.frames)
DelayMs = CLng(ConfigObj.animation_settings.delay_ms)

ScreenshotMode = "B"
On Error Resume Next
ScreenshotMode = UCase(CStr(ConfigObj.screenshot_mode))
On Error GoTo 0
If ScreenshotMode = "" Then ScreenshotMode = "B"
Call LogMsg("截图方案模式: " & ScreenshotMode)

Call LogMsg("画质参数: Width=" & ImageWidth & ", Height=" & ImageHeight & ", NFrames=" & NFrames)

' 4K (3840x2160) 已禁用: SaveImage3 离屏导出在 4K 下实机定案大面积断带 (AI_GUIDE.md 坑册)。
' 旧配置/手改配置在此钳到 1080P, 与 config_gui.collect_config 的钳制双保险。
If ImageWidth >= 3840 Or ImageHeight >= 2160 Then
    Call LogMsg("WARN: 画质 " & ImageWidth & "x" & ImageHeight & " 触发 4K 禁用守卫, 钳到 1920x1080")
    ImageWidth = 1920
    ImageHeight = 1080
End If

' 2.5 陈旧产物清理: 先删本次将重导出的 per-plot 文件, 防止导出中断时旧图混入本次报告
Dim kk
For i = 0 To ConfigObj.plots.length - 1
    Set pObj = HTML.parentWindow.getArrayItem(ConfigObj.plots, i)
    kk = CStr(pObj.key)
    If kk <> "" Then
        DeleteIfPresent TempDir, kk & ".png"
        DeleteIfPresent TempDir, kk & ".gif"
        DeleteIfPresent ModeADir, kk & ".png"
        DeleteIfPresent ModeBDir, kk & ".png"
        DeleteIfPresent ModeBDir, "scale_" & kk & ".png"
        DeleteIfPresent ModeBDir, "model_" & kk & ".png"
    End If
Next
DeleteIfPresent TempDir, "solid_model.png"
DeleteIfPresent TempDir, "solid_model_cad.png"
DeleteIfPresent TempDir, "mesh_model.png"
DeleteIfPresent TempDir, "material_basic.png"
DeleteIfPresent TempDir, "material_process.png"
DeleteIfPresent TempDir, "material_viscosity.png"
DeleteIfPresent TempDir, "material_pvt.png"
DeleteIfPresent TempDir, "last_output_path.txt"
DeleteIfPresent TempDir, "gui_run_done.txt"
DeleteIfPresent TempDir, "material_info.json"
DeleteIfPresent TempDir, "material_fields.json"
DeleteIfPresent TempDir, "mesh_summary.json"
DeleteIfPresent TempDir, "peak_values.json"
DeleteIfPresent TempDir, "clamp_force_info.json"
DeleteIfPresent TempDir, "manifest.json"
DeleteIfPresent TempDir, "solid_model_cover.jpg"
DeleteIfPresent TempDir, "mesh_model_cropped.png"
DeleteIfPresent TempDir, "clamp_force_xy_curve_data.txt"
DeleteIfPresent TempDir, "inj_pressure_xy_curve_data.txt"
Call LogMsg("陈旧产物清理完成")
Call LogMsg("=== RUN START " & Year(Now) & "-" & Right("0" & Month(Now),2) & "-" & Right("0" & Day(Now),2) & " " & Hour(Now) & ":" & Right("0" & Minute(Now),2) & ":" & Right("0" & Second(Now),2) & " mode=" & ScreenshotMode & " res=" & ImageWidth & "x" & ImageHeight & " ===")
On Error Resume Next
FSO.CopyFile ConfigPath, TempDir & "\config_used.json", True
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

' 3. 连接 Moldflow Synergy 实例
Set SynergyGetter = Nothing
Set Synergy = Nothing
Set StudyDoc = Nothing
Set MeshSummary = Nothing
Set PlotManager = Nothing
Set Viewer = Nothing
Set DiagnosisManager = Nothing

Dim ComMethod
ComMethod = "none"
saEnv = WshShell.ExpandEnvironmentStrings("%SAInstance%")
Call LogMsg("SAInstance : " & saEnv)

On Error Resume Next
If saEnv <> "" And saEnv <> "%SAInstance%" Then
    Set SynergyGetter = GetObject(saEnv)
    If Not SynergyGetter Is Nothing Then
        Set Synergy = SynergyGetter.GetSASynergy
        ComMethod = "SAInstance"
        Call LogMsg("已通过 SAInstance 连接已打开的 Moldflow Synergy GUI 实例")
    End If
End If
On Error GoTo 0

If Synergy Is Nothing Then
    On Error Resume Next
    Set Synergy = GetObject(, "Synergy.Synergy")
    If Synergy Is Nothing Then
        Set Synergy = CreateObject("Synergy.Synergy")
        ComMethod = "CreateObject"
    Call LogMsg("通过 CreateObject 新建 Synergy 实例")
    Else
        ComMethod = "GetObject"
        Call LogMsg("通过 GetObject 连接已打开的 Synergy 实例")
    End If
    On Error GoTo 0
End If

If Synergy Is Nothing Then
    MsgBox "未连接到 Autodesk Moldflow Synergy，请确认 Moldflow 已启动", 16, "错误"
    Call LogMsg("ERROR: 未连接到 Synergy")
    WScript.Quit 1
End If

Set StudyDoc = Synergy.StudyDoc()
If StudyDoc Is Nothing Then
    MsgBox "当前没有打开的方案(Study)，请先在 Moldflow 中打开一个方案", 48, "提示"
    Call LogMsg("ERROR: 未获取到 StudyDoc")
    WScript.Quit 1
End If

Set PlotManager = Synergy.PlotManager()
Set Viewer = Synergy.Viewer()
Set DiagnosisManager = Synergy.DiagnosisManager()

StudyName = StudyDoc.StudyName
MeshTypeRaw = UCase(CStr(StudyDoc.MeshType))
Dim PartName, PartCADNames
PartName = ""
On Error Resume Next
Set PartCADNames = StudyDoc.GetPartCadNames
If Not PartCADNames Is Nothing Then
    If PartCADNames.Size() > 0 Then PartName = CStr(PartCADNames.Val(0))
End If
On Error GoTo 0
Call LogMsg("方案: " & StudyName & ", 原始网格类型: " & MeshTypeRaw & ", 零件名: " & PartName)

' 4. 网格统计信息提取 (3D / 双层面 / 中性面)
On Error Resume Next
Set MeshSummary = DiagnosisManager.GetMeshSummary(False)
On Error GoTo 0

DetectedMeshType = "DualDomain"
If InStr(MeshTypeRaw, "3D") > 0 Or InStr(MeshTypeRaw, "TET") > 0 Then
    DetectedMeshType = "3D"
ElseIf InStr(MeshTypeRaw, "MID") > 0 Then
    DetectedMeshType = "Midplane"
ElseIf Not MeshSummary Is Nothing Then
    Dim MatchVal, RecipVal, MVol, TriCount, NodeCount, FreeE, ManiE, NonManiE, UnorientE, InterE, OverE
    TriCount = SafeGetLong(MeshSummary, "TrianglesCount", 0)
    NodeCount = SafeGetLong(MeshSummary, "NodesCount", 0)
    MVol = Round(SafeGetDbl(MeshSummary, "MeshVolume", 0), 3)
    FreeE = SafeGetLong(MeshSummary, "FreeEdgesCount", 0)
    ManiE = SafeGetLong(MeshSummary, "ManifoldEdgesCount", 0)
    NonManiE = SafeGetLong(MeshSummary, "NonManifoldEdgesCount", 0)
    UnorientE = SafeGetLong(MeshSummary, "Unoriented", 0)
    InterE = SafeGetLong(MeshSummary, "IntersectionElements", 0)
    OverE = SafeGetLong(MeshSummary, "OverlapElements", 0)
    
    MatchVal = SafeGetDbl(MeshSummary, "MatchRatio", 0)
    If MatchVal > 0 And MatchVal <= 1.0 Then MatchVal = MatchVal * 100.0
    MatchVal = Round(MatchVal, 1)

    RecipVal = SafeGetDbl(MeshSummary, "ReciprocalMatchRatio", 0)
    If RecipVal > 0 And RecipVal <= 1.0 Then RecipVal = RecipVal * 100.0
    RecipVal = Round(RecipVal, 1)

    Dim MSA, MSAStr
    MSA = SafeGetDbl(MeshSummary, "SurfaceArea", -1.0)
    If MSA < 0 Then
        MSAStr = "null"
    Else
        MSAStr = CStr(Round(MSA, 2))
    End If

    Dim MeshJson
    MeshJson = "{" & vbCrLf & _
               "  ""mesh_type"": """ & DetectedMeshType & """," & vbCrLf & _
               "  ""triangles"": " & CStr(TriCount) & "," & vbCrLf & _
               "  ""nodes"": " & CStr(NodeCount) & "," & vbCrLf & _
               "  ""tetras"": " & CStr(SafeGetLong(MeshSummary, "TetrasCount", 0)) & "," & vbCrLf & _
               "  ""beams"": " & CStr(SafeGetLong(MeshSummary, "BeamsCount", 0)) & "," & vbCrLf & _
               "  ""connectivity_regions"": " & CStr(SafeGetLong(MeshSummary, "ConnectivityRegions", 1)) & "," & vbCrLf & _
               "  ""unvisible_triangles"": " & CStr(SafeGetLong(MeshSummary, "ZeroAreaTrianglesCount", 0)) & "," & vbCrLf & _
               "  ""volume"": " & CStr(MVol) & "," & vbCrLf & _
               "  ""surface_area"": " & MSAStr & "," & vbCrLf & _
               "  ""max_aspect_ratio"": " & CStr(Round(SafeGetDbl(MeshSummary, "MaxAspectRatio", 0), 2)) & "," & vbCrLf & _
               "  ""ave_aspect_ratio"": " & CStr(Round(SafeGetDbl(MeshSummary, "AveAspectRatio", 0), 2)) & "," & vbCrLf & _
               "  ""min_aspect_ratio"": " & CStr(Round(SafeGetDbl(MeshSummary, "MinAspectRatio", 0), 2)) & "," & vbCrLf & _
               "  ""free_edges"": " & CStr(FreeE) & "," & vbCrLf & _
               "  ""manifold_edges"": " & CStr(ManiE) & "," & vbCrLf & _
               "  ""non_manifold_edges"": " & CStr(NonManiE) & "," & vbCrLf & _
               "  ""unoriented"": " & CStr(UnorientE) & "," & vbCrLf & _
               "  ""intersection_elements"": " & CStr(InterE) & "," & vbCrLf & _
               "  ""overlap_elements"": " & CStr(OverE) & "," & vbCrLf & _
               "  ""match_ratio"": " & CStr(MatchVal) & "," & vbCrLf & _
               "  ""reciprocal_match_ratio"": " & CStr(RecipVal) & vbCrLf & _
               "}"
    WriteUtf8TextFile TempDir & "\mesh_summary.json", MeshJson
    Call LogMsg("网格统计数据已导出到 mesh_summary.json")
End If

' 预清理视口 (导出纯模型图前隐藏所有结果图)
    ' 确保视口只显示原始模型
    Dim pActive
    Set pActive = PlotManager.GetFirstPlot()
    While Not pActive Is Nothing
        Viewer.HidePlot pActive
        Set pActive = PlotManager.GetNextPlot(pActive)
    Wend
    ' 4.1 导出纯模型本体截图 (Slide 1 封面专用)
    ' 用户裁决 (2026-09-08, 图层管理器截图): 截模型本体时打开「CAD 几何」所有
    ' 子层、关闭其它所有图层 (网格节点/网格单元), 得到无网格线的干净实体。
    ' 层识别经探针取证: 逐层 Name 写入 run.log (On Error 试探, 禁止臆造 API)。
    ' 双导出兜底: solid_model_cad.png (仅亮 CAD 层, 首选) + solid_model.png
    ' (现行网格自适应, Fusion 隐藏 T 曾全白故保留), Python 侧择优选用。
    Dim LayerManager, L1, I_type, TypeToHide, TypeToShow, LayerName, HasCadLayer
    Set LayerManager = Synergy.LayerManager()
    If Not LayerManager Is Nothing Then
        ' 探针取证: 逐层层名 (取不到则记空, 便于实机核对层树)
        Set L1 = LayerManager.GetFirst()
        Do While Not L1 Is Nothing
            LayerName = ""
            On Error Resume Next
            LayerName = CStr(L1.Name)
            On Error GoTo 0
            Call LogMsg("[LayerProbe] layer name=""" & LayerName & """")
            Set L1 = LayerManager.GetNext(L1)
        Loop
        ' 第一遍: 只亮「CAD 几何」层 (层名含 CAD, 不区分大小写), 其余层隐藏网格与标注类型
        HasCadLayer = False
        Set L1 = LayerManager.GetFirst()
        While Not L1 Is Nothing
            LayerName = ""
            On Error Resume Next
            LayerName = CStr(L1.Name)
            On Error GoTo 0
            If InStr(1, LayerName, "CAD", vbTextCompare) > 0 Then
                HasCadLayer = True
                For Each I_type In Array("C", "S", "R", "STL", "BD")
                    On Error Resume Next
                    LayerManager.SetTypeVisible L1, I_type, True
                    On Error GoTo 0
                Next
            Else
                For Each I_type In Array("N", "B", "T", "NBC", "SBC", "LCS", "TE")
                    On Error Resume Next
                    LayerManager.SetTypeVisible L1, I_type, False
                    On Error GoTo 0
                Next
            End If
            Set L1 = LayerManager.GetNext(L1)
        Wend
        If HasCadLayer Then
            On Error Resume Next
            Viewer.Fit ' 模型全景入框
            On Error GoTo 0
            ' Fit 为异步生效: 立即 SaveImage3 曾截到模型贴边/出框的中间态 (实发"模型割裂/不完整")
            Call SleepSec(1)
            Viewer.SaveImage3 TempDir & "\solid_model_cad.png", ImageWidth, ImageHeight, True, False, False, False, False, False, False, False, False
            If FSO.FileExists(TempDir & "\solid_model_cad.png") Then
                Call LogMsg("CAD 几何图层模型本体已导出: solid_model_cad.png (仅亮 CAD 层)")
            Else
                Call LogMsg("WARN: solid_model_cad.png 导出失败, 封面将回退 solid_model.png")
            End If
        Else
            Call LogMsg("WARN: 未识别到名称含 CAD 的图层 (层名探针见上方日志), 封面沿用 solid_model.png 路径")
        End If
        ' 第二遍: 现行网格自适应导出 solid_model.png (兜底, 逻辑与历史版本一致)
        ' 实机取证 (2026-09-08): Fusion/中性面方案的模型本体就是三角形单元 (T),
        ' 无条件隐藏 T 会导出全白空图, 仅 3D 实体方案才隐藏 T/TE。
        If InStr(MeshTypeRaw, "3D") > 0 Or InStr(MeshTypeRaw, "TET") > 0 Then
            TypeToHide = Array("N", "B", "T", "NBC", "SBC", "LCS", "TE")
        Else
            TypeToHide = Array("N", "B", "NBC", "SBC", "LCS")
        End If
        TypeToShow = Array("C", "S", "R", "STL", "BD")
        Set L1 = LayerManager.GetFirst()
        While Not L1 Is Nothing
            For Each I_type In TypeToHide
                LayerManager.SetTypeVisible L1, I_type, False
            Next
            For Each I_type In TypeToShow
                LayerManager.SetTypeVisible L1, I_type, True
            Next
            Set L1 = LayerManager.GetNext(L1)
        Wend
        On Error Resume Next
        Viewer.Fit ' 模型全景入框 (此导出点无结果图与数据条, Fit 安全; 修复封面局部放大裁切)
        On Error GoTo 0
        Call SleepSec(1)
        Viewer.SaveImage3 TempDir & "\solid_model.png", ImageWidth, ImageHeight, True, False, False, False, False, False, False, False, False
        If FSO.FileExists(TempDir & "\solid_model.png") Then
            FSO.CopyFile TempDir & "\solid_model.png", ModeADir & "\solid_model.png", True
            FSO.CopyFile TempDir & "\solid_model.png", ModeBDir & "\solid_model.png", True
            Call LogMsg("网格自适应模型截图已导出: solid_model.png (封面兜底)")
        End If
        ' 恢复网格显示供 Slide 2 网格质量页使用
        Set L1 = LayerManager.GetFirst()
        While Not L1 Is Nothing
            For Each I_type In Array("N", "T", "TE")
                LayerManager.SetTypeVisible L1, I_type, True
            Next
            Set L1 = LayerManager.GetNext(L1)
        Wend
    End If

    On Error Resume Next
    Viewer.Fit
    On Error GoTo 0
    Call SleepSec(1)
    Viewer.SaveImage3 TempDir & "\mesh_model.png", ImageWidth, ImageHeight, True, False, False, False, False, False, False, False, False
    If FSO.FileExists(TempDir & "\mesh_model.png") Then
        FSO.CopyFile TempDir & "\mesh_model.png", ModeADir & "\mesh_model.png", True
        FSO.CopyFile TempDir & "\mesh_model.png", ModeBDir & "\mesh_model.png", True
    End If
    Call LogMsg("网格模型图已导出: mesh_model.png")

' 5. 材料属性提取 (Slide 3: 基本信息/推荐工艺/粘度/PVT) — 取"本方案实际使用材料"
' 实机取证 (2026-09-08, study 097_1): GetFirstProperty(21000) 返回属性表第一项
' "通用 PP : 通用默认" (ID=1, 未用残留), 而分析实际用材是 "Novodur HH-106 :
' INEOS Styrolution" (ID=2, 求解日志 097_1_____~1.out 明确记载) — 旧代码永远取错材料
' (材料页数据缺失 + 曲线是通用 PP 的)。
' 锁定算法 (三级): 1) 求解日志精确匹配材料全名; 2) 日志不可用 → 属性表中非"通用默认"
' 的最后一项 (用户换料后新属性追加在表尾); 3) 仅通用默认 → 用它 (方案确实用默认料)。
' MaterialSelector 悬挂调用仍禁用 (2026-09-07 实机取证挂起 16 分钟)。
Dim PropEd, Prop
Dim MatDBCode, MatUseIdx, MatChosenName, MatCount, Mi, DbTry, PropAny
Dim MatIds(30), MatNames(30)
MatDBCode = 0 : MatUseIdx = 0 : MatChosenName = "" : MatCount = 0
Set PropEd = Synergy.PropertyEditor()
Set Prop = Nothing

' 5.0 枚举属性表: 热塑性(21000) 优先, 空则热固性(20030)
For Each DbTry In Array(21000, 20030)
    On Error Resume Next
    Set PropAny = PropEd.GetFirstProperty(DbTry)
    If Err.Number <> 0 Then
        Err.Clear
        Set PropAny = Nothing
    End If
    Do While Not PropAny Is Nothing And MatCount < 30
        MatCount = MatCount + 1
        MatIds(MatCount) = PropAny.ID
        MatNames(MatCount) = CStr(PropAny.Name)
        Set PropAny = PropEd.GetNextPropertyOfType(PropAny)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
    Loop
    On Error GoTo 0
    If MatCount > 0 Then
        MatDBCode = DbTry
        Exit For
    End If
Next

If MatCount = 0 Then
    Call LogMsg("ERROR: 材料属性表为空 (21000/20030 均无属性), 材料页将标注数据缺失")
Else
    ' 5.1 求解日志精确匹配 (最新修改时间的 out 日志优先)
    Dim LogMatchIdx
    LogMatchIdx = FindMaterialInSolverLogs()
    If LogMatchIdx > 0 Then
        MatUseIdx = MatIds(LogMatchIdx)
        MatChosenName = MatNames(LogMatchIdx)
        Call LogMsg("材料锁定(求解日志匹配): [" & MatChosenName & "] ID=" & MatUseIdx & ", 库=" & MatDBCode)
    Else
        ' 5.2 启发式: 非"通用默认"的最后一项; 全为通用默认则用第一项
        Dim HeurIdx
        HeurIdx = 0
        For Mi = 1 To MatCount
            If Not IsGenericDefaultName(MatNames(Mi)) Then HeurIdx = Mi
        Next
        If HeurIdx = 0 Then HeurIdx = 1
        MatUseIdx = MatIds(HeurIdx)
        MatChosenName = MatNames(HeurIdx)
        Call LogMsg("材料锁定(启发式, 求解日志未匹配): [" & MatChosenName & "] ID=" & MatUseIdx & ", 库=" & MatDBCode)
    End If
    For Mi = 1 To MatCount
        Dim CandMark
        CandMark = ""
        If MatIds(Mi) = MatUseIdx Then CandMark = "  <== 选用"
        Call LogMsg("材料属性表候选 #" & Mi & ": ID=" & MatIds(Mi) & " [" & MatNames(Mi) & "]" & CandMark)
    Next
    On Error Resume Next
    Set Prop = PropEd.FindProperty(MatDBCode, MatUseIdx)
    If Err.Number <> 0 Then
        Err.Clear
        Set Prop = Nothing
    End If
    On Error GoTo 0
    If Prop Is Nothing Then
        Call LogMsg("ERROR: 选定材料属性读取失败 (FindProperty(" & MatDBCode & "," & MatUseIdx & ")), 材料页将标注数据缺失")
    End If
End If

' 5.1 官方字段枚举原始导出 (material_fields.json)
Dim MatPropName, FieldIdRaw, FieldDesc, FieldValsRaw, Vi, Vn, ValStr, FieldsJson, FieldGuard, PropIdStr
PropIdStr = ""
MatPropName = ""
FieldsJson = ""
If Not Prop Is Nothing Then
    On Error Resume Next
    MatPropName = CStr(Prop.Name)
    PropIdStr = CStr(Prop.ID)
    FieldGuard = 0
    FieldIdRaw = 0
    FieldIdRaw = Prop.GetFirstField()
    Do While FieldIdRaw > 0 And FieldGuard < 500
        FieldGuard = FieldGuard + 1
        FieldDesc = ""
        FieldDesc = CStr(Prop.GetFieldDescription(FieldIdRaw))
        ValStr = ""
        Set FieldValsRaw = Prop.FieldValues(FieldIdRaw)
        If Not FieldValsRaw Is Nothing Then
            Vn = FieldValsRaw.Size()
            For Vi = 0 To Vn - 1
                If Vi > 0 Then ValStr = ValStr & "|"
                ValStr = ValStr & CStr(FieldValsRaw.Val(Vi))
            Next
        End If
        If Len(FieldsJson) > 0 Then FieldsJson = FieldsJson & "," & vbCrLf
        FieldsJson = FieldsJson & "    {""id"": " & CStr(FieldIdRaw) & ", ""desc"": """ & EscapeJson(FieldDesc) & """, ""values"": """ & EscapeJson(ValStr) & """}"
        FieldIdRaw = Prop.GetNextField(FieldIdRaw)
    Loop
    If Err.Number <> 0 Then
        Call LogMsg("WARN: 材料字段枚举异常: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
    Call LogMsg("材料属性字段枚举完成: " & FieldGuard & " 个字段 (prop_name=" & MatPropName & ")")
End If

Dim MatTrade, MatManuf, ColonPos, MatCandidates
MatTrade = "" : MatManuf = ""
ColonPos = InStrRev(MatPropName, " : ")
If ColonPos > 1 Then
    MatTrade = Trim(Left(MatPropName, ColonPos - 1))
    MatManuf = Trim(Mid(MatPropName, ColonPos + 3))
End If
MatCandidates = ""
For Mi = 1 To MatCount
    If Len(MatCandidates) > 0 Then MatCandidates = MatCandidates & "; "
    MatCandidates = MatCandidates & MatIds(Mi) & ":" & MatNames(Mi)
Next
Dim MatFieldsJson
MatFieldsJson = "{" & vbCrLf & _
                "  ""prop_name"": """ & EscapeJson(MatPropName) & """," & vbCrLf & _
                "  ""material_name"": """ & EscapeJson(MatChosenName) & """," & vbCrLf & _
                "  ""trade_name"": """ & EscapeJson(MatTrade) & """," & vbCrLf & _
                "  ""manufacturer"": """ & EscapeJson(MatManuf) & """," & vbCrLf & _
                "  ""material_id"": """ & EscapeJson(CStr(MatUseIdx)) & """," & vbCrLf & _
                "  ""prop_type"": " & CStr(MatDBCode) & "," & vbCrLf & _
                "  ""prop_id"": """ & EscapeJson(PropIdStr) & """," & vbCrLf & _
                "  ""candidates"": """ & EscapeJson(MatCandidates) & """," & vbCrLf & _
                "  ""fields"": [" & vbCrLf & FieldsJson & vbCrLf & _
                "  ]" & vbCrLf & _
                "}"
WriteUtf8TextFile TempDir & "\material_fields.json", MatFieldsJson
Call LogMsg("材料字段原始数据已导出: material_fields.json (材料=[" & MatChosenName & "], ID=" & MatUseIdx & ", 库=" & CStr(MatDBCode) & ")")

' 5.3 材料曲线图: CreateMaterialPlot(库代码, 属性ID, 曲线字段ID)
'     实机取证 (2026-09-08): 第二参 = 属性 ID — 传 1 得通用 PP 曲线 (39946B),
'     传 2 得 Novodur 曲线 (39510B); 传 0 静默无产出 (2023 实机) — 禁用 0。
If MatDBCode > 0 And MatUseIdx > 0 Then
    If ExportMaterialPlotSafe(MatDBCode, MatUseIdx, 1310, TempDir & "\material_viscosity.png") Then
        Call LogMsg("粘度曲线已导出: material_viscosity.png (材料: " & MatChosenName & ")")
    Else
        Call LogMsg("ERROR: 粘度曲线导出失败 (所有候选 ID 均未产出文件)")
    End If
    If ExportMaterialPlotSafe(MatDBCode, MatUseIdx, 1004, TempDir & "\material_pvt.png") Then
        Call LogMsg("PVT 曲线已导出: material_pvt.png (材料: " & MatChosenName & ")")
    Else
        Call LogMsg("ERROR: PVT 曲线导出失败 (所有候选 ID 均未产出文件)")
    End If
Else
    Call LogMsg("WARN: 未找到当前方案材料属性, 跳过粘度/PVT 曲线导出")
End If
' 6. 循环提取每个结果项的图
Set PlotsArray = ConfigObj.plots
ExportCount = 0
Dim PeakClampVal, PeakClampTime, PeakClampOK
Dim PeakInjVal, PeakInjTime, PeakInjOK
PeakClampOK = False : PeakInjOK = False
PeakClampVal = Null : PeakClampTime = Null
PeakInjVal = Null : PeakInjTime = Null

For i = 0 To PlotsArray.length - 1
    Set pObj = HTML.parentWindow.getArrayItem(PlotsArray, i)
    pKey = CStr(pObj.key)
    pName = CStr(pObj.plot_name)
    pType = CStr(pObj.type)
    pEnabled = CBool(pObj.enabled)

    If pEnabled Then
        Set PlotObj = FindPlotRobust(PlotManager, pName)
        If Not PlotObj Is Nothing Then
            Viewer.ShowPlot PlotObj
                ' 强制重绘: 实机取证 SaveImage/SavePlotScaleImage 存在陈旧帧 (ShowPlot 后视口未重绘,
                ' 抓到上一张图的画面, 数据条串台) — 与 GIF 导出路径的 Regenerate 同源
                On Error Resume Next
                PlotObj.Regenerate
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0

                ' 防陈旧帧 (官方导出模式, 2026-09-08 实机验证可用):
                ' 跳到最后一帧并等待视口重绘, 确保截到本张结果的最终时刻画面
                Dim FrameTotal
                FrameTotal = 0
                On Error Resume Next
                FrameTotal = PlotObj.GetNumberOfFrames()
                If Err.Number <> 0 Then
                    Err.Clear
                    FrameTotal = 0
                End If
                If FrameTotal > 1 Then
                    Viewer.ShowPlotFrame PlotObj, FrameTotal - 1
                    If Err.Number <> 0 Then Err.Clear
                End If
                On Error GoTo 0
                Call SleepSec(1)

            Dim hasCustomRot, orig_rx, orig_ry, orig_rz
            hasCustomRot = False
            If pKey = "warpage_z" Then
                hasCustomRot = True
                On Error Resume Next
                orig_rx = Viewer.GetRotationX
                orig_ry = Viewer.GetRotationY
                orig_rz = Viewer.GetRotationZ
                Call Viewer.Rotate(-90, 0, 0)
                Viewer.Fit
                Call LogMsg("[视角调整] warpage_z 已自动旋转为垂直状态 (-90, 0, 0)")
                On Error GoTo 0
            End If
            
            ' 切勿调用 Viewer.Fit()! 100% 会导致模型与左侧色带标尺挤压重叠!

            If pType = "gif" Then
                ' 充填动画导出 (帧数由配置指定)
                Call LogMsg("导出充填动画 (帧数: " & NFrames & ")...")
                If PlotObj.GetAnimationType() = 0 Then
                    PlotObj.SetNumberOfAnimationFrames NFrames
                Else
                    PlotObj.SetNumberOfFrames NFrames
                End If
                PlotObj.Regenerate
                Call SleepSec(1)
                
                GifPath = TempDir & "\" & pKey & ".gif"
                Viewer.SaveAnimation GifPath
                If FSO.FileExists(GifPath) Then
                    FSO.CopyFile GifPath, ModeADir & "\" & pKey & ".gif", True
                    FSO.CopyFile GifPath, ModeBDir & "\" & pKey & ".gif", True
                End If
                Call LogMsg("充填动画已导出: " & pKey & ".gif")
                ExportCount = ExportCount + 1
            Else
                ' 普通结果图导出 (支持 1080P/2K)
                ImgPath = TempDir & "\" & pKey & ".png"
                ' ---------------- 方案 A: 视口真实抓取 (含完整彩条标尺) ----------------
                If ScreenshotMode = "A" Or ScreenshotMode = "B" Or ScreenshotMode = "BOTH" Or ScreenshotMode = "ALL" Then
                    On Error Resume Next
                    Viewer.SaveImage ModeADir & "\" & pKey & ".png"
                    If Err.Number <> 0 Then
                        Call LogMsg("WARN: 方案 A 视口截图异常: " & Err.Description)
                        Err.Clear
                    Else
                        Call LogMsg("[方案 A] 视口截图完成: " & pName & " -> mode_a\" & pKey & ".png")
                    End If
                    On Error GoTo 0
                End If

                ' ---------------- 方案 B: 标尺与模型独立导出 (智能拼合) ----------------
                If ScreenshotMode = "B" Or ScreenshotMode = "BOTH" Or ScreenshotMode = "ALL" Then
                    On Error Resume Next
                    If InStr(pKey, "_xy") = 0 Then Viewer.SavePlotScaleImage ModeBDir & "\scale_" & pKey & ".png"
                    Viewer.SaveImage3 ModeBDir & "\model_" & pKey & ".png", ImageWidth, ImageHeight, True, False, False, False, False, False, False, False, False
                    If Err.Number <> 0 Then
                        Call LogMsg("WARN: 方案 B 独立导出异常: " & Err.Description)
                        Err.Clear
                    Else
                        Call LogMsg("[方案 B] 分别导出完成: " & pName & " -> mode_b\scale_" & pKey & ".png & model_" & pKey & ".png")
                    End If
                    On Error GoTo 0
                End If

                ImgPath = TempDir & "\" & pKey & ".png"
                Viewer.SaveImage ImgPath
                Call LogMsg("结果图已提取: " & pName & " -> " & pKey & ".png")
                ExportCount = ExportCount + 1
            
            If hasCustomRot Then
                On Error Resume Next
                Call Viewer.Rotate(orig_rx, orig_ry, orig_rz)
                Viewer.Fit
                Call LogMsg("[视角恢复] 已恢复原有视角")
                On Error GoTo 0
            End If

            If pKey = "clamp_force_xy" Or pKey = "inj_pressure_xy" Then
                ' 官方曲线数据导出 (Plot.SaveXYPlotCurveData): 值+时刻由 Python 从 txt 解析
                On Error Resume Next
                Call PlotObj.SaveXYPlotCurveData(TempDir & "\" & pKey & "_curve_data.txt")
                If Err.Number <> 0 Then
                    Call LogMsg("WARN: XY 曲线数据导出失败 (" & pKey & "): " & Err.Description)
                    Err.Clear
                Else
                    Call LogMsg("XY 曲线数据已导出: " & pKey & "_curve_data.txt")
                End If
                On Error GoTo 0
            End If

            If pKey = "clamp_force_xy" Then
                ' Plot.GetDepValues/GetIndpValues 实机已证伪 — 值走 GetMaxValue, 时刻走曲线 txt
                On Error Resume Next
                Dim maxClampVal : maxClampVal = PlotObj.GetMaxValue
                If Err.Number = 0 Then
                    If IsNumeric(maxClampVal) Then
                        If CDbl(maxClampVal) > 0 Then
                            PeakClampVal = Round(CDbl(maxClampVal), 1)
                            PeakClampOK = True
                        End If
                    End If
                End If
                Err.Clear
                On Error GoTo 0
                Dim clampInfoText
                clampInfoText = "{" & vbCrLf & _
                                "  ""cae_max_clamp_force"": " & JsonNum(PeakClampVal) & "," & vbCrLf & _
                                "  ""peak_time"": " & JsonNum(PeakClampTime) & "," & vbCrLf & _
                                "  ""data_complete"": " & LCase(CStr(PeakClampOK)) & vbCrLf & _
                                "}"
                WriteUtf8TextFile TempDir & "\clamp_force_info.json", clampInfoText
                Call LogMsg("锁模力峰值信息已输出: clamp_force_info.json (CAE最大锁模力: " & JsonNum(PeakClampVal) & " T)")
            End If

            If pKey = "inj_pressure_xy" Then
                On Error Resume Next
                Dim maxInjVal : maxInjVal = PlotObj.GetMaxValue
                If Err.Number = 0 Then
                    If IsNumeric(maxInjVal) Then
                        If CDbl(maxInjVal) > 0 Then
                            PeakInjVal = Round(CDbl(maxInjVal), 3)
                            PeakInjOK = True
                        End If
                    End If
                End If
                Err.Clear
                On Error GoTo 0
                Call LogMsg("注射压力峰值提取完成 (GetMaxValue): value=" & JsonNum(PeakInjVal) & ", time=" & JsonNum(PeakInjTime))
            End If
End If
        Else
            Call LogMsg("未启用 (跳过): " & pName)
        End If
    End If
Next

Call LogMsg("共提取结果图 " & ExportCount & " 张")

' 6.9 XY 探针峰值数据导出 (GetMaxValue 值 + 曲线 txt; null-on-error, 绝不写 0 充数)
Dim PeakJson
PeakJson = "{" & vbCrLf & _
           "  ""clamp_force"": {""peak_value"": " & JsonNum(PeakClampVal) & ", ""peak_time"": " & JsonNum(PeakClampTime) & "}," & vbCrLf & _
           "  ""inj_pressure"": {""peak_value"": " & JsonNum(PeakInjVal) & ", ""peak_time"": " & JsonNum(PeakInjTime) & "}" & vbCrLf & _
           "}"
WriteUtf8TextFile TempDir & "\peak_values.json", PeakJson
Call LogMsg("XY 峰值数据已导出: peak_values.json (clamp_force=" & JsonNum(PeakClampVal) & ", inj_pressure=" & JsonNum(PeakInjVal) & ")")

' 7. 生成 manifest.json 元数据
TodayStr = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)
ManifestText = "{" & vbCrLf & _
               "  ""study_name"": """ & EscapeJson(StudyName) & """," & vbCrLf & _
               "  ""part_name"": """ & EscapeJson(PartName) & """," & vbCrLf & _
               "  ""mesh_type"": """ & DetectedMeshType & """," & vbCrLf & _
               "  ""moldflow_version"": ""2023""," & vbCrLf & _
               "  ""com_method"": """ & ComMethod & """," & vbCrLf & _
               "  ""date"": """ & TodayStr & """" & vbCrLf & _
               "}"
WriteUtf8TextFile TempDir & "\manifest.json", ManifestText

' 8. 调用 Python 后台生成 PPT (窗口样式 0 = 隐藏窗口)
PyCmd = "python """ & BaseDir & "\core\pptx_builder.py"" --config """ & ConfigPath & """ --data-dir """ & TempDir & """ --log-file """ & TempDir & "\run.log"""
Call LogMsg(" PPT : " & PyCmd)

WshShell.Environment("PROCESS")("PYTHONIOENCODING") = "gbk"
ret = WshShell.Run(PyCmd, 0, True)
Call LogMsg("PPT : " & ret)

Dim LastOutPath, OutMsg
LastOutPath = ""
If FSO.FileExists(TempDir & "\last_output_path.txt") Then
    LastOutPath = ReadUtf8TextFile(TempDir & "\last_output_path.txt")
End If

If ret = 0 Then
    Call LogMsg("=== Python 执行完毕，PPT 生成 ===")
    OutMsg = "模流分析 PPT 全部生成完毕" & vbCrLf & vbCrLf & _
             "方案: " & StudyName & vbCrLf & _
             "网格类型: " & DetectedMeshType & vbCrLf & _
             "报告路径: " & vbCrLf & LastOutPath & vbCrLf & vbCrLf & _
             "即将为您打开 PowerPoint 预览"
    MsgBox OutMsg, 64, "模流分析报告"
Else
    Call LogMsg("ERROR: PPT 生成失败: " & ret)
MsgBox "PPT 生成失败。" & vbCrLf & _
             "退出状态: " & ret & " (1=环境问题如模板缺失/Python缺失, 2=构建失败, 其他=未知)" & vbCrLf & _
             "怎么办: 查看日志尾部定位原因" & vbCrLf & _
             "日志: " & TempDir & "\run.log", 48, "提示"
End If

' ==============================================================================
' 辅助函数
' ==============================================================================

Function FindPlotRobust(PlotMgr, MainName)
    ' 只认精确名 (2026-09-08 用户裁决: 别名/模糊匹配/数据集ID猜测彻底移除 —
    ' 曾致"顶出时的体积收缩率"抢匹配真结果"体积收缩率")。
    ' 名字必须与方案中结果名一字不差 (配置界面可从动态结果清单照抄)。
    ' 找不到 → 返回 Nothing, 调用方按缺失跳过并记日志, 绝不猜。
    Set FindPlotRobust = Nothing
    On Error Resume Next
    Set FindPlotRobust = PlotMgr.FindPlotByName(MainName)
    If Err.Number <> 0 Then
        Err.Clear
        Set FindPlotRobust = Nothing
    End If
    On Error GoTo 0
End Function

Function SafeGetLong(obj, propName, defaultVal)
    On Error Resume Next
    Dim val
    val = Eval("obj." & propName)
    If Err.Number <> 0 Or IsEmpty(val) Then
        SafeGetLong = defaultVal
    Else
        SafeGetLong = CLng(val)
    End If
    On Error GoTo 0
End Function

Function SafeGetDbl(obj, propName, defaultVal)
    On Error Resume Next
    Dim val
    val = Eval("obj." & propName)
    If Err.Number <> 0 Or IsEmpty(val) Then
        SafeGetDbl = defaultVal
    Else
        SafeGetDbl = CDbl(val)
    End If
    On Error GoTo 0
End Function

Sub DeleteIfPresent(dirPath, fileName)
    On Error Resume Next
    Dim p
    p = dirPath & "\" & fileName
    If FSO.FileExists(p) Then FSO.DeleteFile p, True
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Function JsonNum(v)
    ' 数值→JSON: Null/Empty/非数值 → null (绝不写 0 充数)
    If IsNull(v) Or IsEmpty(v) Or Not IsNumeric(v) Then
        JsonNum = "null"
    Else
        JsonNum = CStr(v)
    End If
End Function


Function EscapeJson(strText)
    Dim s
    s = Replace(strText, "\", "\\")
    s = Replace(s, """", "\""")
    EscapeJson = s
End Function

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

Sub WriteUtf8TextFile(Path, Text)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText Text
    stream.SaveToFile Path, 2 ' 2 = overwrite
    stream.Close
End Sub

Sub LogMsg(msg)
    On Error Resume Next
    Dim f
    Set f = FSO.OpenTextFile(TempDir & "\run.log", 8, True)
    f.WriteLine Now & " - " & msg
    f.Close
    On Error GoTo 0
End Sub

Sub SleepSec(n)
    ' 宏宿主 (Moldflow 内运行) 无 WScript.Sleep: 借外部 ping 进程等待, 期间消息泵不阻塞
    On Error Resume Next
    WshShell.Run "cmd.exe /c ping -n " & (CLng(n) + 1) & " 127.0.0.1 >nul 2>&1", 0, True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Function ReadFileLatin1(path)
    ' 按字节原样读为 latin1 字符串 (保持字节一一对应, 供 GBK 字节级匹配)
    Dim stream
    ReadFileLatin1 = ""
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "iso-8859-1"
    If Err.Number = 0 Then ReadFileLatin1 = stream.ReadText
    stream.Close
    If Err.Number <> 0 Then
        Err.Clear
        ReadFileLatin1 = ""
    End If
    On Error GoTo 0
End Function

Function GbkToLatin1(s)
    ' 字符串 → GBK 字节 → latin1 字符串 (与 ReadFileLatin1 的日志字节做 InStr 匹配)
    Dim stream, bytes
    GbkToLatin1 = ""
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "gbk"
    stream.Open
    stream.WriteText s
    stream.Position = 0
    stream.Type = 1
    bytes = stream.Read
    stream.Close
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "iso-8859-1"
    If Err.Number = 0 Then GbkToLatin1 = stream.ReadText
    stream.Close
    If Err.Number <> 0 Then
        Err.Clear
        GbkToLatin1 = ""
    End If
    On Error GoTo 0
End Function

Function IsGenericDefaultName(nm)
    ' Moldflow 自带通用默认材料记录标识 (中/英文界面)
    Dim u
    u = UCase(CStr(nm & ""))
    IsGenericDefaultName = (InStr(CStr(nm), "通用默认") > 0 Or InStr(u, "GENERIC DEFAULT") > 0 Or InStr(u, "GENERIC PP") > 0)
End Function

Function FindMaterialInSolverLogs()
    ' 在工程目录 <方案基名>~*.out 求解日志中按字节精确匹配材料属性全名,
    ' 返回命中的候选下标 (1..MatCount), 未命中 0。多个日志命中时取修改时间最新的。
    ' 实机依据: 097_1_____~1.out 内有整行 "Novodur HH-106 : INEOS Styrolution"。
    Dim Proj, projDir, baseName, folder, f, logText, nameLatin, Mi, bestIdx, bestDate, bestFile
    FindMaterialInSolverLogs = 0
    bestIdx = 0 : bestDate = 0 : bestFile = ""
    On Error Resume Next
    Set Proj = Synergy.Project()
    If Err.Number <> 0 Or Proj Is Nothing Then
        Err.Clear
        Exit Function
    End If
    projDir = CStr(Proj.Path)
    On Error GoTo 0
    If Len(projDir) = 0 Or Not FSO.FolderExists(projDir) Then Exit Function
    baseName = StudyName
    If LCase(Right(baseName, 4)) = ".sdy" Then baseName = Left(baseName, Len(baseName) - 4)
    If Len(baseName) = 0 Then Exit Function
    On Error Resume Next
    Set folder = FSO.GetFolder(projDir)
    If Err.Number <> 0 Or folder Is Nothing Then
        Err.Clear
        Exit Function
    End If
    For Each f In folder.Files
        If LCase(Right(f.Name, 4)) = ".out" And InStr(f.Name, "~") > 0 And InStr(f.Name, baseName) > 0 Then
            logText = ReadFileLatin1(f.Path)
            If Len(logText) > 0 Then
                For Mi = 1 To MatCount
                    nameLatin = GbkToLatin1(MatNames(Mi))
                    If Len(nameLatin) > 0 And InStr(logText, nameLatin) > 0 Then
                        If f.DateLastModified >= bestDate Then
                            bestIdx = Mi
                            bestDate = f.DateLastModified
                            bestFile = f.Name
                        End If
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    If Err.Number <> 0 Then
        Err.Clear
        bestIdx = 0
    End If
    On Error GoTo 0
    If bestIdx > 0 Then
        Call LogMsg("求解日志材料匹配: [" & MatNames(bestIdx) & "] <- " & bestFile)
    End If
    FindMaterialInSolverLogs = bestIdx
End Function

Function ExportMaterialPlotSafe(dbCode, idx, fieldId, outPath)
    ' 材料曲线导出: 首选 idx (锁定的真材料 ID), 失败重试一次后回退 ID=1 (保底出图)。
    ' 返回 True = 文件已产出。0 永不使用 (2023 实机静默无产出)。
    Dim MP, attemptIdx, ok
    ExportMaterialPlotSafe = False
    For Each attemptIdx In Array(CLng(idx), 1)
        If attemptIdx > 0 Then
            ok = False
            On Error Resume Next
            Set MP = PlotManager.CreateMaterialPlot(CLng(dbCode), CLng(attemptIdx), CLng(fieldId))
            If Err.Number <> 0 Then
                Call LogMsg("WARN: CreateMaterialPlot(" & dbCode & "," & attemptIdx & "," & fieldId & ") 异常: " & Err.Description)
                Err.Clear
                Set MP = Nothing
            End If
            If Not MP Is Nothing Then
                MP.SaveImage outPath
                If Err.Number <> 0 Then Err.Clear
                If Not FSO.FileExists(outPath) Then
                    Call SleepSec(1)
                    MP.SaveImage outPath
                    If Err.Number <> 0 Then Err.Clear
                End If
                If FSO.FileExists(outPath) Then ok = True
            End If
            On Error GoTo 0
            If ok Then
                ExportMaterialPlotSafe = True
                If CLng(attemptIdx) <> CLng(idx) Then
                    Call LogMsg("WARN: 材料曲线首选 ID=" & idx & " 导出失败, 已回退 ID=1 (通用默认) — 曲线可能与实际用材不符, 请人工复核")
                End If
                Exit Function
            End If
        End If
    Next
End Function
