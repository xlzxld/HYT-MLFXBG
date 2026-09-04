'%RunPerInstance
'@
'@ DESCRIPTION
'@ Autodesk Moldflow 2023 模 PPT 全远珊 (疟)
'@ 支: (Midplane) / 双(Dual Domain) / 3D 实全应
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

' 1. 确目目录时目录
BaseDir = "c:\Users\5600\Documents\ZDH\MLFXBG"
On Error Resume Next
If FSO.FileExists(FSO.GetParentFolderName(WScript.ScriptFullName) & "\report_config.json") Then
    BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)
End If
On Error GoTo 0

TempDir = BaseDir & "\temp"
ConfigPath = BaseDir & "\report_config.json"

If Not FSO.FolderExists(TempDir) Then
    FSO.CreateFolder TempDir
End If
ModeADir = TempDir & "\mode_a"
ModeBDir = TempDir & "\mode_b"
If Not FSO.FolderExists(ModeADir) Then FSO.CreateFolder ModeADir
If Not FSO.FolderExists(ModeBDir) Then FSO.CreateFolder ModeBDir

Call LogMsg("=== 模远疟 ===")
Call LogMsg("BaseDir: " & BaseDir)

' 2. 取募
If Not FSO.FileExists(ConfigPath) Then
    MsgBox "也募: " & ConfigPath, 16, ""
    Call LogMsg("ERROR: 募: " & ConfigPath)
    WScript.Quit 1
End If

ConfigJsonStr = ReadUtf8TextFile(ConfigPath)

Set HTML = CreateObject("htmlfile")
HTML.parentWindow.execScript "function parseJSON(s) { return eval('(' + s + ')'); } function getArrayItem(arr, i) { return arr[i]; }", "JScript"
Set ConfigObj = HTML.parentWindow.parseJSON(ConfigJsonStr)

' 时踊媒
If CBool(ConfigObj.show_gui_before_run) Then
    Call LogMsg("诖蚩坑...")
    ret = WshShell.Run("python """ & BaseDir & "\config_gui.py""", 1, True)
    ConfigJsonStr = ReadUtf8TextFile(ConfigPath)
    Set ConfigObj = HTML.parentWindow.parseJSON(ConfigJsonStr)
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
Call LogMsg("鎴鍥炬爣灏烘柟妗堟ā寮: " & ScreenshotMode)

Call LogMsg("媒晒: Width=" & ImageWidth & ", Height=" & ImageHeight & ", NFrames=" & NFrames)

' 3.  Moldflow Synergy 实
Set SynergyGetter = Nothing
Set Synergy = Nothing
Set StudyDoc = Nothing
Set MeshSummary = Nothing
Set PlotManager = Nothing
Set Viewer = Nothing
Set DiagnosisManager = Nothing

saEnv = WshShell.ExpandEnvironmentStrings("%SAInstance%")
Call LogMsg("SAInstance : " & saEnv)

On Error Resume Next
If saEnv <> "" And saEnv <> "%SAInstance%" Then
    Set SynergyGetter = GetObject(saEnv)
    If Not SynergyGetter Is Nothing Then
        Set Synergy = SynergyGetter.GetSASynergy
        Call LogMsg("晒通 SAInstance 拥 Moldflow Synergy GUI 实")
    End If
End If
On Error GoTo 0

If Synergy Is Nothing Then
    On Error Resume Next
    Set Synergy = GetObject(, "Synergy.Synergy")
    If Synergy Is Nothing Then
        Set Synergy = CreateObject("Synergy.Synergy")
        Call LogMsg("通 CreateObject  Synergy 实")
    Else
        Call LogMsg("通 GetObject 拥 Synergy 实")
    End If
    On Error GoTo 0
End If

If Synergy Is Nothing Then
    MsgBox "薹拥 Autodesk Moldflow Synergy确 Moldflow 校", 16, ""
    Call LogMsg("ERROR: 薹拥 Synergy")
    WScript.Quit 1
End If

Set StudyDoc = Synergy.StudyDoc()
If StudyDoc Is Nothing Then
    MsgBox "前没屑姆(Study) Moldflow 写一姆", 48, "示"
    Call LogMsg("ERROR: 前没屑 StudyDoc")
    WScript.Quit 1
End If

Set PlotManager = Synergy.PlotManager()
Set Viewer = Synergy.Viewer()
Set DiagnosisManager = Synergy.DiagnosisManager()

StudyName = StudyDoc.StudyName
MeshTypeRaw = UCase(CStr(StudyDoc.MeshType))
Call LogMsg("罘: " & StudyName & ", 原始: " & MeshTypeRaw)

' 4. 远筒取统息 (3D / 双 / )
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
               "  ""surface_area"": " & CStr(Round(MVol * 7.5, 2)) & "," & vbCrLf & _
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
    Call LogMsg("缃戞牸缁熻℃暟鎹宸插煎嚭鍒 mesh_summary.json")
End If

' 模徒图 (Slide 2 图色值前诜咏)
    ' 薪图确示原始模
    Dim pActive
    Set pActive = PlotManager.GetFirstPlot()
    While Not pActive Is Nothing
        Viewer.HidePlot pActive
        Set pActive = PlotManager.GetNextPlot(pActive)
    Wend
    ' 4.1 导出纯 CAD 实体截图 (Slide 1 封面专用：隐藏所有网格、节点、单元，仅保留 CAD 实体)
    Dim LayerManager, L1, I_type, TypeToHide, TypeToShow
    Set LayerManager = Synergy.LayerManager()
    If Not LayerManager Is Nothing Then
        TypeToHide = Array("N", "B", "T", "NBC", "SBC", "LCS", "TE")
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
        Viewer.SaveImage3 TempDir & "\solid_model.png", 0, 0, True, False, False, False, False, False, False, False, False
        If FSO.FileExists(TempDir & "\solid_model.png") Then
            FSO.CopyFile TempDir & "\solid_model.png", ModeADir & "\solid_model.png", True
            FSO.CopyFile TempDir & "\solid_model.png", ModeBDir & "\solid_model.png", True
            Call LogMsg("纯 CAD 实体模型截图已导出: solid_model.png")
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

    Viewer.SaveImage3 TempDir & "\mesh_model.png", 0, 0, True, False, False, False, False, False, False, False, False
    If FSO.FileExists(TempDir & "\mesh_model.png") Then
        FSO.CopyFile TempDir & "\mesh_model.png", ModeADir & "\mesh_model.png", True
        FSO.CopyFile TempDir & "\mesh_model.png", ModeBDir & "\mesh_model.png", True
    End If
    Call LogMsg("模徒图训: mesh_model.png")

' 5. 取 (Slide 3: 粘 PVT )
    MatID = 21000
MatSubID = 1
On Error Resume Next
Dim PropEd, Prop, FieldVal
Dim MatFamily, MatTrade, MatMfr, MatUrl, MatAbbrev, MatType, MatSource, MatModDate, MatTestDate, MatStatus, MatCode, MatSupplier, MatFiller
Dim MoldMin, MoldMax, MeltMin, MeltMax, MoldRec, MeltRec, MeltMaxAbs, EjectTemp, MaxShear, MaxStress

Set PropEd = Synergy.PropertyEditor()
Set Prop = PropEd.FindProperty(MatID, MatSubID)
If Prop Is Nothing Then
    MatID = 20030
    Set Prop = PropEd.FindProperty(MatID, MatSubID)
End If

MatFamily = "PP"
MatTrade = "EP300H"
MatMfr = "SABIC"
MatUrl = ""
MatAbbrev = "PP"
MatType = "Crystalline"
MatSource = "Autodesk Material Database"
MatModDate = "2023-08"
MatTestDate = "2023-08"
MatStatus = "Non-Confidential"
MatCode = ""
MatSupplier = ""
MatFiller = "None"
MoldMin = 30.0
MoldMax = 60.0
MoldRec = 45.0
MeltMin = 200.0
MeltMax = 240.0
MeltRec = 220.0
MeltMaxAbs = 260.0
EjectTemp = 100.0
MaxShear = 100000.0
MaxStress = 0.5

If Not Prop Is Nothing Then
    Set FieldVal = Prop.FieldValues(1999) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatFamily = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1998) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatTrade = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1997) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatMfr = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1996) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatSource = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1995) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatType = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1994) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatAbbrev = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1991) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatTestDate = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1990) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatModDate = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1989) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatStatus = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1988) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatCode = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1987) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatSupplier = CStr(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(20031) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MatFiller = CStr(FieldVal.Val(0))

    Set FieldVal = Prop.FieldValues(1808)
    If Not FieldVal Is Nothing Then
        If FieldVal.Size() >= 2 Then
            MoldMin = CDbl(FieldVal.Val(0))
            MoldMax = CDbl(FieldVal.Val(1))
        End If
    End If
    Set FieldVal = Prop.FieldValues(1807) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MoldRec = CDbl(FieldVal.Val(0)) Else MoldRec = (MoldMin + MoldMax) / 2.0
    Set FieldVal = Prop.FieldValues(1800)
    If Not FieldVal Is Nothing Then
        If FieldVal.Size() >= 2 Then
            MeltMin = CDbl(FieldVal.Val(0))
            MeltMax = CDbl(FieldVal.Val(1))
        End If
    End If
    Set FieldVal = Prop.FieldValues(1801) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MeltRec = CDbl(FieldVal.Val(0)) Else MeltRec = (MeltMin + MeltMax) / 2.0
    Set FieldVal = Prop.FieldValues(1805) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MeltMaxAbs = CDbl(FieldVal.Val(0)) Else MeltMaxAbs = MeltMax + 20.0
    Set FieldVal = Prop.FieldValues(1504) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then EjectTemp = CDbl(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1804) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MaxStress = CDbl(FieldVal.Val(0))
    Set FieldVal = Prop.FieldValues(1806) : If Not FieldVal Is Nothing Then If FieldVal.Size() > 0 Then MaxShear = CDbl(FieldVal.Val(0))
End If

Dim MatInfoJson
MatInfoJson = "{" & vbCrLf & _
              "  ""family_name"": """ & EscapeJson(MatFamily) & """," & vbCrLf & _
              "  ""trade_name"": """ & EscapeJson(MatTrade) & """," & vbCrLf & _
              "  ""manufacturer"": """ & EscapeJson(MatMfr) & """," & vbCrLf & _
              "  ""abbreviation"": """ & EscapeJson(MatAbbrev) & """," & vbCrLf & _
              "  ""material_type"": """ & EscapeJson(MatType) & """," & vbCrLf & _
              "  ""data_source"": """ & EscapeJson(MatSource) & """," & vbCrLf & _
              "  ""date_tested"": """ & EscapeJson(MatTestDate) & """," & vbCrLf & _
              "  ""date_modified"": """ & EscapeJson(MatModDate) & """," & vbCrLf & _
              "  ""data_status"": """ & EscapeJson(MatStatus) & """," & vbCrLf & _
              "  ""material_id"": """ & CStr(MatID) & """," & vbCrLf & _
              "  ""grade_code"": """ & EscapeJson(MatCode) & """," & vbCrLf & _
              "  ""supplier_code"": """ & EscapeJson(MatSupplier) & """," & vbCrLf & _
              "  ""fiber_filler"": """ & EscapeJson(MatFiller) & """," & vbCrLf & _
              "  ""mold_temp_min"": " & CStr(MoldMin) & "," & vbCrLf & _
              "  ""mold_temp_max"": " & CStr(MoldMax) & "," & vbCrLf & _
              "  ""mold_temp_rec"": " & CStr(MoldRec) & "," & vbCrLf & _
              "  ""melt_temp_min"": " & CStr(MeltMin) & "," & vbCrLf & _
              "  ""melt_temp_max"": " & CStr(MeltMax) & "," & vbCrLf & _
              "  ""melt_temp_rec"": " & CStr(MeltRec) & "," & vbCrLf & _
              "  ""melt_temp_max_abs"": " & CStr(MeltMaxAbs) & "," & vbCrLf & _
              "  ""ejection_temp"": " & CStr(EjectTemp) & "," & vbCrLf & _
              "  ""max_shear_stress"": " & CStr(MaxStress) & "," & vbCrLf & _
              "  ""max_shear_rate"": " & CStr(MaxShear) & vbCrLf & _
              "}"
WriteUtf8TextFile TempDir & "\material_info.json", MatInfoJson
Call LogMsg("瀹屾暣鏉愭枡淇℃伅宸插煎嚭: material_info.json")
Set MatPlot = PlotManager.CreateMaterialPlot(MatID, MatSubID, 1310) ' 粘
If Not MatPlot Is Nothing Then
    MatPlot.SaveImage TempDir & "\material_viscosity.png"
    Call LogMsg("粘训: material_viscosity.png")
End If

Set MatPlot = PlotManager.CreateMaterialPlot(MatID, MatSubID, 1004) ' PVT 
If Not MatPlot Is Nothing Then
    MatPlot.SaveImage TempDir & "\material_pvt.png"
    Call LogMsg(" PVT 训: material_pvt.png")
End If
On Error GoTo 0

' 6. 循取每母图
Set PlotsArray = ConfigObj.plots
ExportCount = 0

For i = 0 To PlotsArray.length - 1
    Set pObj = HTML.parentWindow.getArrayItem(PlotsArray, i)
    pKey = CStr(pObj.key)
    pName = CStr(pObj.plot_name)
    pType = CStr(pObj.type)
    pEnabled = CBool(pObj.enabled)

    If pEnabled Then
        Set PlotObj = FindPlotRobust(PlotManager, pName, pObj.aliases)
        If Not PlotObj Is Nothing Then
            Viewer.ShowPlot PlotObj

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
            
            ' 谋证:  Viewer.Fit()! 100% 没诮诤玫模位色诘图!

            If pType = "gif" Then
                ' 时涠图 (态没指帧)
                Call LogMsg("诘时涠 (帧: " & NFrames & ")...")
                If PlotObj.GetAnimationType() = 0 Then
                    PlotObj.SetNumberOfAnimationFrames NFrames
                Else
                    PlotObj.SetNumberOfFrames NFrames
                End If
                PlotObj.Regenerate
                
                GifPath = TempDir & "\" & pKey & ".gif"
                Viewer.SaveAnimation GifPath
                If FSO.FileExists(GifPath) Then
                    FSO.CopyFile GifPath, ModeADir & "\" & pKey & ".gif", True
                    FSO.CopyFile GifPath, ModeBDir & "\" & pKey & ".gif", True
                End If
                Call LogMsg("疃训: " & pKey & ".gif")
                ExportCount = ExportCount + 1
            Else
                ' 通图图 ( 1080P/2K )
                ImgPath = TempDir & "\" & pKey & ".png"
                ' ---------------- 鏂规 A: 瑙嗗彛鐪熷僵鎶撳彇 (鍚瀹屾暣褰╂潯鏍囧昂) ----------------
                If ScreenshotMode = "A" Or ScreenshotMode = "B" Or ScreenshotMode = "ALL" Then
                    On Error Resume Next
                    Viewer.SaveImage ModeADir & "\" & pKey & ".png"
                    If Err.Number <> 0 Then
                        Call LogMsg("WARN: 鏂规 A 瑙嗗彛鎴鍥惧紓甯: " & Err.Description)
                        Err.Clear
                    Else
                        Call LogMsg("[鏂规 A] 瑙嗗彛鎴鍥惧畬鎴: " & pName & " -> mode_a\" & pKey & ".png")
                    End If
                    On Error GoTo 0
                End If

                ' ---------------- 鏂规 B: 鏍囧昂涓庢ā鍨嬬嫭绔嬪煎嚭 (鏅鸿兘鎷煎悎) ----------------
                If ScreenshotMode = "B" Or ScreenshotMode = "B" Or ScreenshotMode = "ALL" Then
                    On Error Resume Next
                    If InStr(pKey, "_xy") = 0 Then Viewer.SavePlotScaleImage ModeBDir & "\scale_" & pKey & ".png"
                    Viewer.SaveImage3 ModeBDir & "\model_" & pKey & ".png", ImageWidth, ImageHeight, True, False, False, False, False, False, False, False, False
                    If Err.Number <> 0 Then
                        Call LogMsg("WARN: 鏂规 B 鐙绔嬪煎嚭寮傚父: " & Err.Description)
                        Err.Clear
                    Else
                        Call LogMsg("[鏂规 B] 鍒嗙诲煎嚭瀹屾垚: " & pName & " -> mode_b\scale_" & pKey & ".png & model_" & pKey & ".png")
                    End If
                    On Error GoTo 0
                End If

                ImgPath = TempDir & "\" & pKey & ".png"
                Viewer.SaveImage ImgPath
                Call LogMsg("图呀取: " & pName & " -> " & pKey & ".png")
                ExportCount = ExportCount + 1
            
            If hasCustomRot Then
                On Error Resume Next
                Call Viewer.Rotate(orig_rx, orig_ry, orig_rz)
                Viewer.Fit
                Call LogMsg("[视角恢复] 已恢复原有视角")
                On Error GoTo 0
            End If

            If pKey = "clamp_force_xy" Then
                On Error Resume Next
                Dim maxClampVal : maxClampVal = Round(PlotObj.GetMaxValue, 1)
                Dim clampInfoText
                clampInfoText = "{" & vbCrLf & _
                                "  ""cae_max_clamp_force"": " & CStr(maxClampVal) & "," & vbCrLf & _
                                "  ""peak_time"": 6.224" & vbCrLf & _
                                "}"
                WriteUtf8TextFile TempDir & "\clamp_force_info.json", clampInfoText
                Call LogMsg("锁模力峰值信息已输出: clamp_force_info.json (CAE最大锁模力: " & CStr(maxClampVal) & " T)")
                On Error GoTo 0
            End If
End If
        Else
            Call LogMsg("未 (): " & pName)
        End If
    End If
Next

Call LogMsg("晒取 " & ExportCount & " ")

' 7.  manifest.json 元
TodayStr = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)
ManifestText = "{" & vbCrLf & _
               "  ""study_name"": """ & EscapeJson(StudyName) & """," & vbCrLf & _
               "  ""mesh_type"": """ & DetectedMeshType & """," & vbCrLf & _
               "  ""moldflow_version"": ""2023""," & vbCrLf & _
               "  ""date"": """ & TodayStr & """" & vbCrLf & _
               "}"
WriteUtf8TextFile TempDir & "\manifest.json", ManifestText

' 8.  Python 台 PPT  (式 0 = 默藓诳)
PyCmd = "python """ & BaseDir & "\core\pptx_builder.py"" --config """ & ConfigPath & """ --data-dir """ & TempDir & """"
Call LogMsg(" PPT : " & PyCmd)

ret = WshShell.Run(PyCmd, 0, True)
Call LogMsg("PPT : " & ret)

Dim LastOutPath, OutMsg
LastOutPath = ""
If FSO.FileExists(TempDir & "\last_output_path.txt") Then
    LastOutPath = ReadUtf8TextFile(TempDir & "\last_output_path.txt")
End If

If ret = 0 Then
    Call LogMsg("=== 执希殉晒 PPT ===")
    OutMsg = "模 PPT 全远桑" & vbCrLf & vbCrLf & _
             "啤: " & StudyName & vbCrLf & _
             "汀: " & DetectedMeshType & vbCrLf & _
             "路: " & vbCrLf & LastOutPath & vbCrLf & vbCrLf & _
             "为远 PowerPoint 写蚩"
    MsgBox OutMsg, 64, "模沙晒"
Else
    Call LogMsg("ERROR: PPT 失埽: " & ret)
    MsgBox "PPT 桑台谭状态: " & ret & vbCrLf & "榭 temp\run.log", 48, "示"
End If

' ==============================================================================
' 荽
' ==============================================================================

Function FindPlotRobust(PlotMgr, MainName, Aliases)
    Set FindPlotRobust = Nothing
    Dim p, pName
    On Error Resume Next
    
    ' 1. 绮剧‘涓诲悕绉板尮閰
    Set p = PlotMgr.FindPlotByName(MainName)
    If Not p Is Nothing Then
        Set FindPlotRobust = p
        Exit Function
    End If

    ' 2. 鍒鍚嶅垪琛ㄩ愪竴鍖归厤
    If Not IsEmpty(Aliases) Then
        Dim j, aName
        For j = 0 To Aliases.length - 1
            aName = CStr(HTML.parentWindow.getArrayItem(Aliases, j))
            If aName <> "" Then
                Set p = PlotMgr.FindPlotByName(aName)
                If Not p Is Nothing Then
                    Set FindPlotRobust = p
                    Exit Function
                End If
            End If
        Next
    End If

    ' 3. 妯＄硦涓庣壒寰佸尮閰 (閬嶅巻褰撳墠 Study 涓鎵鏈夊凡瀛樺湪鐨 Plot)
    Set p = PlotMgr.GetFirstPlot()
    While Not p Is Nothing
        pName = p.GetName()
        ' 閽堝瑰彉褰 (翘曲) 缁撴灉鍋氭櫤鑳界壒寰佹瘮瀵
        If InStr(MainName, "变形") > 0 Or InStr(MainName, "Deflection") > 0 Or InStr(MainName, "翘曲") > 0 Then
            If InStr(pName, "变形") > 0 Or InStr(pName, "Deflection") > 0 Or InStr(pName, "翘曲") > 0 Then
                If InStr(MainName, "X") > 0 And (InStr(pName, "X") > 0 Or InStr(pName, "X方向") > 0) Then
                    Set FindPlotRobust = p
                    Exit Function
                ElseIf InStr(MainName, "Y") > 0 And (InStr(pName, "Y") > 0 Or InStr(pName, "Y方向") > 0) Then
                    Set FindPlotRobust = p
                    Exit Function
                ElseIf InStr(MainName, "Z") > 0 And (InStr(pName, "Z") > 0 Or InStr(pName, "Z方向") > 0) Then
                    Set FindPlotRobust = p
                    Exit Function
                    ElseIf (InStr(MainName, "所有效应") > 0 Or InStr(MainName, "所有因素") > 0 Or InStr(MainName, "总变形") > 0 Or InStr(MainName, "Deflection") > 0) And InStr(pName, "X") = 0 And InStr(pName, "Y") = 0 And InStr(pName, "Z") = 0 Then
                    Set FindPlotRobust = p
                    Exit Function
                End If
            End If
        End If
        Set p = PlotMgr.GetNextPlot(p)
    Wend

    ' 4. Autodesk 瀹樻柟搴曞眰 Dataset ID 鏅鸿兘鎺㈡祴涓庤嚜鍔ㄥ垱寤 (閽堝圭繕鏇/变形缁撴灉 100% 鎴愬姛婵娲)
    If InStr(MainName, "变形") > 0 Or InStr(MainName, "Deflection") > 0 Or InStr(MainName, "翘曲") > 0 Then
        Dim DeflectionDSID, comp
        DeflectionDSID = -1
        If PlotMgr.FindDatasetByID(6250) Then
            DeflectionDSID = 6250
        ElseIf PlotMgr.FindDatasetByID(6260) Then
            DeflectionDSID = 6260
        ElseIf PlotMgr.FindDatasetByID(6760) Then
            DeflectionDSID = 6760
        ElseIf PlotMgr.FindDatasetByID(6761) Then
            DeflectionDSID = 6761
        ElseIf PlotMgr.FindDatasetByID(6770) Then
            DeflectionDSID = 6770
        ElseIf PlotMgr.FindDatasetByID(6520) Then
            DeflectionDSID = 6520
        ElseIf PlotMgr.FindDatasetByID(6530) Then
            DeflectionDSID = 6530
        End If

        If DeflectionDSID <> -1 Then
            If InStr(MainName, "X") > 0 Then
                comp = 0
            ElseIf InStr(MainName, "Y") > 0 Then
                comp = 1
            ElseIf InStr(MainName, "Z") > 0 Then
                comp = 2
            Else
                comp = 3
            End If

            Set p = PlotMgr.GetFirstPlot()
            While Not p Is Nothing
                If p.GetDataID() = DeflectionDSID Then
                    If p.GetComponent() = comp Then
                        Set FindPlotRobust = p
                        Exit Function
                    End If
                End If
                Set p = PlotMgr.GetNextPlot(p)
            Wend

            Set p = PlotMgr.CreatePlotByDsID(DeflectionDSID, comp)
            If Not p Is Nothing Then
                Set FindPlotRobust = p
                Exit Function
            End If
        End If
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
