<!-- /autoplan restore point: C:/Users/5600/.gstack/projects/xlzxld-HYT-MLFXBG/master-autoplan-restore-20260906-024216.md -->

# MLFXBG 模流报告自动化 — 全方位优化计划

> 起草依据：2026-09-06 全量源码审查（config_gui.py / core/*.py / AutoReport.vbs / report_config.json / AI_GUIDE.md）。
> 用户诉求："项目勉强能用但问题多，全方位优化，包括但不限于用户交互体验差、截图不够清晰。"

## 一、现状问题清单（全部带代码证据）

### A. 数据正确性（最严重：报告里可能出现假数据/错数据）

| # | 问题 | 证据 |
|---|------|------|
| A1 | `mesh_summary.json` 缺失或损坏时，网格统计**静默使用硬编码的另一套假数据**（61978 三角形、94.7% 匹配率、9662.27 cm² 等Defaults），直接印进正式报告 | pptx_builder.py:236-258（`d.get("triangles", 61978)` 等 18 处默认值） |
| A2 | `material_info.json` 缺失时，材料卡片**静默渲染假材料**（"EP300H"/"SABIC"/2023-08），印进报告第 3 页 | pptx_builder.py:343-368（`d.get("trade_name", "EP300H")` 等 20+ 处） |
| A3 | XY 曲线探针标注数值**硬编码本项目当前方案的数据**（锁模力 320.5T/6.224s、压力 21.45MPa/5.092s），换方案后标注框显示错误数值；config 里的 `clamp_force_settings`/`inj_pressure_settings` 从未被 image_processor 读取 | image_processor.py:302-331（默认参数），merge_scale_and_model:368-371 无参调用 |
| A4 | VBS 方案 B 导出条件写成 `ScreenshotMode = "B" Or ScreenshotMode = "B"`（重复条件），BOTH/ALL 模式下 mode_b 的 scale_/model_ 图**永远不导出**，双方案对比功能名存实亡 | AutoReport.vbs:423（对照 :410 的 A/B/ALL 判断） |
| A5 | 幻灯片 14~16：config 勾选翘曲 X/Y/Z 分量后图像先被放置、随后被无差别清除——用户勾选被静默忽略；且该行为硬编码而非配置驱动 | pptx_builder.py:682-683（跳过放置）、:813-829（无条件删图）；report_config.json:240/269/298（enabled:false） |
| A6 | 附加结果页 `add_picture` 用固定宽高（9.70×5.95 英寸）不锁纵横比，非恰好比例的图**被拉伸变形** | pptx_builder.py:845-849 |
| A7 | `generate_solid_cad_model` 失败（返回 False）后，紧接着 `shutil.copy2(solid_out, ...)` 抛 FileNotFoundError 且无人捕获，**整个流水线崩溃** | image_processor.py:537-539 |
| A8 | manifest.json 里 `moldflow_version` 硬编码 "2023"，从不读取真实版本 | AutoReport.vbs:475 |

### B. 截图清晰度（用户点名问题）

| # | 问题 | 证据 |
|---|------|------|
| B1 | 方案 B 拼合画布**固定 1896×1200 px**（约 194 DPI），所有 3D 云图无论导出多高分辨率都被降采样到 1130 px 高度，2K/4K 导出被浪费，投影/高分屏上发糊 | image_processor.py:442-450（`target_h = 1200`、`model_avail_h = target_h - 70`） |
| B2 | 截图分辨率 `0x0`（视口原生）时，清晰度完全取决于运行时 Moldflow 窗口大小；GUI 推荐文案写"1080P 高清推荐"但配置默认却是 0x0，**默认配置与推荐自相矛盾**；且 solid_model/mesh_model 两处硬编码 `SaveImage3 ..., 0, 0, ...` 永远不吃分辨率配置 | report_config.json:9-10；config_gui.py:91-95；AutoReport.vbs:217,233,426 |
| B3 | 左侧色带放大后叠加 UnsharpMask 锐化（radius 1.2 / percent 160），降采样+锐化组合引入伪影，数值刻度发毛 | image_processor.py:453-461 |
| B4 | GIF 逐帧独立 256 色量化且 `method=Image.Resampling.LANCZOS`（值=1，实为 MAXCOVERAGE 量化）——参数语义错误；逐帧调色板不一致会造成色彩跳动；50 帧×高分屏全部驻留内存 | gif_enhancer.py:20-26,65 |

### C. 交互体验（用户点名问题）

| # | 问题 | 证据 |
|---|------|------|
| C1 | **全程黑箱**：GUI 点"生成"后窗口立即关闭 → wscript 静默执行 → Python 以隐藏窗口运行（`Run(PyCmd, 0, True)`）且 stdout/stderr **无任何重定向**，用户只能干等；成功/失败仅靠最后一枚 MsgBox，失败只给退出码 + "请看 run.log"，而 run.log 里根本没有 Python 的报错内容 | config_gui.py:372-379；AutoReport.vbs:481-504 |
| C2 | GUI 无前置校验：模板路径为空/不存在照常启动，Moldflow 导出跑完几分钟后 Python 才抛 FileNotFoundError——**最贵的失败顺序** | config_gui.py:343-347（无校验）；pptx_builder.py:894-896 |
| C3 | AutoReport.vbs 不存在时点击"生成"**静默无反应**，无任何提示 | config_gui.py:377-379 |
| C4 | 输出文件名直接拼 raw study_name，实际产物出现 `054_5_hf_2026_8_2603060(1)_____-2026-09-06-...` 这类乱码式文件名 | pptx_builder.py:860；本次冒烟实测产物文件名 |
| C5 | run.log 编码混杂（VBS GBK 写入 + 未来 Python UTF-8 追加），记事本打开可能乱码 | AutoReport.vbs LogMsg 实现 + C1 改动引入的风险 |

### D. 健壮性 / 一致性

| # | 问题 | 证据 |
|---|------|------|
| D1 | BOTH 模式下方案 A 抛异常 → 方案 B 完全不执行（无隔离） | pptx_builder.py:918-932（无 try 包裹） |
| D2 | manifest.json 损坏 → `json.load` 未捕获直接崩 | pptx_builder.py:901-903 |
| D3 | material 字段为 null 时 `float(None)` TypeError 崩溃 | pptx_builder.py:359-368 |
| D4 | `from gif_enhancer import ...` 仅在以脚本方式运行时成立，包方式导入时静默跳过 GIF 优化（顶层 import 已有双 fallback，此处没有） | pptx_builder.py:701-708 |
| D5 | 视口要素提取的裁剪框按 1920 宽硬编码（50..350 / w-260..w / 800..1800），换分辨率抓取错位 | image_processor.py:118-158 |
| D6 | 多处 `except: pass` / 裸 except 吞异常无上下文（违反项目契约 R-3.1） | image_processor.py:395,402,410,420；pptx_builder.py:734,782,963 |

## 二、目标状态（Dream State）

1. 任何数据来源缺失 → 报告中明确标注"数据缺失"，绝不出现别套方案的数据；
2. 标注数值、锁模力等全部单源来自 config/导出数据，换方案零手工改代码；
3. 默认配置下 3D 云图在报告里达到 ≥220 DPI 有效分辨率（2K 导出不降采样，4K 按目标自适应），GIF 无色彩跳动；
4. 从点击"生成"到拿到结果全程有感知：配置校验前置、过程日志可查、成败弹窗带具体原因与文件位置；
5. A/B/BOTH 三种模式行为一致且正确；
6. 全部改动附零依赖回归脚本，门禁（语法+排版+冒烟）保持全绿。

## 三、实施任务（按优先级排序）

> ⚠️ **本节已被评审区（文末）的「计划修订（v2 任务清单）」取代并进一步修订**——v1 保留仅作历史记录，实施以评审区 v2 及后续阶段修订为准。

### P0 数据正确性（先做，防止报告出错）
- T1 (A1/A2): 删除全部"假数据兜底"默认值 → 数据缺失时该字段填 `"（数据缺失）"` 并在日志 ERROR 记录；报告照常生成但绝不伪造。涉及 pptx_builder.py `get_formatted_mesh_text` / `render_material_dialog_cards`。
- T2 (A3): 标注数值单源化：`merge_scale_and_model` 增加 `annotations` 参数（由 pptx_builder 从 config `clamp_force_settings`/`inj_pressure_settings` + clamp_force_info.json 传入），删除 image_processor 中的硬编码默认方案数据。
- T3 (A4): 修复 VBS:423 重复条件为正确的多模式判断（B/BOTH/ALL 语义对齐 :410），并将 "ALL"/"BOTH" 统一为 config_gui 实际产生的值。
- T4 (A5): 14~16 页行为配置驱动：`enabled=false` → 保留现状（清图留标题）；`enabled=true` → 正常放置图像且不删。
- T5 (A6): 附加页 `add_picture` 改为等比适配（复用 `fit_picture_in_safe_box` 的纵横比逻辑）。
- T6 (A7): solid_model 生成失败时跳过 copy 并告警，不崩溃。
- T7 (A8): manifest 版本号从 Synergy/StudyDoc 实际读取（读不到则写 "unknown"，不硬编码）。

### P1 截图清晰度
- T8 (B1): 拼合画布分辨率可配置 `image_settings.composite_target`（默认按 220 DPI 目标 ≈ 2147×1353；上限取源图高度，绝不放大模糊），色带与模型按同一目标高度缩放。
- T9 (B2): 默认 `image_settings` 改为 2560×1440 并同步 GUI 推荐项；solid_model/mesh_model 的 SaveImage3 同样传入配置分辨率。
- T10 (B3): 源分辨率足够时跳过 UnsharpMask；仅当色带被放大时轻度锐化。
- T11 (B4): GIF 量化改用全局调色板（首帧自适应 + 全帧复用）+ 正确的 `Image.Quantize` 方法常量；流式处理帧序列降低内存峰值。

### P2 交互体验
- T12 (C1/C5): 统一日志：VBS 启动 Python 改为 stdout/stderr 重定向追加写入 `temp\run.log`（UTF-8 统一），最终 MsgBox 附带"日志位置"，失败时直接展示 run.log 最后 5 行。
- T13 (C2): GUI 保存/生成前校验：模板文件存在性、输出目录可写、至少勾选 1 项结果；不通过则 messagebox 拦截。
- T14 (C3): VBS 缺失时 messagebox 报错而非静默。
- T15 (C4): 文件名净化：study_name 去除路径非法字符与连续下划线/括号噪声。

### P3 健壮性
- T16 (D1): BOTH 模式按方案隔离 try/except，单一方案失败不拖垮另一方案，汇总两者成败。
- T17 (D2/D3): manifest/material/mesh JSON 读取统一带编码+异常防护，null 字段安全转换。
- T18 (D4): gif_enhancer 导入补齐与顶层一致的双 fallback。
- T19 (D5): 视口裁剪框按图像尺寸比例缩放（以 1920 宽为基准的比例系数）。
- T20 (D6): 清理全部裸 except：至少 print 带文件名与异常上下文（符合 R-3.1）。

### P4 测试与文档
- T21: 新增零依赖回归脚本 `tests/run_checks.py`（stdlib only）：覆盖 T1（缺失数据→"数据缺失"文案）、T2（config 值传入标注）、T5/T8（纵横比与画布计算纯函数）、T15（文件名净化）。登记进 AGENTS.md §2 Test 行（按契约模式 B 幂等更新，先 diff 后写）。
- T22: 更新 使用说明.md（新默认分辨率、日志查看方式）与 AI_GUIDE.md（新坑：BOTH 模式条件、假数据兜底教训）。

## 四、明确不在本次范围（NOT in scope）

- 将 AutoReport.vbs 重写为 Python/COM（架构级重写，风险与工作量另立项）；
- PPT 模板版式重设计（模板是既定交付物，不动）；
- Moldflow COM API 行为本身的变更（仅调参数）；
- 新增第三方依赖（维持 Pillow/pptx/numpy + stdlib）。

## 五、已存在、直接复用的能力

- `fit_picture_in_safe_box`（纵横比适配）→ T5 复用其逻辑；
- `load_truetype_font` 多级降级 → 新增绘制沿用；
- `safe_float` 模式（get_formatted_mesh_text 内）→ T17 推广为通用工具函数；
- VBS LogMsg + run.log 通道 → T12 在其上扩展而非新建日志体系；
- AI_GUIDE.md 十大坑 → 全部改动遵循，不引入新踩坑。

## 六、验证与回归

- 每任务完成后跑门禁：`python -m py_compile ...`（Lint）、`python -m ruff format --check ...`（Format）、`python core/pptx_builder.py --config report_config.json --data-dir temp`（冒烟，退出码 0 + [SUCCESS]）；
- T21 回归脚本纳入 Test 门禁；
- VBS 改动（T3/T7/T9/T12）无法单测：用静态条件断言脚本检查 + 人工在 Moldflow 实机验证 BOTH 模式（列入手动验证清单）。

---

# /autoplan 评审区（流水线产出，以下内容为评审结论与计划修订）

## Phase 1 — CEO 策略评审（SELECTIVE_EXPANSION，自动决策模式）

### 前提挑战（0A）
| 前提 | 评估 | 处置 |
|---|---|---|
| P1 问题清单反映真实痛点 | ✅ 全部 file:line 证据经两声部独立抽查属实（1 处引用偏移 report_config.json:8-9 已更正） | 保留 |
| P2 "9662.27 cm² 只是死兜底默认值" | ❌ **被证伪**：AutoReport.vbs:176 `surface_area = MVol * 7.5` 是活路径伪造公式，temp/mesh_summary.json 实测 9662.26 = 1288.302×7.5 —— 当前每份报告第 2 页印的都是编造的表面积 | 计划修订：T1 扩展到 VBS 层 |
| P3 "配置分辨率影响全部导出清晰度" | ⚠️ 部分错误：`Viewer.SaveImage`（色带/视口图）无尺寸参数，只有 `SaveImage3`（模型图）吃 ImageWidth/Height；色带清晰度受 Moldflow 窗口尺寸封顶 | 计划修订：T9 增加色带导出通道 + 目标状态措辞修正 |
| P4 保持 VBS+Python 双层架构不重写 | ✅ 合理（AI_GUIDE 既定模式 + COM 移植风险），但须写明**移植触发条件** | 保留 + NOT-in-scope 补充 |
| P5 模板为既定交付物不动 | ✅ 合理，但后果（模板钉死布局）须显式声明 + 哈希变更告警 | 保留 + T26 |
| P6 用户"全方位优化"的诉求由 22 项问题覆盖 | ✅ 两声部补充 8 类新发现后仍属同一优化范畴，无需重构问题定义 | 保留 |

### 梦想态映射（0C）
```
现状                          本计划                         12 个月理想
VBS+Python 双层脚本           全层假数据清除                  分析完成→全自动出报告
报告含编造数据/陈旧图          分辨率上限抬升+单源数据          零接触、可追溯、任何机器可跑
黑箱运行、失败只有退出码       全程可观测、配置驱动             进度可视化+自动分发
无测试                        零依赖回归脚本+门禁             COM 层可测(移植后)
```

### 实施备选方案（0C-bis）
| 方案 | 摘要 | 工作量(人/CC) | 风险 | 完整度 | 决策 |
|---|---|---|---|---|---|
| A 最小可行 | 只修用户点名的两项：日志弹窗 + 配置分辨率 | S/~30min | 低 | 4/10 | 否（A1-A3 假数据依旧进报告） |
| **B 系统性加固（计划全文）** | P0-P4 全任务，无架构变更 | M/~2h CC | 低-中 | 9/10 | **✅ 采纳（P1 完整度优先）** |
| C 理想架构 | pywin32 重写 COM 层 + pytest + 进度 UI | XL/天级 | 高（COM 实机验证难、新依赖违 R-3.5） | 10/10 | 否 → 移植触发条件写入 NOT-in-scope |

### 模式确认（0F）
SELECTIVE_EXPANSION（autoplan 缺省）。复杂度检查：触达 9 文件（6 代码 + 测试 + 2 文档）> 8 → 判定：用户诉求本身横跨 GUI/VBS/Python 三层，无法再少；核心代码仅 6 文件。通过。

### 选择性扩展机会（0D 扫描 → 自动决策）
| # | 提案 | 工作量 | 决策 | 理由 |
|---|---|---|---|---|
| E1 | GUI 进度窗口（实时 tail run.log） | M | **延期 TODOS** | T12 日志+弹窗解决 80% 痛点；进度窗涉及时序复杂度 |
| E2 | PyInstaller 打包单文件 exe | L | **延期 TODOS** | 部署改善但引入构建链与杀软误报风险 |
| E3 | GUI 历史报告列表页 | S | 跳过 | YAGNI，last_output_path.txt 已够用 |
| E4 | 英文报告模板支持 | L | 跳过 | 用户群为中文工程师 |
| E5 | mesh/material 缺失时 GUI 前置提示 | S | **并入 T1/T13** | 爆炸半径内、<1 天 |

### 双声部输出（0.5，Codex 缺位 → [subagent-only] 降级）

**CLAUDE SUBAGENT（CEO — 战略独立审查）要点**：P0 打错了靶层——假数据源头在 VBS 层（材料假种子 VBS:255-277 且必写 JSON:341、表面积 =MVol×7.5 活路径伪造 VBS:176、pptx_builder.py:774-775 同类默认值），T1 只清 Python 层没用；陈旧产物无失效机制（ref_triad/ref_title 缓存继承上月方案，图像回退链吃到旧图）与 P0 同罪；标注数值应从 COM 导出取（GetMaxValue 已有雏形 VBS:452）而非手工维护 config；≥220 DPI 目标对色带不可达（SaveImage 无尺寸参数）；"数据缺失也照常出报告"是单方面产品决策，应配置化 + 完成时确认弹窗；VBS 四次打补丁前应有移植触发条件；模板钉死须声明；报告 ~25MB 无体积预算；T19 应先于 T9；"为何不用 Synergy 自带报告"须写进文档；平台漂移风险（COM ID 硬编码、BaseDir 机器绑定）无监控注记。**判定：优先级方向对，但 P0 靶层与 P1 前提需修订。**

**CLAUDE SUBAGENT（外部声部 — 独立挑战）要点**：A4 真因是 GUI 产生 "BOTH" 而 VBS:410/:423 只认 A/B/ALL → BOTH 模式两分支都不执行，且图像回退链**静默吃上一次运行的陈旧图**（数据正确性问题，不只是缺功能）；T2 参数穿透三层不如让 VBS 直接导出峰值 JSON（顺带修掉 mode A 第二标注点 image_processor.py:579-592 被遗漏的问题）；AutoReport.vbs 是 UTF-8/GBK 混合编码地雷（641 行 UTF-8 + 34 行 GBK-only，MsgBox 字面量已乱码），四个 VBS 任务动手前必须先规范化编码；WshShell.Run 无法重定向 stdout，T12 须 cmd /c 包装；冒烟门禁不可重复（自动开 PPT + 写死模板路径）需 --no-open + --output；show_gui_before_run=true 时 GUI→VBS→GUI 递归会双开竞争；composite_target 新配置键没接 GUI 且 2147×1353 不是 1.58:1；create_gif_from_frames 是死代码（管线只调 optimize_existing_gif）。**判定：需要修改（5 处破坏 P0/P1 承诺）。**

**CEO 双声部 — 共识表**（Codex 列因缺位记 N/A，单声部关键发现仍按规则采纳）：
```
维度                                  Claude  Codex  共识
1. 前提是否成立                        修订3项   N/A   前提P2/P3证伪→已修订
2. 是否解对问题                        确认      N/A   确认（方向正确）
3. 范围校准                            确认+2延期 N/A  确认（E1/E2 延期）
4. 备选方案充分性                      确认      N/A   确认（F9 补录为已评估拒绝）
5. 平台/竞争风险覆盖                   补充3项   N/A   单声部采纳（F10/F11→T22）
6. 6个月轨迹                           确认      N/A   确认（F7 移植触发条件已加）
```

**两声部独立命中的同一批问题（高置信信号 → 直接采纳）**：① 假数据源头在 VBS 层；② BOTH 模式真因 + 陈旧图回退；③ 峰值数值应从导出数据取而非手工 config；④ 色带分辨率天花板；⑤ VBS 编码地雷需前置处理；⑥ T19 先于 T9。

### 评审节发现（Section 1-10，Section 11 → Phase 2）

**S1 架构**：依赖图 VBS→temp/*→Python→PPTX，配置三方读同一 JSON（单源 ✅）。新数据流 2 条（峰值 JSON、日志重定向）四路径已核：nil（字段缺失→"数据缺失"）/空（safe 转换）/错误（防护读取）。发现 1 项：输出文件名未防路径注入（study_name 含 `..\` 或分隔符可逃逸输出目录）→ 并入 T15。单点故障：AutoReport.vbs 本身（本计划不动，记入 12 月债）。回滚：全部 git revert 可逆；VBS 为二进制安全存储（GBK 字节保留）。
**S2 错误与救援**：注册表见下节。当前代码 3 个 CRITICAL GAP（A1/A2 假数据静默、A7 未捕获崩溃、D1 双方案连锁失败）全部由 T1/T6/T16 修复。
**S3 安全与威胁**：本地单用户工具，无网络面。发现 2 项：① 路径注入（→T15，Likelihood 低/Impact 中）；② VBS 用 htmlfile eval 解析 config JSON（AutoReport.vbs:55），恶意 config 可执行 JScript——本地文件、用户自控，Likelihood 低/Impact 中 → **延期 TODOS**（重写 VBS JSON 解析属 VBS 手术）。无新依赖、无密钥。
**S4 数据流与交互边界**：GUI 双击保存（幂等 ✅）、模板空（→T13 拦截）、PPT 占用文件（已有时间戳重试 ✅）、GIF 缺失（已有回退链 ✅）、全不勾选（→T13 要求 ≥1）。发现：show_gui_before_run 递归双开 → 并入 T13。
**S5 代码质量**：safe_float 模式重复 → T17 推广；build_single_report 370 行上帝函数 → **延期 TODOS**（本计划只做最小插入，避免大重构混入）；T8 曾拟新增配置键 → 按 P5 简化为复用 image_settings.height，不新增键。
**S6 测试**：新路径可测矩阵——纯函数（画布计算/纵横比/文件名净化/缺失文案）单测覆盖；VBS 模式分发用静态断言脚本（扫 `= "B" Or ScreenshotMode = "B"` 类模式）做回归；损坏 JSON fixtures（逐一喂坏 mesh/material/manifest/peaks → 报告仍生成且标"数据缺失"）；2am-Friday 测试 = 现有端到端冒烟；chaos = 只读输出目录。VBS 行为变更仍需 Moldflow 实机手动清单（T0/T3/T9/T12）。
**S7 性能**：拼合画布升至 ~2138×1353×3B ≈ 8.7MB 内存可忽略；GIF 流式处理（T11）控内存峰值；最慢路径不变（PPTX 保存大 blob）。发现：报告体积 ~25MB 且画布升级后更大 → 新增 T25 体积预算。
**S8 可观测**：T12 即本节主体；增强：日志行加时间戳前缀、失败弹窗附 run.log 尾部 5 行。runbook → T22 更新使用说明。
**S9 部署与回滚**：无部署设施（本地脚本）。回滚 = git revert（含 VBS）。验证清单 = 门禁 + BOTH 模式实机手动验证。发现：陈旧 temp 产物跨运行污染 → T23（已列 P0）。
**S10 长期轨迹**：技术债净变化 = -3（假数据类、静默失败类、不可诊断类）/ +1（峰值 JSON 新契约须文档化）；可逆性 4/5；知识集中 → T22 补"为何不用 Synergy 自带报告"+ BaseDir 机器绑定 + 移植触发条件。

### 错误与救援注册表（改动后状态）
```
方法/路径                        | 可能出错                    | 异常/处置                        | 用户所见
--------------------------------|---------------------------|--------------------------------|---------
get_formatted_mesh_text         | JSON 缺失/损坏/字段 null   | 捕获→"数据缺失"占位+ERROR 日志   | 明确标注，无假数
render_material_dialog_cards    | 同上                       | 同上（safe_float(None)→缺失）    | 同上
VBS 材料导出(T1)                | COM FindProperty 失败      | 写 null+data_complete=false      | 弹窗列出缺失项
merge_scale_and_model(T2)       | peak_values.json 缺失      | 不绘制探针+WARN（config override 可用但标过期） | 无错误标注
process_all_mode_b_plots(T2)    | 同上（第二调用点）          | 同上                             | 同上
build_report 按方案隔离(T16)    | 单方案异常                  | 捕获→另一方案继续→汇总弹窗        | 分方案成败+日志
solid_model 复制(T6)            | 生成失败文件不存在          | 跳过+WARN                        | 封面用旧图/日志说明
run.log 重定向(T12)             | cmd 包装失败                | 退回仅 VBS 日志+弹窗注记          | 弹窗提示日志受限
GIF 优化(T11)                   | 帧读取失败                  | 捕获→保留原 GIF+WARN              | 动图未优化其余正常
```

### 失败模式注册表（当前代码 → 修复后）
```
路径                 | 失败模式            | 修复前        | 修复后          | 测试
--------------------|--------------------|--------------|-----------------|--------
mesh_summary 缺失    | 假网格数据进报告     | 静默假数 ❌CRITICAL | "数据缺失"标注 | 回归✓
material_info 缺失   | 假材料卡进报告       | 静默假数 ❌CRITICAL | 同上+确认弹窗  | 回归✓
surface_area         | ×7.5 编造公式       | 活路径假数 ❌CRITICAL | 真实值或缺失  | 回归✓
标注数值             | 换方案后探针错值     | 静默错数 ❌CRITICAL | 导出数据单源   | 回归✓
BOTH 模式            | 两分支都不导出+旧图  | 静默降级 ❌CRITICAL | 正确分发+清旧图 | 静态断言+实机
ref_triad 缓存       | 继承上月方案坐标     | 静默错图 ❌HIGH    | 每次重新生成   | 回归✓
```

### 计划修订（v2 任务清单——取代原 §三，声部发现全部吸收）

**P0 数据正确性（含 VBS 层）**
- **T0（新，前置）**：AutoReport.vbs 编码规范化——全文统一 GBK/ANSI，修复已乱码的 MsgBox 字面量（如 :503"桑台谭"），约定后续编辑工具链必须按 GBK 读写。**先于一切 VBS 任务。**
- **T1（修订）**：全层假数据清除——Python 侧删除全部假默认值（含 :774-775 锁模力默认）；VBS 侧 material/mesh 导出失败时写 null/省略字段 + `data_complete` 标志；**删除 VBS:176 的 MVol×7.5 编造公式**（无真实值则字段缺失）；新增 `on_missing_data: annotate|fail` 配置（默认 annotate）；完成弹窗列出全部缺失字段要求确认。
- **T2（重构）**：峰值数据单源化改为**从导出层取**——VBS 用 `PlotObj.GetMaxValue`（已有雏形 :452）+ 峰值时间导出 `peak_values.json`（锁模力+注射压力两项）；image_processor **两个调用点**（merge_scale_and_model 与 :579-592 直标循环）统一从 data_dir 读取；config 旧数值降级为 override（启用时日志 WARN"手工覆盖导出值"）。不再做三层参数穿透。
- **T3（重构）**：BOTH/ALL 模式修复——VBS:410/:423 条件改为完整模式判断（A/B/BOTH/ALL 全集），GUI 的 "BOTH" 与 VBS 的 "ALL" 语义统一。
- **T23（新）**：陈旧产物失效——每次运行开始清理上次 mode_a/mode_b 的 per-plot PNG（或按 manifest 日期拒绝陈旧图）；ref_triad/ref_title 每次运行重新生成；pptx_builder 图像回退链加新鲜度校验。

**P1 截图清晰度（顺序修订：先比例化再提分辨率）**
- **T8-pre（原 T19，提前）**：视口裁剪框按图像尺寸比例缩放（以 1920 宽为基准系数），**必须先于 T9**（分辨率变化会让硬编码框错位并触发 T23 的陈旧回退）。
- **T8（简化）**：拼合画布高度取 `image_settings.height`（默认 2560×1440 时画布 ≈ 2273×1438，精确 1.58:1），源图分辨率不足则用源高、绝不放大；**不新增配置键**；实施前先做实机实验：0×0 vs 2560×1440 的 model_ 导出像素级对比，验证 SaveImage3 高分辨率确实渲染更多细节（写入计划的前提核实）。
- **T9（修订）**：默认导出 2560×1440（GUI 同步默认项与文案）；solid/mesh 的 SaveImage3 同传配置分辨率；色带尝试 `SavePlotScaleImage` 高分辨率导出通道；目标状态措辞修正为"模型区域 ≥220 DPI，色带受导出通道能力限制"。
- **T10**：锐化条件化（源足够时跳过）。（不变）
- **T11（修订）**：GIF 全局调色板 + 正确 `Image.Quantize` 常量 + 流式帧处理；scope 注记：`create_gif_from_frames` 当前不在管线调用链（死路径，本次标记不删）。

**P2 交互体验**
- **T24（新）**：pptx_builder CLI 增加 `--no-open`；门禁冒烟命令改为 `--output temp\smoke.pptx --no-open`（可重复、不弹窗、产物进 temp）。
- **T12（修订）**：日志重定向改用 `cmd /c "python ... >> run.log 2>&1"` 包装（WshShell.Run 无重定向能力），python 路径显式解析；日志行加时间戳前缀；失败弹窗附 run.log 尾部 5 行 + 完整路径。
- **T13（修订）**：GUI 前置校验（模板存在/目录可写/≥1 勾选）+ **GUI→VBS→GUI 递归防护**（GUI 唤醒 VBS 时写入临时标志，VBS 跳过再弹 GUI）+ 完成弹窗列出缺失数据字段（承接 T1）。
- **T14**：VBS 缺失时 messagebox 报错。（不变）
- **T15（修订）**：文件名净化——非法字符、连续下划线、路径分隔符与 `..\` 注入防护。

**P3 健壮性**
- **T16**：BOTH 模式按方案隔离 try/except + 分方案成败汇总。（不变）
- **T17**：JSON 读取统一防护 + null 安全转换（推广 safe_float）。（不变）
- **T18**：gif_enhancer 导入双 fallback。（不变）
- **T20**：裸 except 清理（带上下文）。（不变）
- **T25（新）**：报告体积预算——目标 ≤15MB，不透明图像按预算转 JPEG/压缩，画布升级后体积不失控。
- **T26（新）**：模板哈希校验——记录模板 sha256，变更时 WARN（模板钉死布局的防呆）。

**P4 测试与文档**
- **T21（扩充）**：零依赖回归脚本 `tests/run_checks.py`：缺失数据文案、峰值单源读取（两调用点）、纵横比、画布计算、文件名净化（含注入用例）、损坏 JSON fixtures（逐文件投毒）、VBS 模式条件静态断言；登记进 AGENTS.md §2 Test 行（契约模式 B 幂等更新）。
- **T22（扩充）**：文档更新（使用说明/AI_GUIDE）+ 写明"为何不用 Synergy 自带报告"（F10）+ pywin32 移植触发条件（F7）+ BaseDir 机器绑定说明（F11）+ 手动降级路径（COM 失效时半手动流程，F11）。

### NOT in scope（修订版）
| 项 | 理由 |
|---|---|
| VBS→pywin32 重写 | 风险与工作量另立项；**触发条件（写入文档）：再出现任一 VBS 层 bug，或 Moldflow 版本升级** |
| PPT 模板重设计 | 模板是既定交付物；T26 哈希告警防呆 |
| 图例从数据重绘（F9，替代栅格拼合） | 需重做核心拼合 IP 且违反"严格对齐官方图"要求；若 T8/T9/T19 后清晰度仍不达标再立项 |
| 新增第三方依赖 | 维持现有栈（R-3.5） |
| Moldflow COM API 行为变更 | 仅调参数 |
| build_single_report 拆分重构 | 与本计划改动混入风险大于收益 → TODOS |
| VBS eval JSON 解析加固（S3-②） | VBS 手术 → TODOS |

### Deferred to TODOS.md
E1 GUI 进度窗口（P2，依赖 T12 落地）；E2 PyInstaller 打包（P3）；build_single_report 拆分（P3）；VBS eval 解析加固（P3，安全）；F9 图例数据化重绘（P3，视 P1 效果）。

### Dream State Delta
本计划完成后：12 项理想态达成 7 项（数据可信 ✅、单源 ✅、可观测 ✅、模式一致 ✅、有回归 ✅、清晰度受限达标 ✅、体积可控 ✅）；未达 3 项（零接触全自动、跨机器可移植、COM 层可测）——均需架构移植，已留触发条件。

### CEO 完成度总结
```
+====================================================================+
|         MEGA PLAN REVIEW — COMPLETION SUMMARY (Phase 1)            |
+====================================================================+
| 模式            | SELECTIVE_EXPANSION（自动决策）                    |
| 系统审计         | 2 提交/无 stash/无 TODO 标记/冷启动无学习库         |
| Step 0          | 方案 B 采纳；扩展 2 延期 2 跳过 1 并入              |
| S1 架构          | 1 发现（路径注入→T15）                             |
| S2 错误          | 9 路径映射，修复前 3 CRITICAL GAP（本计划全修）      |
| S3 安全          | 2 发现（1 修 T15，1 延期）                          |
| S4 边界          | 6 边界映射，1 未处理（递归→T13）                    |
| S5 质量          | 2 发现（DRY→T17；上帝函数→延期）                    |
| S6 测试          | 矩阵产出，0 缺口（VBS 实机清单显式化）              |
| S7 性能          | 1 发现（体积预算→T25）                             |
| S8 可观测        | 0 缺口（T12+增强）                                 |
| S9 部署          | 1 发现（陈旧产物→T23）                             |
| S10 轨迹         | 可逆性 4/5；债 -3/+1；3 项延期                      |
| S11 设计         | → Phase 2 专属评审                                 |
| 双声部           | [subagent-only]；6 维度确认 5、修订前提 2           |
| Lake Score       | 9/9 决策选择完整选项                                |
+====================================================================+
```

## Phase 2 — 设计评审（UI 范围命中，SELECTIVE_EXPANSION 自动决策）

> Mockup 生成降级说明：gstack designer 二进制存在但需 OPENAI_API_KEY（未配置），按规程回退 ASCII 线框稿。

### Step 0 设计范围评估
- **初评：4/10**（独立设计声部评分，本人复核同意）——数据状态工程扎实，但运行期零交互设计、报告内视觉状态变化未规格化、对话框无文案规格。
- 10 分样板：每个用户可见状态（含报告页）都有具体呈现规格；运行全程有进度与取消；每处文案逐字定稿。
- DESIGN.md：不存在（内部单用户工具，Tkinter 原生 vista 主题即设计系统；补设计系统文档属过度工程，P3 判定不加）。
- 现有设计资产：Tkinter vista 主题、蓝白配色（#0066CC 主按钮）、微软雅黑字体、消息图标约定——T27/T28 沿用。

### 双声部（[subagent-only]，Codex 缺位）
**CLAUDE SUBAGENT（设计 — 独立评审）**：8 项发现（F1 GUI 层级缺失于计划、F2 缺失数据的报告内呈现未定义（mesh 表格对齐破坏+"适合X分析"结论行仍打印健康声明、材料卡 14 框全缺）、F3 14-16 空白标题页是客户可见的破损页、F4 运行期无进度无取消【critical】、F5 GUI 生命周期态（配置损坏静默崩溃）、F6 缺图分支把模板里上个零件的图留在报告里【high】、F7 对话框无文案规格（messagebox 不可复制等宽日志不可读）、F8 T25 盲目 JPEG 与 T10 锐化互搏（线条图振铃）+GIF 占预算大头；元发现：v1 任务清单未标废弃）。**评分 4/10。**

**设计 litmus 记分卡**（APP UI 分类；Codex 列 N/A）：
```
检查项                                   Claude  Codex  共识
1 首屏品牌/功能可识别                     YES     N/A   确认
2 存在强视觉锚点（运行按钮）              YES     N/A   确认
3 只扫标题即可理解（T27 后）              YES     N/A   确认（修订后）
4 每区域一职（T27 后）                    YES     N/A   确认（修订后）
5 卡片确有必要                            n/a     N/A   不适用
6 动效改善层级                            n/a     N/A   不适用
7 去除装饰阴影仍有质感                    YES     N/A   确认
```

### 7 遍评审（评分 → 修订后）

**Pass 1 信息架构：4/10 → 8/10**
发现：GUI 主行动按钮沉在 600px 滚动区之下；模式单选用营销文案（"★ 方案 B：…绝不遮挡…"）而非决策语言；一次性配置（模板/输出目录）占据黄金首屏。报告页层级模板钉死（已声明 NOT-in-scope）。
**自动决策（P1/P5）**：新增 **T28 GUI 层级修正**——顶部常驻行动栏（生成/保存），模式单选一行化改工具文案，setup 字段降级"高级设置"折叠。线框：
```
┌─ 模流分析报告生成器 ────────────────────────────┐
│ [★保存配置并立即生成报告] [仅保存]      [关闭]  │ ← 常驻行动栏
│ 模式: (●)方案B-推荐(分图拼合) ( )方案A ( )双方案 │ ← 一行工具文案
│ ┌ 高级设置 ▼ (模板/输出/画质/GIF 参数) ────────┐ │
│ └──────────────────────────────────────────────┘ │
│ 分析结果勾选区 (Notebook, 占剩余空间)            │
└──────────────────────────────────────────────────┘
运行状态窗（T27，生成期间）:
┌─ 模流分析报告生成器 - 运行中 ──────────┐
│ 正在导出: 注射位置处压力 (12/31)        │
│ ┌─────────────────────────────────────┐│
│ │[03:21:05][方案B] 导出完成: press.png││ ← 等宽日志 tail
│ └─────────────────────────────────────┘│
│ [打开日志]                    [取消生成]│
└─────────────────────────────────────────┘
```

**Pass 2 交互状态覆盖：3/10 → 8/10**
```
特性            | LOADING            | EMPTY              | ERROR                    | SUCCESS          | PARTIAL
---------------|--------------------|--------------------|--------------------------|------------------|----------
生成运行        | T27 状态窗实时步骤  | T13 拦截(≥1勾选)   | T12 失败对话框(可滚动日志) | 弹窗+自动打开     | 弹窗列缺失字段(T13)
配置加载        | —                  | 首次运行默认+引导   | 配置损坏→弹窗+默认兜底(T28)| 静默             | —
保存            | —                  | —                  | 写失败弹窗(T28)           | "已保存"提示     | —
报告页数据      | —                  | —                  | "数据缺失"短token+"—"     | 正常渲染         | 封面完整性章
报告页缺图      | —                  | —                  | "（图像缺失）"占位+ERROR   | 正常             | —
```
**自动决策**：全部采纳进 T1/T12/T13/T23/T28（F2/F5/F6/F7）。具体：mesh 表缺失字段用短 token "—" 保对齐 + 结论行改"数据不完整，结论仅供参考"；材料卡渲染但每框标"数据缺失"；封面右下加"数据完整性：部分缺失"角章（部分成功在交付物内自明，MsgBox 受众只有操作者，转发的客户只看得到报告）。

**Pass 3 用户旅程：4/0 → 8/10**
```
步骤 | 用户行为         | 情绪     | 计划支撑
1    | 打开工具         | 平静     | 配置记忆(T13 兜底)
2    | 点"生成"         | 期待     | T13 即时校验反馈
3    | 等待数分钟       | 焦虑峰值 | T27 状态窗: 可见步骤+可取消 ←原计划断裂点已接
4    | 失败             | 沮丧     | 失败对话框给下一步动作按钮
5    | 成功             | 解脱     | 弹窗+报告自动打开
6    | 转发客户         | 信任/心虚 | 封面完整性章 ←新增支撑
```
5 秒层：状态窗即时显示当前步骤；5 分钟层：任意时刻可取消（进程树 taskkill /T，Popen 句柄设计记录在 T27）；5 年层：诚实报告累积信任。

**Pass 4 AI Slop 风险：8/10**
APP UI 分类无硬拒模式；唯一违例 = 模式单选的营销文案（"utility language" 原则）→ T28 修正。计划 UI 描述全部具体（数值/字面 token），无泛型模式。通过。

**Pass 5 设计系统对齐：N/A**
无 DESIGN.md（见 Step 0）。新组件（状态窗/对话框/角章）均复用 Tkinter 原生词汇与既有蓝白配色，不引入新视觉词汇。通过。

**Pass 6 响应式与无障碍：5/10 → 7/10**
桌面-only，无响应式问题（minsize 已有）。A11y：T28 补 Tab 焦点顺序、Enter=生成 快捷键、按钮最小尺寸 ≥ 标准触达；对比度沿用现有 #0066CC/白 ✓。屏幕阅读器走 Tk 默认。通过。

**Pass 7 未决设计决策（全部当场裁决）**
| 决策 | 裁决 | 去向 |
|---|---|---|
| mesh 表缺失呈现 | 短 token "—" + 结论行替换 | T1 |
| 材料卡 14 框全缺 | 渲染+逐框"数据缺失"+封面角章 | T1 |
| 缺图分支 | 占位/移除+ERROR，绝不留模板旧图 | T23 |
| 对话框规格 | 自定义对话框规格逐字定稿（可滚动等宽+按钮） | T12/T13 |
| T25 压缩策略 | 分级：拼合 PNG 保留/封面 JPEG q90/GIF 帧×分辨率上限，冒烟产物实测定阈值 | T25 |
| 14-16 空白页 | **TASTE DECISION → 终门**：设计声部主张隐藏，用户历史明确要求"保留页面与标题"；默认保留 + 新配置 hide_blank_pages(default false) | T4 |
| 取消进程语义 | Popen 句柄 + taskkill /T 进程树 | T27 |

### 设计阶段任务（新增/修订）
- **T27（新，P1）**：运行状态窗（Text tail run.log via after() 轮询）+ 取消按钮（Popen 句柄+taskkill /T）。
- **T28（新，P1）**：GUI 层级修正（常驻行动栏/单选一行化/高级设置折叠/工具文案）+ 配置损坏兜底 + Tab 顺序与 Enter 绑定。
- **T1/T23/T25/T12/T13 修订**：吸收 Pass 2/7 裁决（缺失呈现、缺图占位、分级压缩、对话框规格）。

### 设计完成度总结
```
+====================================================================+
|      DESIGN PLAN REVIEW — COMPLETION SUMMARY (Phase 2)             |
+====================================================================+
| Pass 1 信息架构  | 4/10 → 8/10（T28 线框落计划）                     |
| Pass 2 状态覆盖  | 3/10 → 8/10（状态表落计划）                        |
| Pass 3 旅程      | 4/10 → 8/10（断点 3/6 接通）                       |
| Pass 4 Slop     | 8/10 → 8/10（T28 文案修正）                        |
| Pass 5 设计系统  | N/A（无 DESIGN.md，声明不加）                      |
| Pass 6 响应/A11y | 5/10 → 7/10（T28 补齐）                            |
| Pass 7 未决      | 7 项全裁决（1 项 TASTE → 终门）                    |
| Mockups         | 0 生成（OPENAI_API_KEY 缺失）→ ASCII 线框 2 幅      |
| 总评分           | 4/10 → 8/10                                       |
+====================================================================+
```

## Phase 2.5 — DX 评审（DX POLISH 模式，自动决策）

**Persona**（从文档推断）：单一工程师所有者 + 按 AGENTS.md 契约工作的 AI 编码代理；终端用户用 GUI 不用终端。

### 双声部（[subagent-only]，Codex 缺位）
**CLAUDE SUBAGENT（DX — 独立评审）18 项发现**，关键项：
- 【严重】F1 无依赖清单无引导：requirements.txt 缺失（把"不新增依赖"与"不固定依赖"混为一谈）；门禁还依赖 ruff 但未追踪。
- 【严重】F2 模板是未声明的代码内先决条件：16 页结构/标记字符串/象限几何无任何文档，全新机器首跑被模板卡死且无从得知模板要求。
- 【高】F3 无环境预检：python 不在 PATH 时 run.log 空白，失败不可诊断。
- 【高】F7 错误注册表缺"怎么办"列；对话框文案"稍后定稿"= 交付半吊子错误处理。
- 【高】F8 peak_values.json 缺失→无标注=换皮静默失败（与 T1 要消灭的同类）。
- 【高】F10 T22 太模糊：AI_GUIDE 速查命令 #2-#4 在 T24 后仍会弹 PowerPoint 且依赖机器绑定模板；使用说明 §四 配置参考缺新键与优先级表。
- 【高】F15 "UTF-8 统一"无机制：cmd 重定向默认发 cp936 字节，VBS LogMsg 是 GBK——文件仍混合编码。
- 【中】F5/F6 CLI 工效：--no-open 与 open_after_export 双名一义；--output 在 BOTH 模式被静默忽略；argparse 无 ALL；冒烟门禁未固定 --mode B（模式回退到易变配置）。
- 【中】F11 "AGENTS.md §2 仍全是占位符"——**与事实不符，部分驳回**：§2 本会话已填充（Lint=py_compile、Format=ruff format --check、Test=冒烟、主干=master、豁免=无）；但其引申点有效：`ruff check`（静态 lint）从未作为门禁，py_compile 仅语法级。
- 【中】F16 run.log 无运行边界/无配置快照；F17 last_output_path.txt 失败时是陈旧谎言；F18 manifest 缺 COM 连接方式记录。
- TTHW：**~120 分钟 → 目标 15 分钟**（requirements + check_env + 模板契约文档 + 确定性冒烟命令）。

**DX 共识表**（Codex 列 N/A）：
```
维度                        Claude  Codex  共识
1. Getting started <5min?    ✗→修   N/A   单声部采纳(T29)
2. CLI 命名可猜测?           部分→修 N/A   单声部采纳(T24扩)
3. 错误信息可行动?           ✗→修   N/A   单声部采纳(三行式对话框)
4. 文档 2 分钟可找到?        ✗→修   N/A   单声部采纳(T22 具体化)
5. 升级路径安全?             ✓      N/A   确认(git revert+配置键增量)
6. 环境无摩擦?              ✗→修   N/A   单声部采纳(check_env)
```

### 开发者旅程图（9 阶段）
| 阶段 | 现状痛点 | 计划支撑 |
|---|---|---|
| 发现 | 仅本人 | — |
| 安装 | 无清单，grep import 猜依赖 | T29 requirements.txt+check_env.py |
| 配置 | 模板要求未知 | T22 模板契约章节 |
| 首跑 | 静默失败不可诊断 | T12 环境探测+三行式错误 |
| 调试 | run.log 混编码无边界 | T12/T15 RUN START 头+GBK 全链路 |
| 日常 | 黑箱等待 | T27 状态窗+取消 |
| 扩展 | v1/v2 任务混淆 | Phase 4 汇聚规范任务表 |
| 升级 | git revert 可逆+配置键增量 | T26 模板哈希 |
| 三周后排障 | 无运行快照 | T12 配置快照 config_used.json+AI_GUIDE 产物判读指南 |

**共情叙事**（第一人称维护者）："周五 17:40，同事催报告。我勾完结果点生成，窗口消失，光标闪了三分钟。弹窗说'状态: 1，请看 run.log'——打开是一堆 GBK 乱码夹着看不懂的英文，最后一行是 python 的半个 traceback。我不知道是模板路径错了还是 Moldflow 掉线了。修复后我希望的是：点生成时看到'正在导出 12/31'，失败时弹窗直接告诉我哪一步、为什么、点哪个按钮。"

### DX 记分卡（8 维度，修订后）
| # | 维度 | 当前 | 修订后 | 依据 |
|---|---|---|---|---|
| 1 | Getting Started | 2 | 8 | T29 引导三件套 |
| 2 | API/CLI 设计 | 5 | 8 | T24 扩展(--open/ALL/BOTH+output) |
| 3 | 错误与调试 | 3 | 8 | 三行式文案+退出码语义+峰值缺失可见 |
| 4 | 文档与学习 | 4 | 8 | T22 逐条重写清单+配置优先级表 |
| 5 | 升级与迁移 | 5 | 7 | 键增量+哈希告警（无版本化迁移，留待） |
| 6 | 开发环境与工具 | 5 | 8 | 门禁已运行+T21 扩展评估 ruff check |
| 7 | 社区与生态 | N/A | N/A | 内部单用户工具，不评分 |
| 8 | DX 度量与反馈 | 3 | 7 | RUN START 头+产物判读指南（用户已拒绝遥测，不加度量） |

**TTHW 评估**：~120 分钟（75-165 区间，受 F1/F2/F3 卡死）→ **目标 15 分钟**：`pip install -r requirements.txt` → `python check_env.py`（验解释器/依赖/模板/标记）→ 打开文档载明的模板 → 门禁冒烟命令（--mode B --no-open 固定）→ 实跑。

### DX 实施清单（新增/修订任务）
- **T29（新，P1）**：requirements.txt（pin pillow==12.3.0 / python-pptx==1.0.2 / numpy==2.5.2 / ruff==0.16.6——固定非新增）+ check_env.py（stdlib 环境预检：解释器/三依赖/模板存在/幻灯片数/标记字符串）；T13 GUI 启动调用。
- **T2 修订**：peak_values 缺失 → 完成弹窗缺失列表 + 报告图面无标注徽记；`on_missing_data: fail` 时致命。
- **T7 修订**：manifest 增记 COM 连接方式（SAInstance/GetObject/CreateObject）与实例值。
- **T8 修订**：0x0 选项语义定义为画布固定 1920 宽并文档化。
- **T12 修订**：对话框文案三行式（问题/原因/怎么办）逐字落计划；退出码语义 0 成功/1 环境/2 构建；cmd 包装内 `set PYTHONIOENCODING=gbk`（**run.log 全链路 GBK**，取代原 UTF-8 设想——P5 显式简单）；每次运行写 `=== RUN START ts mode res template=..hash=.. ===` 头 + 快照 temp/config_used.json。
- **T22 修订**：逐条列出待重写命令（AI_GUIDE 速查 #2-#4 加 --no-open --mode B --output temp\smoke.pptx）；使用说明 §四 配置参考补 on_missing_data/hide_blank_pages/inj_pressure_settings 及导出值>config 优先级表；新增"模板契约"与"产物判读指南"（面向维护者/代理，放 AI_GUIDE）。
- **T24 修订**：--open/--no-open 命名对齐并文档化；--mode 补 ALL 选项；--output 与 BOTH 组合时显式拒绝或按方案区分；冒烟门禁固定 --mode B。
- **T28 修订**：高级设置折叠面板暴露 on_missing_data / hide_blank_pages。
- **T20/T4 修订**：image_processor --dir 默认值改相对/必填；BaseDir 回退触发时 WARN。
- **T21 修订**：评估 `ruff check` 纳入 Lint 门禁（先跑基线、差异登记豁免清单，只登记不放宽）。

### DX 完成度总结
> **Phase 2.5 complete.** DX overall: 4/10 → 7.7/10（8 维均分，第 7 维 N/A 除外）。TTHW: ~120 min → 15 min。Codex: 缺位。Claude subagent: 18 项（1 项部分驳回）。Consensus: 2/6 确认，4 项单声部采纳 → 已并入任务。Passing to Phase 3.

## Phase 3 — 工程评审（必经门禁，评审最终修订版；[subagent-only]）

### Step 0 范围挑战
逐条核对了计划引用的实际代码（含 temp/run.log 字节级验证 GBK/UTF-8 混合编码）。范围不缩减（P2）。复杂度：现触达 10 文件（+requirements.txt/check_env.py/tests），新增组件仅 1 个 stdlib 脚本——通过。**工程声部判定：按现稿不可实施，7 处承载性修订必须先落**（全部采纳，见下）。

### 架构依赖图（新增组件与既有关系）
```
[Moldflow Synergy COM]                            [GUI config_gui.py (Tkinter)]
        │ COM 导出                                        │ Popen(wscript) + 句柄/取消
        ▼                                                 ▼
┌─ AutoReport.vbs (GBK) ──────────────────────────► wscript ──► cmd/python(core/pptx_builder.py)
│  写: temp/manifest.json mesh_summary.json                │ --log-file 方式挂接 GBK 日志 (T12 修订)
│      material_info.json peak_values.json(新增,null-on-error) ▼
│      config_used.json 快照(新增)                 core/image_processor.py (拼合,峰值单源读取)
│  清: 上次 mode_a/mode_b/根目录 per-plot PNG+GIF (T23 扩)        │
└──────────────────────────────────────────────► temp/*.{png,gif,json} ◄── run.log (全链路 GBK,
        │                                                │                    RUN START 头+快照)
        ▼                                                ▼
   core/pptx_builder.py ──blip 替换──► 模板 PPTX ──► 交付 .pptx ◄── tests/run_checks.py (函数级回归)
                                                       check_env.py (stdlib 预检,GUI 启动调用)
```
耦合评估：峰值数据流 VBS→JSON→image_processor 两调用点，方向单一无环 ✅；运行状态工件 6 项（递归标志/run.log/config_used/last_output_path/manifest/freshness）**无属主** → A4 修订：T12/T13 定义生命周期（标志 consume-once+时间戳过期）。

### 工程双声部共识表（Codex 缺位 N/A）
```
维度                        Claude  Codex  共识
1. 架构健全?                修订后✓  N/A   5 处承载修订采纳
2. 测试覆盖充分?            修订后✓  N/A   T21 补 5 区块(A10-A14/B1-B4)
3. 性能风险已覆盖?          ✓       N/A   确认
4. 安全威胁已覆盖?          ✓       N/A   确认(日志注入降级 low,T15 足够)
5. 错误路径已处理?          修订后✓  N/A   peak=0 假数据洞补上
6. 部署风险可控?            修订后✓  N/A   T0 单独提交规则
```

### 评审发现与修订（工程声部 13 项，全部采纳）
| # | 严重度 | 发现 | 修订去向 |
|---|---|---|---|
| A1 | critical | T2 峰值时间无导出路径（GetMaxValue 只给值不给时刻，需 GetIndpValues/GetDepValues 数组扫描，未验证）；且原雏形在 On Error Resume Next 内，COM 失败写 `"cae_max_clamp_force": 0`——合法 JSON 里的伪造 0T，恰是 T1 要消灭的病 | **T2 重写**：时刻提取=GetIndpValues/GetDepValues 扫描（实机验证条目进手动清单 D-2）；COM 任何错误→写 null+data_complete:false，**绝不写 0**；T21 加 0/null fixture (A3) |
| A2 | high | T27 taskkill 缺 `/F`（WM_CLOSE 对隐藏进程树无效=取消 visibly 无效）；无 poll() 守卫（PID 重用风险）；取消后残留半截 PPTX 与陈旧 last_output_path；杀不死 Moldflow 本体（继续写 temp）；宏路径启动无状态窗 | **T27 修订**：`taskkill /PID <pid> /T /F`+poll 守卫；取消→删半截输出+失效 last_output_path；"宏路径无状态窗/取消"显式声明范围 |
| A3 | high | T12 cmd /c 引号陷阱（Program Files 空格+中文路径+外层引号剥离）→ 失败=空白 run.log=又回到 F3 症状 | **T12 重写**：放弃 cmd 包装，改为 pptx_builder 增加 `--log-file` 参数（Python 侧挂接 GBK 文件句柄 5 行），env 经 WshShell.Environment("PROCESS") 继承 |
| A4 | medium | 6 个运行状态工件无生命周期属主；崩溃残留递归标志会永久压制 GUI | T12/T13 补生命周期段：标志 consume-once+时间戳过期+VBS 读后即删 |
| B1 | critical | T0 不可机械恢复的乱码（"桑台谭"原意无从得知）——修复=发明用户文案，违反 R-0.2；一次 UTF-8 重写即毁全文；零门禁覆盖 | **T0 修订**：先机械反 mojibake（GBK 字节→按 UTF-8 解码回正）；不可恢复字面量标"待用户确认"绝不臆造；T21 加 VBS 编码+结构门禁（B1/B2）；**T0 单独提交，先于一切逻辑改动** |
| B2 | high | T23 清理范围漏 temp 根目录 per-plot PNG 与 GIF（VBS:436-437 每模式都写根；Slide 4 吃 19MB 根 GIF）；整体清空 mode_b 会误杀 ref_triad/ref_title 且 image_processor.py:526-529 的 exists 缓存条件不改就永不重生成 | **T23 修订**：三位置+*.gif 全清理；同步改 :526-529 缓存条件 |
| B3 | medium | T13 "输出目录可写"检查 os.access 在 Windows 不可靠 | 改为真实临时文件创建/删除探测 |
| B4 | medium | tail 活动文件可能切半个双字节 GBK 字符 → 状态窗 after() 循环崩 | errors="replace" 逐块解码；T21 A10 |
| C1 | high | 门禁离开本机不可复现（temp/ gitignore+模板机器绑定）；T21 的"报告仍生成"fixtures 隐含需要模板 | T21 fixtures 定级为**函数级**（无需模板）；端到端冒烟声明为机器本地 |
| C2 | high | 最危险的 2am 故障（T0/T3/T12 改坏 VBS 使 wscript 拒解析）无任何门禁 | T21 B1/B2 静态门禁（GBK 可解码+无内嵌 UTF-8 中文+块平衡计数） |
| C3 | medium | 四处改动零测试故事：GUI 标志生命周期/tail 解码/T12 命令串/T23 新鲜度/T28 load_config 投毒 | T21 A10-A14 全补 |
| C4 | medium | T2 时刻 API、T7 版本 API、T9 通道均为 Moldflow 侧未验证假设，无实验安排 | 全部进手动清单 D-2/D-4 时限 spike；失败→记录降级路径 |
| E1 | high | **T8-pre 前提错误**：配置分辨率只进 SaveImage3 模型导出，视口抓图（裁剪框的输入）走 SaveImage 与配置无关；且 Moldflow 图例铬层是固定逻辑像素，按比例缩放反而可能破坏当前已调好的框 | **T8-pre 降级**为"前提核实实验"（现窗口尺寸下对比现框 vs 比例框的像素级效果），验证通过才升级为阻塞任务 |
| E2 | medium | T9 SavePlotScaleImage 参数化变体是 API 臆测（现存调用仅路径参数） | 时限 1h spike+书面降级路径（色带维持视口封顶——目标状态措辞已兼容） |
| E3 | medium | T11 首帧调色板：充填动画首帧近白，全局色板将欠采样温度彩虹→逐帧劣化 | 调色板改从**代表性中帧**（或多帧采样并集）计算 |
| E4 | cosmetic | 计划内画布数字三处漂移（2147/2138/2273）；--output×BOTH 悬而未决；线框"(12/31)"进度无任务实现 | 汇聚时统一为一个规范任务表（Phase 4 aggregator）；--output×BOTH=显式拒绝；进度计数要么砍出线框要么 VBS 加计数器导出（采纳：T12 附 per-plot 进度行 "[i/N]"，VBS LogMsg 已有 ExportCount 可扩展） |
| D1 | low | 日志注入被高估：StudyName 来自 Windows 文件名（禁 \/:*?"<>| 与换行），EscapeJson 已处理 JSON 面 | 威胁模型按 low 记录；T15 维持 |

### 失败模式注册表（新增行）
```
peak_values COM 出错   | 合法 JSON 伪造 0T      | 现稿❌CRITICAL → null+data_complete | T21 A3 fixture
取消生成               | 半截 PPTX+陈旧指针      | 现稿❌HIGH → 删输出+失效指针+poll守卫 | 手动清单 D-5
递归标志崩溃残留       | GUI 永久跳过            | 现稿❌MEDIUM → consume-once+过期    | T21 A11
T0 误写 UTF-8          | wscript 乱码全毁        | 现稿❌CRITICAL → 编码+结构静态门禁   | T21 B1/B2
tail 切半字符          | 状态窗崩溃              | 现稿❌MEDIUM → replace 解码         | T21 A10
```

### What already exists（工程面补充）
- `WshShell.Environment("PROCESS")`（VBS 内建）替代 cmd set 链；PIL `Image.quantize(palette=...)` 支持跨帧共用色板（T11 机制现成）；`Popen.poll()` stdlib；`EscapeJson`（VBS:638）已覆盖 JSON 注入面。

### 工程完成度总结
```
+====================================================================+
|      ENG PLAN REVIEW — COMPLETION SUMMARY (Phase 3)                |
+====================================================================+
| S1 架构    | 依赖图产出；6 工件生命周期补属主；耦合无环            |
| S2 质量    | 0 新 DRY/命名问题（计划内修订已吸收）                 |
| S3 测试    | 测试图+工件落盘（xlzxld-master-test-plan-20260906）；  |
|            | 函数级 14 用例+VBS 静态门禁 4+实机清单 5；C1-C3 全修   |
| S4 性能    | tail 轮询/内存/JPEG 均无风险                          |
| 安全       | 攻击面≈0 增量；D1 降级 low；标志文件加过期            |
| 判定       | 修订后可实施；最危险任务 = T0（单提交+门禁兜底）        |
| 未决       | --output×BOTH=显式拒绝；进度计数=T12 扩展              |
+====================================================================+
```
