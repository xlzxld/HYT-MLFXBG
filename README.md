# HYT-MLFXBG — Moldflow 模流分析报告 (PPT) 自动化生成系统

> 一键把 Autodesk Moldflow 的分析结果（30+ 项云图 / XY 曲线 / 充填 GIF）自动导出、
> 拼合、回填并渲染成中文定制母版的 PPT 报告。
> 适用版本：Moldflow 2023 / 2021 / 2019；支持中性面 / 双层面 / 3D 网格自适应。
> 开源仓库：https://github.com/xlzxld/HYT-MLFXBG

---

## 一、 快速开始

```powershell
# 0. 环境自检 (依赖 / 配置 / 模板标记, 有问题会给出修复指引)
python check_env.py

# 1. 依赖安装 (全新机器)
pip install -r requirements.txt

# 2. 日常使用: 打开可视化配置中心, 选好项后一键生成
python config_gui.py          # 或直接双击 start.bat
```

在 Moldflow 内使用时，运行宏 `AutoReport.vbs`（Tools → Macro → Play Macro），
或直接双击 `start.bat` 走 GUI。四种入口与完整操作说明见 [`使用说明.md`](使用说明.md)。

---

## 二、 目录结构

| 路径 | 作用 |
|---|---|
| `AutoReport.vbs` | Moldflow 宏主入口（GBK 编码，COM 调度 + 结果导出） |
| `config_gui.py` | 可视化配置中心（Tkinter） |
| `core/pptx_builder.py` | PPTX 生成核心引擎（python-pptx） |
| `core/image_processor.py` | 方案 B 图像流水线（色带提取 / 白边裁剪 / 拼合 / 探针） |
| `core/gif_enhancer.py` | 充填动图调色板与体积优化 |
| `report_config.json` | 系统主配置（母版路径 / 截图模式 / 结果项与页码） |
| `templates/` | 报告母版 PPTX（唯一入库母版，配置以相对路径引用） |
| `tests/run_checks.py` | 零依赖回归测试（无需 Moldflow 与模板） |
| `check_env.py` | 环境预检脚本 |
| `temp/` | 运行期中间产物（**已忽略，不入库**） |
| `enforcement/` | 契约机械执法层母版包（见下） |

---

## 三、 文档地图

| 文档 | 读者 | 内容 |
|---|---|---|
| [`使用说明.md`](使用说明.md) | 最终用户（模流工程师） | 功能详解、四种使用方式、配置详解、FAQ |
| [`AI_GUIDE.md`](AI_GUIDE.md) | AI / 接手维护者 | 架构机制、核心模块职责、**10 大技术坑**、命令速查、数据契约与产物判读 |
| [`AGENTS.md`](AGENTS.md) | 所有 AI 编码助手 | 项目契约：铁律 / 行为契约 / 验证门禁 §2 / 红线 / 触发词映射 |
| [`AUDIT-SPEC.md`](AUDIT-SPEC.md) | 说"体检"时 | 只读扫描细则（四大靶心 + P0~P3 分级） |
| [`BOOTSTRAP.md`](BOOTSTRAP.md) | 新项目部署时 | 三件套 + `enforcement/` 的部署与适配流程 |
| [`PROJECT_SESSIONS.md`](PROJECT_SESSIONS.md) | AI 新会话启动 | 会话追踪与断点索引 |
| [`PLAN.md`](PLAN.md) | 历史存档 | 2026-09-06 全方位优化计划与分阶段评审结论 |
| [`enforcement/README.md`](enforcement/README.md) | 维护者 | 执法包四步安装与覆盖对照表 |

---

## 四、 契约与机械执法 (enforcement)

本项目遵循 `AGENTS.md` v2.1.1：**文档管行为，钩子兜底线**。
`enforcement/` 是执法层的**母版包**；其 4 个文件已按 `enforcement/README.md` 落地到本仓库：

| 落地位置 | 作用 | 覆盖的契约规则 |
|---|---|---|
| `.pre-commit-config.yaml` | gitleaks 密钥扫描 + commitlint 提交信息校验 | R-0.4 / R-3.8 |
| `commitlint.config.js` | Conventional Commits 规则 | R-3.8 |
| `Makefile` | `make verify` 聚合门禁（变量已按 §2 填好） | R-0.1 |
| `.github/workflows/gate.yml` | PR 上的 CI 门禁（分支 `master`） | §2 主干保护 |

**待手动激活（涉及新增依赖，尚未执行）**：

```powershell
pip install pre-commit
pre-commit install
pre-commit install --hook-type commit-msg
npm i -D @commitlint/cli @commitlint/config-conventional   # commitlint 钩子以 npx --no 运行
```

**本机 Windows 未安装 `make`**，因此本地闭环请按 `AGENTS.md` §2 逐条执行；
`make verify` 由 CI 在 ubuntu runner 上调用（如需本地一份，`choco install make`）。

> 红线的机械覆盖是**部分**的：gitleaks 覆盖密钥类，主干保护靠 branch protection；
> R-3.1 吞异常、R-3.3 调试残留等仍由提示词纪律 + 人工审查兜底。

---

## 五、 验证门禁

```powershell
# ① 语法/静态检查
python -m py_compile config_gui.py core/gif_enhancer.py core/image_processor.py core/pptx_builder.py check_env.py

# ② 零依赖回归 (无需 Moldflow 与模板)
python tests/run_checks.py

# ③ 格式化检查
python -m ruff format --check config_gui.py core/gif_enhancer.py core/image_processor.py core/pptx_builder.py

# ④ 方案 B 端到端冒烟 (退出码: 0=成功, 1=环境问题, 2=构建失败)
python core/pptx_builder.py --config report_config.json --data-dir temp --mode B --no-open --output temp\smoke.pptx
```

未跑通上述门禁不得声称"已完成"（`AGENTS.md` R-0.1）。

---

## 六、 已知平台绑定

- `AutoReport.vbs` 必须保存为 **GBK/ANSI**；以 UTF-8 覆盖会导致中文字符串判据全部失效。
- `AutoReport.vbs` 内有 BaseDir 绝对路径兜底（脚本目录无 `report_config.json` 时生效），换机器需检查。
- 4K（3840×2160）离屏导出在实机上大面积断带，默认锁 2K；详见 `AI_GUIDE.md` 坑册。
- 模板几何硬编码于 `SLIDE_SAFE_BOXES`，换母版会触发 sha256 告警，须人工核对版式。
