# 规范部署与适配生成器 (BOOTSTRAP.md)

> **版本 v2.1.1** | 与同目录 AGENTS.md v2.1.1、AUDIT-SPEC.md 配套。
> **单源原则**：本文件**不内嵌**契约模板。母版 = 同目录的 `AGENTS.md` + `AUDIT-SPEC.md`。契约只有一个真身，禁止再复制出第三份。
> **使用方法**：将本目录三个文件整体拷贝到目标项目后说"适配 / 初始化规范"；或提供本目录路径给 AI 并说"读取通用母版，自动扫描当前项目，生成定制化 AGENTS.md。有无法确定的配置再问我。"

---

## 阶段 0：探测目标项目特征（静默执行，不输出废话）

1. **语言与运行时**：检查根目录 `package.json` (Node/TS)、`Cargo.toml` (Rust)、`go.mod` (Go)、`pyproject.toml`/`requirements.txt` (Python)、`pom.xml`/`build.gradle` (Java)、`*.csproj` (C#)。
2. **真实门禁命令**：
   - 测试：`npm test` / `pytest` / `go test ./...` / `cargo test` / `python -m unittest discover`
   - 静态检查：`npm run lint` / `ruff check` / `golangci-lint run` / `cargo clippy`
   - 格式化：**只登记 `--check` 等价命令**（`prettier --check` / `black --check` / `cargo fmt --check`），禁止登记写模式命令
   - 构建：`npm run build` / `cargo build`
3. **主干分支**：`git branch` / `git remote show origin`。
4. **裸源码兜底**：无任何 manifest 时按文件后缀推断语言（`.py`→Python、`.ts`/`.js`→Node、`.go`→Go、`.rs`→Rust），并在 §2 如实登记探测依据，禁止臆造。

## 阶段 1：模式判定（三选一，禁止混淆）

- **模式 A（新项目部署）**：目标项目根目录**无** AGENTS.md → 拷贝母版**三件套**（`AGENTS.md` + `AUDIT-SPEC.md` + `BOOTSTRAP.md`）及 `enforcement/` 目录到项目根目录，再执行阶段 3 填充。
- **模式 B（已有契约适配）**：目标项目根目录**已有** AGENTS.md → **仅重算其 §2 表格**，输出修改前后 diff，等待确认后才写入；禁止全文覆盖，禁止触碰 §0 / §1 / §3 / §4 / §5。
- **模式 C（契约版本升级）**：目标项目已有**旧版本**本契约 → 三件套与 `enforcement/` 整体同步升级，仅保留并重算其 §2 项目取值，输出版本级 diff，确认后写入；禁止静默丢弃项目定制。

## 阶段 2：人机对齐门禁（极简）

仅当存在**多重工具冲突**（既装 Jest 又装 Vitest、多语言 Monorepo）才提问，最多 1~2 个选择题；严禁拿显而易见的信息打扰用户。

## 阶段 3：填充与校验

1. 用阶段 0 探测结果填充 AGENTS.md §2；探测不到的项如实写"无"或向用户提问，**禁止臆造**。
2. 校验（模式 A / B / C 均适用）：AGENTS.md 行数 < 80（`wc -l`，LF 计）；§2 每条命令在终端实测可执行（附退出码）。
3. 目标项目存在基线失败（存量 lint 警告、跳过测试）时，登记进"已知豁免清单"行——**只登记，不放宽验证条件**。

## 阶段 4：完成反馈（如实陈述，禁止绝对化承诺）

输出三件事：部署/更新了哪些文件；§2 各项取值的来源（探测文件 / 用户确认）；遗留未决项清单。

> 禁止使用"AI 将全自动遵守此契约"等话术——契约降低违规概率，机械执法兜住底线：
> 密钥泄露 → gitleaks pre-commit；Conventional Commits → commitlint；主干保护 → branch protection；闭环验证 → 一条 `verify` 聚合命令（如 `make verify`）。
