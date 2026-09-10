# 执法包 (Enforcement)

> 对应 AGENTS.md v2.1.1 的机械执法层：**文档管行为，钩子兜底线**。
> 红线中可机器判定且已落地的项（密钥、提交格式、门禁、主干保护）在此落地；其余红线（如 R-3.1 吞异常、R-3.3 调试残留）由提示词纪律与人工审查兜底。

## 文件去向

| 模板 | 拷贝到目标项目 | 作用 |
|---|---|---|
| `.pre-commit-config.yaml` | 项目根 | pre-commit 框架入口：gitleaks 密钥扫描 + commitlint 提交信息校验 |
| `commitlint.config.js` | 项目根 | Conventional Commits 规则 |
| `Makefile` | 项目根 | `make verify` 聚合门禁（AI 闭环验证的唯一入口） |
| `gate.yml` | `.github/workflows/gate.yml` | PR 上的 CI 门禁 |

## 四步安装

1. **拷贝**：按上表把四个文件复制到目标项目对应位置。
2. **挂钩**：`pip install pre-commit && pre-commit install && pre-commit install --hook-type commit-msg`；commitlint 钩子以 `npx --no` 运行（不联网下载），须先本地安装：`npm i -D @commitlint/cli @commitlint/config-conventional`
3. **填变量**：编辑 Makefile 顶部，与 AGENTS.md §2 登记的命令保持一致：
   - `FMT_CHECK_CMD`（如 `npx prettier --check .` / `ruff format --check .`）
   - 大仓存量格式问题时改用增量检查（与 §2"仅检查本次改动文件"对齐）：`make verify CHANGED="$(git diff --name-only)"`，仅对改动文件执行 fmt-check
   - `LINT_CMD`（如 `npm run lint` / `ruff check .`）
   - `TEST_CMD`（如 `npm test` / `pytest`）
   - §2 登记"无"的项（无 format / lint / test / build）：对应变量填 `skip`，`make verify` 输出 skip 并通过；**留空视为未配置，报错退出**（防误配静默放行）
4. **分支保护**（GitHub：Settings → Branches → Add branch protection rule）：
   - Branch name pattern 填主干分支名
   - 勾选 Require a pull request before merging
   - 勾选 Require status checks to pass → 选中 `gate`
   - 勾选 Block force pushes

## 日常使用

- AI 完成修改后执行 `make verify`，全绿才算闭环（AGENTS.md R-0.1）
- 升级 gitleaks：`pre-commit autoupdate`
- 离线环境：步骤 2 已将 commitlint 本地化（`npx --no` 不访问网络，无需额外处理）；若偏好在线按需下载，把 `.pre-commit-config.yaml` 中 entry 改为 `npx -- commitlint --edit`

## 覆盖对照

| 契约规则 | 执法手段 |
|---|---|
| R-0.4 禁止泄露密钥 | gitleaks（密钥类部分覆盖；内网域名 / 连接串须自定义规则或人工审查） |
| R-3.8 原子提交 / Conventional Commits | commitlint（commit-msg 钩子） |
| R-0.1 未验证不交付 | `make verify` 聚合门禁 |
| 主干保护（§2） | branch protection + gate CI |
