# make verify —— AGENTS.md §2 验证门禁的机械执法层（闭环验证的唯一入口）
# 用法：按目标项目 AGENTS.md §2 登记的命令填充以下变量，然后执行 `make verify`
# §2 登记"无"的项：对应变量填 skip（留空 = 未配置，报错退出，防误配静默放行）
# 增量格式检查（大仓推荐）：make verify CHANGED="$(git diff --name-only)"，仅对改动文件执行 fmt-check
# ---------------------------------------------------------------------------
# 本项目 (HYT-MLFXBG) 取值来源：AGENTS.md §2，改动须与 §2 同批同步
# 注意：本机 Windows 未安装 make，本地闭环请逐条执行 §2 三条命令
#       （py_compile / tests/run_checks.py / ruff format --check + 冒烟）；
#       `make verify` 由 .github/workflows/gate.yml 在 ubuntu runner 上调用。
FMT_CHECK_CMD ?= python -m ruff format --check config_gui.py core/gif_enhancer.py core/image_processor.py core/pptx_builder.py check_env.py tests/run_checks.py
LINT_CMD ?= python -m py_compile config_gui.py core/gif_enhancer.py core/image_processor.py core/pptx_builder.py check_env.py tests/run_checks.py
TEST_CMD ?= python tests/run_checks.py
BUILD_CMD ?= skip
CHANGED ?=

.PHONY: verify fmt-check lint test build
verify: fmt-check lint test build

fmt-check:
	@if [ -z "$(FMT_CHECK_CMD)" ]; then echo "FMT_CHECK_CMD 未配置（见 enforcement/README.md）"; exit 1; fi
	@if [ "$(FMT_CHECK_CMD)" = "skip" ]; then echo "skip: fmt-check（§2 登记“无”）"; \
	elif [ -n "$(CHANGED)" ]; then $(FMT_CHECK_CMD) $(CHANGED); \
	else $(FMT_CHECK_CMD); fi

lint:
	@if [ -z "$(LINT_CMD)" ]; then echo "LINT_CMD 未配置（见 enforcement/README.md）"; exit 1; fi
	@if [ "$(LINT_CMD)" = "skip" ]; then echo "skip: lint（§2 登记“无”）"; else $(LINT_CMD); fi

test:
	@if [ -z "$(TEST_CMD)" ]; then echo "TEST_CMD 未配置（见 enforcement/README.md）"; exit 1; fi
	@if [ "$(TEST_CMD)" = "skip" ]; then echo "skip: test（§2 登记“无”）"; else $(TEST_CMD); fi

build:
	@if [ "$(BUILD_CMD)" = "skip" ]; then echo "skip: no build configured"; else $(BUILD_CMD); fi

