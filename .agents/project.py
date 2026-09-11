# -*- coding: utf-8 -*-
"""本项目自己的门禁命令（tsc.py verify 会按顺序执行这里登记的 4 条命令）。

取值来源：AGENTS.md §2（check-config 校验同源）。
"""
__all__ = ["FMT_CHECK_CMD", "LINT_CMD", "TEST_CMD", "BUILD_CMD"]

FMT_CHECK_CMD = "python -m ruff format --check config_gui.py core/gif_enhancer.py core/image_processor.py core/pptx_builder.py"
LINT_CMD = "python -m py_compile config_gui.py core/gif_enhancer.py core/image_processor.py core/pptx_builder.py check_env.py"
TEST_CMD = "python tests/run_checks.py"
BUILD_CMD = None
