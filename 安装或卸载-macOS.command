#!/bin/bash
# 矢印 VecStamp - macOS 安装 / 卸载（双击运行；首次可能需要右键 → 打开）
cd "$(dirname "$0")" || exit 1
exec /bin/bash "./installer/macos/setup.sh" "$@"
