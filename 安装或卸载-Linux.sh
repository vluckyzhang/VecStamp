#!/bin/bash
# 矢印 VecStamp - Linux 安装 / 卸载（LibreOffice 扩展）。用法：bash 安装或卸载-Linux.sh
cd "$(dirname "$0")" || exit 1
exec /bin/bash "./installer/linux/setup.sh" "$@"
