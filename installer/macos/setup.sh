#!/bin/bash
# 矢印 VecStamp - macOS 安装 / 卸载程序
# Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>
# Released under the MIT License.
#
# 由根目录的「安装或卸载-macOS.command」调用，也可以在终端运行：
#   bash installer/macos/setup.sh [install|uninstall]
#
# 只在当前用户目录下操作：
#   · 加载项：~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/VecStamp.ppam
#   · 辅助脚本：~/Library/Application Scripts/com.microsoft.Powerpoint/VecStamp.scpt
#     （PowerPoint for Mac 的 VBA 运行在沙盒中，保存对话框、写出文件、EMZ 压缩、
#       VSDX 打包、取色器和 SVG 备用转换都需要它）
# 不修改任何系统安全设置。

set -u

VERSION="1.2.1"
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$HERE/../.." && pwd)"
PPAM_SRC="$ROOT/dist/VecStamp.ppam"
TEMPLATE="$ROOT/dist/VecStamp-template.pptm"
BAS="$ROOT/dist/VecStamp.bas"
SCRIPT_SRC="$HERE/VecStamp.applescript"
OFFICE_GC="$HOME/Library/Group Containers/UBF8T346G9.Office"
SCRIPT_DIR="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
SCRIPT_DST="$SCRIPT_DIR/VecStamp.scpt"
SANDBOX_TMP="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/VecStampTemp"

RED=$'\033[31m'; GREEN=$'\033[32m'; YELLOW=$'\033[33m'; CYAN=$'\033[36m'; DIM=$'\033[2m'; OFF=$'\033[0m'

say()  { printf '%s\n' "$*"; }
info() { printf '%s%s%s\n' "$CYAN" "$*" "$OFF"; }
ok()   { printf '%s%s%s\n' "$GREEN" "$*" "$OFF"; }
warn() { printf '%s%s%s\n' "$YELLOW" "$*" "$OFF"; }
err()  { printf '%s%s%s\n' "$RED" "$*" "$OFF"; }

ask_yes_no() {   # ask_yes_no "question" default(y|n)
  local hint ans
  if [ "$2" = "y" ]; then hint="[Y/n]"; else hint="[y/N]"; fi
  read -r -p "$1 $hint " ans
  [ -z "$ans" ] && ans="$2"
  case "$ans" in [Yy]*) return 0 ;; *) return 1 ;; esac
}

banner() {
  clear
  say ""
  printf '%s   ════════════════════════════════════════════%s\n' "$RED" "$OFF"
  printf '%s     矢印 VecStamp  ·  PowerPoint 矢量导出加载项%s\n' "$RED" "$OFF"
  say   "     版本 $VERSION  ·  macOS 安装器"
  printf '%s     Copyright (c) 2026 vluckyzhang  ·  MIT License%s\n' "$DIM" "$OFF"
  printf '%s   ════════════════════════════════════════════%s\n' "$RED" "$OFF"
  say ""
}

addin_dir() {
  local d
  for d in "$OFFICE_GC/User Content.localized/Add-Ins.localized" "$OFFICE_GC/User Content/Add-Ins"; do
    if [ -d "$d" ]; then printf '%s' "$d"; return; fi
  done
  printf '%s' "$OFFICE_GC/User Content/Add-Ins"
}

installed_path() {
  local d
  for d in "$OFFICE_GC/User Content.localized/Add-Ins.localized" "$OFFICE_GC/User Content/Add-Ins"; do
    if [ -f "$d/VecStamp.ppam" ]; then printf '%s' "$d/VecStamp.ppam"; return; fi
  done
}

powerpoint_running() { pgrep -x "Microsoft PowerPoint" >/dev/null 2>&1; }

wait_for_powerpoint_quit() {
  while powerpoint_running; do
    warn "PowerPoint 正在运行。请先保存并退出 PowerPoint（⌘Q），然后按回车继续；输入 q 取消。"
    read -r ans
    [ "$ans" = "q" ] && return 1
  done
  return 0
}

show_manual_build() {
  say ""
  info "还没有 dist/VecStamp.ppam。请先用 PowerPoint for Mac 生成一次（约 1 分钟）："
  say "  1. 用 PowerPoint 打开：$TEMPLATE"
  say "  2. 菜单「工具 → 宏 → Visual Basic 编辑器」，然后「文件 → 导入文件…」选择："
  say "     $BAS"
  say "  3. 菜单「调试 → 编译 VBAProject」，没有提示即代码正常；关闭 VBA 编辑器"
  say "  4. 「文件 → 另存为…」，文件格式选「PowerPoint 加载项 (.ppam)」，"
  say "     保存为 $ROOT/dist/VecStamp.ppam"
  say "  5. 退出 PowerPoint，重新运行本安装器"
  say ""
  say "也可以直接使用在 Windows 上生成、或从项目 Release 下载的 VecStamp.ppam，放到 dist 文件夹即可。"
}

do_install() {
  if [ ! -f "$PPAM_SRC" ]; then
    show_manual_build
    return 1
  fi
  wait_for_powerpoint_quit || return 1

  local dir dst
  dir="$(addin_dir)"
  dst="$dir/VecStamp.ppam"
  info "[1/2] 复制加载项到 Office 加载项文件夹…"
  mkdir -p "$dir" || { err "无法创建文件夹：$dir"; return 1; }
  cp -f "$PPAM_SRC" "$dst" || { err "复制失败：$dst"; return 1; }
  say "      $dst"

  info "[2/2] 安装 macOS 辅助脚本（保存对话框、写出文件、EMZ / VSDX 打包、取色器）…"
  mkdir -p "$SCRIPT_DIR" || { err "无法创建文件夹：$SCRIPT_DIR"; return 1; }
  if ! /usr/bin/osacompile -o "$SCRIPT_DST" "$SCRIPT_SRC"; then
    err "编译 AppleScript 失败。"
    return 1
  fi
  say "      $SCRIPT_DST"

  say ""
  ok "文件已就位。最后一步需要在 PowerPoint 里启用一次加载项："
  say "  1. 打开 PowerPoint，菜单「工具 → PowerPoint 加载项…」"
  say "  2. 点「+」，选择上面的 VecStamp.ppam（访达中按 ⌘⇧G 可以粘贴路径）"
  say "  3. 确认 VecStamp 已勾选，点「确定」；如果询问是否启用宏，选「启用宏」"
  say "  4. 功能区出现「矢印」选项卡即可使用。以后打开 PowerPoint 会自动加载。"
  say ""
  if ask_yes_no "现在在访达中显示 VecStamp.ppam 并打开 PowerPoint 吗？" y; then
    open -R "$dst"
    open -a "Microsoft PowerPoint" 2>/dev/null || warn "没有找到 Microsoft PowerPoint，请手动打开。"
  fi
  return 0
}

do_uninstall() {
  ask_yes_no "确定要卸载矢印 VecStamp 吗？" n || return 0
  wait_for_powerpoint_quit || return 1
  local d removed=0
  for d in "$OFFICE_GC/User Content.localized/Add-Ins.localized" "$OFFICE_GC/User Content/Add-Ins"; do
    if [ -f "$d/VecStamp.ppam" ]; then rm -f "$d/VecStamp.ppam" && say "  已删除 $d/VecStamp.ppam"; removed=1; fi
  done
  if [ -f "$SCRIPT_DST" ]; then rm -f "$SCRIPT_DST" && say "  已删除 $SCRIPT_DST"; fi
  if [ -d "$SANDBOX_TMP" ]; then rm -rf "$SANDBOX_TMP" && say "  已删除临时文件夹"; fi
  say ""
  ok "卸载完成。"
  if [ "$removed" = "1" ]; then
    say "下次打开 PowerPoint 时，如果「工具 → PowerPoint 加载项…」里还列着 VecStamp，选中后点「−」移除即可。"
  fi
  return 0
}

status_line() {
  local p
  p="$(installed_path)"
  if [ -n "$p" ]; then
    ok "   当前状态：已安装（$p）"
  else
    warn "   当前状态：未安装"
  fi
}

case "${1:-menu}" in
  install)   banner; do_install;   exit $? ;;
  uninstall) banner; do_uninstall; exit $? ;;
esac

while true; do
  banner
  status_line
  say ""
  say "   [1] 安装 / 更新"
  say "   [2] 卸载"
  say "   [0] 退出"
  say ""
  read -r -p "   请输入数字后回车：" choice
  say ""
  case "$choice" in
    1) do_install ;;
    2) do_uninstall ;;
    0) exit 0 ;;
    *) continue ;;
  esac
  say ""
  read -r -p "按回车键返回菜单" _
done
