#!/bin/bash
# 矢印 VecStamp - Linux 安装 / 卸载程序（LibreOffice Impress / Draw 扩展）
# Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>
# Released under the MIT License.
#
# Microsoft PowerPoint 没有 Linux 版，Linux 上的矢印以 LibreOffice 扩展的形式提供，
# 功能与 PowerPoint 版一致（EMF / EMZ / SVG / WMF / PDF / VSDX / PNG、背景、逐页导出）。
#
# 由根目录的「安装或卸载-Linux.sh」调用，也可以直接运行：
#   bash installer/linux/setup.sh [install|uninstall|status]
# 只为当前用户安装（unopkg 用户模式），不需要 sudo，也不修改系统设置。

set -u

VERSION="1.2.1"
EXT_ID="com.vluckyzhang.vecstamp"
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$HERE/../.." && pwd)"
OXT="$ROOT/dist/VecStamp-LibreOffice.oxt"
SETTINGS_DIR="${XDG_CONFIG_HOME:-$HOME/.config}/VecStamp"

RED=$'\033[31m'; GREEN=$'\033[32m'; YELLOW=$'\033[33m'; CYAN=$'\033[36m'; DIM=$'\033[2m'; OFF=$'\033[0m'
say()  { printf '%s\n' "$*"; }
info() { printf '%s%s%s\n' "$CYAN" "$*" "$OFF"; }
ok()   { printf '%s%s%s\n' "$GREEN" "$*" "$OFF"; }
warn() { printf '%s%s%s\n' "$YELLOW" "$*" "$OFF"; }
err()  { printf '%s%s%s\n' "$RED" "$*" "$OFF"; }

ask_yes_no() {
  local hint ans
  if [ "$2" = "y" ]; then hint="[Y/n]"; else hint="[y/N]"; fi
  read -r -p "$1 $hint " ans
  [ -z "$ans" ] && ans="$2"
  case "$ans" in [Yy]*) return 0 ;; *) return 1 ;; esac
}

banner() {
  [ -t 1 ] && clear
  say ""
  printf '%s   ════════════════════════════════════════════%s\n' "$RED" "$OFF"
  printf '%s     矢印 VecStamp  ·  LibreOffice 矢量导出扩展%s\n' "$RED" "$OFF"
  say   "     版本 $VERSION  ·  Linux 安装器"
  printf '%s     Copyright (c) 2026 vluckyzhang  ·  MIT License%s\n' "$DIM" "$OFF"
  printf '%s   ════════════════════════════════════════════%s\n' "$RED" "$OFF"
  say ""
}

UNOPKG=()
find_unopkg() {
  local c
  if command -v unopkg >/dev/null 2>&1; then UNOPKG=(unopkg); return 0; fi
  for c in /usr/lib/libreoffice/program/unopkg /usr/lib64/libreoffice/program/unopkg \
           /opt/libreoffice*/program/unopkg /usr/local/lib/libreoffice/program/unopkg; do
    if [ -x "$c" ]; then UNOPKG=("$c"); return 0; fi
  done
  if command -v libreoffice.unopkg >/dev/null 2>&1; then UNOPKG=(libreoffice.unopkg); return 0; fi
  if command -v flatpak >/dev/null 2>&1 && flatpak info org.libreoffice.LibreOffice >/dev/null 2>&1; then
    UNOPKG=(flatpak run --command=unopkg org.libreoffice.LibreOffice); return 0
  fi
  return 1
}

has_python_support() {
  # /usr/bin/unopkg is often a wrapper script, so look in the usual program folders too.
  local exe d
  exe="$(readlink -f "$(command -v "${UNOPKG[0]}" 2>/dev/null || echo "${UNOPKG[0]}")")"
  for d in "$(dirname "$exe")" /usr/lib/libreoffice/program /usr/lib64/libreoffice/program \
           /opt/libreoffice*/program /usr/local/lib/libreoffice/program; do
    if [ -e "$d/pythonloader.py" ] || [ -e "$d/libpythonloaderlo.so" ]; then return 0; fi
  done
  return 1
}

python_support_hint() {
  case "${UNOPKG[0]}" in flatpak|libreoffice.unopkg) return 0 ;; esac   # bundled with Python
  if ! has_python_support; then
    warn "没有找到 LibreOffice 的 Python 支持，矢印需要它才能运行。请先安装："
    say  "  Debian / Ubuntu / Deepin / UOS：sudo apt install python3-uno libreoffice-script-provider-python"
    say  "  Fedora / openEuler：           sudo dnf install libreoffice-pyuno"
    say  "  openSUSE：                     sudo zypper install libreoffice-pyuno"
    say  "  Arch / Manjaro：               LibreOffice 自带，无需额外安装"
    say  ""
    ask_yes_no "仍然继续安装吗？" n || return 1
  fi
  return 0
}

office_running() { pgrep -x soffice.bin >/dev/null 2>&1 || pgrep -x soffice >/dev/null 2>&1; }

wait_for_office_quit() {
  local ans
  while office_running; do
    warn "LibreOffice 正在运行。请保存并关闭所有 LibreOffice 窗口，然后按回车继续；输入 q 取消。"
    read -r ans
    [ "$ans" = "q" ] && return 1
  done
  return 0
}

# Read the whole listing first: "grep -q" would close the pipe early, and an interrupted
# unopkg leaves a stale lock file behind in the user profile.
is_installed() {
  local listing
  listing="$("${UNOPKG[@]}" list 2>/dev/null)"
  case "$listing" in *"Identifier: $EXT_ID"*) return 0 ;; *) return 1 ;; esac
}

lock_hint() {
  local lock
  for lock in "${XDG_CONFIG_HOME:-$HOME/.config}"/libreoffice/*/.lock; do
    [ -e "$lock" ] || continue
    office_running && continue
    warn "提示：LibreOffice 没有运行，但锁文件仍在：$lock"
    say  "  这通常是上次异常退出留下的。确认所有 LibreOffice 窗口都已关闭后，可以删除它再重试。"
  done
}

do_install() {
  if [ "$(id -u)" = "0" ]; then
    err "请不要用 root / sudo 运行：扩展只为当前用户安装。"
    return 1
  fi
  if [ ! -f "$OXT" ]; then
    err "缺少 $OXT ，请先完整解压项目，或运行 python3 build/build.py 生成。"
    return 1
  fi
  find_unopkg || { err "没有找到 LibreOffice（unopkg）。请先安装 LibreOffice 6.4 或更高版本。"; return 1; }
  python_support_hint || return 1
  wait_for_office_quit || return 1
  info "正在安装 LibreOffice 扩展（${UNOPKG[*]} add）…"
  if ! "${UNOPKG[@]}" add -f "$OXT"; then
    err "安装失败。"
    lock_hint
    return 1
  fi
  say ""
  ok "安装完成！"
  say "  打开 LibreOffice Impress 或 Draw，菜单栏会出现「矢印」，同时多一条「矢印」工具栏"
  say "  （如果没看到工具栏：菜单「视图 → 工具栏 → 矢印」）。"
  say "  设置保存在：$SETTINGS_DIR"
  return 0
}

do_uninstall() {
  find_unopkg || { err "没有找到 LibreOffice（unopkg）。"; return 1; }
  ask_yes_no "确定要卸载矢印 VecStamp 吗？" n || return 0
  wait_for_office_quit || return 1
  if is_installed; then
    "${UNOPKG[@]}" remove "$EXT_ID" && say "  已移除 LibreOffice 扩展 $EXT_ID"
  else
    say "  扩展没有安装。"
  fi
  if [ -d "$SETTINGS_DIR" ] && ask_yes_no "同时删除导出设置（$SETTINGS_DIR）吗？" y; then
    rm -rf "$SETTINGS_DIR" && say "  已删除导出设置"
  fi
  say ""
  ok "卸载完成。"
  return 0
}

status_line() {
  if find_unopkg && is_installed; then
    ok "   当前状态：已安装（LibreOffice 扩展 $EXT_ID）"
  elif [ "${#UNOPKG[@]}" -eq 0 ]; then
    warn "   当前状态：没有找到 LibreOffice"
  else
    warn "   当前状态：未安装"
  fi
}

case "${1:-menu}" in
  install)   banner; do_install;   exit $? ;;
  uninstall) banner; do_uninstall; exit $? ;;
  status)    status_line; exit 0 ;;
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
