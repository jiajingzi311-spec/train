# -*- coding: utf-8 -*-
"""
一键打包脚本（带数据保护）
================================
打包前自动备份 dist/MRDN系统/data（你已配置好的节点、截图、输入文档等），
打包后自动恢复，确保系统版本更新时不会清空/重置你的配置数据。

安全机制（2026-11-10 加固）：
  - 备份目录带时间戳（_backups/<kind>_<时间戳>/），**每次打包互不覆盖**。
  - 打包失败时（--clean 可能已误删部分 data），**立即用备份抢救恢复**。
  - 打包成功后才会清理本次备份，失败时备份永久保留，可手动恢复。

用法：
  1. 双击「打包.bat」即可（推荐）
  2. 或命令行：python 打包.py
"""
import datetime
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SPEC = os.path.join(ROOT, 'MRDN系统.spec')
DIST = os.path.join(ROOT, 'dist', 'MRDN系统')
DATA_DIR = os.path.join(DIST, 'data')
BACKUP_ROOT = os.path.join(ROOT, '_backups')   # 所有备份统一放这里，按时间戳分目录
TOOLS_SRC = os.path.join(ROOT, 'tools')
TOOLS_DST = os.path.join(DIST, 'tools')
README = os.path.join(ROOT, '使用说明.md')


def log(msg):
    print(msg, flush=True)


def _stamp():
    # 精确到微秒，避免同秒内多次备份目录名冲突
    return datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')


def _backup(src, kind):
    """把 src 目录备份到 _backups/<kind>_<时间戳>/，返回备份目录路径。

    无源目录时返回 None（跳过备份）。时间戳目录保证每次打包互不覆盖，
    即使后续打包失败，历史完整备份也仍在。
    """
    if not os.path.isdir(src):
        log('[1/4] 未发现已有%s，跳过备份' % kind)
        return None
    os.makedirs(BACKUP_ROOT, exist_ok=True)
    dst = os.path.join(BACKUP_ROOT, '%s_%s' % (kind, _stamp()))
    # 跳过 Excel/WPS 的锁文件（~$ 开头），避免文件被占用时备份失败
    shutil.copytree(src, dst,
                    ignore=lambda d, names: [n for n in names if n.startswith('~$')])
    n = sum(len(fs) for _, _, fs in os.walk(dst))
    log('[1/4] 已备份%s（%d 个文件） -> %s' % (kind, n, dst))
    return dst


def _restore(src, dst, kind):
    """把备份目录 src 合并恢复到 dst（dirs_exist_ok=True，能补回被误删的文件）。"""
    if not src or not os.path.isdir(src):
        log('[4/4] 无%s备份，跳过恢复' % kind)
        return
    log('[4/4] 恢复%s...' % kind)
    os.makedirs(dst, exist_ok=True)
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, dirs_exist_ok=True)
        else:
            shutil.copy2(s, d)
    log('     %s已恢复。' % kind)


def backup_data():
    """打包前备份数据目录，返回备份路径（无数据时返回 None）。"""
    return _backup(DATA_DIR, 'data')


def backup_tools():
    """打包前备份 dist 里的 tools 工具目录，返回备份路径（无工具时返回 None）。"""
    return _backup(TOOLS_DST, 'tools')


def do_build():
    log('[2/4] 开始打包（PyInstaller）... 约需 1-2 分钟')
    r = subprocess.run(
        [sys.executable, '-m', 'PyInstaller', SPEC, '--clean', '--noconfirm'],
        cwd=ROOT,
    )
    if r.returncode != 0:
        log('!! 打包失败，退出码 %d' % r.returncode)
        sys.exit(r.returncode)
    log('[2/4] 打包完成')


def copy_extra():
    log('[3/4] 复制 tools 工具和使用说明...')
    tools_dst = os.path.join(DIST, 'tools')
    if os.path.isdir(TOOLS_SRC):
        os.makedirs(tools_dst, exist_ok=True)
        for f in os.listdir(TOOLS_SRC):
            src = os.path.join(TOOLS_SRC, f)
            if os.path.isfile(src):
                shutil.copy2(src, os.path.join(tools_dst, f))
    if os.path.isfile(README):
        shutil.copy2(README, os.path.join(DIST, os.path.basename(README)))


def restore_data(backup_dir):
    _restore(backup_dir, DATA_DIR, '用户数据')


def restore_tools(backup_dir):
    _restore(backup_dir, TOOLS_DST, 'tools 工具')


def main():
    log('=' * 56)
    log('MRDN系统 一键打包（自动保护已配置数据）')
    log('=' * 56)
    data_backup = backup_data()
    tools_backup = backup_tools()
    try:
        do_build()
        copy_extra()
        restore_data(data_backup)
        restore_tools(tools_backup)
    except SystemExit:
        # do_build 失败：--clean 可能已删除部分 data，立即用备份抢救
        log('!! 打包失败，正在用备份抢救用户数据（防止 --clean 误删）...')
        restore_data(data_backup)
        restore_tools(tools_backup)
        log('   抢救完成，完整备份仍保留在 _backups/ 下，数据安全。')
        raise
    except Exception as e:
        log('!! 出错：%s' % e)
        log('   备份仍保留在 _backups/ 下（时间戳目录），你的数据是安全的。')
        raise
    # 打包成功：数据已恢复到 dist，清理本次备份，避免 _backups 目录膨胀
    for b in (data_backup, tools_backup):
        if b and os.path.isdir(b):
            shutil.rmtree(b)
    log('')
    log('✅ 打包完成！dist/MRDN系统/ 已是最新版，数据保留完整。')
    log('   可以把整个 dist/MRDN系统 文件夹分发给同事了。')


if __name__ == '__main__':
    main()
