# -*- coding: utf-8 -*-
"""
一键打包脚本（带数据保护）
================================
打包前自动备份 dist/MRDN系统/data（你已配置好的节点、截图、输入文档等），
打包后自动恢复，确保系统版本更新时不会清空/重置你的配置数据。

用法：
  1. 双击「打包.bat」即可（推荐）
  2. 或命令行：python 打包.py
"""
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SPEC = os.path.join(ROOT, 'MRDN系统.spec')
DIST = os.path.join(ROOT, 'dist', 'MRDN系统')
DATA_DIR = os.path.join(DIST, 'data')
BACKUP_DIR = os.path.join(ROOT, '_data_backup')
TOOLS_SRC = os.path.join(ROOT, 'tools')
README = os.path.join(ROOT, '使用说明.md')


def log(msg):
    print(msg, flush=True)


def backup_data():
    """打包前备份数据目录，返回是否成功备份"""
    if not os.path.isdir(DATA_DIR):
        log('[1/4] 未发现已有用户数据，跳过备份')
        return False
    if os.path.isdir(BACKUP_DIR):
        shutil.rmtree(BACKUP_DIR)
    shutil.copytree(DATA_DIR, BACKUP_DIR)
    # 统计备份内容，方便用户确认
    n = sum(len(fs) for _, _, fs in os.walk(BACKUP_DIR))
    log(f'[1/4] 已备份用户数据（{n} 个文件） -> _data_backup/')
    return True


def do_build():
    log('[2/4] 开始打包（PyInstaller）... 约需 1-2 分钟')
    r = subprocess.run(
        [sys.executable, '-m', 'PyInstaller', SPEC, '--clean', '--noconfirm'],
        cwd=ROOT,
    )
    if r.returncode != 0:
        log(f'!! 打包失败，退出码 {r.returncode}（用户数据仍安全保存在 _data_backup/）')
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


def restore_data():
    if not os.path.isdir(BACKUP_DIR):
        log('[4/4] 无备份数据，跳过恢复')
        return
    log('[4/4] 恢复用户数据...')
    os.makedirs(DATA_DIR, exist_ok=True)
    for item in os.listdir(BACKUP_DIR):
        s = os.path.join(BACKUP_DIR, item)
        d = os.path.join(DATA_DIR, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, dirs_exist_ok=True)
        else:
            shutil.copy2(s, d)
    shutil.rmtree(BACKUP_DIR)
    log('     用户数据已恢复，配置未丢失。')


def main():
    log('=' * 56)
    log('MRDN系统 一键打包（自动保护已配置数据）')
    log('=' * 56)
    had_data = backup_data()
    try:
        do_build()
        copy_extra()
        restore_data()
    except SystemExit:
        raise
    except Exception as e:
        log(f'!! 出错：{e}')
        if had_data and os.path.isdir(BACKUP_DIR):
            log('   备份仍保存在 _data_backup/，你的数据是安全的。')
        raise
    log('')
    log('✅ 打包完成！dist/MRDN系统/ 已是最新版，数据保留完整。')
    log('   可以把整个 dist/MRDN系统 文件夹分发给同事了。')


if __name__ == '__main__':
    main()
