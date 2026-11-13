# -*- coding: utf-8 -*-
"""
MRDN 接口下载（mrdn_download）
=================================
功能：
    从接口（预签名下载 URL / 接口地址）下载文件，保存到固定目录
    （默认 data/ML_design/exp，相对 exe 的上一级），保存前先清空该目录下所有文件。
    文件名优先从响应头 Content-Disposition 推断，其次从 URL 推断。
依赖：
    pip install requests
"""


# ============================================================================
#  ⚙ MRDN 系统集成参数对照表（配置系统「工具输入关联」时，照此填「参数名」）
# ============================================================================
#   参数名（填到系统）       对应文件 / 含义
#   --url                    下载链接（预签名 URL / 接口地址，必填）
#   --save_dir               保存目录（可选，默认 ../data/ML_design/exp 相对 exe）
#   --referer                防盗链 Referer（可选，如 https://share.example.com/）
#   （--out_dir 为系统自动传入的输出目录，可选；传了会额外复制一份过去便于系统预览）
# ============================================================================

import os
import re
import sys
import shutil
import argparse
from urllib.parse import urlparse, parse_qs, unquote

import requests
import urllib3

# 企业内网常为自签名证书，关闭告警
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 默认保存目录：exe 位于 tools/，上一级 ../data/ML_design/exp
DEFAULT_SAVE_DIR = os.path.join(BASE_DIR, '..', 'data', 'ML_design', 'exp')

DEFAULT_HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36'),
}


def _configure_stdout():
    """Windows 打包后默认 GBK，中文 print 会崩；统一转 UTF-8。"""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args():
    p = argparse.ArgumentParser(description="MRDN 接口下载")
    p.add_argument("--url", default=None, help="下载链接（预签名 URL / 接口地址）")
    p.add_argument("--save_dir", default=None, help="保存目录（默认 ../data/ML_design/exp）")
    p.add_argument("--referer", default=None, help="防盗链 Referer（可选）")
    p.add_argument("--out_dir", default=None, help="输出目录（系统自动传入，可选）")
    return p.parse_args()


def _extract_filename(text):
    """从一段 Content-Disposition 文本里解析文件名，支持 RFC5987 filename* 与普通 filename。"""
    if not text:
        return None
    m = re.search(r"filename\*\s*=\s*(?:utf-8|UTF-8)''([^;]+)", text)
    if m:
        name = unquote(m.group(1).strip())
        if name:
            return os.path.basename(name)
    m = re.search(r'filename\s*=\s*"?([^";]+)"?', text)
    if m:
        name = m.group(1).strip()
        if name:
            return os.path.basename(name)
    return None


def guess_filename(url, resp):
    """推断下载文件名：响应头 Content-Disposition -> URL 参数 -> URL path -> 默认。"""
    cd = resp.headers.get('Content-Disposition', '') if resp is not None else ''
    name = _extract_filename(cd)
    if name:
        return name
    qs = parse_qs(urlparse(url).query)
    rcd = (qs.get('response-content-disposition') or [''])[0]
    name = _extract_filename(rcd)
    if name:
        return name
    name = os.path.basename(unquote(urlparse(url).path))
    if name and name != '/':
        return name
    return '下载文件.xlsx'


def clear_dir(d):
    """清空目录下所有文件和子目录（不删目录本身）。"""
    if not os.path.isdir(d):
        return
    for fn in os.listdir(d):
        p = os.path.join(d, fn)
        try:
            if os.path.isfile(p) or os.path.islink(p):
                os.remove(p)
            elif os.path.isdir(p):
                shutil.rmtree(p)
        except Exception:
            pass


def main():
    _configure_stdout()
    args = parse_args()

    url = (args.url or '').strip()
    if not url:
        print("[错误] 未提供下载链接，请通过 --url 传入")
        sys.exit(1)

    save_dir = os.path.abspath(args.save_dir or DEFAULT_SAVE_DIR)
    os.makedirs(save_dir, exist_ok=True)

    headers = dict(DEFAULT_HEADERS)
    if args.referer:
        headers['Referer'] = args.referer

    print("正在下载：", url)
    try:
        resp = requests.get(url, headers=headers, verify=False, timeout=120, allow_redirects=True)
        resp.raise_for_status()
    except Exception as e:
        print("[错误] 下载失败：%s" % e)
        sys.exit(1)

    filename = guess_filename(url, resp)
    # 清理 Windows 非法文件名字符
    filename = re.sub(r'[\\/:*?"<>|]', '_', filename).strip()
    if not filename:
        filename = '下载文件.xlsx'

    # 保存前先清空目标目录
    clear_dir(save_dir)
    dst = os.path.join(save_dir, filename)
    with open(dst, 'wb') as f:
        f.write(resp.content)

    print("已保存文件：", dst)
    print("文件大小：%d 字节" % len(resp.content))

    # MRDN 兼容：若系统传了 --out_dir，额外复制一份过去，便于系统收集/预览
    if args.out_dir:
        out_dir = os.path.abspath(args.out_dir)
        os.makedirs(out_dir, exist_ok=True)
        shutil.copy2(dst, os.path.join(out_dir, filename))
        print("已复制到输出目录：", out_dir)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        _configure_stdout()
        print("[错误] 运行失败：%s" % e)
        sys.exit(1)
