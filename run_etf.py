#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETF 跟踪定时任务入口——编排抓取(fetch) + 更新(update)流程。

用法:
  python run_etf.py [--force]
  python run_etf.py --backfill 起 止 [--only 代码]

环境变量继承给子进程: ETF_LIST_PATH, ETF_TRACK_PATH, ETF_DATA_PATH
"""
import sys
import os
import logging
import subprocess
from datetime import date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("run-etf")


def run_script(script: str, args: list):
    """调用同级脚本。"""
    py = sys.executable
    path = os.path.join(BASE_DIR, script)
    cmd = [py, path] + args
    log.info("执行: %s", " ".join(cmd))
    r = subprocess.run(cmd, cwd=BASE_DIR, capture_output=False, text=True)
    if r.returncode != 0:
        raise RuntimeError("%s 退出码 %d" % (script, r.returncode))


def main():
    force = "--force" in sys.argv
    backfill = "--backfill" in sys.argv

    # 日常模式先判断交易日
    if not backfill:
        today = date.today()
        log.info("运行日期: %s (%s)", today, "一二三四五六日"[today.weekday()])
        if not force and today.weekday() >= 5:
            log.info("今天是周末，非交易日，跳过。")
            return

    # 构造子进程参数
    fetch_args = []
    update_args = []
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--backfill":
            fetch_args += ["--backfill", args[i + 1], args[i + 2]]
            update_args += ["--backfill", args[i + 1], args[i + 2]]
            i += 3
        elif a == "--only":
            fetch_args += ["--only", args[i + 1]]
            i += 2
        else:
            i += 1    # --force 等入口参数不传子进程

    if backfill:
        # update 脚本只需 --backfill 起 止(--only 只影响抓取)
        update_args = update_args[:3] if "--only" in sys.argv else update_args

    # step 1: 抓取
    run_script("etf_fetch.py", fetch_args)
    # step 2: 更新Excel
    run_script("etf_update.py", update_args)
    log.info("ETF跟踪任务完成。")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.error("任务失败: %s", e)
        sys.exit(1)
