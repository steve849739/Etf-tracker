#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETF 公示份额抓取层——纯数据获取，输出 JSON，不碰 Excel（仅 openpyxl 读清单 xlsx）。

数据来源:
  - 上交所: query.sse.com.cn/commonQuery.do (sqlId=COMMON_SSE_ZQPZ_ETFZL_XXPL_ETFGM_SEARCH_L)
  - 深交所: www.szse.cn/api/report/ShowReport/data (CATALOGID=scsj_fund_jjgm&jjlb=ETF)

用法:
  python etf_fetch.py                          -> daily 模式: 各基金最新交易日份额
  python etf_fetch.py --backfill 起 止 [--only 代码] -> 历史区间全量

输出: data/etf_data.json
环境变量: ETF_LIST_PATH(清单路径) ETF_DATA_PATH(JSON输出路径, 默认 data/etf_data.json)
"""
import sys
import os
import json
import re
import time
import logging
import random
from datetime import datetime, date, timedelta

import requests
import openpyxl    # 仅读清单，轻量依赖

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LIST_PATH = os.environ.get("ETF_LIST_PATH", os.path.join(BASE_DIR, "etf清单.xlsx"))
DATA_PATH = os.environ.get("ETF_DATA_PATH", os.path.join(DATA_DIR, "etf_data.json"))
LIST_SHEET = "ETF清单"
BASE_DATE = "2025-12-31"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
HEADERS_SSE = {"User-Agent": UA,
               "Referer": "https://www.sse.com.cn/market/funddata/volumn/etfvolumn/",
               "Accept": "*/*"}
HEADERS_SZSE = {"User-Agent": UA,
                "Referer": "https://www.szse.cn/market/fund/volume/etf/index.html",
                "Accept": "*/*"}
SSE_API_URL = "https://query.sse.com.cn/commonQuery.do?jsonCallBack=callback"
SSE_SQLID = "COMMON_SSE_ZQPZ_ETFZL_XXPL_ETFGM_SEARCH_L"
SZSE_API_URL = "https://www.szse.cn/api/report/ShowReport/data"
SZSE_CATALOGID = "scsj_fund_jjgm"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("etf-fetch")


def load_funds():
    """读取清单: [{name, code, source, url}]。"""
    wb = openpyxl.load_workbook(LIST_PATH)
    ws = wb[LIST_SHEET]
    funds = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * 4
        name, code, source, url = row[0], row[1], row[2], row[3]
        if not code or not str(code).strip():
            continue
        funds.append({"name": str(name).strip() if name else "",
                      "code": str(code).strip(),
                      "source": str(source).strip() if source else "",
                      "url": str(url).strip() if url else ""})
    return funds


def fetch_sse(code: str) -> tuple:
    """上交所: 最新交易日 (date, share万份)。"""
    params = {"isPagination": "true", "pageHelp.pageSize": "100", "pageHelp.pageNo": "1",
              "pageHelp.beginPage": "1", "pageHelp.cacheSize": "1", "pageHelp.endPage": "1",
              "sqlId": SSE_SQLID, "STAT_DATE": ""}
    resp = requests.post(SSE_API_URL, data=params, headers=HEADERS_SSE, timeout=30)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    m = re.match(r"^[^(]*\((.*)\)\s*;?\s*$", resp.text.strip(), re.S)
    if not m:
        raise RuntimeError("上交所接口返回无法解析: " + resp.text[:200])
    data = json.loads(m.group(1))
    rows = data.get("result") or data.get("pageHelp", {}).get("data") or []
    if not rows:
        raise RuntimeError("上交所接口未返回数据")
    t = next((r for r in rows if str(r.get("SEC_CODE", "")).strip() == code), None)
    if not t:
        raise RuntimeError("最新交易日(%s)未找到 %s" % (rows[0].get("STAT_DATE"), code))
    return t["STAT_DATE"], float(t["TOT_VOL"])


def fetch_szse(code: str) -> tuple:
    """深交所: 最新有数据交易日 (date, share万份)。回退最多10天。"""
    d = date.today() - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    for _ in range(10):
        ds = d.strftime("%Y-%m-%d")
        params = {"SHOWTYPE": "JSON", "CATALOGID": SZSE_CATALOGID, "jjlb": "ETF",
                  "txtStart": ds, "txtEnd": ds, "txtDm": code,
                  "tab1PAGENO": "1", "tab1PAGESIZE": "20",
                  "RANDOM": "%.10f" % random.random()}
        resp = requests.get(SZSE_API_URL, params=params, headers=HEADERS_SZSE, timeout=30)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        data = resp.json()
        rows = (data[0].get("data") or []) if isinstance(data, list) and data else []
        for r in rows:
            if str(r.get("fund_code", "")).strip() == code:
                return ds, float(str(r.get("current_size", "")).replace(",", ""))
        d -= timedelta(days=1)
    raise RuntimeError("深交所近10天未找到 %s" % code)


def fetch_sse_history(code: str, start: str, end: str) -> dict:
    """上交所历史: 逐工作日查询, 返回 {日期: 份额}。节日自动跳过。"""
    result = {}
    d = datetime.strptime(start, "%Y-%m-%d").date()
    end_d = datetime.strptime(end, "%Y-%m-%d").date()
    while d <= end_d:
        if d.weekday() < 5:
            ds = d.strftime("%Y-%m-%d")
            params = {"isPagination": "true", "pageHelp.pageSize": "200",
                      "pageHelp.pageNo": "1", "pageHelp.beginPage": "1",
                      "pageHelp.cacheSize": "1", "pageHelp.endPage": "1",
                      "sqlId": SSE_SQLID, "STAT_DATE": ds}
            for attempt in range(3):
                try:
                    resp = requests.post(SSE_API_URL, data=params, headers=HEADERS_SSE, timeout=30)
                    resp.raise_for_status()
                    resp.encoding = "utf-8"
                    m = re.match(r"^[^(]*\((.*)\)\s*;?\s*$", resp.text.strip(), re.S)
                    if m:
                        data = json.loads(m.group(1))
                        rows = data.get("result") or data.get("pageHelp", {}).get("data") or []
                        for r in rows:
                            if str(r.get("SEC_CODE", "")).strip() == code:
                                result[ds] = float(r["TOT_VOL"])
                                break
                    break
                except Exception as e:
                    if attempt == 2:
                        log.warning("[%s] %s 查询失败(3次): %s", code, ds, e)
                    else:
                        time.sleep(1.5)
            time.sleep(0.3)
        d += timedelta(days=1)
    return result


def fetch_szse_history(code: str, start: str, end: str) -> dict:
    """深交所历史: 日期范围+翻页, 返回 {日期: 份额}。"""
    result = {}
    page = 1
    while True:
        params = {"SHOWTYPE": "JSON", "CATALOGID": SZSE_CATALOGID, "jjlb": "ETF",
                  "txtStart": start, "txtEnd": end, "txtDm": code,
                  "tab1PAGENO": str(page), "tab1PAGESIZE": "100",
                  "RANDOM": "%.10f" % random.random()}
        resp = requests.get(SZSE_API_URL, params=params, headers=HEADERS_SZSE, timeout=30)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        data = resp.json()
        meta = data[0]["metadata"]
        rows = data[0].get("data") or []
        for r in rows:
            if str(r.get("fund_code", "")).strip() == code:
                result[r["size_date"]] = float(str(r.get("current_size", "")).replace(",", ""))
        if not rows or len(result) >= int(meta.get("recordcount") or 0):
            break
        page += 1
    return result


def fetch_base_share(fund: dict):
    """查询 BASE_DATE 份额作为基期, 失败返回 None。"""
    try:
        if fund["source"] == "上交所":
            h = fetch_sse_history(fund["code"], BASE_DATE, BASE_DATE)
        elif fund["source"] == "深交所":
            h = fetch_szse_history(fund["code"], BASE_DATE, BASE_DATE)
        else:
            return None
        return h.get(BASE_DATE)
    except Exception:
        return None


def main():
    # ---- 参数解析 ----
    mode = "daily"
    start, end = None, None
    only_code = None
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == "--backfill" and i + 2 < len(args):
            mode, start, end = "backfill", args[i + 1], args[i + 2]
            i += 3
        elif args[i] == "--only" and i + 1 < len(args):
            only_code = args[i + 1]; i += 2
        else:
            i += 1

    os.makedirs(DATA_DIR, exist_ok=True)
    funds = load_funds()
    if only_code:
        funds = [f for f in funds if f["code"] == only_code]

    fetched = []
    for fund in funds:
        code, source = fund["code"], fund["source"]
        try:
            base = fetch_base_share(fund)
            log.info("[%s] 基期(%s): %s", code, BASE_DATE,
                     ("%.2f" % base) if base else "未获取")
        except Exception as e:
            log.warning("[%s] 基期查询失败: %s", code, e)
            base = None

        item = {"name": fund["name"], "code": code, "source": source,
                "base_share": base, "base_date": BASE_DATE}
        try:
            if mode == "daily":
                if source == "上交所":
                    sd, sh = fetch_sse(code)
                elif source == "深交所":
                    sd, sh = fetch_szse(code)
                else:
                    log.error("[%s] 未知来源 %r, 跳过", code, source)
                    item["latest"] = None
                    fetched.append(item)
                    continue
                item["latest"] = {"date": sd, "share": sh}
                log.info("[%s] 最新 %s 份额=%.2f 万份", code, sd, sh)
            else:   # backfill
                if source == "上交所":
                    hist = fetch_sse_history(code, start, end)
                elif source == "深交所":
                    hist = fetch_szse_history(code, start, end)
                else:
                    log.error("[%s] 未知来源 %r", code, source)
                    item["history"] = {}
                    fetched.append(item)
                    continue
                item["history"] = hist if hist else {}
                log.info("[%s] %s~%s 取到 %d 天", code, start, end, len(item["history"]))
        except Exception as e:
            log.error("[%s] 抓取失败: %s", code, e)
            item["_error"] = str(e)
            if mode == "daily":
                item["latest"] = None
            else:
                item["history"] = {}
        fetched.append(item)

    result = {"fetched_at": datetime.now().isoformat(), "mode": mode, "funds": fetched}
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    log.info("数据已写入 %s (%d 基金)", DATA_PATH, len(fetched))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.error("抓取失败: %s", e)
        sys.exit(1)
