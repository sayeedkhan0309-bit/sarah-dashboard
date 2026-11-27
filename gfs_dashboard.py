#!/usr/bin/env python3
import asyncio, json, re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

SCREEN=("ADVANCEGFSCASHSTOCKS","https://chartink.com/screener/advancegfscashstocks","stocks")
OUT_XLSX=Path("gfs_dashboard_latest.xlsx"); OUT_JSON=Path("gfs_dashboard_latest.json")

async def rendered_table(page):
    tables=await page.locator("table").all(); candidates=[]
    for t in tables:
        try:
            h=await t.locator("thead tr th").all_inner_texts()
            if not h: h=await t.locator("tr").first.locator("th,td").all_inner_texts()
            trs=await t.locator("tbody tr").all()
            if not trs: trs=await t.locator("tr").all()
            h=[re.sub(r"\s+"," ",x).strip() for x in h]
            rows=[]
            for tr in trs:
                c=await tr.locator("td").all_inner_texts()
                if not c: c=await tr.locator("th,td").all_inner_texts()
                c=[re.sub(r"\s+"," ",x).strip() for x in c]
                if c and any(c): rows.append(c)
            if h and rows:
                w=len(h); rows=[r[:w]+[""]*max(0,w-len(r)) for r in rows]
                candidates.append((len(rows),h,rows))
        except Exception: pass
    if not candidates: return [],[],None
    _,h,rows=max(candidates,key=lambda x:x[0])
    return [dict(zip(h,r)) for r in rows],h,"rendered_html_table"

async def json_fallback(page):
    captured=[]
    async def on_response(r):
        try:
            if r.request.method.upper()=="POST" and "screener/process" in r.url.lower() and "json" in (r.headers.get("content-type") or "").lower():
                b=await r.json()
                if isinstance(b,dict) and isinstance(b.get("data"),list): captured.append(b)
        except Exception: pass
    page.on("response",on_response)
    try:
        for _ in range(12):
            if captured: break
            await page.wait_for_timeout(2000)
    finally: page.remove_listener("response",on_response)
    if not captured: return [],[],None
    rec=[r for r in captured[-1]["data"] if isinstance(r,dict)]
    fields=list(dict.fromkeys(k for r in rec for k in r.keys()))
    return rec,fields,"chartink_json_response"

async def run(page):
    title,url,kind=SCREEN
    try:
        await page.goto(url,wait_until="domcontentloaded",timeout=60000)
        for _ in range(24):
            rec,fields,method=await rendered_table(page)
            if rec: return {"title":title,"url":url,"kind":kind,"status":"OK","method":method,"records":rec,"raw_fields":fields}
            await page.wait_for_timeout(2500)
        rec,fields,method=await json_fallback(page)
        if rec: return {"title":title,"url":url,"kind":kind,"status":"OK","method":method,"records":rec,"raw_fields":fields}
        return {"title":title,"url":url,"kind":kind,"status":"NO_RESULT_TABLE_CAPTURED","method":None,"records":[],"raw_fields":[]}
    except PlaywrightTimeoutError as e:
        return {"title":title,"url":url,"kind":kind,"status":"PAGE_TIMEOUT","error":str(e),"records":[],"raw_fields":[]}
    except Exception as e:
        return {"title":title,"url":url,"kind":kind,"status":"ERROR","error":str(e),"records":[],"raw_fields":[]}

async def main():
    ts=datetime.now(ZoneInfo("Asia/Kolkata")).isoformat(timespec="seconds")
    async with async_playwright() as pw:
        browser=await pw.chromium.launch(headless=True)
        ctx=await browser.new_context(viewport={"width":1440,"height":1200},user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0 Safari/537.36")
        page=await ctx.new_page()
        print(f"Running {SCREEN[0]} ...",flush=True)
        result=await run(page); result["extracted_at"]=ts
        print(f"{result['status']} | rows={len(result['records'])} | fields={len(result['raw_fields'])} | method={result.get('method')}",flush=True)
        await browser.close()
    with pd.ExcelWriter(OUT_XLSX,engine="openpyxl") as w:
        pd.DataFrame([{"Field":"Extracted at","Value":ts},{"Field":"Timezone","Value":"Asia/Kolkata"},{"Field":"Screener","Value":SCREEN[0]},{"Field":"Status","Value":result["status"]},{"Field":"Capture method","Value":result.get("method","")},{"Field":"Rows","Value":len(result["records"])},{"Field":"Fields","Value":len(result["raw_fields"])}]).to_excel(w,sheet_name="Dashboard",index=False)
        sheet=re.sub(r"[:\\/?*\[\]]","_",SCREEN[0])[:31]
        pd.DataFrame(result["records"],columns=result["raw_fields"]).to_excel(w,sheet_name=sheet,index=False)
    OUT_JSON.write_text(json.dumps({"extracted_at":ts,"timezone":"Asia/Kolkata","source":"Chartink saved screener","screeners":[result]},ensure_ascii=False,indent=2),encoding="utf-8")
    if result["status"]!="OK": raise SystemExit(f"Chartink result capture failed: {result['status']}")

if __name__=="__main__": asyncio.run(main())
