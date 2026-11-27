#!/usr/bin/env python3
import asyncio, json, re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

SCREEN=("ADVANCEGFSCASHSTOCKS","https://chartink.com/screener/advancegfscashstocks","stocks")
OUT_XLSX=Path("gfs_dashboard_latest.xlsx")
OUT_JSON=Path("gfs_dashboard_latest.json")


def clean(v):
    return re.sub(r"\s+", " ", str(v)).strip()


def records_from_json(obj):
    if isinstance(obj, dict):
        for key in ("data", "results", "rows"):
            val = obj.get(key)
            if isinstance(val, list) and val and all(isinstance(x, dict) for x in val):
                return val
        for val in obj.values():
            found = records_from_json(val)
            if found:
                return found
    elif isinstance(obj, list) and obj and all(isinstance(x, dict) for x in obj):
        return obj
    return []


async def rendered_table(page):
    candidates=[]
    for t in await page.locator("table").all():
        try:
            h=await t.locator("thead tr th").all_inner_texts()
            if not h:
                h=await t.locator("tr").first.locator("th,td").all_inner_texts()
            trs=await t.locator("tbody tr").all()
            if not trs:
                trs=await t.locator("tr").all()
            h=[clean(x) for x in h]
            rows=[]
            for tr in trs:
                c=await tr.locator("td").all_inner_texts()
                if not c:
                    c=await tr.locator("th,td").all_inner_texts()
                c=[clean(x) for x in c]
                if c and any(c): rows.append(c)
            if h and rows:
                w=len(h)
                rows=[r[:w]+[""]*max(0,w-len(r)) for r in rows]
                candidates.append((len(rows),h,rows))
        except Exception:
            pass
    if not candidates:
        return [],[],None
    _,h,rows=max(candidates,key=lambda x:x[0])
    return [dict(zip(h,r)) for r in rows],h,"rendered_html_table"


async def run(page):
    title,url,kind=SCREEN
    captured=[]

    # IMPORTANT: install the response listener BEFORE page.goto().
    async def on_response(resp):
        try:
            u=resp.url.lower()
            ct=(await resp.header_value("content-type") or "").lower()
            if "screener/process" in u or ("json" in ct and ("chartink.com" in u)):
                if "json" in ct:
                    data=await resp.json()
                    rec=records_from_json(data)
                    if rec:
                        captured.append((rec,"chartink_json_response",resp.url))
        except Exception:
            pass

    page.on("response", on_response)
    try:
        await page.goto(url,wait_until="domcontentloaded",timeout=60000)

        # Give Chartink time to execute its client-side scan request.
        for _ in range(30):
            if captured:
                rec,method,source=captured[-1]
                fields=list(dict.fromkeys(k for r in rec for k in r.keys()))
                return {"title":title,"url":url,"kind":kind,"status":"OK","method":method,"source":source,"records":rec,"raw_fields":fields}
            rec,fields,method=await rendered_table(page)
            if rec:
                return {"title":title,"url":url,"kind":kind,"status":"OK","method":method,"records":rec,"raw_fields":fields}
            await page.wait_for_timeout(2000)

        if captured:
            rec,method,source=captured[-1]
            fields=list(dict.fromkeys(k for r in rec for k in r.keys()))
            return {"title":title,"url":url,"kind":kind,"status":"OK","method":method,"source":source,"records":rec,"raw_fields":fields}

        return {"title":title,"url":url,"kind":kind,"status":"NO_RESULT_TABLE_CAPTURED","method":None,"records":[],"raw_fields":[]}
    except PlaywrightTimeoutError as e:
        return {"title":title,"url":url,"kind":kind,"status":"PAGE_TIMEOUT","error":str(e),"records":[],"raw_fields":[]}
    except Exception as e:
        return {"title":title,"url":url,"kind":kind,"status":"ERROR","error":str(e),"records":[],"raw_fields":[]}
    finally:
        page.remove_listener("response",on_response)


async def main():
    ts=datetime.now(ZoneInfo("Asia/Kolkata")).isoformat(timespec="seconds")
    async with async_playwright() as pw:
        browser=await pw.chromium.launch(headless=True)
        ctx=await browser.new_context(
            viewport={"width":1440,"height":1200},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
        )
        page=await ctx.new_page()
        print(f"Running {SCREEN[0]} ...",flush=True)
        result=await run(page)
        result["extracted_at"]=ts
        print(f"{result['status']} | rows={len(result['records'])} | fields={len(result['raw_fields'])} | method={result.get('method')}",flush=True)
        await browser.close()

    with pd.ExcelWriter(OUT_XLSX,engine="openpyxl") as w:
        pd.DataFrame([{
            "Field":"Extracted at","Value":ts
        },{
            "Field":"Timezone","Value":"Asia/Kolkata"
        },{
            "Field":"Screener","Value":SCREEN[0]
        },{
            "Field":"Status","Value":result["status"]
        },{
            "Field":"Capture method","Value":result.get("method","")
        },{
            "Field":"Rows","Value":len(result["records"])
        },{
            "Field":"Fields","Value":len(result["raw_fields"])
        }]).to_excel(w,sheet_name="Dashboard",index=False)
        sheet=re.sub(r"[:\\/?*\[\]]","_",SCREEN[0])[:31]
        pd.DataFrame(result["records"],columns=result["raw_fields"]).to_excel(w,sheet_name=sheet,index=False)

    OUT_JSON.write_text(json.dumps({
        "extracted_at":ts,
        "timezone":"Asia/Kolkata",
        "source":"Chartink saved screener",
        "screeners":[result]
    },ensure_ascii=False,indent=2),encoding="utf-8")

    if result["status"]!="OK":
        raise SystemExit(f"Chartink result capture failed: {result['status']}")


if __name__=="__main__":
    asyncio.run(main())
