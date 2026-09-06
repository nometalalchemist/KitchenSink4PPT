"""Headless check for the ppt-fix2 pass: console cleanliness across all seven
languages, horizontal overflow at three widths for en/ko/ja/zh/es, and the shape
of the three things that changed (trade counter, tip card, family row)."""
import asyncio, json, os, subprocess, sys, time, urllib.request

import websockets

CHROME = r"C:\Users\nykal\AppData\Local\ms-playwright\chromium-1234\chrome-win64\chrome.exe"
PAGE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "docs", "index.html")).replace("\\", "/")
PORT = 9347
WIDTHS = [390, 768, 1280]
OVER_LANGS = ["en", "ko", "ja", "zh", "es"]
ALL_LANGS = ["en", "ko", "ja", "zh", "de", "fr", "es"]


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.i = 0
        self.events = []

    async def send(self, method, params=None):
        self.i += 1
        mid = self.i
        await self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(await self.ws.recv())
            if msg.get("id") == mid:
                if "error" in msg:
                    raise RuntimeError(method + ": " + json.dumps(msg["error"]))
                return msg.get("result", {})
            if "method" in msg:
                self.events.append(msg)

    async def drain(self, seconds=0.6):
        end = time.time() + seconds
        while time.time() < end:
            try:
                msg = json.loads(await asyncio.wait_for(self.ws.recv(), timeout=0.15))
            except asyncio.TimeoutError:
                continue
            if "method" in msg:
                self.events.append(msg)

    async def ev(self, expr):
        r = await self.send("Runtime.evaluate", {
            "expression": expr, "returnByValue": True, "awaitPromise": True})
        if "exceptionDetails" in r:
            return {"__error__": r["exceptionDetails"].get("text", "")}
        return r["result"].get("value")


def console_issues(events):
    bad = []
    for e in events:
        m = e["method"]
        if m == "Runtime.exceptionThrown":
            d = e["params"]["exceptionDetails"]
            bad.append("EXCEPTION: " + (d.get("exception", {}).get("description") or d.get("text", "")))
        elif m == "Runtime.consoleAPICalled":
            txt = " ".join(str(a.get("value", a.get("description", "?"))) for a in e["params"]["args"])
            bad.append("console.%s: %s" % (e["params"]["type"], txt))
        elif m == "Log.entryAdded":
            en = e["params"]["entry"]
            if "fonts.googleapis" in (en.get("url") or "") or "fonts.gstatic" in (en.get("url") or ""):
                continue  # offline font fetch, not a page fault
            bad.append("log[%s/%s]: %s" % (en.get("level"), en.get("source"), en.get("text")))
    return bad


OVERFLOW_JS = """(() => {
  const de = document.documentElement;
  const over = de.scrollWidth - de.clientWidth;
  const culprits = [];
  if (over > 0) {
    document.querySelectorAll('body *').forEach(el => {
      const r = el.getBoundingClientRect();
      if (r.width === 0) return;
      if (r.right > de.clientWidth + 1 || r.left < -1) {
        const p = el.parentElement;
        let scrolls = false;
        for (let n = el; n && n !== document.body; n = n.parentElement) {
          const ov = getComputedStyle(n).overflowX;
          if (ov === 'auto' || ov === 'scroll') { scrolls = true; break; }
        }
        if (!scrolls) culprits.push({
          tag: el.tagName.toLowerCase(),
          cls: (el.className && el.className.toString().slice(0, 46)) || '',
          key: el.getAttribute('data-i18n') || '',
          right: Math.round(r.right), left: Math.round(r.left)
        });
      }
    });
  }
  return {over: over, culprits: culprits.slice(0, 8)};
})()"""

SHAPE_JS = """(() => {
  const q = s => document.querySelector(s);
  const inst = q('#install');
  const eff = q('#efficiency');
  const fold = q('#install details.devnote');
  const row = q('.famrow');
  const cards = row ? Array.from(row.children) : [];
  const rowStyle = row ? getComputedStyle(row) : null;
  return {
    fold_in_install: !!fold,
    fold_plate: fold ? (fold.querySelector('.devnote-label')||{}).textContent : null,
    fold_summary: fold ? fold.querySelector('summary').textContent.trim() : null,
    fold_terms: fold ? fold.querySelectorAll('.term').length : 0,
    devnote_in_efficiency: !!(eff && eff.querySelector('.devnote')),
    fold_before_footnote: fold ? !!(fold.nextElementSibling &&
        fold.nextElementSibling.classList.contains('foot-note')) : false,
    tipcard: !!q('#install .tipcard'),
    tip_title: q('#install .tipcard h3') ? q('#install .tipcard h3').textContent.trim() : null,
    famrow: !!row,
    fam_cards: cards.map(a => ({h3: a.querySelector('h3').textContent.trim(),
                                href: a.getAttribute('href'),
                                w: Math.round(a.getBoundingClientRect().width)})),
    fam_overflow_x: rowStyle ? rowStyle.overflowX : null,
    fam_wraps: rowStyle ? rowStyle.flexWrap : null,
    fam_scrolls_now: row ? row.scrollWidth > row.clientWidth : null,
    glosses: Array.from(document.querySelectorAll('.gloss')).map(g => g.textContent.trim()),
    // the Web card slots in later: clone a card, confirm the row scrolls sideways
    // instead of wrapping and that the page itself still does not move
    web_card_probe: (() => {
      if (!row) return null;
      const de = document.documentElement;
      const clone = cards[0].cloneNode(true);
      clone.id = '';
      row.appendChild(clone);
      const r = {rowScrolls: row.scrollWidth > row.clientWidth,
                 rowTops: new Set(Array.from(row.children)
                          .map(c => Math.round(c.getBoundingClientRect().top))).size,
                 pageOver: de.scrollWidth - de.clientWidth};
      row.removeChild(clone);
      return r;
    })()
  };
})()"""


async def main():
    proc = subprocess.Popen([
        CHROME, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--allow-file-access-from-files", "--remote-debugging-port=%d" % PORT,
        "--user-data-dir=" + os.path.join(os.path.dirname(__file__), "..", ".chrome-verify"),
        "about:blank",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        for _ in range(60):
            try:
                urllib.request.urlopen("http://127.0.0.1:%d/json/version" % PORT, timeout=1).read()
                break
            except Exception:
                time.sleep(0.4)
        else:
            sys.exit("chromium did not come up")

        req = urllib.request.Request("http://127.0.0.1:%d/json/new?about:blank" % PORT, method="PUT")
        tid = json.loads(urllib.request.urlopen(req, timeout=10).read().decode())
        out = {}
        async with websockets.connect(tid["webSocketDebuggerUrl"], max_size=None) as ws:
            c = CDP(ws)
            await c.send("Runtime.enable")
            await c.send("Log.enable")
            await c.send("Page.enable")
            await c.send("Emulation.setDeviceMetricsOverride",
                         {"width": 1280, "height": 900, "deviceScaleFactor": 1, "mobile": False})
            await c.send("Page.navigate", {"url": "file:///" + PAGE})
            await c.drain(2.5)
            out["console_on_load"] = console_issues(c.events)
            c.events = []

            out["shape"] = await c.ev(SHAPE_JS)

            lang_console = {}
            for lg in ALL_LANGS:
                await c.ev("document.querySelector('[data-lang=\"%s\"]').click()" % lg)
                await c.drain(0.35)
                st = await c.ev("""(() => ({
                  htmlLang: document.documentElement.lang,
                  unresolved: Array.from(document.querySelectorAll('[data-i18n]'))
                              .filter(e => !e.textContent.trim()).map(e => e.getAttribute('data-i18n')),
                  literalKeys: Array.from(document.querySelectorAll('[data-i18n]'))
                              .filter(e => e.textContent.trim() === e.getAttribute('data-i18n'))
                              .map(e => e.getAttribute('data-i18n')),
                  emdash: (document.body.innerText.match(/\\u2014/g)||[]).length,
                  glosses: Array.from(document.querySelectorAll('.gloss')).map(g=>g.textContent.trim()),
                  fam: Array.from(document.querySelectorAll('.famrow .counter p'))
                        .map(p=>p.textContent.trim().slice(0,34))
                }))()""")
                st["console"] = console_issues(c.events)
                c.events = []
                lang_console[lg] = st
            out["langs"] = lang_console

            overflow = {}
            for lg in OVER_LANGS:
                await c.ev("document.querySelector('[data-lang=\"%s\"]').click()" % lg)
                await c.drain(0.2)
                for w in WIDTHS:
                    await c.send("Emulation.setDeviceMetricsOverride",
                                 {"width": w, "height": 900, "deviceScaleFactor": 1,
                                  "mobile": w <= 768})
                    await c.drain(0.25)
                    overflow["%s@%d" % (lg, w)] = await c.ev(OVERFLOW_JS)
                await c.send("Emulation.setDeviceMetricsOverride",
                             {"width": 1280, "height": 900, "deviceScaleFactor": 1, "mobile": False})
            out["overflow"] = overflow
            out["console_tail"] = console_issues(c.events)
        print(json.dumps(out, indent=1, ensure_ascii=False))
    finally:
        proc.terminate()


asyncio.run(main())
