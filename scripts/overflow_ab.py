"""A/B horizontal overflow: clean HEAD vs the working tree, all seven languages
at three widths, so the known de/fr section-head fragility can be told apart from
anything this pass introduced."""
import asyncio, json, os, subprocess, sys, time, urllib.request

import websockets

CHROME = r"C:\Users\nykal\AppData\Local\ms-playwright\chromium-1234\chrome-win64\chrome.exe"
HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = r"C:/Users/nykal/AppData/Local/Temp/claude/C--Users-nykal-Documents-Obsidian-Vault/a932dcae-7d2b-4d5f-9125-a25cb4f4f7b5/scratchpad"
PAGES = {
    "baseline": SCRATCH + "/ppt_baseline/index.html",
    "working": os.path.abspath(os.path.join(HERE, "..", "docs", "index.html")).replace("\\", "/"),
}
PORT = 9351
LANGS = ["en", "ko", "ja", "zh", "de", "fr", "es"]
WIDTHS = [390, 768, 1280]

JS = """(() => {
  const de = document.documentElement;
  const over = de.scrollWidth - de.clientWidth;
  const culprits = [];
  if (over > 0) {
    document.querySelectorAll('body *').forEach(el => {
      const r = el.getBoundingClientRect();
      if (r.width === 0) return;
      if (r.right > de.clientWidth + 1) {
        let scrolls = false;
        for (let n = el; n && n !== document.body; n = n.parentElement) {
          const ov = getComputedStyle(n).overflowX;
          if (ov === 'auto' || ov === 'scroll') { scrolls = true; break; }
        }
        if (!scrolls) culprits.push(el.tagName.toLowerCase() + '.' +
          ((el.className||'').toString().split(' ')[0]) + '#' +
          (el.getAttribute('data-i18n')||'-') + '@' + Math.round(r.right));
      }
    });
  }
  return {over: over, culprits: culprits.slice(0, 5)};
})()"""


class CDP:
    def __init__(self, ws):
        self.ws = ws
        self.i = 0

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

    async def ev(self, expr):
        r = await self.send("Runtime.evaluate",
                            {"expression": expr, "returnByValue": True, "awaitPromise": True})
        return r["result"].get("value")


async def main():
    proc = subprocess.Popen([
        CHROME, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
        "--allow-file-access-from-files", "--remote-debugging-port=%d" % PORT,
        "--user-data-dir=" + os.path.join(HERE, "..", ".chrome-ab"), "about:blank",
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

        out = {}
        for label, path in PAGES.items():
            req = urllib.request.Request("http://127.0.0.1:%d/json/new?about:blank" % PORT, method="PUT")
            tid = json.loads(urllib.request.urlopen(req, timeout=10).read().decode())
            async with websockets.connect(tid["webSocketDebuggerUrl"], max_size=None) as ws:
                c = CDP(ws)
                await c.send("Runtime.enable")
                await c.send("Page.enable")
                await c.send("Page.navigate", {"url": "file:///" + path})
                await asyncio.sleep(2.0)
                for lg in LANGS:
                    await c.ev("document.querySelector('[data-lang=\"%s\"]').click()" % lg)
                    await asyncio.sleep(0.2)
                    for w in WIDTHS:
                        await c.send("Emulation.setDeviceMetricsOverride",
                                     {"width": w, "height": 900, "deviceScaleFactor": 1,
                                      "mobile": w <= 768})
                        await asyncio.sleep(0.22)
                        out.setdefault("%s@%d" % (lg, w), {})[label] = await c.ev(JS)
            urllib.request.urlopen("http://127.0.0.1:%d/json/close/%s" % (PORT, tid["id"]), timeout=5).read()

        worse = []
        for k in sorted(out):
            b = out[k]["baseline"]["over"]
            w = out[k]["working"]["over"]
            flag = "SAME" if b == w else ("BETTER" if w < b else "WORSE")
            if flag == "WORSE":
                worse.append(k)
            print("%-10s baseline=%-5d working=%-5d %s %s" % (k, b, w, flag,
                  out[k]["working"]["culprits"] if w else ""))
        print("\nregressions:", worse or "none")
    finally:
        proc.terminate()


asyncio.run(main())
