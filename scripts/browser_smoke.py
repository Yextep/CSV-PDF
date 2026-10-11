"""Prueba la interfaz móvil, los temas y los controles con un puente Android simulado."""
import functools
import http.server
import os
from pathlib import Path
import shutil
import threading
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "android/app/src/main/assets"
MOCK = r"""
const mock={ready:true,busy:false,selection:[],history:[],themeMode:'system',theme:'light',options:{}};
window.LumaNative={theme:()=>mock.theme,request:(action,payload)=>{
  const data=JSON.parse(payload);
  const update=()=>setTimeout(()=>Luma.onEvent('state',structuredClone(mock)),10);
  if(action==='state')update();
  if(action==='theme'){mock.themeMode=data.mode;mock.theme=data.mode==='dark'?'dark':'light';setTimeout(()=>Luma.onEvent('theme',{mode:data.mode,resolved:mock.theme}),10);}
  if(action==='demo'){mock.selection=[{id:'a',name:'Café.xlsx',bytes:1200,relative:'Café.xlsx'},{id:'b',name:'Bienvenida.txt',bytes:300,relative:'Bienvenida.txt'}];update();}
  if(action==='remove'){mock.selection=mock.selection.filter(f=>f.id!==data.id);update();}
  if(action==='clearSelection'){mock.selection=[];update();}
  if(action==='options')mock.options=data;
  if(action==='convert'){mock.selection=[];mock.history=[{id:'s',time:Date.now(),total:2,outputCount:2,outputs:['Café.csv','Bienvenida.pdf'],results:[],errors:0}];update();setTimeout(()=>Luma.onEvent('done',{id:'s'}),40);}
  if(action==='delete'){mock.history=[];update();}
}};
"""


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *_):
        pass


def main():
    handler = functools.partial(QuietHandler, directory=str(ASSETS))
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    (ROOT / "docs").mkdir(exist_ok=True)
    try:
        with sync_playwright() as playwright:
            launch = {"headless": True, "args": ["--no-sandbox", "--disable-dev-shm-usage"]}
            if not os.environ.get("CI") and shutil.which("chromium"):
                launch["executable_path"] = shutil.which("chromium")
                launch["args"] += ["--single-process", "--no-zygote", "--disable-gpu"]
            browser = playwright.chromium.launch(**launch)
            for width, height in ((320,844), (360,844), (390,844), (412,844), (720,844), (320,640), (390,680)):
                page = browser.new_page(viewport={"width": width, "height": height}, device_scale_factor=1, bypass_csp=True)
                errors = []
                page.on("pageerror", lambda error: errors.append(str(error)))
                page.add_init_script(MOCK)
                page.goto(f"http://127.0.0.1:{server.server_port}/index.html")
                page.wait_for_function("state.ready")
                assert page.locator("#convert").is_disabled()
                assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
                assert page.evaluate("Math.abs(innerHeight-document.querySelector('.bottom-nav').getBoundingClientRect().bottom-16)<2")
                assert page.evaluate("document.querySelector('#pick-files').getBoundingClientRect().bottom <= document.querySelector('.bottom-nav').getBoundingClientRect().top - 4")
                for mode in ("light", "dark"):
                    page.evaluate(f"send('theme',{{mode:'{mode}'}})")
                    page.wait_for_function(f"document.documentElement.dataset.theme === '{mode}'")
                    if width == 390 and height == 844:
                        page.wait_for_timeout(300)
                        page.screenshot(path=str(ROOT / f"docs/luma-{mode}.png"), full_page=True)
                page.locator("#demo").click()
                page.wait_for_function("state.selection.length===2")
                assert page.locator("#convert").is_enabled()
                assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
                page.locator('[data-kind="excel"]').click()
                assert "1 archivo" in page.locator("#convert").inner_text()
                page.locator('[data-kind="both"]').click()
                page.locator("#convert").click()
                page.wait_for_function("!document.querySelector('#modal-wrap').hidden")
                assert page.locator(".output-row").count() == 2
                page.locator(".modal-close").click()
                page.locator('[data-page="history"]').click()
                assert page.locator(".history-card").count() == 1
                page.locator('[data-page="settings"]').click()
                page.locator('[data-theme="light"]').click()
                page.wait_for_function("document.documentElement.dataset.theme==='light'")
                page.locator("#clear-history").click()
                page.locator('[data-confirm-delete="all"]').click()
                page.wait_for_function("state.history.length===0")
                page.evaluate("state.selection=[{id:'x',name:'<img src=x onerror=alert(1)>.txt',bytes:1}];render();navigate('convert')")
                assert page.locator("#selection img").count() == 0
                page.evaluate("state.selection=Array.from({length:8},(_,i)=>({id:String(i),name:'Archivo '+i+'.txt',bytes:10}));render()")
                assert page.locator("#selection .file-row").count() == 4
                page.locator('[data-more]').click()
                assert page.locator("#selection .file-row").count() == 8
                assert not errors, errors
                page.close()
            browser.close()
    finally:
        server.shutdown()
    print("Interfaz verificada en 5 anchos y 3 alturas: botones visibles, navegación, temas, selección, historial y diálogos.")


if __name__ == "__main__":
    main()
