import asyncio
from playwright.async_api import async_playwright
URL='http://127.0.0.1:5173/'
async def main():
  async with async_playwright() as p:
    b=await p.chromium.launch(); pg=await (await b.new_context()).new_page(); errs=[]
    pg.on('pageerror', lambda e: errs.append(str(e)))
    await pg.goto(URL); await pg.fill('[name=email]','luismatagalvao@gmail.com'); await pg.fill('[name=senha]','teste123'); await pg.click('button[type=submit]'); await pg.wait_for_timeout(500)
    await pg.goto(URL+'#/dados'); await pg.wait_for_timeout(300)
    await pg.set_input_files('#iXls','test/export.xlsx'); await pg.wait_for_timeout(2500); print(await pg.inner_text('.sheet-body p')); await pg.click('#cfOk'); await pg.wait_for_timeout(2500)
    r = await pg.evaluate("(()=>{const L=window.__sfg.S.list;let s=0;for(const t of L) if(t.status==='pago') s+=t.valor+(t.juros||0);return [L.length, Math.round(s*100)/100, L.filter(t=>t.status!=='pago').length, window.__sfg.S.orcamento.quadras.K1.Material, window.__sfg.S.checklist.lotes['K1|38'], window.__sfg.S.estrutura.quadras.length]})()")
    print(r, errs); await b.close()
asyncio.run(main())
