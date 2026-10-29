import asyncio, sys
from playwright.async_api import async_playwright
OUT='/tmp/claude-0/-home-claude/e34b4822-ec90-5509-87e2-1ec7436f6334/scratchpad/shots'
import os; os.makedirs(OUT, exist_ok=True)
URL='http://127.0.0.1:5173/'
async def main():
  async with async_playwright() as p:
    b = await p.chromium.launch()
    ctx = await b.new_context(viewport={'width':1360,'height':900})
    pg = await ctx.new_page(); errs=[]
    pg.on('pageerror', lambda e: errs.append('PAGEERR '+str(e)))
    pg.on('console', lambda m: errs.append('CONSOLE '+m.text) if m.type=='error' else None)
    await pg.goto(URL); await pg.wait_for_timeout(500)
    # login errado e certo
    await pg.fill('[name=email]','luismatagalvao@gmail.com'); await pg.fill('[name=senha]','x'); await pg.click('button[type=submit]'); await pg.wait_for_timeout(300)
    print('login err:', await pg.inner_text('#loginErr'))
    await pg.fill('[name=senha]','teste123'); await pg.click('button[type=submit]'); await pg.wait_for_timeout(600)
    await pg.screenshot(path=OUT+'/01_vazio.png')
    # importar
    await pg.goto(URL+'#/dados'); await pg.wait_for_timeout(300)
    await pg.set_input_files('#iXls','test/SFG.xlsx'); await pg.wait_for_timeout(3000)
    await pg.screenshot(path=OUT+'/02_confirm.png')
    await pg.click('#cfOk'); await pg.wait_for_timeout(3000)
    print('import status:', await pg.inner_text('#impStatus'))
    print('n titulos', await pg.evaluate('window.__sfg.S.list.length'))
    await pg.goto(URL+'#/inicio'); await pg.wait_for_timeout(800)
    await pg.screenshot(path=OUT+'/03_inicio.png', full_page=True)
    for v in ['pagar','receber','extrato','obras','obras/K1','obras/K1/38','fluxo','checklist','cadastros']:
      await pg.goto(URL+'#/'+v); await pg.wait_for_timeout(700)
      await pg.screenshot(path=OUT+'/04_'+v.replace('/','_')+'.png', full_page=True)
    # saldo
    print('saldo:', await pg.evaluate("(()=>{let s=0;for(const t of window.__sfg.S.list) if(t.status==='pago') s+=t.valor+(t.juros||0); return s})()"))
    print('ERRS', errs[:10])
    await b.close()
asyncio.run(main())
