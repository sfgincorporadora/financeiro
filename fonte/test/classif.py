import asyncio, json, os
from playwright.async_api import async_playwright
OUT='/tmp/claude-0/-home-claude/e34b4822-ec90-5509-87e2-1ec7436f6334/scratchpad/shots_cl'; os.makedirs(OUT, exist_ok=True)
URL='http://127.0.0.1:5173/'
J=lambda pg,s: pg.evaluate(s)
async def main():
  async with async_playwright() as p:
    b = await p.chromium.launch()
    for vp,nm in [({'width':1440,'height':900},'d'),({'width':390,'height':844},'m')]:
      ctx = await b.new_context(viewport=vp); pg = await ctx.new_page(); errs=[]
      pg.on('pageerror', lambda e: errs.append('PAGEERR '+str(e)))
      pg.on('console', lambda m: errs.append('CONSOLE '+m.text) if m.type=='error' else None)
      await pg.goto(URL); await pg.wait_for_timeout(300)
      await pg.evaluate("localStorage.clear()"); await pg.goto(URL); await pg.wait_for_timeout(300)
      await pg.fill('[name=email]','luismatagalvao@gmail.com'); await pg.fill('[name=senha]','teste123'); await pg.click('button[type=submit]'); await pg.wait_for_timeout(600)
      await pg.goto(URL+'#/dados'); await pg.wait_for_timeout(300)
      await pg.set_input_files('#iXls','test/SFG.xlsx'); await pg.wait_for_timeout(2500); await pg.click('#cfOk'); await pg.wait_for_timeout(2500)
      n0 = await J(pg,"window.__sfg.S.list.length")
      await pg.goto(URL+'#/inicio'); await pg.wait_for_timeout(600)
      print(nm,'alertas', await J(pg,"[...document.querySelectorAll('.al span')].map(x=>x.textContent)"))
      await pg.goto(URL+'#/arrumar'); await pg.wait_for_timeout(800)
      await pg.screenshot(path=f'{OUT}/{nm}_empresa.png', full_page=True)
      tabs = await J(pg,"[...document.querySelectorAll('.rev-tabs button')].map(b=>b.textContent)"); print(nm,'tabs',tabs)
      grp = await J(pg,"[...document.querySelectorAll('.arr-g .arr-txt')].slice(0,12).map(x=>x.innerText.replace(/\\n/g,' | '))"); print(nm,'grupos'); [print('  ',g) for g in grp]
      # filtro strada e juntar
      await pg.fill('#rvQ','strada'); await pg.wait_for_timeout(600)
      before = await J(pg,"window.__sfg.S.list.filter(t=>/strada/i.test(t.item)).map(t=>t.valor).reduce((a,b)=>a+b,0)")
      await pg.click('#rvTodos'); await pg.wait_for_timeout(200); await pg.click('#cfOk'); await pg.wait_for_timeout(1200)
      after = await J(pg,"(()=>{const xs=window.__sfg.S.list.filter(t=>/strada/i.test(t.item));return [xs.length, xs.reduce((a,t)=>a+t.valor,0), xs.slice(0,4).map(t=>[t.item,t.valor,t.ambito,t.quadra,t.emp])]})()")
      print(nm,'strada antes',round(before,2),'depois',after)
      # regras
      await pg.click('[data-v=regras]'); await pg.wait_for_timeout(300); await pg.click('#rgSug'); await pg.wait_for_timeout(600)
      await pg.screenshot(path=f'{OUT}/{nm}_regras.png', full_page=True)
      bts = await pg.query_selector_all('[data-apl]'); print(nm,'regras aplicaveis', len(bts))
      while True:
        bts = await pg.query_selector_all('[data-apl]')
        if not bts: break
        await bts[0].click(); await pg.wait_for_timeout(200); await pg.click('#cfOk'); await pg.wait_for_timeout(900)
      cim = await J(pg,"(()=>{const xs=window.__sfg.S.list.filter(t=>/^cim /i.test(t.item));return [xs.length, xs.slice(0,3).map(t=>[t.item,t.valor,t.ambito,t.emp,t.destino])]})()")
      print(nm,'cim',cim)
      cimento = await J(pg,"window.__sfg.S.list.filter(t=>/\bcimento/i.test(t.item)&&t.ambito==='empresa').length"); print(nm,'cimento virou empresa (deve ser 0):',cimento)
      # rateio
      await pg.click('[data-v=rateio]'); await pg.wait_for_timeout(500)
      await pg.screenshot(path=f'{OUT}/{nm}_rateio.png', full_page=True)
      print(nm,'rateio', (await pg.inner_text('#revBox'))[:600].replace('\n',' | '))
      await pg.select_option('#rtCrit','custo'); await pg.wait_for_timeout(400)
      # aportes, nf, contas
      for v in ['aporte','nf','contas','forn','anexo']:
        await pg.click(f'[data-v={v}]'); await pg.wait_for_timeout(400); await pg.screenshot(path=f'{OUT}/{nm}_{v}.png', full_page=True)
      await pg.click('[data-v=nf]'); await pg.wait_for_timeout(300)
      await pg.fill('[data-nf="0"]','1234'); await pg.click('[data-snf="0"]'); await pg.wait_for_timeout(400)
      await pg.click('[data-sem="0"]'); await pg.wait_for_timeout(400)
      print(nm,'nf ok', await J(pg,"window.__sfg.S.list.filter(t=>t.nfNum==='1234').length"), await J(pg,"window.__sfg.S.list.filter(t=>t.semNota).length"))
      # conta: cadastrar numero
      await pg.click('[data-v=contas]'); await pg.wait_for_timeout(300)
      await pg.click('#revBox [data-act=cadConta][data-nome]'); await pg.wait_for_timeout(300)
      await pg.fill('#fx [name=banco]','BTG Pactual (208)'); await pg.fill('#fx [name=ag]','0050'); await pg.fill('#fx [name=num]','1234567-8'); await pg.click('#fx button.primary'); await pg.wait_for_timeout(500)
      print(nm,'conta', await J(pg,"JSON.stringify(window.__sfg.E().contas[0])"))
      # transferir
      ncont = await J(pg,"window.__sfg.E().contas.length"); print(nm,'contas',ncont)
      if ncont>=2:
        await pg.click('#revBox [data-act=transferir]'); await pg.wait_for_timeout(300)
        await pg.fill('#fx [name=valor]','1.000,00'); print(nm,'hint', await pg.inner_text('#trHint')); await pg.click('#fx button.primary'); await pg.wait_for_timeout(500)
        print(nm,'transf', await J(pg,"window.__sfg.S.list.filter(t=>t.destino==='Transferência'||t.aporteTipo==='intragrupo').map(t=>[t.item,t.valor,t.conta,t.emp])"))
      # novo lançamento empresa + regra automática
      await pg.click('.top [data-act=novo]' if nm=='d' else '.fab'); await pg.wait_for_timeout(300)
      await pg.fill('#fNovo [name=item]','Claude - Novembro/26'); await pg.dispatch_event('#fNovo [name=item]','change'); await pg.wait_for_timeout(300)
      print(nm,'form ambito', await J(pg,"[document.querySelector('#fNovo [name=ambito]').value, document.querySelector('#fNovo').classList.contains('modo-empresa'), getComputedStyle(document.querySelector('#fNovo [name=quadra]').closest('label')).display]"))
      await pg.fill('#fNovo [name=valor]','120,00'); await pg.wait_for_timeout(200)
      await pg.screenshot(path=f'{OUT}/{nm}_novo_empresa.png', full_page=True)
      if await pg.query_selector('#fNovo [name=emp]'):
        await pg.select_option('#fNovo [name=emp]', index=1)
      await pg.click('#fNovo button[type=submit]'); await pg.wait_for_timeout(500)
      print(nm,'novo', await J(pg,"JSON.stringify(window.__sfg.S.list.filter(t=>t.item.startsWith('Claude - Novembro')).map(t=>[t.valor,t.ambito,t.quadra,t.lote,t.emp,t.destino]))"))
      # aporte no form
      await pg.click('.top [data-act=novo]' if nm=='d' else '.fab'); await pg.wait_for_timeout(300)
      await pg.click('#segTipo [data-v=entrada]'); await pg.select_option('#fNovo [name=destino]','Aporte'); await pg.wait_for_timeout(100)
      print(nm,'rowAporte visivel', await J(pg,"!document.querySelector('#rowAporte').classList.contains('hidden')"))
      await pg.click('#fNovo [data-close]')
      # telas obras
      await pg.goto(URL+'#/obras/K3/1'); await pg.wait_for_timeout(600); await pg.screenshot(path=f'{OUT}/{nm}_lote.png', full_page=True)
      print(nm,'lote kpi', (await pg.inner_text('.kpis'))[:200].replace('\n',' | '))
      await pg.goto(URL+'#/obras/K3'); await pg.wait_for_timeout(600)
      await pg.goto(URL+'#/pagar'); await pg.wait_for_timeout(500)
      if await pg.query_selector('#fobra'):
        await pg.select_option('#fobra','__emp'); await pg.wait_for_timeout(400); print(nm,'filtro empresa', await pg.inner_text('#ltot'))
      await pg.goto(URL+'#/demonstrativos'); await pg.wait_for_timeout(600)
      await pg.goto(URL+'#/empresas'); await pg.wait_for_timeout(600)
      print(nm,'n', n0, await J(pg,"window.__sfg.S.list.length"))
      print(nm,'ERRS', errs[:8])
      await ctx.close()
    await b.close()
asyncio.run(main())
