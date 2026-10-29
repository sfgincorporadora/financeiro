import asyncio, json
from playwright.async_api import async_playwright
OUT='/tmp/claude-0/-home-claude/e34b4822-ec90-5509-87e2-1ec7436f6334/scratchpad/shots2'
import os; os.makedirs(OUT, exist_ok=True)
URL='http://127.0.0.1:5173/'
J=lambda pg,s: pg.evaluate(s)
async def login(pg, email='luismatagalvao@gmail.com'):
    await pg.goto(URL); await pg.wait_for_timeout(300)
    await pg.fill('[name=email]',email); await pg.fill('[name=senha]','teste123'); await pg.click('button[type=submit]'); await pg.wait_for_timeout(600)
async def main():
  async with async_playwright() as p:
    b = await p.chromium.launch()
    ctx = await b.new_context(viewport={'width':1440,'height':900}, accept_downloads=True)
    pg = await ctx.new_page(); errs=[]
    pg.on('pageerror', lambda e: errs.append('PAGEERR '+str(e)))
    pg.on('console', lambda m: errs.append('CONSOLE '+m.text) if m.type=='error' else None)
    await login(pg)
    await pg.goto(URL+'#/dados'); await pg.wait_for_timeout(300)
    await pg.set_input_files('#iXls','test/SFG.xlsx'); await pg.wait_for_timeout(2500); await pg.click('#cfOk'); await pg.wait_for_timeout(2500)
    for v in ['inicio','vendas','cronograma','simulador','fornecedores','relatorios','usuarios','conciliacao']:
        await pg.goto(URL+'#/'+v); await pg.wait_for_timeout(700); await pg.screenshot(path=f'{OUT}/{v}.png', full_page=True)
    # conciliação
    await pg.set_input_files('#cfile','test/ofx_btg.ofx'); await pg.wait_for_timeout(800)
    await pg.screenshot(path=f'{OUT}/conc2.png', full_page=True)
    print('conc rows', await J(pg,"[...document.querySelectorAll('.ftbl tbody tr')].map(r=>r.innerText.replace(/\\s+/g,' ').slice(0,140))"))
    await pg.click('[data-act=cBx] >> nth=0'); await pg.wait_for_timeout(200); await pg.click('#cfOk'); await pg.wait_for_timeout(500)
    print('MDO', await J(pg,"(()=>{const t=window.__sfg.S.list.find(t=>t.item.startsWith('MDO + Despesas - H3 1ª (3/8)'));return [t.status,t.pag,t.conta,t.conc]})()"))
    btns = await pg.query_selector_all('[data-act=cOk]'); print('auto btns', len(btns))
    if btns: await btns[0].click(); await pg.wait_for_timeout(400)
    await pg.click('[data-act=cNovo]'); await pg.wait_for_timeout(300)
    await pg.screenshot(path=f'{OUT}/conc_novo.png')
    await pg.click('#fNovo button[type=submit]'); await pg.wait_for_timeout(500)
    print('tarifa', await J(pg,"(()=>{const t=window.__sfg.S.list.find(t=>t.item.startsWith('TARIFA'));return t&&[t.status,t.pag,t.valor,t.conc,t.destino]})()"))
    print('conc rows after', await J(pg,"[...document.querySelectorAll('.ftbl tbody tr')].map(r=>r.innerText.replace(/\\s+/g,' ').slice(0,90))"))
    # vendas
    await pg.goto(URL+'#/vendas'); await pg.wait_for_timeout(500)
    await pg.click('.kcard[data-k="K3|1"]'); await pg.wait_for_timeout(300)
    await pg.select_option('#fV [name=etapa]','registro'); await pg.fill('#fV [name=cliente]','Cliente Teste'); await pg.click('#fV button.primary'); await pg.wait_for_timeout(400)
    print('venda', await J(pg,"JSON.stringify(window.__sfg.S.vendas.lotes['K3|1'])"))
    # medição
    await pg.goto(URL+'#/cronograma'); await pg.wait_for_timeout(500)
    await pg.click('[data-act=medicao][data-q=H3]'); await pg.wait_for_timeout(300)
    await pg.select_option('#fM [name=etapa]','Alvenaria'); await pg.fill('#fM [name=pct]','15'); await pg.click('#fM button.primary'); await pg.wait_for_timeout(400)
    await pg.screenshot(path=f'{OUT}/crono2.png', full_page=True)
    # simulador
    await pg.goto(URL+'#/simulador'); await pg.wait_for_timeout(500)
    print('sim', (await pg.inner_text('#simRes'))[:200].replace('\n',' | '))
    await pg.fill('#fSim [name=nome]','Araripina teste'); await pg.fill('#fSim [name=lotes]','60'); await pg.click('#fSim button[type=submit]'); await pg.wait_for_timeout(400)
    print('cenarios', await J(pg,"Object.values(window.__sfg.S.simulacoes.cenarios).map(c=>[c.nome,c.lotes])"))
    # fornecedores: vincular Mendonça e qtd
    await pg.goto(URL+'#/fornecedores'); await pg.wait_for_timeout(300)
    await pg.click('[data-act=fornNovo]'); await pg.fill('#fF [name=nome]','Mendonça'); await pg.click('#fF button.primary'); await pg.wait_for_timeout(500)
    await pg.fill('#vtxt','Mendonça'); await pg.wait_for_timeout(200); print('vprev', await pg.inner_text('#vPrev'))
    await pg.click('#vBtn'); await pg.wait_for_timeout(200); await pg.click('#cfOk'); await pg.wait_for_timeout(500)
    # compra com quantidade
    for d,(qt,val) in [('2026-11-10',(200,'7.000,00')),('2026-12-10',(200,'7.400,00'))]:
        await pg.click('.top [data-act=novo]'); await pg.wait_for_timeout(200)
        await pg.fill('#fNovo [name=item]','Cimento CP-II'); await pg.fill('#fNovo [name=forn]','Mendonça'); await pg.fill('#fNovo [name=valor]',val); await pg.fill('#fNovo [name=venc]',d)
        await pg.select_option('#fNovo [name=quadra]','H3'); await pg.wait_for_timeout(100)
        for l in ['1','2']: await pg.click(f'#fNovo .pick:has(input[value="{l}"])')
        await pg.fill('#fNovo [name=qtd]',str(qt)); await pg.fill('#fNovo [name=unid]','sc'); await pg.click('#fNovo button[type=submit]'); await pg.wait_for_timeout(400)
    await pg.goto(URL+'#/fornecedores'); await pg.wait_for_timeout(400)
    await pg.click('[data-act=forn][data-n="Mendonça"]'); await pg.wait_for_timeout(500)
    await pg.screenshot(path=f'{OUT}/forn.png', full_page=False)
    print('preco', await J(pg,"[...document.querySelectorAll('.sheet .tbl tbody tr')].slice(0,3).map(r=>r.innerText.replace(/\\s+/g,' '))"))
    await pg.keyboard.press('Escape'); await pg.wait_for_timeout(300)
    # relatório PDF
    await pg.goto(URL+'#/relatorios'); await pg.wait_for_timeout(400)
    async with pg.expect_download() as dl: await pg.click('#rPdf')
    d=await dl.value; await d.save_as(OUT+'/rel.pdf'); print('pdf', d.suggested_filename)
    # ics
    await pg.goto(URL+'#/inicio'); await pg.wait_for_timeout(400)
    async with pg.expect_download() as dl: await pg.click('[data-act=ics]')
    d=await dl.value; await d.save_as(OUT+'/v.ics'); print('ics', d.suggested_filename, open(OUT+'/v.ics').read().count('BEGIN:VEVENT'))
    # usuários: leitor
    await pg.goto(URL+'#/usuarios'); await pg.wait_for_timeout(300)
    await pg.fill('#fU [name=email]','socio@teste.com'); await pg.click('#fU button.primary'); await pg.wait_for_timeout(400)
    print('acesso', await J(pg,"JSON.stringify(window.__sfg.S.acesso)"))
    st = await J(pg,"localStorage.getItem('fakefs')")
    c2 = await b.new_context(viewport={'width':390,'height':844}, is_mobile=True, has_touch=True); p2=await c2.new_page()
    await p2.goto(URL); await p2.evaluate("s=>localStorage.setItem('fakefs',s)", st); await login(p2,'socio@teste.com'); await p2.wait_for_timeout(800)
    print('leitor ro', await p2.evaluate("document.body.classList.contains('ro')"), 'fab visivel', await p2.is_visible('.fab'), 'n', await p2.evaluate('window.__sfg.S.list.length'))
    await p2.screenshot(path=f'{OUT}/m_leitor_inicio.png')
    for v in ['vendas','mais','conciliacao']:
        await p2.goto(URL+'#/'+v); await p2.wait_for_timeout(500); await p2.screenshot(path=f'{OUT}/m_{v}.png'); print('overflow',v, await p2.evaluate('document.documentElement.scrollWidth'))
    print('ERRS', errs[:8])
    await b.close()
asyncio.run(main())
