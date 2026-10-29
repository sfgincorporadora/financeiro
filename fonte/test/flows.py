import asyncio
from playwright.async_api import async_playwright
OUT='/tmp/claude-0/-home-claude/e34b4822-ec90-5509-87e2-1ec7436f6334/scratchpad/shots'
URL='http://127.0.0.1:5173/'
J = lambda pg, s: pg.evaluate(s)
async def main():
  async with async_playwright() as p:
    b = await p.chromium.launch()
    ctx = await b.new_context(viewport={'width':1360,'height':900}, accept_downloads=True)
    pg = await ctx.new_page(); errs=[]
    pg.on('pageerror', lambda e: errs.append('PAGEERR '+str(e)))
    pg.on('console', lambda m: errs.append('CONSOLE '+m.text) if m.type=='error' else None)
    await pg.goto(URL); await pg.wait_for_timeout(400)
    await pg.fill('[name=email]','luismatagalvao@gmail.com'); await pg.fill('[name=senha]','teste123'); await pg.click('button[type=submit]'); await pg.wait_for_timeout(500)
    await pg.goto(URL+'#/dados'); await pg.wait_for_timeout(300)
    await pg.set_input_files('#iXls','test/SFG.xlsx'); await pg.wait_for_timeout(2500); await pg.click('#cfOk'); await pg.wait_for_timeout(2500)
    n0 = await J(pg,'window.__sfg.S.list.length'); print('após import', n0)
    # Início agrupado
    await pg.goto(URL+'#/inicio'); await pg.wait_for_timeout(700); await pg.screenshot(path=OUT+'/10_inicio.png')
    # Novo: 10 parcelas x 3 lotes K3
    await pg.click('.top [data-act=novo]'); await pg.wait_for_timeout(300)
    await pg.fill('#fNovo [name=item]','Cimento Votoran'); await pg.fill('#fNovo [name=forn]','Mendonça'); await pg.select_option('#fNovo [name=destino]','Material')
    await pg.fill('#fNovo [name=valor]','3.000,00'); await pg.select_option('#fNovo [name=quadra]','K3'); await pg.wait_for_timeout(100)
    for l in ['19','20','21']: await pg.click(f'#fNovo .pick:has(input[value="{l}"])')
    await pg.fill('#fNovo [name=venc]','2026-11-10'); await pg.fill('#fNovo [name=parcelas]','10'); await pg.wait_for_timeout(200)
    print('preview:', (await pg.inner_text('#prevNovo')).replace('\n',' | '))
    await pg.screenshot(path=OUT+'/11_novo.png')
    await pg.click('#fNovo button[type=submit]'); await pg.wait_for_timeout(400)
    if await pg.query_selector('#cfOk'): await pg.click('#cfOk')
    await pg.wait_for_timeout(600)
    novos = await J(pg,"window.__sfg.S.list.filter(t=>t.item.startsWith('Cimento Votoran'))")
    print('criados', len(novos), 'soma', round(sum(t['valor'] for t in novos),2), novos[0]['item'], novos[-1]['item'], novos[-1]['venc'])
    # Lista a pagar com busca e grupo
    await pg.goto(URL+'#/pagar'); await pg.wait_for_timeout(500)
    await pg.fill('#fq','Cimento'); await pg.wait_for_timeout(400)
    await pg.screenshot(path=OUT+'/12_pagar_grp.png')
    await pg.click('.tr.grp >> nth=0'); await pg.wait_for_timeout(300); await pg.screenshot(path=OUT+'/13_grupo.png')
    await pg.click('#gBx'); await pg.wait_for_timeout(300)
    await pg.fill('#fBx [name=juros]','30,00'); await pg.click('#fBx button[type=submit]'); await pg.wait_for_timeout(500)
    pagos = await J(pg,"window.__sfg.S.list.filter(t=>t.item.startsWith('Cimento Votoran (1/10)')).map(t=>[t.status,t.pag,t.conta,t.juros])")
    print('baixa grupo', pagos)
    # detalhe -> estornar 1
    tid = await J(pg,"window.__sfg.S.list.find(t=>t.item==='Cimento Votoran (1/10) - K3 19').id")
    await pg.evaluate(f"document.querySelector('#view').insertAdjacentHTML('beforeend','<button id=tmpb data-act=abrir data-id={tid}>x</button>')"); await pg.click('#tmpb'); await pg.wait_for_timeout(300)
    await pg.screenshot(path=OUT+'/14_detalhe.png')
    await pg.click('#dEst'); await pg.wait_for_timeout(200); await pg.click('#cfOk'); await pg.wait_for_timeout(400)
    print('estorno', await J(pg,f"window.__sfg.S.byId.get('{tid}').status"), await J(pg,f"window.__sfg.S.byId.get('{tid}').juros"))
    # editar: muda vencimento para outro mês (move bucket)
    await pg.evaluate(f"document.querySelector('#tmpb')?.remove(); document.querySelector('#view').insertAdjacentHTML('beforeend','<button id=tmpb data-act=abrir data-id={tid}>x</button>')"); await pg.click('#tmpb'); await pg.wait_for_timeout(300)
    await pg.click('#dEd'); await pg.wait_for_timeout(300); await pg.fill('#fEd [name=venc]','2027-03-15'); await pg.fill('#fEd [name=valor]','999,90'); await pg.click('#fEd button[type=submit]'); await pg.wait_for_timeout(500)
    t = await J(pg,f"window.__sfg.S.byId.get('{tid}')"); print('editado', t['venc'], t['valor'], t['mes'])
    dup = await J(pg,f"window.__sfg.S.list.filter(x=>x.id==='{tid}').length"); print('sem duplicata', dup)
    # excluir
    await pg.evaluate(f"document.querySelector('#tmpb')?.remove(); document.querySelector('#view').insertAdjacentHTML('beforeend','<button id=tmpb data-act=abrir data-id={tid}>x</button>')"); await pg.click('#tmpb'); await pg.wait_for_timeout(300)
    await pg.click('#dDel'); await pg.wait_for_timeout(200); await pg.click('#cfOk'); await pg.wait_for_timeout(400)
    print('excluído', await J(pg,f"window.__sfg.S.byId.has('{tid}')"))
    # Lançamento realizado (aporte)
    await pg.click('.top [data-act=novo]'); await pg.wait_for_timeout(300)
    await pg.click('#segTipo [data-v=entrada]'); await pg.check('#fNovo [name=realizado]'); await pg.fill('#fNovo [name=item]','Aporte Eduardo'); await pg.select_option('#fNovo [name=destino]','Aporte'); await pg.fill('#fNovo [name=valor]','50000'); await pg.click('#fNovo button[type=submit]'); await pg.wait_for_timeout(500)
    sd = await J(pg,"(()=>{let s=0;for(const t of window.__sfg.S.list) if(t.status==='pago') s+=t.valor+(t.juros||0); return Math.round(s*100)/100})()"); print('saldo após aporte', sd)
    # checklist toggle
    await pg.goto(URL+'#/checklist/K1'); await pg.wait_for_timeout(500)
    await pg.click('[data-act=ck][data-k="K1|38"][data-i="0"]'); await pg.wait_for_timeout(300)
    print('checklist K1|38', await J(pg,"window.__sfg.S.checklist.lotes['K1|38']"))
    await pg.screenshot(path=OUT+'/15_check.png')
    # orçamento
    await pg.goto(URL+'#/obras/K1'); await pg.wait_for_timeout(500); await pg.click('#bOrc'); await pg.wait_for_timeout(300)
    await pg.fill('#fOrc [name="Material"]','12.000,00'); await pg.click('#fOrc button.primary'); await pg.wait_for_timeout(400)
    print('orc K1 material', await J(pg,"window.__sfg.S.orcamento.quadras.K1.Material"), 'MO', await J(pg,"window.__sfg.S.orcamento.quadras.K1['Mão de Obra']"))
    await pg.screenshot(path=OUT+'/16_quadra.png')
    # cadastro quadra
    await pg.goto(URL+'#/cadastros'); await pg.wait_for_timeout(400); await pg.click('[data-act=cadQuadra]:not([data-q])'); await pg.wait_for_timeout(200)
    await pg.fill('#fx [name=q]','K4'); await pg.fill('#fx [name=l]','1-10, 12, 14.1'); await pg.click('#fx button.primary'); await pg.wait_for_timeout(300)
    print('K4', await J(pg,"window.__sfg.S.estrutura.lotes.K4"))
    # export
    await pg.goto(URL+'#/dados'); await pg.wait_for_timeout(300)
    async with pg.expect_download() as dl: await pg.click('#bXls')
    d = await dl.value; await d.save_as('test/export.xlsx'); print('export', d.suggested_filename)
    # reload persistência
    await pg.reload(); await pg.wait_for_timeout(1200); print('após reload', await J(pg,'window.__sfg.S.list.length'))
    print('ERRS', errs[:10])
    # mobile
    m = await b.new_context(viewport={'width':390,'height':844}, is_mobile=True, has_touch=True)
    await m.add_init_script("localStorage.setItem('x','1')")
    mp = await m.new_page()
    # copy storage
    st = await pg.evaluate("JSON.stringify({a:localStorage.getItem('fakefs'),u:localStorage.getItem('fakeuser')})")
    await mp.goto(URL); await mp.evaluate(f"(()=>{{const o={st};localStorage.setItem('fakefs',o.a);localStorage.setItem('fakeuser',o.u)}})()"); await mp.reload(); await mp.wait_for_timeout(1200)
    for v in ['inicio','pagar','obras/K3','mais']:
      await mp.goto(URL+'#/'+v); await mp.wait_for_timeout(600); await mp.screenshot(path=OUT+'/20_m_'+v.replace('/','_')+'.png')
      print('m overflow', v, await mp.evaluate('document.documentElement.scrollWidth'))
    await mp.click('.fab'); await mp.wait_for_timeout(400); await mp.screenshot(path=OUT+'/21_m_novo.png')
    # usuário não autorizado
    o = await b.new_context(); op = await o.new_page(); oe=[]
    await op.goto(URL); await op.evaluate(f"localStorage.setItem('fakefs',{st}.a ? JSON.stringify(JSON.parse({st}.a)) : '{{}}')") if False else None
    await op.fill('[name=email]','outro@gmail.com'); await op.fill('[name=senha]','teste123'); await op.click('button[type=submit]'); await op.wait_for_timeout(800)
    print('outro usuário toast:', await op.inner_text('#toast'), '| titulos', await op.evaluate('window.__sfg.S.list.length'))
    await b.close()
asyncio.run(main())
