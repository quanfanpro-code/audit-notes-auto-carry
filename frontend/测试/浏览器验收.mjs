import { chromium, selectors } from 'playwright-core'
import { spawn } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'

const 根 = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..')
selectors.setTestIdAttribute('data-test')
const 参数 = process.argv.slice(2)
const 证据 = path.resolve(参数[参数.indexOf('--evidence') + 1] || '')
if (!参数.includes('--evidence')) throw new Error('必须提供新的证据目录')
if (fs.existsSync(path.join(证据, '服务就绪.json'))) throw new Error('证据目录已经使用，请另选新目录')
fs.mkdirSync(证据, { recursive: true })
const 日志 = fs.openSync(path.join(证据, '服务日志.txt'), 'a')
const 服务 = spawn('C:\\Users\\27651\\AppData\\Local\\Programs\\Python\\Python314\\python.exe', [path.join(根,'测试','GUI验收服务.py'), '--目录', 证据], { cwd: 根, windowsHide: true, env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONDONTWRITEBYTECODE: '1' }, stdio: ['ignore', 日志, 日志] })
const 等 = ms => new Promise(r => setTimeout(r,ms))
let 清单, 浏览器
const 结果 = { 性能: {}, 任务: [], 页面错误: [], 操作: [], 显示检查: [] }
async function 接口(路径, 内容) {
  const r = await fetch(清单.地址+路径, { method: 内容 === undefined ? 'GET' : 'POST', headers: { 'Content-Type':'application/json', 'X-Audit-Token':清单.令牌 }, ...(内容 === undefined ? {} : { body:JSON.stringify(内容) }) })
  const j = await r.json()
  if (!r.ok || j.ok === false) throw new Error(JSON.stringify(j))
  return j
}
try {
  for(let n=0;n<300;n++) { if(fs.existsSync(path.join(证据,'服务就绪.json')))break; if(服务.exitCode!==null)throw new Error('测试服务提前退出'); await 等(100) }
  清单=JSON.parse(fs.readFileSync(path.join(证据,'服务就绪.json'),'utf-8'))
  浏览器=await chromium.launch({executablePath:'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',headless:true})
  const 环境=await 浏览器.newContext({viewport:{width:1440,height:900},deviceScaleFactor:1})
  const 页=await 环境.newPage()
  页.on('pageerror',e=>结果.页面错误.push(e.message))
  await 页.goto(清单.地址+'/桌面'); await 页.getByRole('heading',{name:'开始一份搬运任务'}).waitFor()
  await 页.screenshot({path:path.join(证据,'01_新建.png'),fullPage:true})
  await 页.getByRole('button',{name:'历史任务',exact:true}).click()
  await 页.getByRole('textbox',{name:'搜索历史任务'}).fill('性能样本_100张')
  await 页.getByRole('button',{name:'查询',exact:true}).click()
  await 页.getByRole('button',{name:'查看任务',exact:true}).waitFor()
  // 历史列表也读取表数；计时前清展示缓存，仍测真实的冷结果首屏。
  await 接口('/验收/清展示缓存',{})
  let 起=performance.now()
  await 页.getByRole('button',{name:'查看任务',exact:true}).click()
  await 页.getByRole('heading',{name:'核查本次结果'}).waitFor()
  await 页.locator('.结果明细 .el-table__body tbody tr').nth(49).waitFor()
  结果.性能.冷首屏毫秒=performance.now()-起
  结果.性能.样本={业务表:100,明细:10000,历史任务:100,每页:50}
  结果.性能.筛选毫秒=[]
  for(let n=0;n<20;n++) {
    const 词='名单'+String(n).padStart(3,'0')
    await 页.getByRole('textbox',{name:'搜索结果'}).fill(词)
    起=performance.now()
    const 响应=页.waitForResponse(r=>decodeURI(r.url()).includes('/结果?')&&r.url().includes('view='))
    await 页.locator('.结果筛选').getByRole('button',{name:'查询',exact:true}).click(); await 响应
    await 页.locator('.结果明细 .el-table__body tbody tr').first().getByText(词,{exact:false}).waitFor()
    结果.性能.筛选毫秒.push(performance.now()-起)
  }
  assert(结果.性能.冷首屏毫秒<=3000, '冷首屏超过3秒')
  assert(结果.性能.筛选毫秒.filter(x=>x<=1000).length>=19,'筛选性能未达到19/20')
  await 页.locator('.结果明细').getByRole('button',{name:'查看',exact:true}).first().click()
  assert(await 页.getByRole('button',{name:'下一项',exact:true}).isEnabled())
  await 页.getByRole('button',{name:'下一项',exact:true}).click()
  await 页.screenshot({path:path.join(证据,'02_结果详情.png'),fullPage:true})
  await 页.keyboard.press('Escape')
  assert.equal(await 页.locator('.详情容器').count(),0)
  结果.操作.push('真实结果筛选20次、上一项下一项及Esc返回')
  await 页.getByRole('button',{name:'清除筛选',exact:true}).click()
  for(const 数 of [20,100]) {
    await 页.locator('.结果明细 .el-pagination .el-select').click()
    await 页.getByRole('option',{name:数+'条/页',exact:true}).click()
    await 页.waitForFunction(n=>document.querySelectorAll('.结果明细 .el-table__body tbody tr').length===n,数)
    结果.操作.push('实际结果分页'+数+'条')
  }

  for(const 场景 of 清单.场景) {
    const 哈希 = p => createHash('sha256').update(fs.readFileSync(p)).digest('hex')
    const 原件 = ['A','B'].map(s => 哈希(场景.输入[s].路径))
    await 页.getByRole('button',{name:'新建任务',exact:true}).click()
    await 接口('/验收/选择队列',{路径:[场景.输入.B.路径,场景.输入.A.路径]})
    await 页.getByTestId('选择B').click()
    if(场景.输入.B.工作表.length)await 页.getByRole('checkbox',{name:场景.输入.B.工作表[0],exact:true}).waitFor({state:'attached'})
    await 页.getByTestId('选择A').click()
    if(场景.输入.A.工作表.length) {
      const 框=页.getByRole('checkbox',{name:场景.输入.A.工作表[0],exact:true})
      await 框.waitFor({state:'attached'})
      if(!await 框.isChecked())await 框.locator('..').click()
      assert(await 框.isChecked())
      assert(await 页.getByRole('checkbox',{name:'隐藏工作表（隐藏，不可选）',exact:true}).isDisabled())
    }
    await 页.getByRole('button',{name:'设置',exact:true}).click()
    await 页.getByRole('heading',{name:'设置',exact:true}).waitFor()
    await 页.getByRole('button',{name:'新建任务',exact:true}).click()
    for(const 侧 of ['A','B']) {
      assert.equal(await 页.getByText(path.basename(场景.输入[侧].路径),{exact:true}).count(),1)
      for(const 表 of 场景.输入[侧].工作表) assert(await 页.getByRole('checkbox',{name:表,exact:true}).isChecked())
    }
    assert(await 页.getByTestId('开始任务').isEnabled())
    const 接受=页.waitForResponse(r=>decodeURI(r.url()).endsWith('/api/桌面/任务')&&r.request().method()==='POST')
    await 页.getByTestId('开始任务').click()
    const 任务=await (await 接受).json()
    await 页.getByRole('heading',{name:'核查本次结果'}).waitFor({timeout:120000})
    const 明细=await 接口('/api/桌面/任务/'+任务.task_id+'/结果?view='+encodeURIComponent('明细'))
    const 资金=明细.条目.find(x=>x.期间==='期初余额'&&x.B本次值.原样文本==='901')
    assert(资金?.A实际值.原样文本==='901',JSON.stringify(明细))
    assert(明细.条目.some(x=>x.A实际值.原样文本==='本次甲'))
    const 文件=(await 接口('/api/桌面/任务/'+任务.task_id+'/文件')).条目
    const 输出=文件.find(x=>x.名称==='Excel新副本'||x.名称==='Word新副本')?.路径
    const 成品哈希=哈希(输出)
    for(let n=0;n<100 && (await 接口('/api/桌面/任务/'+任务.task_id+'/进度')).活动;n++)await 等(50)
    assert.equal(await 页.getByRole('button',{name:'继续此任务',exact:true}).count(),0)
    const 拒绝=await fetch(清单.地址+'/api/桌面/任务/'+任务.task_id+'/继续',{method:'POST',headers:{'Content-Type':'application/json','X-Audit-Token':清单.令牌},body:JSON.stringify({请求标识:'成品禁止继续'+任务.task_id})})
    assert.equal(拒绝.status,409)
    assert.equal((await 接口('/api/桌面/任务/'+任务.task_id+'/进度')).执行次数,1)
    assert.equal(哈希(输出),成品哈希)
    assert.deepEqual(['A','B'].map(s=>哈希(场景.输入[s].路径)),原件)
    结果.任务.push({...场景,...任务,输出,设置往返保留:true,拒绝续跑:拒绝.status,成果未变:true,成品哈希,原件未变:true})
    await 页.getByRole('button',{name:'再做一份',exact:true}).click()
    assert(await 页.getByTestId('开始任务').isDisabled())
    assert.equal(await 页.locator('.文件卡 .文件路径').count(),0)
    await 页.getByRole('button',{name:'历史任务',exact:true}).click()
    await 页.getByRole('textbox',{name:'搜索历史任务'}).fill(任务.task_id)
    await 页.getByRole('button',{name:'查询',exact:true}).click()
    await 页.getByRole('button',{name:'查看任务',exact:true}).click()
    await 页.getByRole('heading',{name:'核查本次结果'}).waitFor()
  }
  await 页.screenshot({path:path.join(证据,'03_真实结果.png'),fullPage:true})
  for(const 页名 of ['识别记录','设置','人工回写']) {
    await 页.getByRole('button',{name:页名,exact:true}).click()
    await 页.getByRole('heading',{name:页名,exact:true}).waitFor()
    await 页.screenshot({path:path.join(证据,'页面_'+页名+'.png'),fullPage:true})
  }
  for(const 密度 of [1,1.25,1.5]) {
    const c=await 浏览器.newContext({viewport:{width:900,height:560},deviceScaleFactor:密度,reducedMotion:'reduce'})
    const p=await c.newPage();await p.goto(清单.地址+'/桌面')
    await p.getByRole('heading',{name:'开始一份搬运任务'}).waitFor()
    for(const 页名 of ['新建任务','历史任务','识别记录','设置','人工回写']) {
      await p.getByRole('button',{name:页名,exact:true}).click()
      const 宽=await p.evaluate(()=>({页面:document.documentElement.scrollWidth,窗口:innerWidth}))
      assert(宽.页面<=宽.窗口+1,页名+'出现整页横向溢出')
      结果.显示检查.push({页名,密度,视口:'900×560',...宽})
    }
    await p.getByRole('button',{name:'历史任务',exact:true}).click()
    await p.getByRole('textbox',{name:'搜索历史任务'}).fill('性能样本_100张')
    await p.getByRole('button',{name:'查询',exact:true}).click()
    await p.getByRole('button',{name:'查看任务',exact:true}).click()
    await p.locator('.结果明细').getByRole('button',{name:'查看',exact:true}).first().click()
    await p.getByRole('dialog',{name:'结果详情'}).waitFor()
    await p.getByRole('button',{name:'下一项',exact:true}).click()
    await p.keyboard.press('Escape')
    await p.getByRole('dialog',{name:'结果详情'}).waitFor({state:'hidden'})
    const 宽=await p.evaluate(()=>({页面:document.documentElement.scrollWidth,窗口:innerWidth}))
    assert(宽.页面<=宽.窗口+1,'窄窗口结果页溢出')
    结果.显示检查.push({页名:'结果与详情抽屉',密度,视口:'900×560',...宽})
    await p.getByRole('button',{name:'新建任务',exact:true}).click()
    await p.keyboard.press('Tab')
    assert(await p.evaluate(()=>document.activeElement.tagName!=='BODY'))
    await p.screenshot({path:path.join(证据,`窄窗口_${密度}.png`),fullPage:true})
    await c.close()
  }
  assert.equal(结果.页面错误.length,0)
  结果.通过=true
} catch(e) { 结果.通过=false;结果.错误=e.stack;process.exitCode=1 }
finally {
  fs.writeFileSync(path.join(证据,'浏览器结果.json'),JSON.stringify(结果,null,2))
  await 浏览器?.close()
  if(清单) {
    try {
      const 概况=await 接口('/api/桌面/概况')
      if(概况.活动task_id) { await 接口('/api/桌面/任务/'+概况.活动task_id+'/中断',{}); for(let n=0;n<600;n++){if(!(await 接口('/api/桌面/概况')).活动task_id)break;await 等(500)} }
      await 接口('/验收/停止',{})
    } catch(e) { fs.writeFileSync(path.join(证据,'停止说明.txt'),String(e));process.exitCode=1 }
  }
  fs.closeSync(日志)
  console.log(JSON.stringify({通过:结果.通过,任务数:结果.任务.length,性能:结果.性能,错误:结果.错误},null,2))
}
