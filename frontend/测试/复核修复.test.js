import { it, expect, vi, beforeEach } from 'vitest'
import { mount, flushPromises } from '@vue/test-utils'
import ElementPlus from 'element-plus'
import App from '../src/App.vue'
import NewTask from '../src/页面/新建任务.vue'
import ResultView from '../src/页面/结果查看.vue'
import { 状态, 创建任务状态 } from '../src/任务状态.js'
import { 请求 } from '../src/请求.js'
vi.mock('../src/请求.js', async original => ({ ...(await original()), 请求: vi.fn() }))
const 选项 = { global: { plugins: [ElementPlus] } }
beforeEach(() => { Object.assign(状态, 创建任务状态()); 请求.mockReset() })
const 按钮 = (w, 名) => w.findAll('button').find(x => x.text() === 名)

it('选好文件工作表和本次设置，去设置再返回仍保留，开始后再做一份才清空', async () => {
  const 默认设置 = { 金标准: 'C:\\金标准.jsonl', 复用: true }
  const 路径 = ['C:\\数据.xlsx', 'C:\\附注.xlsx']
  let 提交
  请求.mockImplementation(async (p, 内容) => {
    if (p === '/api/桌面/概况') return { 默认设置, 活动task_id: null }
    if (p === '/api/桌面/设置') return { 模型: {}, 默认设置 }
    if (p === '/api/选文件') return { 路径: 路径.shift() }
    if (p === '/api/工作表') return { 工作表: [{ 名称: '资金', 可见: true }, { 名称: '另一表', 可见: true }] }
    if (p === '/api/桌面/任务') { 提交 = 内容; return { task_id: '任务_新' } }
    if (p.endsWith('/文件')) return { 条目: [] }
    if (p.includes('/结果?')) return { 条目: [], 总数: 0, 提示: [] }
    return { task_id: '任务_新', 类型: '自动搬运', 步骤: 'done', A: {}, B: {} }
  })
  const w = mount(App, 选项)
  try {
    await flushPromises()
    for (const 侧 of ['B', 'A']) {
      await w.get(`[data-test="选择${侧}"]`).trigger('click'); await flushPromises()
      const 表 = w.get(`[aria-label="${侧}端工作表"]`).findAll('input')
      await 表[0].setValue(true)
    }
    await w.get('[aria-label="自动复用"]').trigger('click')
    await 按钮(w, '设置').trigger('click'); await flushPromises()
    await 按钮(w, '新建任务').trigger('click'); await flushPromises()
    expect(w.text()).toContain('数据.xlsx'); expect(w.text()).toContain('附注.xlsx')
    expect(w.get('[aria-label="A端工作表"] input').element.checked).toBe(true)
    expect(w.get('[aria-label="B端工作表"] input').element.checked).toBe(true)
    expect(w.get('[data-test="开始任务"]').element.disabled).toBe(false)
    await w.get('[data-test="开始任务"]').trigger('click'); await flushPromises()
    expect(提交.原始输入.A).toEqual({ 路径: 'C:\\附注.xlsx', 工作表: ['资金'] })
    expect(提交.原始输入.B).toEqual({ 路径: 'C:\\数据.xlsx', 工作表: ['资金'] })
    expect(提交.原始输入.复用).toBe(false)
    await 按钮(w, '再做一份').trigger('click'); await flushPromises()
    expect(w.text()).not.toContain('数据.xlsx'); expect(w.text()).not.toContain('附注.xlsx')
  } finally { w.unmount() }
})

it('新文件工作表读取被离页取消，回来能重新读取，不遗留无法选择的草稿', async () => {
  状态.概况 = { 默认设置: { 金标准: 'C:\\金标准.jsonl' } }
  let 旧返回
  let 次 = 0
  请求.mockImplementation(async p => p === '/api/选文件' ? { 路径: 'C:\\附注.xlsx' } : ++次 === 1
    ? new Promise(r => { 旧返回 = r }) : { 工作表: [{ 名称: '资金', 可见: true }] })
  const w = mount(NewTask, 选项)
  await w.get('[data-test="选择A"]').trigger('click'); await flushPromises(); w.unmount()
  const 返回 = mount(NewTask, 选项)
  try {
    await flushPromises()
    旧返回({ 工作表: [{ 名称: '过时表', 可见: true }] }); await flushPromises()
    expect(返回.text()).toContain('附注.xlsx'); expect(返回.text()).toContain('资金')
    expect(返回.text()).not.toContain('过时表')
  } finally { 返回.unmount() }
})

it('用原文件新建时离页，旧的分步读取不能再覆盖返回后新选的B文件', async () => {
  状态.概况 = { 默认设置: { 金标准: 'C:\\金标准.jsonl' } }
  状态.用原文件新建({ A: { 路径: 'C:\\原A.xlsx' }, B: { 路径: 'C:\\原B.xlsx' } })
  let 放行旧
  let 首次 = true
  请求.mockImplementation(async (p, 内容) => {
    if (p === '/api/选文件') return { 路径: 'C:\\新B.xlsx' }
    if (首次) { 首次 = false; return new Promise(r => { 放行旧 = r }) }
    return { 工作表: [{ 名称: 内容.路径.includes('新B') ? '新资金' : '原资金', 可见: true }] }
  })
  const w = mount(NewTask, 选项)
  await flushPromises(); w.unmount()
  const 返回 = mount(NewTask, 选项)
  try {
    await flushPromises()
    await 返回.get('[data-test="选择B"]').trigger('click'); await flushPromises()
    放行旧({ 工作表: [{ 名称: '旧A表', 可见: true }] }); await flushPromises()
    expect(返回.text()).toContain('新B.xlsx')
    expect(返回.text()).not.toContain('原B.xlsx')
    expect(返回.get('[aria-label="B端工作表"]').text()).toContain('新资金')
  } finally { 返回.unmount() }
})

it.each([false, true])('业务表旧查询或旧加载更多晚到不能覆盖新筛选，追加=%s', async 追加 => {
  let 旧返回
  const 表 = 名 => ({ id: 名, 组键: 名, 侧: 'A', 科目: 名, 表名: '余额', 计数: {} })
  请求.mockImplementation(async p => {
    if (p.endsWith('/文件')) return { 条目: [] }
    const q = new URL(p, 'http://localhost').searchParams
    if (q.get('view') !== '表') return { 条目: [], 提示: [], 总数: 0 }
    if (q.get('搜索') === '新科目') return { 条目: [表('新科目')], 总数: 1 }
    if (q.get('搜索') === '旧科目' || Number(q.get('offset')) > 0) return new Promise(r => { 旧返回 = r })
    return { 条目: [表('第一批')], 总数: 101 }
  })
  const w = mount(ResultView, { ...选项, props: { 任务: { task_id: '任务_筛选', 执行次数: 1 } } })
  try {
    await flushPromises()
    const 框 = w.get('input[aria-label="搜索业务表"]')
    if (追加) await 按钮(w, '加载更多').trigger('click')
    else { await 框.setValue('旧科目'); await w.get('.业务表 form').trigger('submit') }
    await flushPromises()
    await 框.setValue('新科目'); await w.get('.业务表 form').trigger('submit'); await flushPromises()
    expect(w.get('.业务表').text()).toContain('新科目')
    旧返回({ 条目: [表('旧科目')], 总数: 101 }); await flushPromises()
    expect(w.get('.业务表').text()).toContain('新科目')
    expect(w.get('.业务表').text()).not.toContain('旧科目')
  } finally { w.unmount() }
})
