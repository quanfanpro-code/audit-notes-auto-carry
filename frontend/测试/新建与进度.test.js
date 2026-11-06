import { it, expect, vi } from 'vitest'
import { mount, flushPromises } from '@vue/test-utils'
import ElementPlus from 'element-plus'
import 新建任务 from '../src/页面/新建任务.vue'
import 任务详情 from '../src/页面/任务详情.vue'
import { 状态 } from '../src/任务状态.js'
import { 请求 } from '../src/请求.js'
vi.mock('../src/请求.js', async original => ({ ...(await original()), 请求: vi.fn() }))

it('换文件后隐藏工作表不可选，唯一可见表自动选中', async () => {
  状态.概况 = { 默认设置: { 金标准: 'C:\\基准.jsonl' } }
  状态.活动任务 = null
  请求.mockImplementation(async (p) => p === '/api/选文件' ? { 路径: 'C:\\报表.xlsx' } : { 工作表: [{ 名称: '货币资金', 可见: true }, { 名称: '隐藏表', 可见: false }] })
  const w = mount(新建任务, { global: { plugins: [ElementPlus] } })
  await w.get('[data-test="选择A"]').trigger('click')
  await flushPromises()
  expect(w.text()).toContain('隐藏表')
  const boxes = w.findAll('input[type="checkbox"]')
  expect(boxes.some(x => x.element.disabled && !x.element.checked)).toBe(true)
  expect(boxes.some(x => !x.element.disabled && x.element.checked)).toBe(true)
  w.unmount()
})

it('选择来源不全不能开始，运行期间不能再开任务', async () => {
  状态.活动任务 = '任务_正在跑'
  const w = mount(新建任务, { global: { plugins: [ElementPlus] } })
  expect(w.get('[data-test="开始任务"]').element.disabled).toBe(true)
  expect(w.text()).toContain('有任务正在处理')
  w.unmount()
})

it('等待答复停止轮询，中断后显示真实终态', async () => {
  vi.useFakeTimers()
  状态.查看任务 = '任务_等待'; 状态.活动任务 = null
  let 中断 = false
  请求.mockImplementation(async p => {
    if (p.endsWith('/答复')) return { 条目: [] }
    if (p.endsWith('/中断')) { 中断 = true; return { ok: true } }
    return { task_id: '任务_等待', 名称: '样本.xlsx', 类型: '自动搬运', A: {}, B: {}, 活动: false,
      等待: !中断, 显示状态: 中断 ? '用户中断' : '等待答复', 允许中断: !中断 }
  })
  const w = mount(任务详情, { global: { plugins: [ElementPlus] } })
  await flushPromises()
  const 次数 = 请求.mock.calls.length
  await vi.advanceTimersByTimeAsync(6000)
  expect(请求.mock.calls.length).toBe(次数)
  const 按钮 = w.findAll('button').find(x => x.text() === '中断任务')
  await 按钮.trigger('click'); await flushPromises()
  expect(w.text()).toContain('用户中断')
  w.unmount(); vi.useRealTimers()
})
