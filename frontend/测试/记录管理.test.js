import { it, expect, vi } from 'vitest'
import { mount, flushPromises } from '@vue/test-utils'
import ElementPlus from 'element-plus'
import RecordPage from '../src/页面/识别记录.vue'
import { 请求 } from '../src/请求.js'
import { 状态 } from '../src/任务状态.js'
vi.mock('../src/请求.js', async original => ({ ...(await original()), 请求: vi.fn() }))

it('冲突没有默认保留项，需主动选择后才可查看处理影响', async () => {
  请求.mockImplementation(async p => p.includes('?') ? { 可用:true, 总数:1, 条目:[{版本:'V1',科目:'货币资金',当前适用:'同一证据有冲突解释'}] } :
    {版本:'V1',存档状态:'有效',记录:{解释:{科目:'货币资金'},来源:[]},冲突解释:[{版本:'V1',解释:{科目:'货币资金'}},{版本:'V2',解释:{科目:'存货'}}],影响:{直接:{任务:[]},间接:{任务:[]}},历史操作:[]})
  const w = mount(RecordPage,{global:{plugins:[ElementPlus]}})
  await flushPromises()
  await w.findAll('button').find(x=>x.text()==='解释与依据').trigger('click'); await flushPromises()
  expect(w.text()).toContain('货币资金'); expect(w.text()).toContain('存货')
  expect(w.findAll('input[type="radio"]').every(x=>!x.element.checked)).toBe(true)
  expect(w.findAll('button').find(x=>x.text()==='查看处理影响').element.disabled).toBe(true)
  w.unmount()
})

it('来源可以按完整原路径查历史，不把同名目录当另一任务', async () => {
  const 路径 = 'C:\\原任务\\任务_来源'
  请求.mockImplementation(async p => p.startsWith('/api/桌面/任务?') ? {条目:[{task_id:'任务_来源',运行目录:路径}]} : p.includes('?') ? {可用:true,总数:1,条目:[{版本:'V1',科目:'货币资金'}]} :
    {版本:'V1',记录:{解释:{},来源:[{任务:路径,侧:'A'}]},冲突解释:[],影响:{直接:{任务:[]},间接:{任务:[]}},历史操作:[]})
  const w = mount(RecordPage,{global:{plugins:[ElementPlus]}})
  await flushPromises()
  await w.findAll('button').find(x=>x.text()==='解释与依据').trigger('click'); await flushPromises()
  await w.findAll('.el-collapse-item__header').find(x=>x.text().includes('来源、历史操作')).trigger('click')
  const 按钮 = w.findAll('button').find(x=>x.text()==='查看来源任务')
  expect(按钮).toBeDefined()
  await 按钮.trigger('click'); await flushPromises()
  expect(状态.页面).toBe('任务详情')
  expect(状态.查看任务).toBe('任务_来源')
  状态.页面 = '识别记录'; 状态.查看任务 = null
  请求.mockResolvedValue({条目:[{task_id:'任务_同名但不同来源',运行目录:'D:\\其他位置\\任务_来源'}]})
  await 按钮.trigger('click'); await flushPromises()
  expect(状态.页面).toBe('识别记录')
  expect(状态.查看任务).toBeNull()
  expect(w.text()).toContain('来源任务不在当前历史目录')
  w.unmount()
})
