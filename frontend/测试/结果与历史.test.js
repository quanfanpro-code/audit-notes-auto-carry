import { it, expect, vi } from 'vitest'
import { mount, flushPromises } from '@vue/test-utils'
import ElementPlus from 'element-plus'
import 结果查看 from '../src/页面/结果查看.vue'
import 历史任务 from '../src/页面/历史任务.vue'
import { 状态 } from '../src/任务状态.js'
import { 请求 } from '../src/请求.js'
vi.mock('../src/请求.js', async original => ({ ...(await original()), 请求: vi.fn() }))
const 值 = 文 => ({ 类型: '文字', 原样文本: 文, 是否取得: true })

it('未取得实际值保持未知，名单前导零不丢，详情可连续查看', async () => {
  window.innerWidth = 1440
  状态.查看任务 = '任务_结果'
  请求.mockImplementation(async p => {
    if (p.endsWith('/文件')) return { 条目: [] }
    if (p.includes('view=%E8%A1%A8')) return { 条目: [{ id: '组1', 组键: '组1', 侧: 'A', 科目: '货币资金', 表名: '余额', 计数: {} }], 总数: 1 }
    return { 结果版本: 'v1', 明细完整性: '最终事实已记录', 提示: [], 计数: { 原值保留: 1 }, 问题条数: 1, 总数: 2, 条目: [
      { id: '格1', 项目: '名单', 期间: '期末余额', A位置: { 工作表: '余额!本期', 格: 'C2' }, B位置: {}, A原值: 值('00123'), B本次值: 值('901'), A实际值: { 类型: '未记录', 是否取得: false }, 最终动作: '原值保留', 原因: [{ 原因: '整组保留' }] },
      { id: '格2', 项目: '合计', 期间: '期末余额', A位置: {}, B位置: {}, A原值: 值('0'), B本次值: 值('0'), A实际值: 值('0'), 最终动作: '数值未变', 原因: [] }
    ] }
  })
  const w = mount(结果查看, { props: { 任务: { task_id: '任务_结果', 执行次数: 1 } }, global: { plugins: [ElementPlus] } })
  await flushPromises()
  expect(w.text()).toContain('00123')
  expect(w.text()).toContain('未取得')
  await w.findAll('button').find(x => x.text() === '查看').trigger('click')
  expect(w.text()).toContain('整组保留')
  await w.findAll('button').find(x => x.text() === '下一项').trigger('click')
  expect(状态.查看状态('任务_结果').选中).toBe('格2')
  w.unmount()
})

it('历史页选择指定任务不改变活动任务', async () => {
  状态.活动任务 = '任务_运行中'
  请求.mockResolvedValue({ 总数: 1, 条目: [{ task_id: '任务_旧', 名称: '以前的报告.docx', 显示状态: '部分保留', A: {}, B: {} }] })
  const w = mount(历史任务, { global: { plugins: [ElementPlus] } })
  await flushPromises()
  await w.findAll('button').find(x => x.text() === '查看任务').trigger('click')
  expect(状态.查看任务).toBe('任务_旧')
  expect(状态.活动任务).toBe('任务_运行中')
  w.unmount()
})

it('从任务返回历史时保留筛选与页码', async () => {
  请求.mockResolvedValue({ 总数: 150, 条目: [] })
  const 选项 = { global: { plugins: [ElementPlus] } }
  const w = mount(历史任务, 选项)
  await flushPromises()
  await w.get('input[aria-label="搜索历史任务"]').setValue('上一份报告')
  await w.get('form').trigger('submit')
  await flushPromises()
  w.unmount()
  请求.mockClear()
  const 返回 = mount(历史任务, 选项)
  await flushPromises()
  expect(返回.get('input[aria-label="搜索历史任务"]').element.value).toBe('上一份报告')
  expect(decodeURIComponent(请求.mock.calls[0][0])).toContain('搜索=上一份报告')
  返回.unmount()
})
