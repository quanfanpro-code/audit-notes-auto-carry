import { it, expect, vi } from 'vitest'
import { mount, flushPromises } from '@vue/test-utils'
import ElementPlus from 'element-plus'
import SettingsPage from '../src/页面/设置.vue'
import ManualPage from '../src/页面/人工回写.vue'
import { 状态 } from '../src/任务状态.js'
import { 请求 } from '../src/请求.js'
vi.mock('../src/请求.js', async original => ({ ...(await original()), 请求: vi.fn() }))

it('保存密钥后输入框不回显，未保存设置在连接失败后保留', async () => {
  请求.mockImplementation(async (p,体) => { if(p.endsWith('/测试')) throw new Error('网络不可用'); return 体 ? {ok:true} : {模型:{url:'https://example.invalid/v1',model:'原模型',密钥已保存:true},默认设置:{金标准:'C:\\金标准.jsonl',字号:14,复用:true}} })
  const w = mount(SettingsPage,{global:{plugins:[ElementPlus]}})
  await flushPromises()
  expect(w.get('input[type="password"]').element.value).toBe('')
  expect(w.text()).toContain('已保存密钥')
  await w.get('[data-test="模型名称"] input').setValue('新模型')
  await w.findAll('button').find(x=>x.text()==='测试已保存连接').trigger('click'); await flushPromises()
  expect(w.get('[data-test="模型名称"] input').element.value).toBe('新模型')
  expect(w.text()).toContain('网络不可用')
  w.unmount()
})

it('人工回写没有开始识别按钮，活动任务期间禁止回写', () => {
  状态.活动任务='任务_正在跑'
  const w = mount(ManualPage,{global:{plugins:[ElementPlus]}})
  expect(w.text()).toContain('带连接的 Word')
  expect(w.text()).not.toContain('开始识别')
  expect(w.findAll('button').find(x=>x.text()==='另存回写副本').element.disabled).toBe(true)
  w.unmount()
})
