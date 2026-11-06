import { it, expect, vi } from 'vitest'
import { mount } from '@vue/test-utils'
import ValueCell from '../src/值展示.vue'

it.each(['00123', '0', '-12345678901234567890.123456789', ''])('复制原样值不改变精度或前导零：%s', async 文 => {
  const writeText = vi.fn().mockResolvedValue(undefined)
  Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
  const w = mount(ValueCell, { props: { 值: { 是否取得: true, 类型: 文 === '' ? '空白' : '文字', 原样文本: 文 } } })
  expect(w.text()).toBe(文 === '' ? '空白' : 文)
  await w.get('button').trigger('click')
  expect(writeText).toHaveBeenCalledWith(文)
  w.unmount()
})

it('未取得的值不提供复制按钮', () => {
  const w = mount(ValueCell, { props: { 值: { 是否取得: false, 类型: '读取失败' } } })
  expect(w.text()).toBe('读取失败')
  expect(w.find('button').exists()).toBe(false)
  w.unmount()
})
