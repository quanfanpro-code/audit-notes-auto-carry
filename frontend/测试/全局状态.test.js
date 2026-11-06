import { describe, it, expect, vi } from 'vitest'
import { 请求, 最新请求 } from '../src/请求.js'
import { 创建任务状态 } from '../src/任务状态.js'

describe('全局请求与任务身份', () => {
  it('HTTP成功也不能掩盖业务失败', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ ok: true, json: async () => ({ ok: false, 错误: '具体故障' }) })))
    await expect(请求('/测试')).rejects.toThrow('具体故障')
  })
  it('保留具体HTTP错误并传会话令牌', async () => {
    window.__AUDIT_TOKEN__ = '本例令牌'
    const f = vi.fn(async () => ({ ok: false, status: 409, json: async () => ({ detail: { 说明: '输入文件已改变' } }) }))
    vi.stubGlobal('fetch', f)
    await expect(请求('/测试', {})).rejects.toThrow('输入文件已改变')
    expect(f.mock.calls[0][1].headers['X-Audit-Token']).toBe('本例令牌')
  })
  it('查看历史不改变活动任务，筛选按任务保留', () => {
    const s = 创建任务状态()
    s.活动任务 = '任务_正在跑'
    s.查看('任务_旧')
    s.查看状态('任务_旧').搜索 = '期末'
    s.查看('任务_新')
    expect(s.活动任务).toBe('任务_正在跑')
    s.查看('任务_旧')
    expect(s.查看状态('任务_旧').搜索).toBe('期末')
  })
  it('旧响应不覆盖新选择', async () => {
    const 管 = 最新请求()
    let 完成旧
    const 旧 = 管.执行(() => new Promise(resolve => { 完成旧 = resolve }))
    const 新 = await 管.执行(async () => '新任务')
    完成旧('旧任务')
    expect(新).toBe('新任务')
    expect(await 旧).toBeUndefined()
  })
})
