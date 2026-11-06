// 所有页面共用的本机请求；不把接口错误或迟到响应当成功。
export async function 请求(路径, 内容, 选项 = {}) {
  const r = await fetch(路径, {
    ...选项,
    method: 内容 === undefined ? 'GET' : 'POST',
    headers: { 'Content-Type': 'application/json', 'X-Audit-Token': window.__AUDIT_TOKEN__ || '' },
    ...(内容 === undefined ? {} : { body: JSON.stringify(内容) })
  })
  const 数据 = await r.json().catch(() => null)
  if (!r.ok || !数据 || 数据.ok === false || (数据.错误 && !数据.task_id)) {
    const 细节 = 数据?.detail
    const 文 = 数据?.错误 || (typeof 细节 === 'string' ? 细节 : 细节?.说明) || 数据?.说明 || `请求未完成（${r.status}）`
    const e = new Error(文)
    e.状态码 = r.status
    throw e
  }
  return 数据
}

export function 参数(内容) {
  return new URLSearchParams(Object.entries(内容).filter(([, v]) => v !== undefined && v !== null && v !== '')).toString()
}

export function 最新请求() {
  let 次数 = 0
  let 控制
  return {
    async 执行(调用) {
      const 本次 = ++次数
      控制?.abort()
      控制 = new AbortController()
      try {
        const 数据 = await 调用(控制.signal)
        return 本次 === 次数 ? 数据 : undefined
      } catch (e) {
        if (本次 !== 次数 || e.name === 'AbortError') return undefined
        throw e
      }
    },
    取消() { 次数++; 控制?.abort() }
  }
}

export const 请求标识 = () => crypto.randomUUID()
