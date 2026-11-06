import { reactive } from 'vue'

export function 创建任务状态() {
  return reactive({
    页面: '新建任务', 查看任务: null, 活动任务: null, 概况: {},
    未保存: {}, 查看记录: {}, 刷新号: 0, 新建预填: null,
    历史查看: { 搜索: '', 筛选: '', 页码: 1, 每页: 50 },
    查看(id) { this.查看任务 = id; this.页面 = '任务详情' },
    用原文件新建(任务) { this.新建预填 = { A: 任务.A, B: 任务.B }; this.页面 = '新建任务' },
    查看状态(id) {
      return this.查看记录[id] ||= { 视图: '明细', 组: '', 搜索: '', 状态: '', 页码: 1, 每页: 50, 选中: null, 表搜索: '', 仅需查看: false }
    }
  })
}
export const 状态 = 创建任务状态()
export const 状态文字 = (步骤) => ({ done: '处理结束', extract: '读取文件', identify_a: '识别 A 端', identify_b: '识别 B 端',
  carry: '生成搬运内容', writeback: '保存新副本', readback: '核对实际结果' }[步骤] || 步骤 || '准备文件')
