<script setup>
import { ref, onMounted, onBeforeUnmount, watch } from 'vue'
import { ElMessageBox } from 'element-plus'
import { 状态 } from './任务状态.js'
import { 请求, 最新请求 } from './请求.js'
import UiIcon from './图标.vue'
import NewTask from './页面/新建任务.vue'
import TaskDetail from './页面/任务详情.vue'
import ResultView from './页面/结果查看.vue'
import HistoryPage from './页面/历史任务.vue'
import RecordPage from './页面/识别记录.vue'
import SettingsPage from './页面/设置.vue'
import ManualPage from './页面/人工回写.vue'
const 收起 = ref(false)
const 错误 = ref('')
const 导航 = ['新建任务', '历史任务', '识别记录', '设置', '人工回写']
async function 切换(页, id) {
  if (页 === 状态.页面 && (!id || id === 状态.查看任务)) return
  if (状态.未保存[状态.页面]) {
    try { await ElMessageBox.confirm('本页内容尚未保存，放弃修改并离开？', '尚未保存', { confirmButtonText: '放弃并离开', cancelButtonText: '返回继续' }) }
    catch { return }
    状态.未保存[状态.页面] = false
  }
  状态.页面 = 页
  if (id) 状态.查看(id)
}
const 概况管 = 最新请求()
let 计时器, 已销毁 = false
async function 刷新概况() {
  clearTimeout(计时器)
  if (已销毁) return
  try {
    const r = await 概况管.执行(signal => 请求('/api/桌面/概况', undefined, { signal }))
    if (!r) return
    状态.概况 = r; 状态.活动任务 = r.活动task_id; 错误.value = ''
    document.documentElement.style.setProperty('--字号', (r.默认设置?.字号 || 14) + 'px')
    if (r.活动task_id) 计时器 = setTimeout(刷新概况, 1000)
  } catch (e) { 错误.value = e.message; 计时器 = setTimeout(刷新概况, 5000) }
}
function 同步草稿() { window.pywebview?.api?.标记未保存(Object.values(状态.未保存).some(Boolean)) }
watch(() => Object.values(状态.未保存).some(Boolean), 同步草稿)
watch(() => 状态.刷新号, 刷新概况)
onMounted(() => { 刷新概况(); window.addEventListener('pywebviewready', 同步草稿) })
onBeforeUnmount(() => { 已销毁 = true; clearTimeout(计时器); 概况管.取消(); window.removeEventListener('pywebviewready', 同步草稿) })
</script>
<template>
  <div class="应用" :class="{ '导航收起': 收起 }">
    <aside class="侧栏">
      <div class="品牌"><span class="品牌标志">附</span><div v-if="!收起"><strong>审计附注搬运</strong><small>本机工作台</small></div></div>
      <nav aria-label="主要功能"><button v-for="页 in 导航" :key="页" :class="{ 选中: 状态.页面 === 页 }" :title="页" :aria-label="页" :aria-current="状态.页面 === 页 ? 'page' : undefined" @click="切换(页)"><UiIcon :名称="页"/><span v-if="!收起">{{ 页 }}</span></button></nav>
      <button class="收起导航" @click="收起 = !收起" :aria-label="收起 ? '展开导航' : '收起导航'"><UiIcon 名称="收起"/><span v-if="!收起">收起导航</span></button>
    </aside>
    <div class="工作区"><header class="顶栏"><span>{{ 状态.页面 }}</span><button v-if="状态.活动任务" class="活动提示" @click="切换('任务详情',状态.活动任务)"><span class="运行点"/>有任务正在处理 · 查看</button><span v-else class="辅助">本机处理 · 原文件保留</span></header>
      <main><el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon/><NewTask v-if="状态.页面 === '新建任务'"/><TaskDetail v-else-if="状态.页面 === '任务详情'"><template #default="{任务}"><ResultView :key="任务.task_id + ':' + 任务.执行次数" :任务="任务"/></template></TaskDetail><HistoryPage v-else-if="状态.页面 === '历史任务'"/><RecordPage v-else-if="状态.页面 === '识别记录'"/><SettingsPage v-else-if="状态.页面 === '设置'"/><ManualPage v-else-if="状态.页面 === '人工回写'"/></main>
    </div>
  </div>
</template>
