<script setup>
import { ref, computed, watch, onBeforeUnmount } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { 请求, 请求标识, 最新请求 } from '../请求.js'
import { 状态, 状态文字 } from '../任务状态.js'
import StatusTag from '../状态标签.vue'
const 任务 = ref(null), 错误 = ref(''), 忙 = ref(false), 材料们 = ref([]), 材料 = ref(null), 草稿 = ref('')
const 管 = 最新请求(), 材料管 = 最新请求()
let 计时器, 已销毁 = false, 失败次数 = 0, 答复请求, 继续请求
const 基址 = computed(() => '/api/桌面/任务/' + encodeURIComponent(状态.查看任务))
const 步们 = ['extract', 'identify_a', 'identify_b', 'carry', 'writeback', 'readback']
const 阶段 = computed(() => 步们.indexOf(任务.value?.步骤))
const 名 = 路径 => (路径 || '').split(/[\\/]/).pop()
async function 刷新() {
  clearTimeout(计时器)
  const id = 状态.查看任务
  if (!id || 已销毁) return
  try {
    const 数据 = await 管.执行(signal => 请求('/api/桌面/任务/' + encodeURIComponent(id), undefined, { signal }))
    if (!数据 || id !== 状态.查看任务) return
    任务.value = 数据; 错误.value = ''; 失败次数 = 0
    if (数据.活动) 状态.活动任务 = id
    else if (状态.活动任务 === id) { 状态.活动任务 = null; 状态.刷新号++ }
    if (数据.等待 && !材料们.value.length) {
      const r = await 请求(基址.value + '/答复')
      if (id !== 状态.查看任务) return
      材料们.value = r.条目
      if (r.条目.length) await 选材料(r.条目[0].id)
    }
    if (数据.活动) 计时器 = setTimeout(刷新, 1000)
  } catch (e) {
    错误.value = '无法取得最新状态：' + e.message
    失败次数++
    计时器 = setTimeout(刷新, Math.min(30000, 2000 * 失败次数))
  }
}
async function 继续任务() {
  忙.value = true; 错误.value = ''
  try {
    继续请求 ||= 请求标识()
    const r = await 请求(基址.value + '/继续', { 请求标识: 继续请求 })
    继续请求 = null
    材料.value = null; 材料们.value = []; 草稿.value = ''; 状态.未保存.任务详情 = false
    状态.活动任务 = r.task_id; 状态.刷新号++; await 刷新()
  } catch (e) { 错误.value = e.message }
  finally { 忙.value = false }
}
async function 中断() {
  忙.value = true
  try { await 请求(基址.value + '/中断', {}); await 刷新(); 状态.刷新号++ }
  catch (e) { 错误.value = e.message }
  finally { 忙.value = false }
}
async function 选材料(id) {
  if (材料.value && 草稿.value !== 材料.value.已有答复) {
    try { await ElMessageBox.confirm('当前答复还没保存，放弃草稿并切换材料？', '尚未保存', { confirmButtonText: '放弃并切换', cancelButtonText: '返回继续' }) }
    catch { return }
  }
  try {
    const r = await 材料管.执行(signal => 请求(基址.value + '/答复/' + id, undefined, { signal }))
    if (r) { 材料.value = r; 草稿.value = r.已有答复; 答复请求 = null }
  } catch (e) { 错误.value = e.message }
}
async function 保存(并继续 = false) {
  if (!材料.value || 忙.value) return
  let 确认 = false
  if (材料.value.已有答复 && 材料.value.已有答复 !== 草稿.value) {
    try { await ElMessageBox.confirm('原答复会先备份，再保存当前内容。确认替换？', '替换已有答复', { confirmButtonText: '备份并替换', cancelButtonText: '返回' }); 确认 = true }
    catch { return }
  }
  if (!答复请求 || 答复请求.文 !== 草稿.value) 答复请求 = { 文: 草稿.value, id: 请求标识() }
  忙.value = true
  try {
    await 请求(基址.value + '/答复/' + 材料.value.id, { 答复: 草稿.value, 原答复哈希: 材料.value.答复哈希, 替换确认: 确认, 请求标识: 答复请求.id })
    材料.value = await 请求(基址.value + '/答复/' + 材料.value.id)
    状态.未保存.任务详情 = false; ElMessage.success('答复已保存')
    if (并继续) await 继续任务()
  } catch (e) { 错误.value = e.message }
  finally { 忙.value = false }
}
async function 复制(文) { try { await navigator.clipboard.writeText(文); ElMessage.success('已复制') } catch { 错误.value = '复制失败，请选中文字后复制' } }
async function 目录() { try { await 请求(基址.value + '/文件/任务目录/打开', {}) } catch (e) { 错误.value = e.message } }
watch(草稿, 文 => { 状态.未保存.任务详情 = !!材料.value && 文 !== 材料.value.已有答复 })
watch(() => 状态.查看任务, () => { clearTimeout(计时器); 管.取消(); 材料管.取消(); 继续请求 = null; 任务.value = null; 材料.value = null; 材料们.value = []; 草稿.value = ''; 刷新() }, { immediate: true })
onBeforeUnmount(() => { 已销毁 = true; clearTimeout(计时器); 管.取消(); 材料管.取消() })
</script>
<template>
  <el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon><el-button @click="刷新">重新连接</el-button></el-alert>
  <el-skeleton v-if="!任务" :rows="6"/>
  <template v-else>
    <section class="内容卡 任务概览" :class="{ 已结束: !任务.活动 && !任务.等待 }"><div class="页标题"><div><h1>{{ 任务.名称 }}</h1><p class="辅助">{{ 任务.类型 }} · {{ 任务.创建时间 }} · 第 {{ 任务.执行次数 || 1 }} 次执行</p></div><StatusTag :值="任务.显示状态"/></div>
      <el-alert v-for="x in 任务.待复核影响 || []" :key="x.请求" :title="'旧识别记录已处理，本任务待复核：' + x.原因" type="warning" :closable="false" show-icon><span>{{ x.依据 }}；受影响表：{{ x.表.map(t => [t.侧,t.工作表,t.表组].filter(Boolean).join(' · ')).join('；') }}</span></el-alert>
      <div class="文件关系"><div><small>B · 提供数据</small><strong :title="任务.B?.路径">{{ 名(任务.B?.路径) || '未记录' }}</strong><span class="辅助">{{ (任务.B?.工作表 || []).join('、') }}</span></div><span class="关系箭头">→</span><div><small>A · 填入新副本</small><strong :title="任务.A?.路径">{{ 名(任务.A?.路径) || 任务.类型 }}</strong><span class="辅助">{{ (任务.A?.工作表 || []).join('、') }}</span></div></div>
      <div class="操作行"><el-button v-if="任务.允许中断" :disabled="忙 || 任务.显示状态 === '正在中断'" @click="中断">{{ 任务.显示状态 === '正在中断' ? '正在等待安全停止' : '中断任务' }}</el-button><el-button v-if="任务.允许继续 && !任务.等待" :type="任务.步骤 === 'done' ? 'default' : 'primary'" :loading="忙" @click="继续任务">继续此任务</el-button><el-button @click="目录">打开任务目录</el-button><el-button v-if="任务.类型 === '自动搬运' && 任务.A?.路径 && 任务.B?.路径" @click="状态.用原文件新建(任务)">用这两份文件新建</el-button><span v-if="任务.允许继续" class="辅助">继续时使用当前模型连接：{{ 状态.概况.模型?.model || '未配置，使用人工答复' }}；沿用本任务文件与设置。</span></div>
      <p v-if="任务.不能继续原因" class="辅助">{{ 任务.不能继续原因 }}</p>

      <p v-if="任务.错误" class="原因区 错误文本">{{ 任务.错误 }}</p>
      <div v-if="任务.活动 || 任务.等待" role="status" aria-live="polite" class="当前步骤">{{ 状态文字(任务.显示状态) }}<span v-if="任务.活动" class="辅助"> · 已完成 {{ 任务.已完成范围 }} 个读取部分，具体耗时取决于文件和模型。</span><span v-if="任务.等待" class="辅助"> · 复制材料取得模型答复，再粘贴到下方继续。</span></div>
      <ol v-if="任务.活动 && 任务.类型 !== '人工回写'" class="阶段列表"><li v-for="(步, i) in 步们" :key="步" :class="{ 当前: i === 阶段, 已过: i < 阶段 }">{{ 状态文字(步) }}</li></ol>
    </section>
    <section v-if="任务.等待" class="内容卡"><div class="详情标题"><h2>等待模型答复</h2><el-select :model-value="材料?.id" aria-label="待答复材料" @change="选材料"><el-option v-for="x in 材料们" :key="x.id" :value="x.id" :label="x.名称"/></el-select></div>
      <template v-if="材料"><p class="辅助 文件路径">{{ 材料.路径 }}</p><div class="操作行"><el-button @click="复制(材料.指令 + '\n\n' + 材料.材料)">复制全部材料</el-button><el-button @click="复制(材料.指令)">只复制指令</el-button><el-button @click="复制(材料.材料)">只复制原文</el-button></div><el-collapse><el-collapse-item title="查看指令和材料" name="材料"><pre class="长文">{{ 材料.指令 }}
{{ 材料.材料 }}</pre></el-collapse-item></el-collapse><label for="答复正文" class="答复标签">粘贴模型答复</label><el-input id="答复正文" v-model="草稿" type="textarea" :rows="8" placeholder="在这里粘贴模型返回的完整答复"/><p class="辅助">保存检查文字格式；继续后仍由原识别流程核验内容，不会跳过业务检查。</p><div class="操作行"><el-button type="primary" :loading="忙" :disabled="!草稿.trim()" @click="保存(true)">保存答复并继续</el-button><el-button :disabled="忙 || !草稿.trim()" @click="保存(false)">只保存答复</el-button></div></template>
    </section>
    <slot v-if="!任务.活动 && !任务.等待" :任务="任务"/>
    <el-collapse v-if="任务.日志?.length"><el-collapse-item title="处理记录" name="日志"><pre class="长文">{{ 任务.日志.join('\n') }}</pre></el-collapse-item></el-collapse>
  </template>
</template>
<style scoped>
.任务概览.已结束{padding:14px 20px}.已结束 .文件关系{padding:8px 12px;margin:8px 0}.已结束 .页标题{margin-bottom:6px}.已结束 .页标题 h1{font-size:18px}.已结束 .页标题 p{margin:4px 0}.任务概览 .页标题{margin-bottom:12px}.文件关系{display:grid;grid-template-columns:1fr 32px 1fr;gap:12px;background:#f7f9fc;border-radius:6px;padding:12px 16px;margin:12px 0 16px}.文件关系>div{display:flex;flex-direction:column;gap:6px;min-width:0}.文件关系 strong{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:14px}.文件关系 small{color:var(--次文);font-size:12px}.关系箭头{align-self:center;color:var(--次文)}.当前步骤{margin-top:16px;font-size:14px}.阶段列表{display:flex;padding:0;list-style:none;gap:16px;color:var(--次文);font-size:12px;margin-bottom:0;flex-wrap:wrap}.阶段列表 li{padding:8px 0;border-bottom:2px solid transparent}.阶段列表 .当前{color:var(--主色);border-color:var(--主色)}.阶段列表 .已过{color:var(--成功)}.答复标签{display:block;margin:12px 0 8px;font-weight:600}.详情标题 .el-select{width:260px}
</style>
