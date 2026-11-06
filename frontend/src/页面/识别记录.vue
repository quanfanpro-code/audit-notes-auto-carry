<script setup>
import { ref, computed, onMounted, onBeforeUnmount } from 'vue'
import { 请求, 参数, 请求标识, 最新请求 } from '../请求.js'
import { 状态 } from '../任务状态.js'
import ExplanationView from '../解释查看.vue'
const 列表 = ref(null), 详情 = ref(null), 搜索 = ref(''), 筛选 = ref(''), 页码 = ref(1), 错误 = ref(''), 忙 = ref(false)
const 确认框 = ref(false), 预览 = ref(null), 原因 = ref(''), 依据 = ref(''), 选择 = ref(''), 操作 = ref(''), 结果 = ref(null)
const 管 = 最新请求(), 详情管 = 最新请求()
let 请求号, 提交内容
const 可提交 = computed(() => !!预览.value && 原因.value.trim() && 依据.value.trim() && !忙.value)
async function 查历史(路径) {
  try {
    const r = await 请求('/api/桌面/任务?' + 参数({ 搜索: 路径.split(/[\\/]/).at(-1), limit: 100 }))
    const 原任务 = r.条目.find(x => x.运行目录 === 路径)
    if (!原任务) { 错误.value = '来源任务不在当前历史目录中，请按记录中的原路径查看。'; return }
    状态.查看(原任务.task_id)
  } catch (e) { 错误.value = e.message }
}
async function 查询(重置 = false) {
  if (重置) 页码.value = 1
  try { const r = await 管.执行(signal => 请求('/api/桌面/记录?' + 参数({ 搜索: 搜索.value, 状态: 筛选.value, offset: (页码.value-1)*50, limit: 50 }), undefined, { signal })); if (r) 列表.value = r }
  catch (e) { 错误.value = e.message }
}
async function 查看(版本) {
  try { const r = await 详情管.执行(signal => 请求('/api/桌面/记录/' + encodeURIComponent(版本), undefined, { signal })); if (r) { 详情.value = r; 选择.value = ''; 错误.value = '' } }
  catch (e) { 错误.value = e.message }
}
async function 准备(类型) {
  if (类型 === '裁决' && !选择.value) { 错误.value = '请明确选择保留哪一份解释，或全部停用。'; return }
  操作.value = 类型; 原因.value = ''; 依据.value = ''; 预览.value = null; 请求号 = 请求标识(); 提交内容 = null; 结果.value = null
  const 目标 = 类型 === '停用' ? [详情.value.版本] : 详情.value.冲突解释.map(x => x.版本).filter(x => 选择.value === '全部停用' || x !== 选择.value)
  try { 预览.value = { ...(await 请求('/api/桌面/记录/影响预览', { 版本: 目标 })), 目标 }; 确认框.value = true }
  catch (e) { 错误.value = e.message }
}
async function 提交() {
  if (!可提交.value) return
  const 内容 = { 原因: 原因.value, 依据: 依据.value, 预览代次: 预览.value.代次, ...(操作.value === '停用' ? { 版本: 预览.value.目标 } : { 证据组: 详情.value.证据组, 预期冲突版本: 详情.value.冲突解释.map(x => x.版本), 保留: 选择.value === '全部停用' ? null : 选择.value, 停用: 预览.value.目标 }) }
  const 文 = JSON.stringify(内容)
  if (提交内容 && 提交内容 !== 文) 请求号 = 请求标识()
  提交内容 = 文; 忙.value = true
  try {
    const 接受 = await 请求('/api/桌面/记录/' + 操作.value, { ...内容, 请求标识: 请求号 })
    结果.value = { ...(await 请求('/api/桌面/记录/请求/' + encodeURIComponent(请求号))), 预览后变化: 接受.预览后变化 }
    确认框.value = false; 错误.value = ''; await 查询(); await 查看(详情.value.版本)
  } catch (e) {
    try { const 已有 = await 请求('/api/桌面/记录/请求/' + encodeURIComponent(请求号)); if (已有.状态) { 结果.value = 已有; 确认框.value = false; await 查询() } else 错误.value = e.message }
    catch { 错误.value = e.message }
  } finally { 忙.value = false }
}
onMounted(查询); onBeforeUnmount(() => { 管.取消(); 详情管.取消() })
</script>
<template><div class="页标题"><div><h1>识别记录</h1><p class="辅助">日常自动核对并复用。确认旧解释有错时，在这里查看依据、停用或处理冲突。</p></div><el-button @click="查询()">刷新</el-button></div>
  <el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon/><el-alert v-if="列表 && !列表.可用" :title="列表.原因 || '当前记录簿不可用，任务会按原规则重新识别'" type="warning" :closable="false" show-icon/>
  <section v-if="结果" class="内容卡"><h2>{{ 结果.状态 }} · 受影响任务待复核</h2><el-alert v-if="结果.预览后变化" title="预览后记录簿已有变化，请以下列实际影响清单为准。" type="warning" :closable="false"/><p class="辅助">这是后端实际登记的影响；历史报告不会自动改写。</p><div v-for="类 in ['直接','间接']" :key="类"><strong>{{ 类 }}影响</strong><div v-for="(x,i) in 结果[类]?.任务表组 || []" :key="i" class="文件路径">{{ x.任务 }} · {{ x.侧 }} · {{ x.工作表 }} · {{ x.表组 }}</div></div></section>
  <section class="内容卡"><p class="文件路径">记录位置：{{ 列表?.目录 || '未取得' }}</p><form class="操作行 记录筛选" @submit.prevent="查询(true)"><el-input v-model="搜索" placeholder="搜索科目、表名或解释版本" aria-label="搜索识别记录" clearable/><el-select v-model="筛选" placeholder="全部记录" clearable aria-label="筛选识别记录" @change="查询(true)"><el-option v-for="x in ['可自动复用','冲突','有效','停用']" :key="x" :label="x === '有效' ? '存档有效' : x" :value="x"/></el-select><el-button native-type="submit">查询</el-button></form>
    <el-table :data="列表?.条目 || []" row-key="版本" empty-text="没有匹配记录。完成首次识别后，符合条件的解释会自动保存。"><el-table-column label="科目 / 表名" min-width="170"><template #default="{row}"><strong>{{ row.科目 || '整份布局' }}</strong><div class="辅助">{{ row.表名 || row.步骤 }}</div></template></el-table-column><el-table-column prop="侧" label="首次侧" width="80"/><el-table-column prop="步骤" label="步骤" width="90"/><el-table-column prop="存档状态" label="存档" width="80"/><el-table-column prop="当前适用" label="按当前规则" min-width="180"/><el-table-column prop="来源数" label="来源数" width="80"/><el-table-column label="查看" width="90"><template #default="{row}"><el-button link type="primary" @click="查看(row.版本)">解释与依据</el-button></template></el-table-column></el-table><div class="分页"><el-pagination v-model:current-page="页码" :page-size="50" :total="列表?.总数 || 0" layout="total, prev, pager, next" @current-change="查询()"/></div>
  </section>
  <section v-if="详情" class="内容卡"><div class="详情标题"><h2>解释与来源</h2><el-button v-if="详情.存档状态 === '有效'" @click="准备('停用')">停用这份解释</el-button></div><p class="文件路径">解释版本：{{ 详情.版本 }} · {{ 详情.当前适用 }}</p><ExplanationView :解释="详情.记录.解释"/>
    <template v-if="详情.冲突解释?.length"><h2>同一证据存在不同解释</h2><p class="辅助">请根据原表和识别依据作决定，程序不会按新旧或使用次数替你选择。</p><div class="冲突比较"><section v-for="x in 详情.冲突解释" :key="x.版本" class="原因区"><strong>{{ x.版本 }}</strong><ExplanationView :解释="x.解释"/></section></div><el-radio-group v-model="选择" aria-label="冲突处理选择"><el-radio v-for="x in 详情.冲突解释" :key="x.版本" :value="x.版本">保留 {{ x.版本 }}</el-radio><el-radio value="全部停用">全部停用，之后重新识别</el-radio></el-radio-group><div class="操作行"><el-button :disabled="!选择" @click="准备('裁决')">查看处理影响</el-button></div></template>
    <el-collapse><el-collapse-item title="识别依据与完整解释" name="证据"><pre class="长文">{{ JSON.stringify({证据:详情.记录.证据,解释:详情.记录.解释}, null, 2) }}</pre></el-collapse-item><el-collapse-item title="来源、历史操作及影响" name="来源"><el-table :data="详情.记录.来源 || []"><el-table-column label="来源任务" min-width="200"><template #default="{row}"><div class="文件路径">{{ row.任务 }}</div><el-button v-if="row.任务" link type="primary" @click="查历史(row.任务)">查看来源任务</el-button></template></el-table-column><el-table-column prop="侧" label="侧" width="60"/><el-table-column prop="工作表" label="工作表" min-width="100"/><el-table-column prop="表组" label="表组" width="100"/></el-table><div v-for="(x,i) in 详情.历史操作" :key="i" class="原因区">{{ x.类型 }}：{{ x.原因 }}；依据：{{ x.依据 }}</div><p class="辅助">直接相关任务 {{ 详情.影响.直接.任务.length }} 份，间接相关任务 {{ 详情.影响.间接.任务.length }} 份。</p></el-collapse-item></el-collapse>
  </section>
  <el-dialog v-model="确认框" :title="操作 === '停用' ? '确认停用解释' : '确认冲突处理'" width="min(720px,90vw)" :close-on-click-modal="false"><el-alert v-if="错误" :title="错误" type="error" :closable="false"/><p class="辅助">影响预览时间：{{ 预览?.时点 }}</p><p>将停用 {{ 预览?.目标.join('、') }}，以及依赖它们的下游解释。</p><div class="长文 原因区"><div v-for="类 in ['直接','间接']" :key="类"><strong>{{ 类 }}影响</strong><p v-for="(x,i) in 预览?.[类]?.任务表组 || []" :key="i">{{ x.任务 }} · {{ x.侧 }} · {{ x.工作表 }} · {{ x.表组 }}</p></div></div><el-form label-position="top"><el-form-item label="处理原因"><el-input v-model="原因" type="textarea" :rows="2"/></el-form-item><el-form-item label="核实依据（原表位置或复核记录）"><el-input v-model="依据" type="textarea" :rows="2"/></el-form-item></el-form><p class="辅助">历史任务会提示待复核；原产物不变。提交后展示实际影响清单。</p><template #footer><el-button :disabled="忙" @click="确认框=false">返回核对</el-button><el-button type="primary" :disabled="!可提交" :loading="忙" @click="提交">确认处理</el-button></template></el-dialog>
</template>
<style scoped>.记录筛选{margin:12px 0}.记录筛选 .el-input{max-width:380px}.记录筛选 .el-select{width:170px}.冲突比较{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:12px;margin-bottom:16px}.冲突比较>section{min-width:0}.el-radio-group{display:flex;flex-wrap:wrap;margin-bottom:12px}.el-form{margin-top:16px}.文件路径{margin:8px 0}</style>
