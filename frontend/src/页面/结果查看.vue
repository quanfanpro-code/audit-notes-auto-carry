<script setup>
import { ref, computed, watch, onMounted, onBeforeUnmount } from 'vue'
import { ElMessage, ElDrawer } from 'element-plus'
import { 请求, 参数, 最新请求 } from '../请求.js'
import { 状态 } from '../任务状态.js'
import StatusTag from '../状态标签.vue'
import ValueCell from '../值展示.vue'
const props = defineProps({ 任务: Object })
const 查看 = 状态.查看状态(props.任务.task_id)
const 数据 = ref(null), 组们 = ref([]), 组总数 = ref(0), 文件 = ref([]), 错误 = ref(''), 加载 = ref(false), 详情 = ref(null), 文件展开 = ref([]), 宽度 = ref(window.innerWidth), 高度 = ref(window.innerHeight)
const 管 = 最新请求(), 组管 = 最新请求()
let 版本 = '', 返回焦点
const 基址 = '/api/桌面/任务/' + encodeURIComponent(props.任务.task_id)
const 抽屉 = computed(() => 宽度.value < 1280)
const 行们 = computed(() => 数据.value?.条目 || [])
const 明细格 = computed(() => 详情.value?.A原值)
const 位置文 = x => x?.工作表 ? `${x.工作表} · ${x.格}` : x?.格 || '未记录'
const 可前 = computed(() => 行们.value.findIndex(x => x.id === 详情.value?.id) > 0 || 查看.页码 > 1)
const 可后 = computed(() => 行们.value.findIndex(x => x.id === 详情.value?.id) < 行们.value.length - 1 || 查看.页码 * 查看.每页 < (数据.value?.总数 || 0))
async function 读取(清版本 = false) {
  if (清版本) 版本 = ''
  加载.value = true
  try {
    const r = await 管.执行(signal => 请求(基址 + '/结果?' + 参数({ view: 查看.视图, 组: 查看.组, 搜索: 查看.搜索, 状态: 查看.状态, offset: (查看.页码-1)*查看.每页, limit: 查看.每页, 结果版本: 版本 }), undefined, { signal }))
    if (r) { 数据.value = r; 版本 = r.结果版本; 错误.value = ''; 详情.value = r.条目.find(x => x.id === 查看.选中) || null }
  } catch (e) {
    if (e.状态码 === 409 && 版本) { 版本 = ''; 查看.页码 = 1; 查看.选中 = null; 错误.value = '结果已更新，请刷新后继续核查。' }
    else 错误.value = e.message
  } finally { 加载.value = false }
}
async function 读组(追加 = false) {
  if (!追加) { 组们.value = []; 组总数.value = 0 }
  try {
    const r = await 组管.执行(signal => 请求(基址 + '/结果?' + 参数({ view: '表', 侧: 'A', 搜索: 查看.表搜索, 仅需查看: 查看.仅需查看 || '', offset: 追加 ? 组们.value.length : 0, limit: 100 }), undefined, { signal }))
    if (r) { 组们.value = 追加 ? [...组们.value, ...r.条目] : r.条目; 组总数.value = r.总数 }
  } catch (e) { 错误.value = e.message }
}
async function 读文件() { try { 文件.value = (await 请求(基址 + '/文件')).条目 } catch (e) { 错误.value = e.message } }
async function 打开(x, 目录 = false) { try { await 请求(基址 + '/文件/' + x.file_id + '/打开', { 打开目录: 目录 }) } catch (e) { 错误.value = e.message } }
function 筛选() { 查看.页码 = 1; 查看.选中 = null; 读取() }
function 选组(x) { 查看.组 = x || ''; 筛选() }
function 选行(x) { 返回焦点 = document.activeElement; 详情.value = x; 查看.选中 = x.id }
function 关闭详情() { 详情.value = null; 查看.选中 = null; 返回焦点?.focus?.() }
async function 相邻(方向) {
  const i = 行们.value.findIndex(x => x.id === 详情.value?.id) + 方向
  if (i >= 0 && i < 行们.value.length) { 选行(行们.value[i]); return }
  查看.页码 += 方向
  await 读取()
  if (行们.value.length) 选行(方向 > 0 ? 行们.value[0] : 行们.value.at(-1))
}
function 尺寸() { 宽度.value = window.innerWidth; 高度.value = window.innerHeight }
function 按键(e) { if (e.key === 'Escape' && 详情.value) 关闭详情() }
async function 复制详情() {
  try { await navigator.clipboard.writeText(JSON.stringify(详情.value, null, 2)); ElMessage.success('已复制本条依据') }
  catch { 错误.value = '复制失败，请选中文字复制' }
}
watch(() => 查看.视图, () => { 查看.组 = ''; 查看.状态 = ''; 筛选() })
onMounted(() => { 读取(); 读组(); 读文件(); window.addEventListener('resize', 尺寸); window.addEventListener('keydown', 按键) })
onBeforeUnmount(() => { 管.取消(); 组管.取消(); window.removeEventListener('resize', 尺寸); window.removeEventListener('keydown', 按键) })
</script>
<template>
  <section class="内容卡 结果结论"><div class="详情标题"><h2>核查本次结果</h2><div class="操作行"><el-button v-for="x in 文件.filter(x => ['Word新副本','Excel新副本'].includes(x.名称))" :key="x.file_id" type="primary" :disabled="!x.存在" @click="打开(x)">打开{{ x.名称.startsWith('Word') ? 'Word' : 'Excel' }}新副本</el-button><el-button @click="状态.新建()">再做一份</el-button></div></div>
    <div class="结果计数" v-if="数据"><button class="问题入口" @click="查看.仅需查看=true;读组(false)">{{ 数据.需查看表数 || 0 }} 张表需要查看</button><span v-for="(数, 名) in 数据.计数" :key="名"><strong>{{ 数 }}</strong> 格{{ 名 }}</span><button class="问题入口" @click="查看.视图 = '问题'">{{ 数据.问题条数 }} 条问题 · 查看</button></div>
    <p class="辅助" v-if="数据">{{ 数据.明细完整性 }}。标题单独列示，问题条数不等于格数。</p>
    <el-alert v-for="提示 in 数据?.提示 || []" :key="提示" :title="提示" type="warning" :closable="false" show-icon/>
    <el-alert v-for="x in 文件.filter(x => x.正式 && x.文件状况 !== '未改动')" :key="x.file_id" :title="x.名称 + '：' + x.文件状况 + '。下方展示的是历史处理记录。'" type="warning" :closable="false" show-icon/>
    <el-collapse v-model="文件展开"><el-collapse-item title="文件、候选副本与材料" name="文件"><div v-for="x in 文件" :key="x.file_id" class="文件行"><div><strong>{{ x.名称 }}</strong><span class="辅助"> · {{ x.候选 ? '候选，不能当作正式成果' : x.正式 ? '正式副本' : x.名称.endsWith('输入原件') ? '输入原件' : '处理材料' }} · {{ x.文件状况 }}</span><div class="文件路径">{{ x.路径 }}</div></div><div class="操作行"><el-button :disabled="!x.存在" @click="打开(x)">打开</el-button><el-button @click="打开(x,true)">所在目录</el-button></div></div><el-button @click="读文件">重新检查文件</el-button></el-collapse-item></el-collapse>
  </section>
  <el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon><el-button @click="读取(true)">刷新结果</el-button></el-alert>
  <div class="结果工作台" :class="{ 有详情: 详情 && !抽屉 }">
    <aside class="内容卡 业务表"><h2>业务表</h2><form @submit.prevent="读组(false)"><el-input v-model="查看.表搜索" placeholder="搜索科目或表名" aria-label="搜索业务表" clearable @clear="读组(false)"/><el-button native-type="submit" link>查询业务表</el-button></form><el-checkbox v-model="查看.仅需查看" @change="读组(false)">只看需要查看的表</el-checkbox><button :class="{ 当前: !查看.组 }" @click="选组('')">全部结果</button><button v-for="x in 组们.filter(x => x.侧 === 'A')" :key="x.id" :class="{ 当前: 查看.组 === x.组键 }" @click="选组(x.组键)"><strong>{{ x.科目 || '未识别科目' }}</strong><span>{{ x.表名 || x.表组 || '未记录表名' }}</span><small>{{ x.物理表?.map(p => p.工作表 || p.区域).join('、') }}</small><small>{{ x.物理表?.length || 1 }} 张主表 / 续表 · {{ x.需要查看 ? '需要查看' : '已记录' }}</small><small v-if="x.问题数">{{ x.问题数 }} 条问题</small></button><el-button v-if="组们.length < 组总数" @click="读组(true)">加载更多</el-button></aside>
    <section class="内容卡 结果明细"><el-tabs v-model="查看.视图"><el-tab-pane v-for="名 in ['明细','问题','标题','公式','复用']" :key="名" :label="名 === '明细' ? '搬运明细' : 名" :name="名"/></el-tabs>
      <form class="操作行 结果筛选" @submit.prevent="筛选"><el-input v-model="查看.搜索" placeholder="搜索项目、期间或原因" aria-label="搜索结果" clearable/><el-select v-if="查看.视图 === '明细'" v-model="查看.状态" placeholder="全部处理方式" aria-label="筛选处理方式" clearable @change="筛选"><el-option v-for="x in ['已写入','数值未变','已清空','原值保留','公式保留','回读有差异','未记录']" :key="x" :label="x" :value="x"/></el-select><el-button native-type="submit">查询</el-button><el-button v-if="查看.搜索 || 查看.状态" @click="查看.搜索='';查看.状态='';筛选()">清除筛选</el-button></form>
      <el-table v-loading="加载" :max-height="Math.max(260, 高度 - 420)" :data="行们" row-key="id" highlight-current-row :current-row-key="查看.选中" empty-text="没有匹配内容。可清除筛选或查看其他结果。" @row-dblclick="选行">
        <template v-if="查看.视图 === '明细'"><el-table-column prop="项目" label="项目" min-width="110"/><el-table-column prop="期间" label="期间" min-width="110"/><el-table-column v-for="列 in ['A原值','B本次值','A实际值']" :key="列" :label="列 === 'A原值' ? 'A 原值' : 列 === 'B本次值' ? 'B 本次值' : 'A 最终读值'" min-width="110" align="right"><template #default="{row}"><ValueCell :值="row[列]"/></template></el-table-column><el-table-column label="处理结果" min-width="110"><template #default="{row}"><StatusTag :值="row.最终动作"/></template></el-table-column></template>
        <template v-else-if="查看.视图 === '问题'"><el-table-column prop="原因" label="问题" min-width="220" show-overflow-tooltip/><el-table-column label="位置 / 范围" min-width="140"><template #default="{row}">{{ row.a格 || row.区域 || row.表组 || '任务整体' }}</template></el-table-column></template>
        <template v-else-if="查看.视图 === '标题'"><el-table-column label="位置" min-width="130"><template #default="{row}">{{ row.区域 || row.工作表 }} {{ row.范围 }}</template></el-table-column><el-table-column label="本次标题动作" min-width="200"><template #default="{row}"><div v-for="(动作,i) in row.动作" :key="i">{{ 动作.old_text }} → {{ 动作.text }}</div><span v-if="!row.动作?.length">原样保留</span></template></el-table-column><el-table-column prop="最终动作" label="处理" width="110"/></template>
        <template v-else-if="查看.视图 === '公式'"><el-table-column prop="a格" label="公式位置" min-width="120"/><el-table-column v-for="列 in ['A计算值','B核对值','差额值']" :key="列" :label="列 === '差额值' ? '差额' : 列" min-width="100" align="right"><template #default="{row}"><ValueCell :值="row[列]"/></template></el-table-column><el-table-column prop="结论" label="核对结论" min-width="130"/></template>
        <template v-else><el-table-column prop="侧" label="侧" width="60"/><el-table-column prop="工作表" label="工作表" min-width="120"/><el-table-column prop="读取部分" label="部分" width="90"/><el-table-column v-for="名 in ['第一遍','第二遍','第三遍']" :key="名" :label="名+'调用'" min-width="100"><template #default="{row}">{{ row.调用?.[名] ?? '未记录' }}</template></el-table-column></template>
        <el-table-column label="详情" width="72" fixed="right"><template #default="{row}"><el-button link type="primary" @click="选行(row)">查看</el-button></template></el-table-column>
      </el-table><p v-if="查看.视图 === '公式'" class="辅助">金额相等只代表本次核对一致，不代表公式本身已被全面证明正确。</p>
      <div class="分页"><el-pagination v-model:current-page="查看.页码" v-model:page-size="查看.每页" :page-sizes="[20,50,100]" :total="数据?.总数 || 0" layout="total, sizes, prev, pager, next" @current-change="读取()" @size-change="筛选"/></div>
    </section>
    <component :is="抽屉 ? ElDrawer : 'aside'" v-if="详情" :model-value="true" :title="'结果详情'" size="420px" class="详情容器" :class="{ 内容卡: !抽屉 }" @close="关闭详情">
      <div class="详情标题"><h2 v-if="!抽屉">结果详情</h2><el-button v-if="!抽屉" @click="关闭详情" aria-label="关闭详情">关闭</el-button></div><div class="操作行"><el-button :disabled="!可前" @click="相邻(-1)">上一项</el-button><el-button :disabled="!可后" @click="相邻(1)">下一项</el-button></div>
      <template v-if="明细格"><dl><div class="详情项"><dt>项目与期间</dt><dd>{{ 详情.项目 }} · {{ 详情.期间 }}</dd></div><div class="详情项"><dt>A 位置</dt><dd>{{ 位置文(详情.A位置) }}</dd></div><div class="详情项"><dt>B 位置</dt><dd>{{ 位置文(详情.B位置) }}</dd></div><div class="详情项" v-for="列 in ['A原值','B本次值','A实际值']" :key="列"><dt>{{ 列 }}</dt><dd><ValueCell :值="详情[列]"/></dd></div><div class="详情项"><dt>最终动作</dt><dd><StatusTag :值="详情.最终动作"/></dd></div></dl><div v-for="(原因,i) in 详情.原因" :key="i" class="原因区"><strong>{{ 原因.原因 }}</strong><p v-if="原因.说明">{{ 原因.说明 }}</p><p v-if="原因.错误">{{ 原因.错误 }}</p></div><p v-if="!详情.原因?.length" class="辅助">该格没有单独记录问题。</p></template>
      <template v-else-if="查看.视图 === '复用'"><p v-if="详情.库?.原因" class="原因区">{{ 详情.库.原因 }}</p><div v-for="类 in ['复用','回退','跳过','首认','发布']" :key="类"><h3 v-if="详情[类]?.length">{{ 类 === '回退' ? '重新识别的原因' : 类 === '跳过' ? '本次未精读' : 类 === '首认' ? '本次首次识别' : 类 }}</h3><p v-for="(项,i) in 详情[类] || []" :key="i" class="原因区">{{ 项.步骤 }} · {{ 项.表组 }}<br/>{{ 项.原因 || 项.版本 || '见已有记录' }}</p></div></template>
      <template v-else><dl><div v-for="(文,键) in 详情" :key="键" class="详情项"><template v-if="!['id','组键'].includes(键) && typeof 文 !== 'object'"><dt>{{ 键 }}</dt><dd>{{ 文 ?? '未记录' }}</dd></template></div></dl></template>
      <el-collapse><el-collapse-item title="查看原始留证" name="依据"><pre class="长文">{{ JSON.stringify(详情,null,2) }}</pre><el-button @click="复制详情">复制本条依据</el-button></el-collapse-item></el-collapse>
    </component>
  </div>
</template>
<style scoped>
.结果结论{padding:12px 20px}.结果结论 :deep(.el-collapse-item__header){height:32px}.结果结论 .辅助{margin:8px 0}.结果计数{display:flex;gap:16px;align-items:center;flex-wrap:wrap;margin:12px 0;font-size:13px}.结果计数 strong{font-size:18px;font-variant-numeric:tabular-nums;margin-right:4px}.问题入口{color:var(--主色);background:none;border:0;padding:4px}.文件行{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:12px}.结果工作台{display:grid;grid-template-columns:220px minmax(0,1fr);gap:16px;align-items:start}.结果工作台.有详情{grid-template-columns:200px minmax(0,1fr) 320px}.业务表{padding:12px;position:sticky;top:72px;max-height:calc(100vh - 96px);overflow:auto}.业务表 h2{padding:4px 8px}.业务表>button{display:flex;flex-direction:column;gap:4px;border:0;background:none;border-left:2px solid transparent;text-align:left;padding:10px 8px;width:100%;color:var(--正文);border-radius:4px;overflow-wrap:anywhere}.业务表>button small{color:var(--次文);font-size:11px}.业务表>button span{font-size:12px}.业务表>button:hover{background:#f4f7fb}.业务表>button.当前{background:#eaf1ff;border-left-color:var(--主色)}.结果明细{min-width:0;padding:12px 16px}.结果筛选{margin-bottom:12px}.结果筛选 .el-input{flex:1;min-width:140px}.结果筛选 .el-select{width:150px}.详情容器{min-width:0}.详情容器 dl{margin-top:16px}.详情容器 .值按钮{white-space:pre-wrap}.详情容器.内容卡{position:sticky;top:72px;max-height:calc(100vh - 96px);overflow:auto;padding:16px}@media(max-width:1100px){.结果工作台{grid-template-columns:1fr}.业务表{position:static;max-height:180px;display:flex;flex-wrap:wrap;gap:4px}.业务表 h2{width:100%;margin:0}.业务表>button{width:auto;max-width:240px}.结果筛选 .el-select{width:140px}.分页{overflow:auto;justify-content:flex-start}}
</style>
