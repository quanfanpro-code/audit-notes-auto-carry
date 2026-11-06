<script setup>
import { ref, toRefs, onMounted, onBeforeUnmount } from 'vue'
import { 请求, 参数, 最新请求 } from '../请求.js'
import { 状态 } from '../任务状态.js'
import StatusTag from '../状态标签.vue'
const { 搜索, 筛选, 页码, 每页 } = toRefs(状态.历史查看)
const 行们 = ref([]), 总数 = ref(0), 加载 = ref(false), 错误 = ref('')
const 管 = 最新请求()
async function 查询(重置 = false) {
  if (重置) 页码.value = 1
  加载.value = true
  try {
    const r = await 管.执行(signal => 请求('/api/桌面/任务?' + 参数({ 搜索: 搜索.value, 状态: 筛选.value, offset: (页码.value - 1) * 每页.value, limit: 每页.value }), undefined, { signal }))
    if (r) { 行们.value = r.条目; 总数.value = r.总数; 错误.value = '' }
  } catch (e) { 错误.value = e.message }
  finally { 加载.value = false }
}
onMounted(查询); onBeforeUnmount(() => 管.取消())
const 名 = p => (p || '').split(/[\\/]/).pop()
</script>
<template><div class="页标题"><div><h1>历史任务</h1><p class="辅助">查看已有结果，或进入指定任务继续处理。查看历史不会改变正在运行的任务。</p></div><el-button @click="查询()">刷新</el-button></div><el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon/>
  <section class="内容卡"><form class="操作行 历史筛选" @submit.prevent="查询(true)"><el-input v-model="搜索" placeholder="搜索文件名或任务" aria-label="搜索历史任务" clearable/><el-select v-model="筛选" placeholder="全部状态" aria-label="筛选历史状态" clearable @change="查询(true)"><el-option v-for="x in ['done','部分保留','公式待核','回读不一致','等待答复','用户中断','任务失败','识别失败','Excel交付失败']" :key="x" :label="x === 'done' ? '处理结束' : x" :value="x"/></el-select><el-button native-type="submit">查询</el-button></form>
    <el-table v-loading="加载" :data="行们" row-key="task_id" empty-text="还没有匹配的任务；可以清除筛选，或开始一份新任务。"><el-table-column label="任务 / A文件" min-width="200"><template #default="{ row }"><strong>{{ row.名称 }}</strong><div class="辅助">{{ row.类型 || '自动搬运' }}</div></template></el-table-column><el-table-column label="B文件" min-width="150"><template #default="{ row }"><span :title="row.B?.路径">{{ 名(row.B?.路径) || '未记录' }}</span></template></el-table-column><el-table-column label="状态" min-width="130"><template #default="{ row }"><StatusTag :值="row.显示状态"/><span v-if="row.待复核影响?.length" class="辅助"> · 识别记录受影响，待复核</span></template></el-table-column><el-table-column label="需查看" width="85"><template #default="{row}">{{ row.需查看表数 ?? '未记录' }}{{ row.需查看表数 == null ? '' : ' 张表' }}</template></el-table-column><el-table-column prop="文件状况" label="正式副本" min-width="140"/><el-table-column prop="创建时间" label="时间" min-width="170"/><el-table-column label="操作" width="110" fixed="right"><template #default="{ row }"><el-button link type="primary" @click="状态.查看(row.task_id)">查看任务</el-button></template></el-table-column></el-table>
    <div class="分页"><el-pagination v-model:current-page="页码" v-model:page-size="每页" :page-sizes="[20,50,100]" :total="总数" layout="total, sizes, prev, pager, next" @current-change="查询()" @size-change="查询(true)"/></div>
  </section>
</template>
<style scoped>.历史筛选{margin-bottom:16px}.历史筛选 .el-input{max-width:360px}.历史筛选 .el-select{width:160px}</style>
