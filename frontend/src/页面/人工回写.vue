<script setup>
import { reactive, ref, computed } from 'vue'
import { 请求, 请求标识 } from '../请求.js'
import { 状态 } from '../任务状态.js'
const 输入 = reactive({ 带连接word: '', excel: '', 输出: '', 结构操作: '' })
const 错误 = ref(''), 忙 = ref(false)
let 上次
const 字段 = [{ 键:'带连接word',名:'带连接的 Word',类型:'Word',说明:'使用原任务提取后生成的带连接 Word。' }, { 键:'excel',名:'已经修改并保存的 Excel',类型:'Excel',说明:'先在 Excel 中保存修改，再选择文件。' }, { 键:'输出',名:'新 Word 副本的位置',类型:'另存为Word',说明:'另存结果，不能覆盖输入原件。' }]
const 可用 = computed(() => !状态.活动任务 && !忙.value && 输入.带连接word && 输入.excel && 输入.输出)
async function 选择(键,类型) { try { const r = await 请求('/api/选文件',{类型}); if(r.路径) 输入[键]=r.路径 } catch(e){错误.value=e.message} }
async function 开始() {
  if(!可用.value)return
  忙.value=true; 错误.value=''
  const 文=JSON.stringify(输入)
  if(!上次 || 上次.文!==文)上次={文,id:请求标识()}
  try { const r=await 请求('/api/桌面/任务',{类型:'人工回写',请求标识:上次.id,原始输入:{...输入}}); 上次=null; 状态.活动任务=r.task_id; 状态.查看(r.task_id); 状态.刷新号++ }
  catch(e){错误.value=e.message}finally{忙.value=false}
}
</script>
<template><div class="页标题"><div><h1>人工回写</h1><p class="辅助">把已经保存的 Excel 修改写回关联 Word 的新副本。此工具不重新识别或配对。</p></div></div><el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon/><el-alert v-if="状态.活动任务" title="有任务正在处理，结束后才能进行人工回写。" type="info" :closable="false" show-icon/>
  <section class="内容卡 人工表单"><div v-for="x in 字段" :key="x.键" class="输入行"><div><h2>{{ x.名 }}</h2><p class="辅助">{{ x.说明 }}</p><p class="文件路径">{{ 输入[x.键] || '尚未选择' }}</p></div><el-button @click="选择(x.键,x.类型)">{{ 输入[x.键] ? '重新选择' : '选择文件' }}</el-button></div><el-collapse><el-collapse-item title="可选：已有结构操作文件" name="结构"><p class="辅助">只使用已经明确给定的结构操作；留空就不插行、不删行。</p><div class="操作行"><span class="文件路径">{{ 输入.结构操作 || '未选择' }}</span><el-button @click="选择('结构操作','结构操作')">选择文件</el-button><el-button v-if="输入.结构操作" @click="输入.结构操作=''">不使用</el-button></div></el-collapse-item></el-collapse><div class="操作行 开始行"><el-button type="primary" :disabled="!可用" :loading="忙" @click="开始">另存回写副本</el-button><span class="辅助">写回开始后请等待安全结束，不提供即时中断或续跑。</span></div></section>
</template>
<style scoped>.人工表单{max-width:1000px}.输入行{display:flex;align-items:center;justify-content:space-between;gap:16px;border-bottom:1px solid var(--边线);padding:16px 0}.输入行:first-child{padding-top:0}.输入行 h2{margin:0}.输入行>div{min-width:0}.开始行{margin-top:20px}</style>
