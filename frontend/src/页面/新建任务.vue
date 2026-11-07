<script setup>
import { reactive, ref, computed, watch, onMounted, onBeforeUnmount } from 'vue'
import { 请求, 请求标识, 最新请求 } from '../请求.js'
import { 状态 } from '../任务状态.js'
import UiIcon from '../图标.vue'
状态.新建草稿 ||= {
  文件: { A: { 路径: '', 工作表: [] }, B: { 路径: '', 工作表: [] } }, 表们: { A: [], B: [] }, 搜索: { A: '', B: '' },
  设置: { 金标准: 状态.概况.默认设置?.金标准 || '', 复用: 状态.概况.默认设置?.复用 !== false, 记录目录: 状态.概况.默认设置?.识别记录目录 || '' },
  已初始化: !!状态.概况.默认设置, 上次提交: null
}
const 草稿 = 状态.新建草稿
const { 文件, 表们, 搜索, 设置 } = 草稿
const 加载 = reactive({ A: false, B: false })
const 错误 = ref('')
const 提交中 = ref(false)
const 管 = { A: 最新请求(), B: 最新请求() }
let 已销毁 = false
watch(() => 状态.概况.默认设置, 值 => { if (值 && !草稿.已初始化) { Object.assign(设置, { 金标准: 值.金标准, 复用: 值.复用, 记录目录: 值.识别记录目录 }); 草稿.已初始化 = true } })
const 可开始 = computed(() => !状态.活动任务 && !提交中.value && !加载.A && !加载.B && 设置.金标准 && ['A', 'B'].every(侧 => 文件[侧].路径 && (!文件[侧].路径.toLowerCase().endsWith('.xlsx') || 文件[侧].工作表.length)))
async function 选择(侧) {
  错误.value = ''; 加载[侧] = true
  try {
    const r = await 请求('/api/选文件', { 类型: '报表' })
    if (!r.路径 || 已销毁) return
    if (!/\.(docx|xlsx)$/i.test(r.路径)) throw new Error('请选择 Word（docx）或 Excel（xlsx）文件')
    await 读取文件(侧, r.路径)
  } catch (e) { 错误.value = e.message }
  finally { 加载[侧] = false }
}
async function 读取文件(侧, 路径, 原选择 = []) {
  文件[侧] = { 路径, 工作表: [] }; 表们[侧] = []; 搜索[侧] = ''
  管[侧].取消()
  if (!路径.toLowerCase().endsWith('.xlsx')) return
  const 表 = await 管[侧].执行(signal => 请求('/api/工作表', { 路径 }, { signal }))
  if (!表) return
  表们[侧] = 表.工作表
  const 可见 = 表.工作表.filter(x => x.可见).map(x => x.名称)
  文件[侧].工作表 = 原选择.length ? 原选择.filter(x => 可见.includes(x)) : 可见.length === 1 ? 可见 : []
}
onMounted(async () => {
  const 原 = 状态.新建预填; 状态.新建预填 = null
  for (const 侧 of ['A', 'B']) {
    if (已销毁) return
    const 恢复 = 原?.[侧] || 文件[侧]
    if (!恢复.路径 || (!原 && 表们[侧].length)) continue
    const 原搜索 = 搜索[侧]
    加载[侧] = true
    try { await 读取文件(侧, 恢复.路径, 恢复.工作表 || []); if (!原 && !已销毁) 搜索[侧] = 原搜索 }
    catch (e) { 错误.value = e.message }
    finally { 加载[侧] = false }
  }
})
async function 选设置(键) {
  try { const r = await 请求('/api/选文件', { 类型: 键 === '记录目录' ? '文件夹' : '金标准' }); if (r.路径) 设置[键] = r.路径 }
  catch (e) { 错误.value = e.message }
}
async function 开始() {
  if (!可开始.value) return
  提交中.value = true; 错误.value = ''
  const 输入 = { A: { ...文件.A }, B: { ...文件.B }, 金标准: 设置.金标准, 复用: 设置.复用, 识别记录目录: 设置.记录目录 || 设置.识别记录目录 || '' }
  const 内容 = JSON.stringify(输入)
  if (!草稿.上次提交 || 草稿.上次提交.内容 !== 内容) 草稿.上次提交 = { 内容, id: 请求标识() }
  try {
    const r = await 请求('/api/桌面/任务', { 请求标识: 草稿.上次提交.id, 原始输入: 输入 })
    if (状态.新建草稿 === 草稿) 状态.新建草稿 = null
    状态.活动任务 = r.task_id; 状态.查看(r.task_id); 状态.刷新号++
  } catch (e) { 错误.value = e.message }
  finally { 提交中.value = false }
}
onBeforeUnmount(() => { 已销毁 = true; 管.A.取消(); 管.B.取消() })
const 文件名 = 路径 => 路径.split(/[\\/]/).pop()
</script>
<template>
  <div class="页标题"><div><h1>开始一份搬运任务</h1><p class="辅助">选择提供数据的 B 端，以及需要填写的 A 端。结果另存为 A 端的新副本。</p></div></div>
  <el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon/>
  <el-alert v-if="状态.活动任务" title="有任务正在处理，结束后可开始下一份。文件可以先选好。" type="info" :closable="false" show-icon/>
  <div class="文件组合">
    <template v-for="侧 in ['B', 'A']" :key="侧">
      <div v-if="侧 === 'A'" class="搬运方向"><UiIcon 名称="方向"/><span>B → A</span></div>
      <section class="内容卡 文件卡">
        <div class="详情标题"><h2><span class="端标">{{ 侧 }}</span>{{ 侧 === 'B' ? '提供数据的文件' : '需要填写的文件' }}</h2><span class="辅助">Word / Excel</span></div>
        <div class="选择区"><UiIcon 名称="文件"/><strong v-if="文件[侧].路径">{{ 文件名(文件[侧].路径) }}</strong><span v-else class="辅助">{{ 侧 === 'B' ? '从这份文件取数据' : '将数据填入这份文件的新副本' }}</span><el-button :data-test="'选择'+侧" :loading="加载[侧]" @click="选择(侧)">{{ 文件[侧].路径 ? '更换文件' : '选择文件' }}</el-button></div>
        <p v-if="文件[侧].路径" class="文件路径">{{ 文件[侧].路径 }}</p>
        <template v-if="文件[侧].路径.toLowerCase().endsWith('.xlsx')">
          <div class="工作表标题"><strong>本次处理的工作表</strong><span class="辅助">已选 {{ 文件[侧].工作表.length }} 张</span></div>
          <div class="操作行"><el-input v-model="搜索[侧]" :aria-label="侧+'端搜索工作表'" placeholder="搜索工作表" clearable/><el-button @click="文件[侧].工作表 = 表们[侧].filter(x => x.可见).map(x => x.名称)">全选可见</el-button><el-button @click="文件[侧].工作表 = []">清空选择</el-button></div>
          <el-skeleton v-if="加载[侧]" :rows="3"/>
          <el-checkbox-group v-else v-model="文件[侧].工作表" class="工作表列表" :aria-label="侧+'端工作表'">
            <el-checkbox v-for="表 in 表们[侧].filter(x => x.名称.includes(搜索[侧]))" :key="表.名称" :value="表.名称" :disabled="!表.可见">{{ 表.名称 }}<span v-if="!表.可见" class="辅助">（隐藏，不可选）</span></el-checkbox>
          </el-checkbox-group>
          <p class="辅助">一次处理同一套报表；未选的工作表保持原样。</p>
        </template>
      </section>
    </template>
  </div>
  <section class="内容卡"><el-collapse><el-collapse-item title="本次设置" name="本次设置"><el-form label-position="top"><el-form-item label="金标准"><div class="操作行 整行"><span class="文件路径">{{ 设置.金标准 || '尚未选择金标准' }}</span><el-button @click="选设置('金标准')">选择文件</el-button></div></el-form-item><el-form-item label="已识别内容自动复用"><el-switch v-model="设置.复用" :disabled="状态.概况.复用锁定" aria-label="自动复用"/><span class="辅助">{{ 状态.概况.复用说明 || '先核对已有记录，符合原规则时自动复用' }}</span></el-form-item><el-form-item label="识别记录位置"><div class="操作行"><span class="文件路径">{{ 设置.记录目录 || '项目下的识别记录' }}</span><el-button :disabled="状态.概况.复用锁定" @click="选设置('记录目录')">选择文件夹</el-button></div></el-form-item></el-form></el-collapse-item></el-collapse></section>
  <div class="新建底栏"><span class="辅助">原文件保留。公式、标题及整组保护沿用程序现有规则。</span><el-button type="primary" :disabled="!可开始" :loading="提交中" data-test="开始任务" @click="开始">开始处理</el-button></div>
</template>
<style scoped>
.文件组合{display:grid;grid-template-columns:minmax(0,1fr) 56px minmax(0,1fr);align-items:start}.文件卡{min-height:240px}.端标{display:inline-grid;place-items:center;width:26px;height:26px;border-radius:6px;background:#eaf1ff;color:#204d94;margin-right:8px}.搬运方向{display:flex;flex-direction:column;align-items:center;gap:8px;padding-top:110px;color:var(--次文);font-size:12px}.选择区{border:1px dashed #b9c7d8;border-radius:6px;background:#fafbfd;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;min-height:150px;padding:16px;overflow-wrap:anywhere}.选择区 svg{color:#59779e;width:28px;height:28px}.工作表标题{display:flex;justify-content:space-between;margin:20px 0 10px;font-size:13px}.工作表列表{display:flex;flex-direction:column;max-height:260px;overflow:auto;padding:8px 0}.工作表列表 .el-checkbox{margin-right:0;min-height:32px}.操作行 .el-input{flex:1;min-width:120px}.新建底栏{display:flex;justify-content:space-between;align-items:center;gap:16px}@media(max-width:1100px){.文件组合{grid-template-columns:1fr 32px 1fr}.文件卡{padding:12px}.操作行{gap:6px}.操作行 .el-input{flex-basis:100%}.文件卡 h2{font-size:14px}.文件卡 .详情标题>.辅助{display:none}}
</style>
