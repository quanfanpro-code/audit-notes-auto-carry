<script setup>
import { reactive, ref, watch, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import { 请求 } from '../请求.js'
import { 状态 } from '../任务状态.js'
const 模型 = reactive({ url: '', model: '', api_key: '', timeout: 300 })
const 默认 = reactive({ 金标准: '', 复用: true, 识别记录目录: '', 字号: 14 })
const 已有密钥 = ref(false), 错误 = ref(''), 测试结果 = ref(''), 保存中 = ref(false), 测试中 = ref(false), 加载 = ref(true), 锁定 = ref(false), 复用说明 = ref('')
let 原文 = ''
const 内容 = () => JSON.stringify({ 模型, 默认 })
async function 读取() {
  try { const r = await 请求('/api/桌面/设置'); Object.assign(模型, r.模型, { api_key: '' }); Object.assign(默认, r.默认设置); 已有密钥.value = !!r.模型.密钥已保存; 锁定.value = r.复用锁定; 复用说明.value = r.复用说明; 原文 = 内容() }
  catch (e) { 错误.value = e.message }
  finally { 加载.value = false }
}
watch([模型, 默认], () => { if (!加载.value) 状态.未保存.设置 = 内容() !== 原文 }, { deep: true })
async function 选择(键) { try { const r = await 请求('/api/选文件', { 类型: 键 === '金标准' ? '金标准' : '文件夹' }); if(r.路径) 默认[键] = r.路径 } catch(e) { 错误.value = e.message } }
async function 保存() {
  保存中.value = true; 错误.value = ''
  try {
    await 请求('/api/桌面/设置', { 模型: { url: 模型.url, model: 模型.model, api_key: 模型.api_key, timeout: 模型.timeout }, 默认设置: { ...默认 } })
    if (模型.api_key) 已有密钥.value = true
    模型.api_key = ''; 原文 = 内容(); 状态.未保存.设置 = false
    document.documentElement.style.setProperty('--字号', 默认.字号 + 'px'); 状态.刷新号++; ElMessage.success('设置已保存')
  } catch(e) { 错误.value = e.message }
  finally { 保存中.value = false }
}
async function 测试() {
  测试中.value = true; 测试结果.value = ''; 错误.value = ''
  try { const r = await 请求('/api/设置/测试', {}); 测试结果.value = r.说明 + '；仅检查连接与访问权限，没有执行识别。' }
  catch(e) { 错误.value = e.message }
  finally { 测试中.value = false }
}
onMounted(读取)
</script>
<template><div class="页标题"><div><h1>设置</h1><p class="辅助">新设置用于之后的任务；正在运行的任务保留启动时的连接与文件设置。</p></div><el-button type="primary" :loading="保存中" :disabled="加载" @click="保存">保存设置</el-button></div><el-alert v-if="错误" :title="错误" type="error" :closable="false" show-icon/><el-alert v-if="测试结果" :title="测试结果" type="success" :closable="false" show-icon/>
  <div class="设置分区"><section class="内容卡"><h2>模型连接</h2><el-form label-position="top" :disabled="加载"><el-form-item label="API 地址"><el-input v-model="模型.url" placeholder="兼容接口地址，以 /v1 等实际地址为准"/></el-form-item><el-form-item label="模型名称" data-test="模型名称"><el-input v-model="模型.model" placeholder="填写服务提供方给出的模型名称"/></el-form-item><el-form-item label="API 密钥"><el-input v-model="模型.api_key" type="password" autocomplete="new-password" :placeholder="已有密钥 ? '已保存密钥；留空则保持原密钥' : '尚未保存密钥'" show-password/><span class="辅助">{{ 已有密钥 ? '已保存密钥，不回显。' : '没有密钥时也可通过人工答复完成识别。' }} 密钥在本机加密保存。</span></el-form-item><el-form-item label="等待答复时间（秒）"><el-input-number v-model="模型.timeout" :min="30" :max="3600" :step="30"/></el-form-item></el-form><div class="操作行"><el-button :loading="测试中" @click="测试">测试已保存连接</el-button><span class="辅助">未保存的修改不参与本次测试</span></div><p class="辅助">模型会收到任务的识别材料；请使用你认可的模型服务。</p></section>
    <section class="内容卡"><h2>常用设置</h2><el-form label-position="top" :disabled="加载"><el-form-item label="金标准"><p class="文件路径">{{ 默认.金标准 || '尚未选择' }}</p><el-button @click="选择('金标准')">选择金标准文件</el-button></el-form-item><el-form-item label="自动复用已有识别"><el-switch v-model="默认.复用" :disabled="锁定" aria-label="默认自动复用"/><span class="辅助">{{ 复用说明 }}</span></el-form-item><el-form-item label="识别记录位置"><p class="文件路径">{{ 默认.识别记录目录 || '项目下的识别记录' }}</p><el-button @click="选择('识别记录目录')">选择文件夹</el-button></el-form-item><el-form-item label="界面字号"><el-radio-group v-model="默认.字号"><el-radio-button v-for="x in [14,15,16,18]" :key="x" :value="x">{{ x }} px</el-radio-button></el-radio-group></el-form-item></el-form><p class="辅助">更换金标准后，旧解释是否适用仍按原程序的版本规则核验。</p></section></div>
</template>
<style scoped>.设置分区{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:16px;max-width:1100px}.el-form-item .文件路径{width:100%;margin:0 0 8px}.el-form-item .辅助{margin-top:6px}.el-switch+.辅助{margin-left:8px}@media(max-width:1100px){.设置分区{grid-template-columns:1fr}}</style>
