<script setup>
import { ElMessage } from 'element-plus'
defineProps({ 值: Object })
async function 复制(值) {
  if (!值?.是否取得) return
  try { await navigator.clipboard.writeText(值.原样文本); ElMessage.success('已原样复制') }
  catch { ElMessage.error('未能复制，请选中文字复制') }
}
</script>
<template><button v-if="值?.是否取得" class="值按钮" :class="{ 金额: 值.类型 === '数字' }" :title="'原样复制：' + (值.类型 === '空白' ? '空白' : 值.原样文本)" @click.stop="复制(值)">{{ 值.类型 === '空白' ? '空白' : 值.原样文本 }}</button><span v-else class="辅助" :title="值?.类型">{{ 值?.类型 === '读取失败' ? '读取失败' : '未取得' }}</span></template>
<style scoped>.值按钮{border:0;background:none;color:inherit;padding:0;max-width:100%;display:inline-block;white-space:nowrap;text-overflow:ellipsis;overflow:hidden;vertical-align:middle;user-select:text}.值按钮.金额{display:block;width:100%;text-align:right}.值按钮:hover{color:var(--主色);text-decoration:underline}</style>
