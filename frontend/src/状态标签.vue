<script setup>
import { computed } from 'vue'
import { 状态文字 } from './任务状态.js'
const p = defineProps({ 值: String })
const 类 = computed(() => /失败|错误|不一致/.test(p.值) ? '错误' : /保留|待核|等待|中断|未记录/.test(p.值) ? '警示' : /done|写入|写回|一致|未变/.test(p.值) ? '成功' : '')
</script>
<template><span class="状态标签" :class="类"><span aria-hidden="true">{{ 类 === '成功' ? '✓' : 类 === '错误' ? '!' : 类 === '警示' ? '·' : '—' }}</span>{{ 状态文字(值) }}</span></template>
