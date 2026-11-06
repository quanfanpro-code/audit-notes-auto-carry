<script setup>
defineProps({ 名称: String })
const 路径 = {
  新建任务: 'M12 5v14M5 12h14', 历史任务: 'M4 5v5h5M4 10a8 8 0 1 1 1 8M12 7v5l3 2',
  识别记录: 'M5 3h14v18H5zM8 7h8M8 11h8M8 15h5', 设置: 'M4 7h16M4 17h16M8 4v6M16 14v6',
  人工回写: 'M4 5h12v14H4zM10 12h10m-4-4 4 4-4 4', 收起: 'm14 6-6 6 6 6',
  文件: 'M5 3h9l5 5v13H5zM14 3v6h5M8 13h8M8 17h5', 方向: 'M4 12h16m-6-6 6 6-6 6'
}
</script>
<template><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path :d="路径[名称] || 路径.文件"/></svg></template>
