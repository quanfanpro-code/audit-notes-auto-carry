// 审计附注自动搬运系统 · 前端脚本（无构建链，原生JS）
const $ = (id) => document.getElementById(id);

// 页签切换
document.querySelectorAll(".页签").forEach((b) => b.addEventListener("click", () => {
  document.querySelectorAll(".页签").forEach((x) => x.classList.remove("当前"));
  b.classList.add("当前");
  document.querySelectorAll("main section").forEach((s) => (s.hidden = s.id !== b.dataset.页));
}));

async function 请求(url, 体) {
  const r = await fetch(url, 体 ? { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(体) } : {});
  return r.json();
}

// 选文件（后端弹 Windows 原生窗口，记真实路径）
document.querySelectorAll("[data-选]").forEach((b) => b.addEventListener("click", async () => {
  const 类型 = b.dataset.选 === "文件夹" ? "文件夹" : "文件";
  const 结果 = await 请求("/api/选文件", { 类型 });
  if (结果.路径) $(b.dataset.选).value = 结果.路径;
}));

// 任务
let 轮询 = null;
$("开始").addEventListener("click", async () => {
  const 输入 = { word: $("word").value, b表: $("b表").value, 金标准: $("金标准").value };
  if ($("b工作表").value.trim()) 输入.b工作表 = $("b工作表").value.trim();
  if (!输入.word || !输入.b表 || !输入.金标准) { $("进度").textContent = "三个文件都要先选好。"; return; }
  $("进度").textContent = "启动中…";
  const 结果 = await 请求("/api/任务", 输入);
  if (结果.错误) { $("进度").textContent = 结果.错误; return; }
  clearInterval(轮询);
  轮询 = setInterval(async () => {
    const p = await 请求("/api/进度");
    const 待 = (p.待答复 || []).length ? "\n卡在材料：" + p.待答复.join("；") + "（把答复.txt 放进该目录后重新开始任务即可续跑）" : "";
    $("进度").textContent = "步骤：" + (p.步骤 || "…") + " ｜ 已完成范围：" + p.已完成范围 + 待 + "\n" + (p.日志 || []).join("\n");
    if (p.步骤 === "done") { clearInterval(轮询); $("进度").textContent += "\n完成。到“结果”页查看。"; }
  }, 2000);
});
$("中断").addEventListener("click", async () => { await 请求("/api/中断", {}); });

// 结果
$("刷新结果").addEventListener("click", async () => {
  const r = await 请求("/api/结果");
  $("结果内容").textContent = r.错误 || JSON.stringify(r, null, 2);
});

// 设置
async function 载入设置() {
  const c = await 请求("/api/设置");
  $("url").value = c.url || ""; $("model").value = c.model || ""; $("timeout").value = c.timeout || 600;
}
$("保存设置").addEventListener("click", async () => {
  const r = await 请求("/api/设置", { url: $("url").value, api_key: $("api_key").value, model: $("model").value, timeout: Number($("timeout").value) });
  $("设置提示").textContent = r.ok ? "已保存" : "失败";
});
$("测试连通").addEventListener("click", async () => {
  const r = await 请求("/api/设置/测试", {});
  $("设置提示").textContent = r.ok ? ("连通，状态码 " + r.状态码) : ("不通：" + (r.错误 || r.状态码));
});

// 人工回写
$("执行回写").addEventListener("click", async () => {
  const 输入 = {
    带连接word: $("带连接word").value, excel: $("回写excel").value,
    输出: $("回写excel").value.replace(/\.xlsx?$/i, "") + "_Word新副本.docx",
    基准表A: $("基准表A").value || undefined,
  };
  if (!输入.带连接word || !输入.excel) { $("回写结果").textContent = "带连接Word和更新后Excel都要先选好。"; return; }
  $("回写结果").textContent = "回写中…";
  const r = await 请求("/api/人工回写", 输入);
  $("回写结果").textContent = r.ok ? JSON.stringify(r.结果, null, 2) : ("失败：" + r.错误);
});

载入设置();
