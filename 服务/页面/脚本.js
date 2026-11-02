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

function 任务输入() {
  const 输入 = { word: $("word").value, b表: $("b表").value, 金标准: $("金标准").value };
  if ($("b工作表").value.trim()) 输入.b工作表 = $("b工作表").value.trim();
  return 输入;
}

function 开始轮询() {
  clearInterval(轮询);
  轮询 = setInterval(async () => {
    const p = await 请求("/api/进度");
    const 待 = (p.待答复 || []).length
      ? "\n卡在材料：" + p.待答复.join("；") + "\n（把答复.txt 放进该目录后点“继续上次任务”）" : "";
    $("进度").textContent = "步骤：" + (p.步骤 || "…") + " ｜ 已完成范围：" + p.已完成范围 + 待 + "\n" + (p.日志 || []).join("\n");
    if (p.步骤 === "done") {
      clearInterval(轮询);
      $("进度").textContent += "\n完成。到“结果”页查看。";
    }
    if (p.步骤 === "回读不一致") {
      clearInterval(轮询);
      $("进度").textContent += "\n回读有差异，不算完成。到“结果”页看差异。";
    }
  }, 2000);
}

async function 发任务(续跑) {
  const 输入 = 任务输入();
  if (!输入.word || !输入.b表 || !输入.金标准) { $("进度").textContent = "三个文件都要先选好。"; return; }
  if (续跑) 输入.续跑 = true;
  $("进度").textContent = 续跑 ? "接着上次跑…" : "启动中…";
  const 结果 = await 请求("/api/任务", 输入);
  if (结果.错误) { $("进度").textContent = 结果.错误; return; }
  开始轮询();
}

$("开始").addEventListener("click", () => 发任务(false));
$("续跑").addEventListener("click", () => 发任务(true));
$("中断").addEventListener("click", async () => { await 请求("/api/中断", {}); });

// 结果
let 结果缓存 = null;

function 渲染结果() {
  if (!结果缓存 || 结果缓存.错误) { $("结果内容").textContent = (结果缓存 || {}).错误 || ""; return; }
  const 摘 = 结果缓存.摘要 || {};
  const 产 = 结果缓存.产物 || {};
  $("结果摘要").textContent = [
    "步骤：" + (摘.步骤 || "—"),
    "对应：" + (摘.对应条数 || 0) + " 条 ｜ 清尾：" + (摘.清尾条数 || 0) + " 条 ｜ 回读差异：" + (摘.回读差异 || 0),
    "表格：" + (产.更新后附注表格 || "—"),
    "Word：" + (产.Word新副本 || "—"),
  ].join("\n");

  const 选 = $("原因筛选");
  const 旧值 = 选.value;
  选.innerHTML = '<option value="">全部未对应</option>';
  Object.keys(摘.未对应 || {}).forEach((原因) => {
    const o = document.createElement("option");
    o.value = 原因;
    o.textContent = 原因 + "（" + 摘.未对应[原因] + "）";
    选.appendChild(o);
  });
  选.value = 旧值;

  const 要显示 = 选.value
    ? 结果缓存.对应.filter((x) => x.状态 !== "对应" && (x.原因 || "未分类") === 选.value)
    : 结果缓存.对应;
  $("结果内容").textContent = JSON.stringify(
    { 对应: 要显示, 问题: 结果缓存.问题, 回读: 结果缓存.回读 }, null, 2);
}

$("刷新结果").addEventListener("click", async () => {
  结果缓存 = await 请求("/api/结果");
  渲染结果();
});
$("原因筛选").addEventListener("change", 渲染结果);
$("打开产物").addEventListener("click", async () => {
  const 目录 = ((结果缓存 || {}).产物 || {}).运行目录;
  const r = await 请求("/api/打开目录", { 路径: 目录 });
  if (!r.ok) $("结果摘要").textContent = "打不开目录：" + (r.错误 || "");
});

// 设置
async function 载入设置() {
  const c = await 请求("/api/设置");
  $("url").value = c.url || "";
  $("model").value = c.model || "";
  $("timeout").value = c.timeout || 600;
}
$("保存设置").addEventListener("click", async () => {
  const r = await 请求("/api/设置", { url: $("url").value, api_key: $("api_key").value, model: $("model").value, timeout: Number($("timeout").value) });
  $("设置提示").textContent = r.ok ? "已保存" : "失败";
});
$("测试连通").addEventListener("click", async () => {
  const r = await 请求("/api/设置/测试", {});
  $("设置提示").textContent = r.ok ? ("连通，状态码 " + r.状态码) : ("不通：" + (r.说明 || r.错误 || r.状态码));
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
