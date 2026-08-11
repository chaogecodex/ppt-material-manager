// 前端回归测试：用 Playwright 直接打开 index.html，拦截后端接口做端到端验证。
// 运行： cd tests/frontend && npm install && npm test
//
// 不依赖真实后端。每个用例自带 mock，断言的是用户实际看到的行为，
// 而不是内部实现，所以重构 index.html 时这些用例仍然有效。

import { chromium } from "playwright";
import { fileURLToPath } from "node:url";
import path from "node:path";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PAGE = "file://" + path.resolve(HERE, "../../index.html");
const API = "**ppt-ai-api-294833-10-1466238220.sh.run.tcloudbase.com";

const pass = [];
const fail = [];
const ok = (name, cond, detail = "") =>
  (cond ? pass : fail).push(name + (detail ? ` [${detail}]` : ""));

const json = (route, body, status = 200) =>
  route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) });

const session = () => ({
  token: "t",
  username: "测试用户",
  expiresAt: new Date(Date.now() + 30 * 864e5).toISOString(),
});

const project = (materials = [], aiResult = null) => ({
  id: "p1",
  title: "t",
  brief: { purpose: "", audience: "", limit: "", instruction: "" },
  materials,
  analysisStatus: aiResult ? "done" : "idle",
  aiResult,
  locked: false,
});

const MATERIALS = [
  { id: "m1", name: "a.pdf", type: "PDF", size: "1 MB", weight: "核心", customerOrder: 1 },
  { id: "m2", name: "b.pdf", type: "PDF", size: "2 MB", weight: "重要", customerOrder: 2 },
];

const AI_RESULT = {
  outline: [{ title: "旧章节XYZ", summary: "旧摘要" }],
  pageEstimate: "12页",
  timeEstimate: "15分钟",
  usedCount: 2,
  conflicts: [],
  orderSuggestions: ["建议把视频放在开场", "<img src=x onerror=alert(1)>"],
};

const seedSession = (page) =>
  page.addInitScript((s) => localStorage.setItem("pptAccessSession", JSON.stringify(s)), session());

const seedProject = (page) =>
  page.addInitScript(() =>
    localStorage.setItem("pptAgentMvp", JSON.stringify({
      schemaVersion: 2, projectId: "p1", projectToken: "ptok", materials: [], currentStep: 1,
    })));

// ---------------------------------------------------------------------- 等待
// 一律等"条件成立"，不等固定时长。固定 sleep 有两个毛病：跑得慢，
// 而且 runner 一负载就可能没睡够 → 偶发红灯。
//
// until: 轮询页面内的判定函数，条件一成立立刻返回。
// 超时留 10s 是给"真的坏了"用的——正常路径根本碰不到。
const until = (page, fn, arg) => page.waitForFunction(fn, arg, { timeout: 10000 });

// settle: 等应用把手头的请求跑完。所有接口都被 route 拦了，
// 所以 networkidle 在这里能精确表示"启动/提交流程结束"。
const settle = (page) => page.waitForLoadState("networkidle");

const visible = (page, sel) => page.waitForSelector(sel, { state: "visible", timeout: 10000 });

// 注意：page.click / textContent / innerHTML 自带 auto-wait，
// 会等元素出现且可操作，所以纯粹为"等元素渲染出来"而写的 sleep 全部删掉了。

// CI 里用 `npx playwright install chromium` 装的默认浏览器；
// 本地若已有预装 Chromium，用 CHROMIUM_PATH 指过去即可，不必重复下载。
const browser = await chromium.launch(
  process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {}
);

// ---------------------------------------------------------------- A1 首屏滚动
{
  const page = await browser.newPage();
  await page.route(API + "/**", (r) => r.abort());
  await page.goto(PAGE);
  // 这条断言的是"什么都别发生"，没有可等的正向条件，
  // 只能等启动流程走完（门禁露出 = 启动的终态）再看有没有被滚走。
  await settle(page);
  await visible(page, "#accessGate");
  const y = await page.evaluate(() => Math.round(window.scrollY));
  ok("A1 首屏不自动滚动", y === 0, `scrollY=${y}`);
  await page.close();
}

// ------------------------------------------------- A2 网络错误提示 + 重试按钮
{
  const page = await browser.newPage();
  await seedSession(page);
  await page.route(API + "/**", (r) => r.abort());
  await page.goto(PAGE);
  await settle(page);
  await page.click("#loadDemoBtn");
  // 直接等断言要看的那句提示出现
  await until(page, () => /网络连接失败/.test(document.querySelector("#busyNoteText")?.textContent || ""));
  const text = (await page.textContent("#busyNoteText")) || "";
  ok("A2a 不暴露 Failed to fetch", !/Failed to fetch/i.test(text), JSON.stringify(text));
  ok("A2b 显示中文网络提示", /网络连接失败/.test(text) && (await page.isVisible("#busyNote")));
  ok("A2c 提供重试按钮", await page.isVisible("#busyRetryBtn"));
  await page.close();
}

// -------------------------------------------------------- A3 旧版本地状态迁移
{
  const page = await browser.newPage();
  await page.addInitScript(() =>
    localStorage.setItem("pptAgentMvp", JSON.stringify({
      demoMode: true,
      materials: [{ id: "old", name: "旧素材.pdf", type: "PDF", size: "1 MB" }],
      purpose: "旧的项目要求", audience: "旧受众", currentStep: 3,
    })));
  await page.route(API + "/**", (r) => r.abort());
  await page.goto(PAGE);
  // 迁移写回 schemaVersion=2 就是迁移完成的精确信号
  await until(page, () => JSON.parse(localStorage.getItem("pptAgentMvp") || "{}").schemaVersion === 2);
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem("pptAgentMvp")));
  ok("A3a 保留项目要求", (await page.inputValue("#purposeInput")) === "旧的项目要求");
  ok("A3b 清除无法补传的幽灵素材", (await page.textContent("#materialCount")) === "0");
  ok("A3c 写入 schemaVersion=2", saved?.schemaVersion === 2, String(saved?.schemaVersion));
  ok("A3d 回到第 1 步", saved?.currentStep === 1, String(saved?.currentStep));
  ok("A3e 只提示一次", /旧版本地演示数据/.test((await page.textContent("#toast")) || ""));
  await page.close();
}

// ------------------------------------ B1 转义 + C2 建议区（含恶意串不可执行）
{
  const page = await browser.newPage();
  let alerted = false;
  page.on("dialog", (d) => { alerted = true; d.dismiss(); });
  await seedSession(page);
  const hostile = [
    { id: "m1", name: "正常.pdf", type: "<img src=x onerror=alert(1)>", size: "<b>1 MB</b>", weight: "核心", customerOrder: 1 },
    { id: "m2", name: "数据.csv", type: "TEXT", size: "2 KB", weight: "重要", customerOrder: 2 },
  ];
  await page.route(API + "/**", (route) => {
    const p = route.request().url().split("tcloudbase.com")[1].split("?")[0];
    if (p === "/auth/session") return json(route, { success: true, user: { username: "测试用户" } });
    if (p === "/projects") return json(route, { success: true, project: project(), accessToken: "ptok" });
    if (/\/analyze$/.test(p)) return json(route, { success: true, project: project(hostile, AI_RESULT), result: AI_RESULT });
    if (/\/files$/.test(p)) return json(route, { success: true, project: project(hostile) });
    return json(route, { success: true, project: project(hostile) });
  });
  await page.goto(PAGE);
  await settle(page);
  await page.click("#loadDemoBtn");
  await until(page, () => document.querySelectorAll("#materialList .file-icon").length === 2);

  const injected = await page.evaluate(() => document.querySelectorAll("#materialList img, #materialList b > b").length);
  ok("B1a 恶意 type/size 未注入 DOM 节点", injected === 0, `injected=${injected}`);
  ok("B1b 恶意串按纯文本渲染", ((await page.textContent(".file-icon")) || "").includes("<img"));
  ok("B1c 未触发 alert", !alerted);

  await page.click('[data-next="2"]');
  await page.click('[data-next="3"]');
  await until(page, () => document.querySelectorAll("#suggestionList li").length === 2);
  const html = await page.innerHTML("#suggestionList");
  ok("C2a 建议区显示", await page.isVisible("#suggestionBox"));
  ok("C2b 渲染两条建议", (await page.evaluate(() => document.querySelectorAll("#suggestionList li").length)) === 2);
  ok("C2c 建议内容已转义", !/<img/.test(html) && /&lt;img/.test(html));
  ok("C2d 建议区未触发 alert", !alerted);
  await page.close();
}

// ---------------------------------------------- 登录态恢复（已有项目时也要恢复）
{
  const page = await browser.newPage();
  await seedSession(page);
  await seedProject(page);
  await page.route(API + "/**", (route) => {
    const p = route.request().url().split("tcloudbase.com")[1].split("?")[0];
    if (p === "/auth/session") return json(route, { success: true, user: { username: "测试用户" } });
    return json(route, { success: true, project: project(MATERIALS) });
  });
  await page.goto(PAGE);
  await until(page, () => document.querySelector("#accessGate")?.hidden === true);
  ok("登录态 门禁隐藏", (await page.evaluate(() => document.querySelector("#accessGate").hidden)) === true);
  ok("登录态 退出按钮可见（accessToken 已恢复）", await page.isVisible("#logoutBtn"));
  await page.close();
}

// ------------------------------- 旧大纲清理：删除素材 / 分析失败 / 保存阶段失败
{
  const page = await browser.newPage();
  let analyzeFails = false;
  let saveFails = false;
  await seedSession(page);
  await seedProject(page);
  await page.exposeFunction("__setAnalyzeFail", (v) => { analyzeFails = v; });
  await page.exposeFunction("__setSaveFail", (v) => { saveFails = v; });
  await page.route(API + "/**", (route) => {
    const req = route.request();
    const p = req.url().split("tcloudbase.com")[1].split("?")[0];
    const m = req.method();
    if (p === "/auth/session") return json(route, { success: true, user: { username: "测试用户" } });
    if (/\/files\/[^/]+$/.test(p) && m === "DELETE") return json(route, { success: true, project: project([MATERIALS[1]]) });
    if (/\/analyze$/.test(p)) return analyzeFails
      ? json(route, { detail: "AI服务暂时不可用" }, 503)
      : json(route, { success: true, project: project(MATERIALS, AI_RESULT), result: AI_RESULT });
    if (/^\/projects\/[^/]+$/.test(p) && m === "PATCH") return saveFails
      ? route.abort()
      : json(route, { success: true, project: project(MATERIALS, AI_RESULT) });
    return json(route, { success: true, project: project(MATERIALS, AI_RESULT) });
  });
  await page.goto(PAGE);
  await settle(page);

  await page.click('[data-next="2"]');
  await page.click('[data-next="3"]');
  const outlineRendered = () =>
    until(page, () => document.querySelector("#outline")?.innerHTML.includes("旧章节XYZ"));
  await outlineRendered();
  ok("前置 大纲已渲染", (await page.innerHTML("#outline")).includes("旧章节XYZ"));

  // 删除素材后旧大纲必须立即消失
  await page.click('.step[data-step="1"]');
  await page.click("[data-remove]");
  // 三条断言的终态一起等：大纲清空 + 建议区隐藏 + 概览重置
  await until(page, () =>
    !document.querySelector("#outline")?.innerHTML.includes("旧章节XYZ") &&
    document.querySelector("#suggestionBox")?.hidden === true &&
    document.querySelector("#pageEstimate")?.textContent === "—");
  ok("删除后 大纲已清空", !(await page.innerHTML("#outline")).includes("旧章节XYZ"));
  ok("删除后 建议区已隐藏", (await page.evaluate(() => document.querySelector("#suggestionBox").hidden)) === true);
  ok("删除后 概览已重置", (await page.textContent("#pageEstimate")) === "—");

  // 重新生成一版大纲作为"旧结果"。
  // 注意：改权重是必要的——不改的话 goStep(3) 会直接复用缓存的 aiResult，
  // 根本不会发起请求，也就测不到失败路径。
  // 三种走法（成功 / analyze 失败 / 保存失败）终态各不相同，
  // 所以这里只等"请求跑完"，具体断言由各调用点自己等。
  const dirtyThenGoToStep3 = async () => {
    await page.click('.step[data-step="2"]');
    await page.click('[data-weight="参考"]');
    await page.click('[data-next="3"]');
    await settle(page);
  };

  await dirtyThenGoToStep3();
  await outlineRendered();
  ok("前置 重新分析后大纲再次渲染", (await page.innerHTML("#outline")).includes("旧章节XYZ"));

  // 分析请求失败后不得残留旧结果
  await page.evaluate(() => window.__setAnalyzeFail(true));
  await dirtyThenGoToStep3();
  await visible(page, "#busyRetryBtn");
  ok("分析失败后 无残留旧大纲", !(await page.innerHTML("#outline")).includes("旧章节XYZ"),
     (await page.innerHTML("#outline")).slice(0, 50));
  ok("分析失败后 提示可见", await page.isVisible("#busyNote"));
  ok("分析失败后 提供重试", await page.isVisible("#busyRetryBtn"));

  // 保存阶段（analyze 之前的 PATCH）失败，同样不得残留旧大纲。
  // 这是 resetOutlineView() 必须放在 saveProject() 之前的原因。
  await page.evaluate(() => window.__setAnalyzeFail(false));
  await dirtyThenGoToStep3();
  await outlineRendered();
  ok("前置 恢复后大纲再次渲染", (await page.innerHTML("#outline")).includes("旧章节XYZ"));
  await page.evaluate(() => window.__setSaveFail(true));
  await dirtyThenGoToStep3();
  await until(page, () => !document.querySelector("#outline")?.innerHTML.includes("旧章节XYZ"));
  ok("保存阶段失败后 无残留旧大纲", !(await page.innerHTML("#outline")).includes("旧章节XYZ"),
     (await page.innerHTML("#outline")).slice(0, 50));
  await page.close();
}

// ------------------------- 项目恢复失败：临时故障保留引用，明确失效才清除
{
  // 网络中断：必须保留 projectId/projectToken，否则用户永久失去云端项目入口
  const page = await browser.newPage();
  await seedSession(page);
  await seedProject(page);
  await page.route(API + "/**", (route) => {
    const p = route.request().url().split("tcloudbase.com")[1].split("?")[0];
    if (p === "/auth/session") return json(route, { success: true, user: { username: "测试用户" } });
    return route.abort();
  });
  await page.goto(PAGE);
  await visible(page, "#busyRetryBtn");
  const kept = await page.evaluate(() => JSON.parse(localStorage.getItem("pptAgentMvp")));
  ok("恢复失败 网络故障保留 projectId", kept?.projectId === "p1", String(kept?.projectId));
  ok("恢复失败 网络故障保留 projectToken", kept?.projectToken === "ptok");
  ok("恢复失败 提示可重试", await page.isVisible("#busyRetryBtn"));
  await page.close();
}
{
  // 服务端 5xx：同样属临时故障
  const page = await browser.newPage();
  await seedSession(page);
  await seedProject(page);
  await page.route(API + "/**", (route) => {
    const p = route.request().url().split("tcloudbase.com")[1].split("?")[0];
    if (p === "/auth/session") return json(route, { success: true, user: { username: "测试用户" } });
    return json(route, { detail: "服务异常" }, 503);
  });
  await page.goto(PAGE);
  await visible(page, "#busyRetryBtn");
  const kept = await page.evaluate(() => JSON.parse(localStorage.getItem("pptAgentMvp")));
  ok("恢复失败 5xx 保留项目引用", kept?.projectId === "p1", String(kept?.projectId));
  await page.close();
}
{
  // 404：项目确实不存在，这时才该清除本地引用
  const page = await browser.newPage();
  await seedSession(page);
  await seedProject(page);
  await page.route(API + "/**", (route) => {
    const p = route.request().url().split("tcloudbase.com")[1].split("?")[0];
    if (p === "/auth/session") return json(route, { success: true, user: { username: "测试用户" } });
    return json(route, { detail: "项目不存在" }, 404);
  });
  await page.goto(PAGE);
  await until(page, () => !JSON.parse(localStorage.getItem("pptAgentMvp") || "{}").projectId);
  const cleared = await page.evaluate(() => JSON.parse(localStorage.getItem("pptAgentMvp")));
  ok("恢复失败 404 清除失效项目引用", !cleared?.projectId, String(cleared?.projectId));
  await page.close();
}

await browser.close();

console.log(`通过 ${pass.length}:`);
pass.forEach((t) => console.log("  ✓ " + t));
if (fail.length) {
  console.log(`\n失败 ${fail.length}:`);
  fail.forEach((t) => console.log("  ✗ " + t));
  process.exit(1);
}
console.log("\n全部通过。");
