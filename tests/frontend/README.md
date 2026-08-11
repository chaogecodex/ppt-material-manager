# 前端回归测试

用 Playwright 直接打开仓库根目录的 `index.html`，拦截后端接口做端到端验证。
不需要启动后端，也不访问真实 CloudBase 环境。

## 运行

```bash
cd tests/frontend
npm ci
npx playwright install chromium   # 首次运行需要
npm test
```

若本机已有预装 Chromium，可跳过下载：

```bash
CHROMIUM_PATH=/path/to/chromium npm test
```

全部通过时退出码为 0，任一用例失败退出码为 1。
CI 中由 `.github/workflows/ci.yml` 的「前端回归」任务自动运行。

## 覆盖范围

| 分组 | 验证内容 |
|---|---|
| A1 | 页面加载后不自动滚动，首屏 hero 不被跳过 |
| A2 | 网络异常显示中文提示而非 `Failed to fetch`，并提供重试 |
| A3 | 旧版本地状态迁移：保留项目要求、清除无法补传的素材、写入 `schemaVersion`、回到第 1 步、只提示一次 |
| B1 | 服务端返回的 `type` / `size` 含恶意串时按纯文本渲染，不注入 DOM、不执行脚本 |
| C2 | AI 顺序建议正确渲染并转义 |
| 登录态 | 已有云端项目时仍恢复 access session，退出按钮可见 |
| 旧大纲清理 | 删除素材后、分析请求失败后、保存阶段失败后，均不残留上一版大纲 |
| 项目恢复 | 网络故障与 5xx 保留本地项目引用；仅 401/403/404 才清除 |

## 设计约定

断言的是**用户实际看到的行为**，不是内部实现细节，因此重构 `index.html`
时这些用例仍然有效。每个用例自带接口 mock，互不影响，可单独阅读。
