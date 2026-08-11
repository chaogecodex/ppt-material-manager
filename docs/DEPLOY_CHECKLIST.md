# 交付检查清单

本仓库**没有自动部署**：前端托管在 CloudBase 静态托管，后端是 CloudBase Run 容器，
两边都需要人工发布。合并到 `main` 不会触发任何上线动作。

照着本清单从头走一遍，可以在上线前后各留下一个可验证的结论，而不是靠"应该没问题"。

相关地址与环境变量说明见 `README.md` 与 `api/DEPLOY.txt`。

---

## 1. 部署前：建立基线 fixture

**在旧版后端上**先建一份数据，用来验证重新部署后数据没丢。这一步必须在部署之前做，
否则就失去了对照意义。

```bash
cd api
python -m pip install -r requirements.txt httpx

PPT_TEST_INVITE_CODE=<一个有效邀请码> \
python tests/online_e2e.py create \
  --base-url https://ppt-ai-api-294833-10-1466238220.sh.run.tcloudbase.com \
  --state /tmp/e2e-state.json
```

预期输出：`{"created": true, "projectId": "...", "files": 2}`

这一步会登录、建项目、上传两个素材、锁定顺序，并把 `projectId` 与访问令牌写进
`/tmp/e2e-state.json`。**这个文件不要删**，第 4 步要用。

---

## 2. 部署后端

1. 打包 `api/`（`Dockerfile` 在该目录下），在 CloudBase Run 重新构建镜像
2. 服务监听端口 **80**
3. 核对环境变量（只核对存在与取值，不要在这一步顺手改密钥）：

   | 变量 | 说明 |
   |---|---|
   | `ZHIPU_API_KEY` | 仅服务端持有，不入仓库 |
   | `INVITE_CODES_JSON` | 覆盖内置邀请码；仅存 PBKDF2 哈希 |
   | `INVITE_EXPIRES_AT` | **本批邀请码统一到期时间（ISO 8601）** |
   | `ACCESS_AUTH_SECRET` | 可选；未设置时从服务器密钥派生 |
   | `ACCESS_SESSION_SECONDS` | 默认 2592000（30 天） |

**`INVITE_EXPIRES_AT` 是最容易出错的一项。** 会话不会越过邀请码到期时间，
所以这个值设早了，用户会在预期之外被踢下线；设晚了，等于邀请码没有按时失效。
交付前按实际约定的有效期核对一次。

---

## 3. 部署前端

1. **先把当前线上的 `index.html` 存一份**，作为回滚用的上一版
2. 上传新的 `index.html` 到静态托管
3. 部署后强制刷新（CDN 有缓存），确认浏览器拿到的确实是新版

### 怎么确认线上是新版

打开正式地址，在 devtools 里查以下特征：

| 特征 | 新版 | 旧版 |
|---|---|---|
| 页面源码含 `busyRetryBtn` | 有 | 无 |
| 页面源码含 `suggestionBox` | 有 | 无 |
| 页面源码含 `schemaVersion` | 有 | 无 |
| 加载完成后 `window.scrollY` | `0` | `624` |

最后一行是最直观的：旧版一打开就会把首屏滚掉，新版不会。

---

## 4. 部署后：验证持久化与真实 AI

```bash
cd api
python tests/online_e2e.py verify \
  --base-url https://ppt-ai-api-294833-10-1466238220.sh.run.tcloudbase.com \
  --state /tmp/e2e-state.json
```

这一步会依次确认：

- 第 1 步建的项目、素材顺序、`locked` 状态在重新部署后**仍然存在**
- 每个素材文件都能下载下来（COS 持久卷没丢数据），且中文文件名编码正确
- CORS 预检返回的 `access-control-allow-origin` 与前端正式域名**精确匹配**
- **真实调用一次 `/analyze`**，断言产出了 outline 且 `analysisStatus == "done"`

预期输出：`{"verified": true, ..., "analysisStatus": "done", "outlineSections": N}`

> **请记录这一步 analyze 实际花了多久。**
> 脚本给了 300 秒超时，但线上网关的容忍度低得多——已知实际出现过 504。
> 这个耗时直接决定下一节「已知风险」的严重程度，是本次交付最值得留下的一个数字。

---

## 5. 真机走查（在微信里，不是桌面浏览器）

自动化测试全部是 mock 后端的，只证明"给定后端这样回，前端行为正确"。
以下必须在真机上确认：

- [ ] 首屏：打开后不自动下滚，能看到大标题和「开始创建PPT」按钮
- [ ] 断网操作：显示中文提示与重试按钮，**不出现** `Failed to fetch`
- [ ] 上传一批**真实体量**的素材做分析，重点观察是否 504
- [ ] 登录态：关掉页面再打开，不应要求重新登录（30 天会话）
- [ ] 退出按钮可见（这一项在修复前是坏的）

---

## 6. 回滚

| 对象 | 回滚方式 |
|---|---|
| 前端 | 重新上传第 3 步保存的上一版 `index.html` |
| 后端 | CloudBase Run 切回上一个镜像版本 |
| 数据 | 用第 1 步的 fixture 跑 `verify`，确认回滚后数据仍完好 |

---

## 已知风险（交付时应主动告知使用方）

**analyze 可能超时。** 已确认耗时会超过 30 秒，并实际出现过 504。

本轮修复改善的是**反馈**：进度条、状态提示、失败后可重试、错误文案中文化。
**根因未修** —— 素材量大时仍可能失败，用户看到的是"服务处理超时，请重试"加一个
重试按钮，而不再是一句英文报错。

根治需要异步任务化（`202 + jobId` + 轮询），属独立一轮工作。
在那之前，建议交付时说明：素材较多时分析可能需要重试。
