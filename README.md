# PPT资料管家

面向微信传播的 PPT 素材编排 H5，前端托管于 CloudBase，后端使用 CloudBase Run 与 COS 持久卷。

- 正式地址：https://ppt-material-manager-ppt-ai-d8gx2g506c70eb54e.webapps.tcloudbase.com/
- 后端服务：https://ppt-ai-api-294833-10-1466238220.sh.run.tcloudbase.com
- CloudBase 环境：`ppt-ai-d8gx2g506c70eb54e`
- 小范围测试使用负责人分发的邀请码登录；本批 10 个邀请码统一有效 30 天
- 邀请码明文不写入仓库或前端代码，后端仅保存 PBKDF2 哈希
- 登录会话最长 30 天，且不会越过邀请码到期时间

## 代码结构

- `index.html`：微信 H5 前端、邀请码登录和 PPT 素材编排界面
- `api/main.py`：CloudBase Run 后端、邀请码校验、项目和素材 API
- `api/tests/`：后端回归与线上端到端测试
- `api/Dockerfile`：CloudBase Run 容器构建入口

## 本地验证

```bash
cd api
python -m pip install -r requirements.txt
python -m unittest discover -s tests -p "test_*.py" -v
```
