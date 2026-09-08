# 盛意旺落地页 · 企微活码留资闭环

落地页提交企业名称 / 手机号 / 合作目的后，服务端从活码池领取一条「联系我」二维码（`lead.id = code_id = state`），弹真码给客户扫。客户加好友后，企微回调按 `State` 关联留资，并给销售写入备注、打标、发欢迎语。

## 本地怎么跑

1. 复制配置：把 `.env.example` 复制为 `.env`（默认可保持 `LEAD_MOCK=1`）
2. 停掉原来的 `python -m http.server 8088`
3. 启动：`python lead_server.py`
4. 打开：http://127.0.0.1:8088/index.html 或 `/index-v3.html`
5. 在「资深顾问」留资区提交表单，应弹出二维码
6. mock 下模拟客户已扫码：

```
POST http://127.0.0.1:8088/api/debug/simulate-scan
{"lead_id":"刚才返回的id"}
```

弹层会轮询状态，扫码成功后显示「顾问已收到您的资料」。

## 你需要在企微后台完成的事项（真闭环）

1. **企业认证**：未认证企业，「客户联系」API 权限受限。
2. **开通客户联系**：客户联系 → 配置「可使用客户联系功能的成员」，把接待销售加进去。记下该成员的 **userid**（不是姓名）。
3. **自建应用**：创建应用，开通「客户联系」权限，拿到 `corp_id` + `secret`。
4. **可信 IP**：把运行 `lead_server.py` 的公网出口 IP 配进应用可信 IP，否则 `get_access_token` 会失败。
5. **回调**：客户联系或自建应用里配置接收事件服务器  
   - URL：`https://channelnotifytest.shengpay.com/wx/callback`  
   - Token / EncodingAESKey：与 `.env` 里 `WECOM_TOKEN`、`WECOM_AES_KEY` 一致  
   - 该域名当前是空 OpenResty，需先把 `deploy/channelnotifytest.openresty.conf` 配上，并保证 8088 上有 `lead_server.py`  
   - 保存前先启动服务，企微会 GET 验签
6. **标签**：预先建「官网留资」以及意向标签（收款/硬件/SaaS/代理/API），把 tag_id 填进 `.env`。
7. 把 `.env` 里 `LEAD_MOCK` 改为 `0`，填齐 `WECOM_CORP_ID` / `WECOM_SECRET` / `WECOM_SALES_USERID` / 回调三项，重启服务。

## 上线注意

- 二维码图片 URL 约 7–14 天过期，服务会在领取时刷新超过 6 天未用的码。
- `config_id` 已写入 `data/leads.db`，请备份该库。
- 手机号按本地密钥混淆存储；生产建议再换成 KMS / 独立密钥。
- 不要把 `.env` 和 `data/` 提交到 git。
