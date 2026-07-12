# SpeedX 面板

本机跑的小服务，自带登录，页面在浏览器里开。

| 路由 | 页面 |
|---|---|
| `/` | 操作员人效 |
| `/exceptions` | 异常原因分析 |
| `/split` | **包裹清单拆分器 + 按车队发邮件** |

---

## 仓库要放的文件

```
speedx_relay.py                    ← 后端（Flask）
speedx_panel.spec                  ← 打包配置（自动收集根目录所有 *.html）
requirements.txt
graph_config.json                  ← Azure client_id（发邮件用）
.github/workflows/build-windows.yml ← 云端打 exe

dashboard_realtime.html            ← 你现有的，拷进来
exception_reason_analysis.html     ← 你现有的，拷进来
fleet_splitter.html                ← 新增
```

> **新增页面不用改 spec** —— spec 会自动把根目录下所有 `*.html` 打进 exe。

## 出 exe

推到 `main` 就自动构建；也可以在 **Actions → build-windows-exe → Run workflow** 手动点。
构建完在那次 run 的 **Artifacts** 里下载 `SpeedX-Panel-windows.zip`，解压得到：

```
SpeedX-Panel.exe
dashboard_realtime.html
exception_reason_analysis.html
fleet_splitter.html
graph_config.json
```

双击 exe → 浏览器自动打开。黑框里会打印每个页面从哪读的：

```
  页面来源：
    dashboard_realtime.html          ← exe 同目录
    exception_reason_analysis.html   ← exe 同目录
    fleet_splitter.html              ← exe 同目录
```

## 改页面不用重新打包

运行时**优先读 exe 同目录的 html**，找不到才用打进 exe 的那份。
所以改完页面，把新的 `.html` 覆盖到 exe 旁边、刷新浏览器就行——不用重新走一遍构建。
（黑框里那行「页面来源」就是给你确认它到底读了哪个文件用的。）

## 发邮件（/split）

新版 Outlook 没有 COM，本地脚本驱动不了它，`.eml` 它也不认——所以走 **Microsoft Graph**。

`graph_config.json`（放 exe 同目录，改配置不用重新打包）：

```json
{
  "client_id": "3cbd3c14-eb3f-4d28-837b-b92bab7001eb",
  "tenant_id": "a7f846ec-8219-438c-8896-9e45cdbbe994",
  "allow_draft": true
}
```

Azure 那边（**已配好并验证通过**）：Single tenant · 无 Redirect URI · **Allow public client flows = Yes** ·
Delegated 权限 `Mail.Send` / `Mail.ReadWrite` / `User.Read`。**不需要 client secret。**

用法：拖 Excel → 填车队邮箱 → 「登录 Outlook 邮箱」（设备码）→ 「建草稿」或「直接发送」。
不存密码（OAuth），登录态只在内存里，关掉 exe 就没了。

`allow_draft: false` = 只要发送权限、不给读邮箱权限（草稿按钮会变灰）。

## 本地开发（可选，需要 Python）

```bash
pip install -r requirements.txt
python speedx_relay.py
```
