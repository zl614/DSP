# 包裹清单拆分器 + 按车队发邮件

单文件本地程序。双击 exe → 浏览器自动打开 → 拖 Excel 进去。

1. 按「Last delivery fleet」把 Excel 拆成**每个车队一个文件**（可打包下载 ZIP）
2. 把各自那份**直接发给对应车队** —— 用你自己的 Outlook 账号

## 仓库文件

```
speedx_relay.py                     后端（Flask，只做发邮件 + 托管页面）
fleet_splitter.html                 页面（拆分逻辑在浏览器里跑）
speedx_panel.spec                   打包配置（自动收集根目录所有 *.html）
requirements.txt                    flask + requests（就这两个）
graph_config.json                   Azure client_id
.github/workflows/main.yml          云端打 exe
```

## 出 exe

推到 `main` 自动构建；也可在 **Actions → build-windows-exe → Run workflow** 手动点。
构建完在那次 run 底部 **Artifacts** 下载 `SpeedX-Panel-windows`，解压得到：

```
SpeedX-Panel.exe
fleet_splitter.html
graph_config.json
```

三个放一起，双击 exe。黑框会打印实际状态：

```
  打开： http://127.0.0.1:5058/
  页面： fleet_splitter.html  ← exe 同目录
  邮件： Graph 已配置 (client_id …7001eb)
```

看到 **`Graph 已配置`** 就说明发送功能是活的。

## 改页面不用重新打包

运行时**优先读 exe 同目录的 html**。改完页面，覆盖 exe 旁边那个 `fleet_splitter.html`、刷新浏览器就行。

## 发邮件

新版 Outlook 砍了 COM，本地脚本驱动不了它，`.eml` 也不认；浏览器又不能连 SMTP、
mailto 不能带附件。所以走 **Microsoft Graph** —— 需要一个后端跑 OAuth，就是这个程序。

`graph_config.json`（放 exe 同目录，改配置不用重新打包）：

```json
{
  "client_id": "3cbd3c14-eb3f-4d28-837b-b92bab7001eb",
  "tenant_id": "a7f846ec-8219-438c-8896-9e45cdbbe994",
  "allow_draft": true
}
```

Azure 应用注册（**已配好并验证通过**）：Single tenant · 无 Redirect URI ·
**Allow public client flows = Yes** · Delegated 权限 `Mail.Send` / `Mail.ReadWrite` / `User.Read`。
**不需要 client secret。**

用法：拖 Excel → 填车队邮箱（可从 Excel 直接粘两列；填过的会记住）→
点「登录 Outlook 邮箱」（设备码，第一次会让你同意授权）→ 点「建草稿」或「直接发送」。

- **建草稿**：每个车队一封、附件已带好，直接出现在你 Outlook 的草稿箱，扫一眼再发
- **直接发送**：全部发出（先弹确认框）

不存密码（OAuth），登录态只在内存里，关掉程序就没了。
`allow_draft: false` = 只给发送权限、不给读邮箱权限（草稿按钮会变灰）。

## 本地开发（可选，需要 Python）

```bash
pip install -r requirements.txt
python speedx_relay.py
```
