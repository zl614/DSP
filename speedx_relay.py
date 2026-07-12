#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SpeedX 包裹清单拆分器 + 按车队发邮件 —— 单文件本地程序
=========================================================
双击 exe 后：① 起一个本地服务 ② 自动打开浏览器。

做两件事：
  1) 把一份 Excel 按「Last delivery fleet」拆成每个车队一个文件（可打包下载 ZIP）
  2) 把各自那份直接发给对应车队 —— 走 Microsoft Graph，用你自己的 Outlook 账号

为什么发邮件非要经过这个程序：
  新版 Outlook 砍掉了 COM 接口，本地脚本（.bat / PowerShell / AppleScript）驱动不了它，
  .eml 它也不认；浏览器又不能直接连 SMTP，mailto 也不能带附件。
  能发信的只剩 Microsoft Graph API —— 它要跑 OAuth，就需要一个后端，就是这个程序。

不存密码：走 OAuth（设备码流程），登录态只在内存里，关掉程序就没了。

本地跑（可选，需要 Python）：
    pip install flask requests
    python speedx_relay.py
打 Windows exe：推到 GitHub，Actions 里的 build-windows-exe 会自动打包。
"""
import os
import sys
import json
import time
import re
import socket
import threading
import webbrowser

from flask import Flask, request, Response
import requests

# ===================== 配置 =====================
CFG = {
    "port": 5058,          # 被占用会自动往后找（5058 → 5059 → …）
}

# ===================== Outlook 邮件（Microsoft Graph）=====================
# OAuth 和发信全部用 requests 手写，不引入新依赖。
#
# 配置方式（优先级从高到低）：
#   1) 环境变量 SPEEDX_GRAPH_CLIENT_ID / SPEEDX_GRAPH_TENANT_ID
#   2) 与本程序（或 exe）同目录的 graph_config.json
#   3) 下面 GRAPH_CFG 里的默认值
# 没配 client_id 也不影响：拆分、下载 ZIP 照常用，只是「发送」按钮是灰的。
GRAPH_CFG = {
    "client_id": "",
    "tenant_id": "common",
    "allow_draft": True,   # True=可建草稿(需 Mail.ReadWrite)；False=只发送(仅需 Mail.Send)
}
GRAPH_LOGIN = "https://login.microsoftonline.com"
GRAPH_API = "https://graph.microsoft.com/v1.0"

GRAPH_TOKEN = {"access": None, "refresh": None, "exp": 0.0, "user": ""}
GRAPH_FLOW = {"device_code": None, "interval": 5, "expires": 0.0}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

PAGE = "fleet_splitter.html"


def _app_dir():
    """exe 时取 exe 所在目录（改配置/改页面都不用重新打包）；开发时取脚本目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(rel):
    """打包后从 _MEIPASS 读，开发时从脚本目录读。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def find_asset(rel):
    """找页面文件：**exe 同目录优先**，其次才是打包进 exe 的那份。
    这样改页面不用重新打包 —— 把新的 html 丢到 exe 旁边、刷新浏览器就行。"""
    side = os.path.join(_app_dir(), rel)
    if os.path.exists(side):
        return side
    return resource_path(rel)


def asset_origin(rel):
    if os.path.exists(os.path.join(_app_dir(), rel)):
        return "exe 同目录"
    return "exe 内置" if os.path.exists(resource_path(rel)) else "缺失 !"


def _graph_scopes():
    s = ["https://graph.microsoft.com/Mail.Send",
         "https://graph.microsoft.com/User.Read",
         "offline_access"]
    if GRAPH_CFG["allow_draft"]:
        s.insert(1, "https://graph.microsoft.com/Mail.ReadWrite")
    return " ".join(s)


def _authority():
    return f"{GRAPH_LOGIN}/{(GRAPH_CFG['tenant_id'] or 'common').strip()}"


def load_graph_cfg():
    cid = os.environ.get("SPEEDX_GRAPH_CLIENT_ID")
    tid = os.environ.get("SPEEDX_GRAPH_TENANT_ID")
    if cid:
        GRAPH_CFG["client_id"] = cid.strip()
    if tid:
        GRAPH_CFG["tenant_id"] = tid.strip()
    p = os.path.join(_app_dir(), "graph_config.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                j = json.load(f)
            if j.get("client_id"):
                GRAPH_CFG["client_id"] = str(j["client_id"]).strip()
            if j.get("tenant_id"):
                GRAPH_CFG["tenant_id"] = str(j["tenant_id"]).strip()
            if "allow_draft" in j:
                GRAPH_CFG["allow_draft"] = bool(j["allow_draft"])
        except Exception as e:
            print(f"[mail] graph_config.json 读取失败：{e}")
    return bool(GRAPH_CFG["client_id"])


def _store_token(j):
    GRAPH_TOKEN["access"] = j.get("access_token")
    GRAPH_TOKEN["refresh"] = j.get("refresh_token") or GRAPH_TOKEN["refresh"]
    GRAPH_TOKEN["exp"] = time.time() + int(j.get("expires_in", 3600)) - 120
    try:
        me = requests.get(f"{GRAPH_API}/me",
                          headers={"Authorization": "Bearer " + GRAPH_TOKEN["access"]},
                          timeout=20).json()
        GRAPH_TOKEN["user"] = me.get("mail") or me.get("userPrincipalName") or ""
    except Exception:
        pass


def graph_token():
    """拿可用的 access token；过期就用 refresh token 续，不用你重新登录。"""
    if GRAPH_TOKEN["access"] and time.time() < GRAPH_TOKEN["exp"]:
        return GRAPH_TOKEN["access"]
    if GRAPH_TOKEN["refresh"]:
        r = requests.post(f"{_authority()}/oauth2/v2.0/token", data={
            "grant_type": "refresh_token",
            "client_id": GRAPH_CFG["client_id"],
            "refresh_token": GRAPH_TOKEN["refresh"],
            "scope": _graph_scopes(),
        }, timeout=30)
        j = r.json()
        if "access_token" in j:
            _store_token(j)
            return GRAPH_TOKEN["access"]
    raise RuntimeError("邮箱未登录或登录已过期，请重新登录。")


def _recips(s):
    return [{"emailAddress": {"address": a.strip()}}
            for a in re.split(r"[;,]", s or "") if a.strip()]


# ===================== Web =====================
app = Flask(__name__)


@app.after_request
def cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return resp


@app.route("/")
@app.route("/split")            # 老链接也留着
def index():
    try:
        with open(find_asset(PAGE), encoding="utf-8") as f:
            return Response(f.read(), mimetype="text/html; charset=utf-8")
    except FileNotFoundError:
        return Response(f"找不到 {PAGE}。把它放到本程序（exe）同目录即可。",
                        status=500, mimetype="text/plain; charset=utf-8")


@app.route("/health")
def health():
    return {"ok": True}


# ---------- 邮件：状态 / 登录 / 发送 ----------
@app.route("/mail/status")
def mail_status():
    return {
        "configured": bool(GRAPH_CFG["client_id"]),
        "signedIn": bool(GRAPH_TOKEN["access"] or GRAPH_TOKEN["refresh"]),
        "user": GRAPH_TOKEN["user"] or "",
        "canDraft": bool(GRAPH_CFG["allow_draft"]),
    }


@app.route("/mail/signin-start", methods=["POST"])
def mail_signin_start():
    if not GRAPH_CFG["client_id"]:
        return {"ok": False, "msg": "还没配置 client_id（把 graph_config.json 放到程序同目录）。"}
    try:
        r = requests.post(f"{_authority()}/oauth2/v2.0/devicecode",
                          data={"client_id": GRAPH_CFG["client_id"], "scope": _graph_scopes()},
                          timeout=30)
        j = r.json()
    except Exception as e:
        return {"ok": False, "msg": f"连不上 Microsoft 登录服务：{e}"}
    if "device_code" not in j:
        return {"ok": False, "msg": j.get("error_description") or str(j)[:300]}
    GRAPH_FLOW["device_code"] = j["device_code"]
    GRAPH_FLOW["interval"] = int(j.get("interval", 5))
    GRAPH_FLOW["expires"] = time.time() + int(j.get("expires_in", 900))
    return {"ok": True,
            "userCode": j.get("user_code"),
            "verificationUri": j.get("verification_uri") or "https://microsoft.com/devicelogin",
            "interval": GRAPH_FLOW["interval"]}


@app.route("/mail/signin-poll", methods=["POST"])
def mail_signin_poll():
    if not GRAPH_FLOW["device_code"]:
        return {"ok": False, "state": "none", "msg": "还没开始登录。"}
    if time.time() > GRAPH_FLOW["expires"]:
        GRAPH_FLOW["device_code"] = None
        return {"ok": False, "state": "expired", "msg": "登录码过期了，重新点一次登录。"}
    try:
        r = requests.post(f"{_authority()}/oauth2/v2.0/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": GRAPH_CFG["client_id"],
            "device_code": GRAPH_FLOW["device_code"],
        }, timeout=30)
        j = r.json()
    except Exception as e:
        return {"ok": False, "state": "error", "msg": str(e)}
    if "access_token" in j:
        GRAPH_FLOW["device_code"] = None
        _store_token(j)
        return {"ok": True, "state": "done", "user": GRAPH_TOKEN["user"]}
    err = j.get("error", "")
    if err in ("authorization_pending", "slow_down"):
        return {"ok": True, "state": "pending"}
    GRAPH_FLOW["device_code"] = None
    return {"ok": False, "state": "error",
            "msg": j.get("error_description") or err or "登录失败"}


@app.route("/mail/signout", methods=["POST"])
def mail_signout():
    GRAPH_TOKEN.update({"access": None, "refresh": None, "exp": 0.0, "user": ""})
    GRAPH_FLOW["device_code"] = None
    return {"ok": True}


@app.route("/mail/send-one", methods=["POST"])
def mail_send_one():
    """发一封（或建一封草稿）。前端循环调用，好显示进度、也好定位是哪封失败。"""
    d = request.get_json(force=True, silent=True) or {}
    mode = (d.get("mode") or "draft").lower()
    to = (d.get("to") or "").strip()
    if not to:
        return {"ok": False, "msg": "没有收件人"}
    if mode == "draft" and not GRAPH_CFG["allow_draft"]:
        return {"ok": False, "msg": "草稿功能没开（需要 Mail.ReadWrite 权限）"}

    file_b64 = d.get("fileB64") or ""
    file_name = d.get("fileName") or "attachment.xlsx"
    if len(file_b64) > 3_000_000:                      # Graph 内联附件上限 ~3MB
        return {"ok": False, "msg": f"{file_name} 超过 3MB，Graph 内联附件放不下"}

    msg = {
        "subject": d.get("subject") or "",
        "body": {"contentType": "Text", "content": d.get("body") or ""},
        "toRecipients": _recips(to),
        "attachments": [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": file_name,
            "contentType": XLSX_MIME,
            "contentBytes": file_b64,
        }],
    }
    cc = (d.get("cc") or "").strip()
    if cc:
        msg["ccRecipients"] = _recips(cc)

    try:
        tok = graph_token()
    except Exception as e:
        return {"ok": False, "needLogin": True, "msg": str(e)}

    h = {"Authorization": "Bearer " + tok, "Content-Type": "application/json"}
    try:
        if mode == "send":
            r = requests.post(f"{GRAPH_API}/me/sendMail", headers=h,
                              json={"message": msg, "saveToSentItems": True}, timeout=120)
            ok = r.status_code in (200, 202)
        else:
            r = requests.post(f"{GRAPH_API}/me/messages", headers=h, json=msg, timeout=120)
            ok = r.status_code in (200, 201)
        if ok:
            return {"ok": True}
        return {"ok": False, "msg": f"Graph {r.status_code}: {r.text[:300]}"}
    except Exception as e:
        return {"ok": False, "msg": f"{type(e).__name__}: {e}"}


# ===================== 启动 =====================
def _free_port(start, tries=10):
    """端口被占（比如你另一个面板也开着）就自动往后找，别直接崩。"""
    for p in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


def _open_browser(port):
    webbrowser.open(f"http://127.0.0.1:{port}/")


if __name__ == "__main__":
    ok_graph = load_graph_cfg()
    port = _free_port(CFG["port"])

    print("=" * 58)
    print("  SpeedX 包裹清单拆分器 + 按车队发邮件")
    print("=" * 58)
    print(f"  打开： http://127.0.0.1:{port}/")
    print(f"  页面： {PAGE}  ← {asset_origin(PAGE)}")
    if ok_graph:
        cid = GRAPH_CFG["client_id"]
        print(f"  邮件： Graph 已配置 (client_id …{cid[-6:]})")
    else:
        print("  邮件： 未配置 client_id → 拆分照常用，但『发送』是灰的。")
        print("         把 graph_config.json 放到本程序同目录即可启用。")
    print("=" * 58)
    print("  关掉这个黑框 = 退出程序。")
    print()

    threading.Timer(1.2, _open_browser, args=(port,)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
