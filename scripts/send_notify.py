"""
Send notification via WeChat API after watchdog restart.

Called by watchdog.bat after ensuring the agent process is alive.
Reads pending_notify.json and the WeChat credentials, then sends a text message.
"""
import json
import os
import sys
import time
import uuid
import urllib.request
import urllib.error

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE = os.path.dirname(SCRIPT_DIR)  # ~/cow
NOTIFY_FILE = os.path.join(SCRIPT_DIR, "pending_notify.json")
CREDENTIALS_FILE = os.path.expanduser("~/.weixin_cow_credentials.json")
BOT_TOKEN_FILE = os.path.join(WORKSPACE, ".bot_token.json")


def load_credentials():
    """Load WeChat credentials from the saved file."""
    cred_paths = [CREDENTIALS_FILE]
    try:
        if os.path.exists(BOT_TOKEN_FILE):
            cred_paths.insert(0, BOT_TOKEN_FILE)
    except Exception:
        pass

    for path in cred_paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                token = data.get("token") or data.get("bot_token", "")
                user_id = data.get("user_id", "")
                base_url = data.get("base_url", "https://ilinkai.weixin.qq.com")
                if token and user_id:
                    return token, user_id, base_url
            except Exception:
                continue

    return None, None, None


def load_notify():
    """Read the pending notification file."""
    if not os.path.exists(NOTIFY_FILE):
        return None
    try:
        with open(NOTIFY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("message", "")
    except Exception as e:
        return f"（通知读取失败: {e}）"


def send_wechat_message(token: str, user_id: str, base_url: str, text: str):
    """Send a text message via WeChat API (direct HTTP, no channel dependency)."""
    url = f"{base_url.rstrip('/')}/ilink/bot/sendmessage"
    
    body = json.dumps({
        "msg": {
            "from_user_id": "",
            "to_user_id": user_id,
            "client_id": uuid.uuid4().hex[:16],
            "message_type": 2,  # BOT
            "message_state": 2,  # FINISH
            "item_list": [{"type": 1, "text_item": {"text": text}}],
            "context_token": "",
        }
    }).encode("utf-8")

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}",
        "AuthorizationType": "ilink_bot_token",
        "X-WECHAT-UIN": "1",
        "iLink-App-Id": "bot",
        "iLink-App-ClientVersion": "131072",
    }

    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            ret = result.get("ret", -1)
            if ret == 0:
                return True, "发送成功"
            else:
                return False, f"API 返回错误: ret={ret} errmsg={result.get('errmsg', '')}"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode('utf-8', errors='ignore')[:200]}"
    except Exception as e:
        return False, f"请求异常: {e}"


def main():
    # Read notification
    message = load_notify()
    if not message:
        print("[send_notify] No pending notification found, skipping")
        clear_notify()
        return 0

    print(f"[send_notify] Sending: {message[:50]}...")

    # Load credentials
    token, user_id, base_url = load_credentials()
    if not token or not user_id:
        print("[send_notify] CREDENTIALS NOT FOUND - cannot send notification")
        print(f"[send_notify] Tried: {CREDENTIALS_FILE}, {BOT_TOKEN_FILE}")
        clear_notify(event_desc="（看门狗重启成功，但发送通知失败：缺少凭据）")
        return 0

    # Send with retry
    for attempt in range(3):
        success, errmsg = send_wechat_message(token, user_id, base_url, message)
        if success:
            print(f"[send_notify] ✅ Notification sent: {message[:50]}")
            clear_notify(event_desc=message)
            return 0
        else:
            print(f"[send_notify] ❌ Attempt {attempt+1}/3 failed: {errmsg}")
            if attempt < 2:
                time.sleep(2)

    # All attempts failed - leave the file with error info
    print("[send_notify] All attempts failed, leaving pending_notify.json for manual check")
    return 1


def clear_notify(event_desc=None):
    """Remove the pending notification file."""
    try:
        if os.path.exists(NOTIFY_FILE):
            os.remove(NOTIFY_FILE)
            print(f"[send_notify] Cleared pending_notify.json")
    except Exception as e:
        print(f"[send_notify] Failed to clear notification: {e}")


if __name__ == "__main__":
    sys.exit(main())
