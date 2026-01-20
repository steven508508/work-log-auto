import os, re, sys, requests, msal, git
from datetime import datetime, timedelta

# 讀取 GitHub Secrets
CLIENT_ID = os.environ.get('MS_CLIENT_ID')
CLIENT_SECRET = os.environ.get('MS_CLIENT_SECRET')
TENANT_ID = os.environ.get('MS_TENANT_ID')
REFRESH_TOKEN = os.environ.get('MS_REFRESH_TOKEN')

# 過濾關鍵字設定
SENSITIVE_KEYWORDS = ["Salary", "Review", "Interview", "Confidential", "Offer", "HR", "Bank"]
PROJECT_MAPPINGS = {
    "Project DeathStar": "Infrastructure Upgrade",
    "Client CocaCola": "Retail Client",
}

def sanitize(event):
    subject = event.get('subject', 'No Subject')
    if event.get('isCancelled'): return None
    
    # 檢查隱私
    if event.get('sensitivity') in ['private', 'personal', 'confidential']: return "🔒 Private Task"
    
    # 關鍵字過濾
    for kw in SENSITIVE_KEYWORDS:
        if kw.lower() in subject.lower(): return "💼 Internal Discussion"
    
    for real, safe in PROJECT_MAPPINGS.items():
        subject = subject.replace(real, safe)
        
    subject = re.sub(r'[\w\.-]+@[\w\.-]+\.\w+', '[Contact]', subject)
    return subject

def check_leaks(content):
    secrets = [CLIENT_SECRET, REFRESH_TOKEN]
    for s in secrets:
        if s and s in content: 
            print("!!! Security Alert: Secret leak detected !!!")
            sys.exit(1)

def main():
    print("--- 開始執行同步 (修正時區版) ---")
    if not REFRESH_TOKEN: 
        print("Missing Refresh Token")
        sys.exit(1)
    
    # 1. 取得 Access Token
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}', client_credential=CLIENT_SECRET)
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=['Calendars.Read', 'Tasks.Read'])
    
    if "access_token" not in result: 
        print(f"Token Error: {result.get('error')}")
        sys.exit(1)
    
    # 2. 設定時間 (強制轉為台灣時間 UTC+8)
    tw_now = datetime.utcnow() + timedelta(hours=8)
    today_str = tw_now.strftime('%Y-%m-%d')
    tomorrow_str = (tw_now + timedelta(days=1)).strftime('%Y-%m-%d')
    
    print(f"台灣時間: {tw_now} (查詢目標日期: {today_str})")

    # 3. 呼叫 Graph API
    url = f"https://graph.microsoft.com/v1.0/me/calendar/events?startDateTime={today_str}T00:00:00&endDateTime={tomorrow_str}T00:00:00&$top=50"
    
    # ★★★ 關鍵修正：將 'Taiwan Standard Time' 改為 'Taipei Standard Time' ★★★
    headers = {
        'Authorization': 'Bearer ' + result['access_token'], 
        'Prefer': 'outlook.timezone="Taipei Standard Time"'
    }
    
    res = requests.get(url, headers=headers)
    print(f"API 回傳狀態碼: {res.status_code}")
    
    if res.status_code != 200:
        print(f"API 錯誤內容: {res.text}")
        sys.exit(1)

    events_data = res.json().get('value', [])
    print(f"共抓取到 {len(events_data)} 個原始行程")

    # 4. 處理資料
    lines = []
    for evt in events_data:
        subject = evt.get('subject', 'No Subject')
        show_as = evt.get('showAs')
        print(f"  - 檢查: [{show_as}] {subject}")
        
        # 如果你想連 Free 的行程都寫入，請把下面這兩行註解掉
        if show_as == 'free':
            print("    -> Skip (Free)")
            continue

        safe_sub = sanitize(evt)
        if safe_sub: 
            start_time = evt['start']['dateTime'][11:16]
            lines.append(f"- **{start_time}**: {safe_sub}")
            print(f"    -> OK (將寫入: {safe_sub})")
        else:
            print("    -> Skip (Sanitize returned None)")

    # 5. 寫入檔案與 Git 上傳
    if lines:
        lines.sort()
        content = f"# {today_str} Work Log\n\n" + "\n".join(lines)
        check_leaks(content)
        
        repo = git.Repo(os.getcwd())
        
        repo.config_writer().set_value("user", "name", "GitHub Action").release()
        repo.config_writer().set_value("user", "email", "action@github.com").release()
        
        log_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = os.path.join(log_dir, f"{today_str}.md")
        
        with open(path, 'w', encoding='utf-8') as f: f.write(content)
        print(f"檔案已建立: {path}")
        
        repo.index.add([path])
        if repo.is_dirty(untracked_files=True):
            repo.index.commit(f"Log: {today_str}")
            origin = repo.remote(name='origin')
            push_info = origin.push()
            print("Git Push 完成。")
        else:
            print("沒有變更需要 Commit。")
    else:
        print("沒有符合條件的行程，跳過寫入。")

if __name__ == "__main__":
    main()
