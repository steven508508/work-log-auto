import os, re, sys, requests, msal, git
from datetime import datetime, timedelta

# 讀取環境變數
CLIENT_ID = os.environ.get('MS_CLIENT_ID')
CLIENT_SECRET = os.environ.get('MS_CLIENT_SECRET')
TENANT_ID = os.environ.get('MS_TENANT_ID')
REFRESH_TOKEN = os.environ.get('MS_REFRESH_TOKEN')
GH_PAT = os.environ.get('GH_PAT') # 萬能鑰匙

SENSITIVE_KEYWORDS = ["Salary", "Review", "Interview", "Confidential", "Offer", "HR", "Bank"]
PROJECT_MAPPINGS = {
    "Project DeathStar": "Infrastructure Upgrade",
    "Client CocaCola": "Retail Client",
}

def sanitize(text):
    if not text: return "No Subject"
    for kw in SENSITIVE_KEYWORDS:
        if kw.lower() in text.lower(): return "💼 Internal Task"
    for real, safe in PROJECT_MAPPINGS.items():
        text = text.replace(real, safe)
    text = re.sub(r'[\w\.-]+@[\w\.-]+\.\w+', '[Contact]', text)
    return text

def check_leaks(content):
    secrets = [CLIENT_SECRET, REFRESH_TOKEN, GH_PAT]
    for s in secrets:
        if s and s in content: 
            print("!!! Security Alert: Secret leak detected !!!")
            sys.exit(1)

def get_calendar_events(access_token, today_str, tomorrow_str):
    print("\n--- [1/2] 正在抓取行事曆 ---")
    url = f"https://graph.microsoft.com/v1.0/me/calendar/events?startDateTime={today_str}T00:00:00&endDateTime={tomorrow_str}T00:00:00&$top=50"
    headers = {'Authorization': 'Bearer ' + access_token, 'Prefer': 'outlook.timezone="Taipei Standard Time"'}
    res = requests.get(url, headers=headers)
    if res.status_code != 200:
        print(f"❌ 行事曆 API 錯誤: {res.text}")
        return []

    events = []
    for evt in res.json().get('value', []):
        if evt.get('isCancelled'): continue
        if evt.get('sensitivity') in ['private', 'personal', 'confidential']:
            events.append(f"- **{evt['start']['dateTime'][11:16]}**: 🔒 Private Meeting")
            continue
        safe_sub = sanitize(evt.get('subject'))
        if evt.get('showAs') == 'free': continue 
        time_str = evt['start']['dateTime'][11:16]
        events.append(f"- **{time_str}**: {safe_sub}")
    print(f"✅ 找到 {len(events)} 個行事曆項目")
    return events

def get_todo_tasks(access_token, target_date_str):
    print("\n--- [2/2] 正在抓取 To-Do ---")
    headers = {'Authorization': 'Bearer ' + access_token}
    lists_res = requests.get("https://graph.microsoft.com/v1.0/me/todo/lists", headers=headers)
    if lists_res.status_code != 200: return []
    
    tasks_found = []
    for task_list in lists_res.json().get('value', []):
        list_id = task_list['id']
        tasks_url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks?$filter=status eq 'completed'"
        tasks_res = requests.get(tasks_url, headers=headers)
        if tasks_res.status_code == 200:
            for task in tasks_res.json().get('value', []):
                completed_obj = task.get('completedDateTime')
                if completed_obj:
                    try:
                        clean_time = completed_obj.get('dateTime', '').split('.')[0]
                        dt_tw = datetime.strptime(clean_time, "%Y-%m-%dT%H:%M:%S") + timedelta(hours=8)
                        if dt_tw.strftime('%Y-%m-%d') == target_date_str:
                            safe_title = sanitize(task.get('title'))
                            tasks_found.append(f"- ✅ **Completed**: {safe_title}")
                    except: pass
    print(f"✅ 找到 {len(tasks_found)} 個完成任務")
    return tasks_found

def main():
    print("--- 開始執行 (混合權限模式) ---")
    if not REFRESH_TOKEN: 
        print("Missing Refresh Token"); sys.exit(1)
    
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}', client_credential=CLIENT_SECRET)
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=['Calendars.Read', 'Tasks.Read'])
    if "access_token" not in result: print(f"Token Error: {result.get('error')}"); sys.exit(1)
    
    tw_now = datetime.utcnow() + timedelta(hours=8)
    today_str = tw_now.strftime('%Y-%m-%d')
    tomorrow_str = (tw_now + timedelta(days=1)).strftime('%Y-%m-%d')
    print(f"目標日期: {today_str}")

    all_lines = get_calendar_events(result['access_token'], today_str, tomorrow_str) + get_todo_tasks(result['access_token'], today_str)

    if all_lines:
        all_lines.sort()
        content = f"# {today_str} Work Log\n\n" + "\n".join(all_lines)
        check_leaks(content)
        
        # === Git 混合操作區塊 ===
        repo = git.Repo(os.getcwd())
        
        # 👇👇👇 務必改成你的帳號 👇👇👇
        repo.config_writer().set_value("user", "name", "steven508508").release()
        repo.config_writer().set_value("user", "email", "82710704+steven508508@users.noreply.github.com").release()
        
        log_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = os.path.join(log_dir, f"{today_str}.md")
        
        with open(path, 'w', encoding='utf-8') as f: f.write(content)
        print(f"檔案建立: {path}")
        
        repo.index.add([path])
        if repo.is_dirty(untracked_files=True):
            repo.index.commit(f"Log: {today_str}")
            
            # ★★★ 終極手段：手動組裝帶有密碼的 URL ★★★
            if GH_PAT:
                # 取得原本的網址 (例如 https://github.com/Kevin/Repo.git)
                origin_url = repo.remotes.origin.url
                
                # 強制移除原本網址中的 https:// 
                clean_url = origin_url.replace("https://", "").split("@")[-1]
                
                # 重新組裝成: https://oauth2:密碼@github.com/Kevin/Repo.git
                # 這種寫法可以直接繞過所有環境變數設定
                auth_url = f"https://oauth2:{GH_PAT}@{clean_url}"
                
                print("正在使用 PAT 強制推送...")
                # 使用 repo.git.push 直接呼叫指令，這是最底層、最不容易出錯的方法
                repo.git.push(auth_url, "HEAD:main")
                print("✅ Git Push 成功！")
            else:
                print("❌ 錯誤：找不到 GH_PAT，無法推送。")
        else:
            print("沒有變更。")
    else:
        print("無資料。")

if __name__ == "__main__":
    main()
