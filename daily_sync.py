import os, re, sys, requests, msal, git
from datetime import datetime, timedelta
# 移除 dateutil，改用內建字串處理，這樣就不會報錯了

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

def sanitize(text):
    if not text: return "No Subject"
    
    for kw in SENSITIVE_KEYWORDS:
        if kw.lower() in text.lower(): return "💼 Internal Task"
    
    for real, safe in PROJECT_MAPPINGS.items():
        text = text.replace(real, safe)
        
    text = re.sub(r'[\w\.-]+@[\w\.-]+\.\w+', '[Contact]', text)
    return text

def check_leaks(content):
    secrets = [CLIENT_SECRET, REFRESH_TOKEN]
    for s in secrets:
        if s and s in content: 
            print("!!! Security Alert: Secret leak detected !!!")
            sys.exit(1)

def get_calendar_events(access_token, today_str, tomorrow_str):
    print("--- 正在抓取行事曆 ---")
    url = f"https://graph.microsoft.com/v1.0/me/calendar/events?startDateTime={today_str}T00:00:00&endDateTime={tomorrow_str}T00:00:00&$top=50"
    headers = {
        'Authorization': 'Bearer ' + access_token, 
        'Prefer': 'outlook.timezone="Taipei Standard Time"'
    }
    res = requests.get(url, headers=headers)
    
    if res.status_code != 200:
        print(f"行事曆 API 錯誤: {res.text}")
        return []

    events = []
    for evt in res.json().get('value', []):
        if evt.get('isCancelled'): continue
        
        # 隱私檢查
        if evt.get('sensitivity') in ['private', 'personal', 'confidential']:
            events.append(f"- **{evt['start']['dateTime'][11:16]}**: 🔒 Private Meeting")
            continue

        safe_sub = sanitize(evt.get('subject'))
        
        # 如果你想保留 Free 的行程，請註解掉下面這行
        #if evt.get('showAs') == 'free': continue 

        time_str = evt['start']['dateTime'][11:16]
        events.append(f"- **{time_str}**: {safe_sub}")
    
    print(f"找到 {len(events)} 個行事曆項目")
    return events

def get_todo_tasks(access_token, target_date_str):
    print("--- 正在抓取 To-Do (已完成項目) ---")
    headers = {'Authorization': 'Bearer ' + access_token}
    
    lists_res = requests.get("https://graph.microsoft.com/v1.0/me/todo/lists", headers=headers)
    if lists_res.status_code != 200:
        print(f"To-Do List API 錯誤: {lists_res.text}")
        return []
    
    tasks_found = []
    
    for task_list in lists_res.json().get('value', []):
        list_id = task_list['id']
        tasks_url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks?$filter=status eq 'completed'"
        tasks_res = requests.get(tasks_url, headers=headers)
        
        if tasks_res.status_code == 200:
            for task in tasks_res.json().get('value', []):
                completed_obj = task.get('completedDateTime')
                if completed_obj:
                    # 使用字串切割 [:10] 取出日期，不需要額外套件
                    c_date = completed_obj.get('dateTime', '')[:10]
                    
                    if c_date == target_date_str:
                        safe_title = sanitize(task.get('title'))
                        tasks_found.append(f"- ✅ **Completed**: {safe_title}")
    
    print(f"找到 {len(tasks_found)} 個已完成任務")
    return tasks_found

def main():
    print("--- 開始執行同步 (行事曆 + To-Do) ---")
    if not REFRESH_TOKEN: 
        print("Missing Refresh Token")
        sys.exit(1)
    
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}', client_credential=CLIENT_SECRET)
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=['Calendars.Read', 'Tasks.Read'])
    
    if "access_token" not in result: 
        print(f"Token Error: {result.get('error')}")
        sys.exit(1)
    
    token = result['access_token']
    
    # 強制 UTC+8
    tw_now = datetime.utcnow() + timedelta(hours=8)
    today_str = tw_now.strftime('%Y-%m-%d')
    tomorrow_str = (tw_now + timedelta(days=1)).strftime('%Y-%m-%d')
    print(f"目標日期: {today_str}")

    calendar_lines = get_calendar_events(token, today_str, tomorrow_str)
    todo_lines = get_todo_tasks(token, today_str)
    
    all_lines = calendar_lines + todo_lines

    if all_lines:
        all_lines.sort()
        content = f"# {today_str} Work Log\n\n## Calendar\n"
        content += "\n".join(calendar_lines) if calendar_lines else "No events."
        
        content += "\n\n## To-Do Tasks\n"
        content += "\n".join(todo_lines) if todo_lines else "No tasks completed."
        
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
            origin.push()
            print("Git Push 完成。")
        else:
            print("沒有變更需要 Commit。")
    else:
        print("今天沒有行事曆也沒有完成的任務。")

if __name__ == "__main__":
    main()
