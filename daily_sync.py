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
    if event.get('isCancelled') or event.get('showAs') not in ['busy', 'oof']: return None
    if event.get('sensitivity') in ['private', 'personal', 'confidential']: return "🔒 Private Task"
    
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
    if not REFRESH_TOKEN: 
        print("Missing Refresh Token")
        sys.exit(1)
    
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f'https://login.microsoftonline.com/{TENANT_ID}', client_credential=CLIENT_SECRET)
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=['Calendars.Read', 'Tasks.Read'])
    
    if "access_token" not in result: 
        print(f"Token Error: {result.get('error')}")
        sys.exit(1)
    
    today = datetime.now().strftime('%Y-%m-%d')
    tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
    url = f"https://graph.microsoft.com/v1.0/me/calendar/events?startDateTime={today}T00:00:00&endDateTime={tomorrow}T00:00:00&$top=50"
    
    # GitHub Actions 環境直接請求即可
    res = requests.get(url, headers={'Authorization': 'Bearer ' + result['access_token'], 'Prefer': 'outlook.timezone="Taiwan Standard Time"'})
    
    lines = []
    for evt in res.json().get('value', []):
        safe_sub = sanitize(evt)
        if safe_sub: lines.append(f"- **{evt['start']['dateTime'][11:16]}**: {safe_sub}")
    
    if lines:
        lines.sort()
        content = f"# {today} Work Log\n\n" + "\n".join(lines)
        check_leaks(content)
        
        repo = git.Repo(os.getcwd())
        log_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = os.path.join(log_dir, f"{today}.md")
        
        with open(path, 'w', encoding='utf-8') as f: f.write(content)
        
        repo.config_writer().set_value("user", "name", "GitHub Action").release()
        repo.config_writer().set_value("user", "email", "action@github.com").release()
        
        repo.index.add([path])
        if repo.is_dirty(untracked_files=True):
            repo.index.commit(f"Log: {today}")
            origin = repo.remote(name='origin')
            origin.push()
            print(f"Successfully logged {len(lines)} events.")
        else:
            print("No changes to commit.")
    else:
        print("No eligible events found today.")

if __name__ == "__main__":
    main()
