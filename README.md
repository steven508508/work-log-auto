[![Daily Sync](https://github.com/steven508508/work-log-auto/actions/workflows/schedule.yml/badge.svg)](https://github.com/steven508508/work-log-auto/actions/workflows/schedule.yml)
It's an useful to connect with MS Graph showing your Outlook Calendar and completed Todo.（old）

# work-log-auto

Automated privacy-safe daily work log generator from Microsoft Outlook Calendar and Microsoft To Do.

`work-log-auto` uses GitHub Actions and Microsoft Graph API to generate a daily Markdown work log from:
- Outlook Calendar events
- Completed Microsoft To Do tasks

It is designed for developers, freelancers, students, and small teams who want an auditable daily work record without manually writing logs every day.

## Features

- Scheduled daily sync with GitHub Actions
- Fetches Outlook Calendar events from Microsoft Graph
- Fetches completed Microsoft To Do tasks
- Converts UTC completion time to Taiwan time
- Sanitizes sensitive meeting titles, contacts, and confidential keywords
- Writes daily logs to `logs/YYYY-MM-DD.md`
- Supports signed Git commits through GPG in GitHub Actions

## Privacy

The project includes a sanitization layer to reduce accidental exposure of:
- email addresses
- confidential keywords
- private or confidential calendar events
- sensitive project/client names

You should review and customize `SENSITIVE_KEYWORDS` and `PROJECT_MAPPINGS` before using this in a public repository.

## Required GitHub Secrets

Set the following repository secrets:

- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_TENANT_ID`
- `MS_REFRESH_TOKEN`
- `GPG_PRIVATE_KEY`
- `GPG_PASSPHRASE`

## Microsoft Graph Permissions

This project uses:

- `Calendars.Read`
- `Tasks.Read`

## Usage

1. Fork this repository.
2. Configure Microsoft Graph OAuth credentials.
3. Add the required GitHub Secrets.
4. Enable GitHub Actions.
5. Run the workflow manually or wait for the scheduled sync.

## Example Output

```md
# 2026-06-02 Work Log

## Calendar
- **09:00**: Internal Task
- **14:30**: Infrastructure Upgrade

## ✅ To-Do Tasks
- ✅ **Completed**: Update deployment checklist (Tasks)


