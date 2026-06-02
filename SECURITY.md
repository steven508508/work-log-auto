# Security Policy

## Sensitive Data

This project may interact with calendar events, Microsoft To Do tasks, OAuth credentials, and GitHub Actions secrets.

Never commit:

- Microsoft Graph refresh tokens
- client secrets
- private calendar data
- private task data
- GPG private keys
- personal work logs containing sensitive information

## Recommended Setup

Use GitHub Actions Secrets for all credentials:

- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `MS_TENANT_ID`
- `MS_REFRESH_TOKEN`
- `GPG_PRIVATE_KEY`
- `GPG_PASSPHRASE`

## Reporting a Vulnerability

Please open a private security advisory or contact the maintainer if you discover a vulnerability related to token handling, log sanitization, or GitHub Actions secrets.
