# Security

## Reporting a vulnerability

Use
[GitHub private vulnerability reporting](https://github.com/Maxim-Mazurok/msforms-api/security/advisories/new).
Do not include credentials, session cookies, private form content, or exploit
details in a public issue.

## Authentication data

`msforms-api` stores its persistent browser profile outside the project:

```text
~/.msforms-api/browser-profile
```

The package does not intentionally log or return authentication cookies,
anti-forgery tokens, or browser storage. Protect the browser profile as you
would any signed-in browser profile. Delete it to remove the local session.

Do not commit:

- `.npmrc` files containing registry credentials.
- `.env` files.
- Browser profiles or storage-state exports.
- Form payloads containing private respondent data.
- Debug logs containing request headers or upload URLs.

## Private API risk

Microsoft Forms respondent APIs are reverse-engineered and unsupported.
Endpoint behavior and authorization requirements can change without notice.
Use least-privilege accounts and review operations before submission.
