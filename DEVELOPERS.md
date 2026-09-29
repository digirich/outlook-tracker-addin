# Tracker Tool Outlook add-in — developer notes

End users should follow **README.md** (install guide). This file is for people hosting or changing the add-in.

## Stack

- Classic Outlook for Windows Office Add-in (Mail Read, VersionOverrides)
- Office.js task pane + ribbon **Add to Tracker**
- Tracker: https://tracker.seoandweb.co.uk  
- Supabase project ref: `zsabmwwtflsonkjnufxc`
- Create-task API: `POST …/functions/v1/create-task` (Bearer user JWT)
- Existing: `POST …/functions/v1/parse-email`

## Configure

Edit `src/config.js`:

- `supabaseUrl`
- `supabaseAnonKey` (anon only — never service_role)
- `createTaskUrl`
- `parseEmailUrl`

Update every `https://localhost:3000` URL in `manifest.xml` (and `AppDomains`) to your HTTPS host.

## Local serve

```bash
npm install
npx office-addin-dev-certs install
npm run serve
```

Outlook requires HTTPS for SourceLocation URLs.

## Sideload options (tech)

- Trusted Add-in Catalog (folder/UNC) — catalog holds **manifest only**; assets still need HTTPS
- Add from file (My add-ins)
- M365 admin → Integrated apps → upload manifest

## Payload (Create Task)

`title`, `description`, `email_from`, `email_to`, `email_date`, optional client/requester/status, conversation/message ids.

## Known gaps

- create-task edge function may still be deploying
- client field is free-text until clients list is wired
- icons are placeholders
