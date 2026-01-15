# Workback + SharePoint + Azure Functions (Option A)

This zip adds a **server-backed** persistence layer and permissions so you can:
- Save all programme / tracker data to **SharePoint (Microsoft Lists)**
- Allow **mixed external users** to open the app and **only tick checkboxes** on pages you permit
- Create an **automatic weekly backup** with a date suffix

The frontend is your existing Vite React app (src/App.jsx etc.). The backend lives in `api/` and is built as **Azure Functions v4** using Microsoft Graph.

---

## 1) Create the SharePoint Lists

In the SharePoint site you want to store data in, create these Lists:

### A) `WorkbackState`
Store the entire app state (one item). Columns:
- **Title** (default) – set value `STATE`
- **DataJson** (Multiple lines of text)
- **LastSavedUtc** (Single line of text) *(optional)*

### B) `WorkbackPermissions`
One row per user per page that they can access. Columns:
- **Title** (default) *(optional)*
- **UserEmail** (Single line of text)
- **UserObjectId** (Single line of text) *(recommended)*
- **PageId** (Single line of text) *(must match `page.id` in your React state)*
- **Role** (Choice or text): `admin`, `editor`, `tickOnly`, `viewer`

> If a user has no matching permission rows, they get `viewer`.

### C) `WorkbackBackups`
One row per weekly backup. Columns:
- **Title** (default) – `backup-YYYY-MM-DD`
- **DataJson** (Multiple lines of text)
- **CreatedUtc** (Single line of text) *(optional)*

---

## 2) Azure AD App Registration (App-Only)

Create an app registration that the Functions backend will use (client credentials).

### API permissions (Microsoft Graph)
Grant **Application** permissions:
- `Sites.ReadWrite.All`

Then **Grant admin consent**.

Create a **client secret**.

---

## 3) Configure Azure Functions (environment variables)

Set these settings in your Function App Configuration (or `api/local.settings.json` for local dev):

- `AAD_TENANT_ID`
- `AAD_CLIENT_ID`
- `AAD_CLIENT_SECRET`
- `SP_SITE_ID` (Graph site id)
- `SP_LIST_STATE` = `WorkbackState`
- `SP_LIST_PERMISSIONS` = `WorkbackPermissions`
- `SP_LIST_BACKUPS` = `WorkbackBackups`

A template is provided at `api/local.settings.json.example`.

---

## 4) Local development

1. Start Functions API (port 5174 is expected by your `vite.config.js` proxy):

```bash
cd api
npm i
func start --port 5174
```

2. Start Vite:

```bash
npm i
npm run dev
```

---

## 5) What the backend exposes

- `GET /api/health`
- `GET /api/bootstrap` – returns saved state + your permissions
- `POST /api/save` – saves the full state (**editor/admin only**)
- `PATCH /api/rows/:rowId/tick` – updates only tick fields (**tickOnly/editor/admin**)
- `POST /api/chat` – placeholder (swap in Azure OpenAI)

Weekly backups run via a timer function:
- `backupWeekly` (Sunday 02:00 UTC) writes a new item to `WorkbackBackups`.

---

## 6) Frontend behavior

- The app still keeps a localStorage copy, but **also** saves to the server.
- On load it tries the server first (`/api/bootstrap`), falling back to localStorage.
- If a user is `tickOnly` on a page, attempts to edit anything except the 4 checkbox fields are blocked.

---

## 7) Next steps for Azure OpenAI

Replace `api/src/functions/chat.js` with your Azure OpenAI call. The current file is a safe placeholder so the UI stays green/neutral.

