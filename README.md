# Install Tracker Tool in Outlook

Add a button in **classic Outlook for Windows** that turns the email you’re reading into a task in [Tracker Tool](https://tracker.seoandweb.co.uk).

---

## Before you start

You need:

1. **Classic Outlook** on Windows (the desktop app you already use — not only Outlook on the web).
2. A **Tracker Tool** login (the same email and password you use on the website).
3. The **Tracker Tool Outlook** package from your IT contact or Desktop folder (`outlook-tracker-addin`).

Someone in the team must already have published the add-in files on a secure (HTTPS) web address and put that address into `manifest.xml`. If that is not done yet, ask them before you install — Outlook will not load the add-in from a plain folder of HTML files alone.

---

## Install (one PC)

### Step 1 — Open My Add-ins

1. Open **Outlook**.
2. Open any email so you are reading a message.
3. On the ribbon, click **Get Add-ins**  
   (on some builds this is under **Home** → **Get Add-ins**, or **…** → **Get Add-ins**).
4. Open the **My add-ins** tab.

### Step 2 — Add from file

1. Near the bottom, open **Add a custom add-in**.
2. Choose **Add from file…**
3. Select the file **`manifest.xml`** from the Tracker Tool Outlook package.
4. Confirm any security prompts.

### Step 3 — Restart Outlook

Close Outlook completely (including the system tray icon if it stays open), then open it again.

### Step 4 — Check it worked

1. Open any email.
2. Look on the ribbon for **Add to Tracker** (often in a **Tracker Tool** group).
3. Click it. A side panel should open with the email subject and a sign-in form.

If you do not see the button, see **Troubleshooting** below.

---

## Install for everyone (office / Microsoft 365 admin)

If you want the button for the whole team without each person adding a file:

1. Sign in to the [Microsoft 365 admin center](https://admin.microsoft.com).
2. Go to **Settings** → **Integrated apps**.
3. Choose to **Upload custom apps** / deploy an Outlook add-in.
4. Upload **`manifest.xml`**.
5. Assign it to the right users or groups.

After Microsoft finishes deploying it (often within a few hours, sometimes sooner), each person should restart Outlook and look for **Add to Tracker**.

---

## How to use it

1. Open the email you want to turn into a task.
2. Click **Add to Tracker**.
3. Sign in with your Tracker Tool email and password (first time only on that PC).
4. Check the title and notes (taken from the email). Adjust client, requester, or status if needed.
5. Click **Create Task**.
6. Open [Tracker Tool](https://tracker.seoandweb.co.uk) — the new task should appear under Active.

---

## Troubleshooting

**I can’t find Get Add-ins / My add-ins**  
Use classic Outlook desktop. In newer builds, try the search box at the top of Outlook and type `add-ins`, or check **File** → **Manage Add-ins**.

**I added the file but there is no button**  
Restart Outlook fully. Confirm you opened a **received message** (the add-in is for reading mail, not for composing a new one). Ask IT that the web address inside `manifest.xml` is live over HTTPS.

**The panel opens but Create Task fails**  
Sign in again. If it still fails, contact whoever set up Tracker Tool / this add-in.

**Outlook says the add-in can’t be loaded**  
Usually the HTTPS site that hosts the add-in files is down, blocked by the network, or the address in `manifest.xml` is wrong. That is an IT fix, not something you change in Outlook settings.

**I’m on a Mac or only Outlook on the web**  
This package is aimed at **classic Outlook for Windows**. Ask if a web or Mac option is planned.

---

## What you’ll see in the package

| File / folder | What it’s for |
|---------------|----------------|
| `manifest.xml` | The file you pick in **Add from file** |
| `src/` | The side-panel screens (hosted on HTTPS for Outlook) |
| `assets/` | Icons on the ribbon |
| `DEVELOPERS.md` | Technical notes for people who host or update the add-in |

You only need `manifest.xml` for the install steps above. The rest is for whoever hosts and maintains Tracker Tool.

---

## Need help?

Ask your Tracker Tool / IT contact, or check the live app at https://tracker.seoandweb.co.uk.
