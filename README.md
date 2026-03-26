# Zendesk Queue Analyser — Google Sheets

> A free, open-source Google Sheets tool for Zendesk admins to measure omnichannel routing queue performance — no coding required.

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform: Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green.svg)

---

## What it does

This tool connects directly to the **Zendesk Queue Events API** and pulls data into a structured Google Sheet. It is aimed at ops managers and Zendesk admins who want to understand how their omnichannel routing queues are performing without needing to export data manually or build a custom integration.

It gives you:

- **⏱️ Wait Times** — how long tickets spent waiting in a queue before an agent accepted them, broken down overall, by queue, by channel, and by day — with charts
- **📊 Queue Size** — how many tickets were sitting in queue at any point, including peak and hourly average, broken down the same way
- **🗂️ Raw Events** — the underlying event-level data for anyone who wants to dig deeper or build their own analysis
- **📋 Queues** — a reference list of all your active queues and their IDs

Everything runs inside Google Sheets via Apps Script. No third-party services are involved. Your data only ever travels between Google and your own Zendesk account.

---

## Requirements

| Requirement | Detail |
|---|---|
| **Zendesk plan** | Suite Professional, Enterprise, or Enterprise Plus — or Support Professional/Enterprise |
| **Zendesk role** | Admin access |
| **Feature** | Omnichannel routing must be enabled with at least one custom queue configured |
| **Data retention** | Zendesk only retains queue event data for the past 90 days |
| **Google account** | Required to use Google Sheets and Apps Script |

---

## Installation

### Option A — Copy the template spreadsheet (easiest)

1. Open the [template spreadsheet](https://docs.google.com/spreadsheets/d/1AbGFur1z8xwKuHEQHoPLE72hPpa91JqHZX-nI2HswwI/edit?usp=sharing) 
2. Go to **File → Make a copy**
3. The script is already embedded — skip straight to [Setup](#setup)

### Option B — Add the script to your own spreadsheet

1. Open any Google Sheet (or create a new blank one)
2. Go to **Extensions → Apps Script**
3. Delete any placeholder code in the editor
4. You need to create **three script files** and one **manifest file**. For each:
   - Click the **+** button next to "Files" in the left panel to add a new file
   - Name it exactly as shown below
   - Paste in the contents from this repository

| File to create | Source in this repo | Notes |
|---|---|---|
| `Code.gs` | [`Code.gs`](./Code.gs) | The default file — rename it if needed |
| `Sidebar.html` | [`Sidebar.html`](./Sidebar.html) | Add as an HTML file |
| `AnalysisDialog.html` | [`AnalysisDialog.html`](./AnalysisDialog.html) | Add as an HTML file |

5. Update the `appsscript.json` manifest — see [Updating the manifest](#updating-the-appsscriptjson-manifest) below
6. Click **Save** (Ctrl/Cmd + S)
7. Close the Apps Script tab and reload your spreadsheet
8. A **🎯 Queue Analyser** menu will appear in the menu bar

---

## Updating the appsscript.json manifest

The `appsscript.json` file controls which Google permissions the script requests. By default, Apps Script hides this file. Here is how to update it:

1. In the Apps Script editor, click the **⚙️ Project Settings** gear icon in the left sidebar
2. Scroll down and check **"Show 'appsscript.json' manifest file in editor"**
3. Go back to the **Editor** (the `<>` icon in the left sidebar)
4. You will now see `appsscript.json` listed in the Files panel — click it
5. Replace the entire contents with the contents of [`appsscript.json`](./appsscript.json) from this repo
6. Save

The manifest sets three OAuth scopes that the script needs to function:

| Scope | Why it's needed |
|---|---|
| `spreadsheets.currentonly` | Read and write data to the active spreadsheet |
| `script.external_request` | Make HTTP requests to the Zendesk API |
| `script.container.ui` | Show the sidebar and modal dialogs |

Without the correct manifest, the script may fail silently or prompt for incorrect permissions.

---

## Setup

### Step 1 — Enter your credentials

1. In your spreadsheet, click **🎯 Queue Analyser → 🔧 Setup credentials**
2. A sidebar will open on the right. Enter:
   - **Subdomain** — just the subdomain portion of your Zendesk URL. If your Zendesk is at `yourcompany.zendesk.com`, enter `yourcompany`
   - **Admin email** — the email address you use to log in to Zendesk as an admin
   - **API token** — see [Finding your API token](#finding-your-api-token) below
3. Click **Test & save connection**
4. The tool will make a test call to your Zendesk account to verify the details before saving. If it fails, you will see a specific error message explaining what went wrong

Your credentials are stored in Google Apps Script User Properties — they are tied to your Google account, not the spreadsheet itself, and are not visible to anyone else who has access to the sheet.

### Step 2 — Load your queues

1. Click **🎯 Queue Analyser → 🔄 Load my queues**
2. The tool fetches all active queues from your Zendesk account and writes them to the **📋 Queues** sheet
3. This step is required before running an analysis — it builds the list of queues you can choose from

### Step 3 — Run an analysis

1. Click **🎯 Queue Analyser → ▶️ Run analysis**
2. A dialog will open. Configure your analysis:
   - **Date range** — select From and To dates using the day/month/year dropdowns. Maximum 90 days
   - **Queues** — tick specific queues to include, or leave all unchecked to analyse every queue
   - **Channels** — toggle Messaging, Talk, and/or Support. At least one must be selected
3. Click **▶️ Run analysis**
4. The tool fetches all matching events from Zendesk, paginating through the full dataset automatically. This may take 15–60 seconds for large date ranges
5. When complete, results appear in the **⏱️ Wait Times**, **📊 Queue Size**, and **🗂️ Raw Events** sheets

---

## Finding your API token

1. In Zendesk, go to **Admin Center** (the grid icon in the left nav)
2. Click **Apps & integrations → APIs → Zendesk API**
3. Click the **Settings** tab and make sure Token access is enabled
4. Click the **API tokens** tab, then **Add API token**
5. Give it a descriptive name (e.g. "Queue Analyser — Google Sheets") and click **Create**
6. Copy the token immediately — Zendesk will not show it again
7. Paste it into the Queue Analyser credentials sidebar

> ⚠️ Treat your API token like a password. Anyone who has it can read and write data on your Zendesk account using your admin permissions. Do not share it or commit it to version control.

---

## Sheets explained

| Sheet | What's in it |
|---|---|
| ⚙️ Setup | Getting started guide and instructions |
| 📋 Queues | All active queues — ID, name, description, priority, groups |
| ⏱️ Wait Times | Wait time analysis with summary, breakdowns, and daily chart |
| 📊 Queue Size | Queue size analysis with summary, breakdowns, and daily chart |
| 🗂️ Raw Events | Every individual queue event returned by the API |

---

## Understanding the metrics

### What is a queue event?

Every time a ticket enters or leaves a queue, Zendesk creates an event record. A single ticket can generate multiple events — for example, if it enters a queue, gets transferred to another queue, and is then accepted by an agent, that creates at least three events. The raw events sheet shows all of these individually.

There are three event types:

| Event type | What it means |
|---|---|
| `INBOUND` | A ticket entered the queue |
| `OUTBOUND` | A ticket left the queue (accepted, transferred, expired, etc.) |
| `ROUTING_STATE_CHANGE` | A ticket was offered to an agent but not yet accepted |

### Wait Times (⏱️)

Wait time is only calculated for **OUTBOUND events** that have both a `start_queue_time` and an `end_queue_time`. This represents the full time a ticket spent waiting in the queue before it left.

```
wait time = end_queue_time − start_queue_time
```

The tool reports:
- **Average wait time** — mean across all qualifying events
- **Maximum wait time** — the longest any single ticket waited
- **Minimum wait time** — the shortest any single ticket waited

These are broken down overall, by queue name, by channel, and by day.

The daily chart shows average and maximum wait time in minutes, with ticket volume as a bar chart on a second axis — so you can see whether wait time spikes correlate with volume spikes.

### Queue Size (📊)

**Minimum and maximum** are taken from `current_channel_queue_size`, which is a snapshot of how many tickets were in the queue at the moment each event occurred. Zeros are excluded from the minimum (a size of 0 just means the queue briefly emptied).

**Average queue size** uses `hourly_channel_queue_size_seconds`. This field accumulates signed seconds — positive for INBOUND events, negative for OUTBOUND. The correct way to calculate an average is:

1. Group all events into hour-long buckets (per queue, per channel)
2. Sum the `hourly_channel_queue_size_seconds` values within each bucket
3. Floor the total at 0 (a negative total means more left than entered that hour)
4. Divide each bucket total by 3600 to get the average number of tickets in queue during that hour
5. Average those hourly values to get the daily or overall average

Simply averaging the raw per-event values will produce near-zero results because the positive and negative values cancel out. The tool handles this correctly.

### Ticket volume

The "Tickets entering queue" figure counts unique `ticket_id` values from INBOUND events per day. This is different from the event count — one ticket can create many events.

---

## Known limitations

### 90-day data retention
Zendesk only retains queue event data for the past 90 days. You cannot analyse periods older than this, regardless of your plan.

### Queue size on the Support channel
In testing, `current_channel_queue_size` and `hourly_channel_queue_size_seconds` consistently return 0 for Support channel events, even when wait times for the same events are calculated correctly. This appears to be a Zendesk product behaviour rather than a bug in this tool — Support tickets are email-based and may not be tracked with real-time queue size metrics in the same way as Messaging and Talk.

If you see all zeros in the Queue Size sheet when analysing Support data, this is likely the cause. Queue size metrics are most meaningful for Messaging and Talk channels. This limitation is currently being investigated with the Zendesk developer community.

### One ticket, multiple events
A ticket that moves between queues will appear multiple times in the raw events and will contribute multiple wait times to the averages. The "tickets processed" count in the Wait Times sheet counts OUTBOUND events with valid wait times, not unique tickets.

### Rate limiting
The Zendesk Queue Events API allows 30 requests per minute. For large date ranges across multiple queues and channels, the tool may briefly hit this limit. If you see a rate limit error, wait a minute and try again. Running analyses for shorter date ranges or fewer queues at a time will reduce the likelihood of hitting the limit.

### Apps Script execution time
Google Apps Script has a 6-minute execution time limit. For very large datasets (many queues, all three channels, over a long period), you may hit this. If the script times out, try narrowing the date range or selecting fewer queues/channels.

---

## Troubleshooting

| Problem | Likely cause | What to do |
|---|---|---|
| "🎯 Queue Analyser" menu not visible | Script not yet authorised, or page not reloaded | Reload the spreadsheet. If prompted, authorise the script |
| Authentication failed (401) | Wrong email address or API token | Re-enter credentials in Setup credentials |
| Access denied (403) | Account is not an admin, or omnichannel routing is not enabled | Check your Zendesk role and routing settings in Admin Center |
| Subdomain not found (404) | Subdomain entered incorrectly | Check the subdomain — enter only `yourcompany`, not `yourcompany.zendesk.com` |
| No queues found | No active custom queues exist | Create at least one queue in Zendesk's omnichannel routing settings |
| No events found | Date range too narrow, wrong channel, or no activity | Widen the date range, check you have selected the right channel(s) |
| Queue Size sheet shows all zeros | Support channel limitation (see above) | Re-run with Messaging or Talk selected |
| Rate limit error (429) | Too many API requests in a short period | Wait 1 minute and try again |
| Ticket IDs showing as dates in Raw Events | Google Sheets auto-formatting | The script forces the Ticket ID column to plain text — if you see this, it may be a one-time display glitch. Try re-running the analysis |
| Script times out | Dataset too large for 6-minute limit | Narrow the date range or select fewer queues and channels |

---

## Privacy & security

- Your Zendesk credentials (subdomain, email, API token) are stored in **Google Apps Script User Properties**. These are tied to your individual Google account and are never stored in the spreadsheet itself. Other users who have access to the spreadsheet cannot see your credentials.
- All API requests go directly from Google's servers to your Zendesk subdomain. No data passes through any third-party service.
- The script requests only the minimum permissions needed to function (see [Updating the manifest](#updating-the-appsscriptjson-manifest)).
- The API token is stored in plaintext in User Properties. For production use in a shared environment, consider creating a dedicated Zendesk service account with read-only API access and using its token.

---

## Contributing

Issues and pull requests are welcome. If you find a bug or have a feature suggestion, please open an issue first so we can discuss it before you invest time in a PR.

---

## Roadmap

A native Zendesk app version of this tool is in development. See [`NATIVE_APP_BLUEPRINT.md`](./native_app_blueprint.md) for the technical specification.

---

## License

MIT — see [LICENSE](./LICENSE) for details.

---

## Built with

- [Zendesk Queue Events API](https://developer.zendesk.com/api-reference/agent-availability/queue-events/queue_events/)
- [Zendesk Queues API](https://developer.zendesk.com/api-reference/agent-availability/omnichannel_routing_queues/omnichannel_routing_queues/)
- [Google Apps Script](https://developers.google.com/apps-script)
