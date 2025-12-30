# Gold Standard Medical Group — Telehealth SMS Sender (Google Apps Script)

A small Apps Script that scans today’s appointment workbook, finds telehealth visits, and texts patients a doxy.me link via ClickSend.



## What it does

* Looks in Google Drive for a Google Sheet named like **`MM-dd-yyyy appointment`** (e.g., `12-29-2025 appointment`).
* Iterates each provider tab **until a stop-tab is reached** (see *Stop tabs* below).
* For rows marked telehealth (column **`O/T`** begins with **`T`**), it:

  * Parses the **`Time`** value and treats the provider’s shift as starting **5 minutes before** the earliest telehealth appointment.
  * Skips rows with missing phone numbers, **`DOB` = `*`**, or where **`Follow Up`** implies the patient was already seen/rescheduled/cancelled.
  * Sends an SMS via ClickSend with the provider’s doxy.me URL.
  * Appends a timestamp to the **`Link Sent`** column.
* Logs any failures to a **`Messaging Errors`** sheet.

> Heads‑up: the script **does not de‑duplicate** by default. If you run it multiple times, it will text those patients multiple times unless your sheet logic prevents it.



## Prerequisites

* A Google Workspace or Gmail account.
* A daily appointment workbook saved as a **Google Sheet** (not Excel) with provider tabs.
* Column headers present on each provider tab (first row):

  * `O/T`, `Time`, `Phone #`, `Link Sent`, `Appt With`, `Follow Up`, `DOB`
* A ClickSend account with an API Username and API Key.
* Provider doxy.me rooms following this pattern: `goldstandard.doxy.me/<firstname>gsmg` (with a special case for "Vivian": `viviangsmg1`).

---

## Install (one‑time)

1. Open **Google Apps Script** (from the appointments Google Sheet: *Extensions → Apps Script*).
2. Paste the entire script into `Code.gs` and save.
3. In the script, review the **`CONFIG`** section:

   * `CLICKSEND_USERNAME` – your ClickSend login (usually email).
   * `CLICKSEND_API_KEY` – your ClickSend API key.
   * `TIME_WINDOW_MINUTES` – how far before the first appointment a provider’s shift is considered “started” (default 5).
   * `STOP_AT_TABS` – list of tab names that mark where to stop processing.

> Note: `SPREADSHEET_ID` is present in `CONFIG` but **not used** by the current code. The script searches Drive for today’s file by name.

4. **Authorize** the script when prompted (Drive, Sheets, and external (UrlFetch) access are required).

---

## Set up time‑based triggers

Create four daily triggers so patients are reminded around common start times.

1. In Apps Script, open **Triggers** (clock icon).
2. Click **+ Add Trigger**.
3. Choose function: **`sendAllProvidersSMS`**.
4. Deployment: **Head**.
5. Event source: **Time‑driven** → **Day timer**.
6. Create triggers for these times: **9:00 AM**, **12:00 PM**, **1:00 PM**, **4:00 PM** (your local time zone).

---

## Test it

* In Apps Script, select **`sendAllProvidersSMS`** and click **Run**.
* Watch the **Logs** for messages like:

  * found today’s sheet
  * provider tabs processed
  * SMS success/HTTP error details
* Check your sheet’s **`Link Sent`** column for appended timestamps.
* If errors occur, open the **`Messaging Errors`** sheet for details.

---

## Stop tabs

The script will process tabs **until** it hits any name that matches (case‑insensitive, substring match) one of these:

```
Therapist A, Therapist B, Therapist C, Therapist D,
Therapist E, Therapist F, Therapist G, Therapist H, Therapist I
```


## How the script decides who to text

* **Telehealth filter:** `O/T` starts with `T` (so `T`, `TF`, `TE`, etc.).
* **Shift start check:** It will only send after the provider’s shift start = earliest telehealth appointment **minus** `TIME_WINDOW_MINUTES`.
* **Skip rules:**

  * No phone number → skip
  * `DOB` is `*` → skip (incomplete intake)
  * `Follow Up` suggests already handled (e.g., contains a date, `cancel`, `resch`, `no f/u`, `seen`, `RS`, `2 wks`, `1 mo`, etc.) → skip

> Time parsing: the script handles Date objects and strings like `9:00`, `9:00 AM`, `14:30`. If no AM/PM, hours `1–7` are treated as PM.


## Error logging

Failures are appended to a **`Messaging Errors`** sheet (created automatically if missing) with:

`Timestamp, Provider Tab, Patient Name, Appointment Time, Phone Number, Error Message`

If there’s no active spreadsheet context, a standalone **`Messaging Errors`** spreadsheet is created in Drive.

