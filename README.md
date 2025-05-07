# Identify-Email-Addin

**Identify Email** brings professional-grade email forensics straight into Outlook’s reading pane, so every employee can spot threats in seconds—while security teams capture evidence with zero overhead.

---

## Plugin Overview

### Identify Email – Outlook Security Add-In  
A lightweight, client-side dashboard that surfaces objective trust signals for every message opened in Outlook.  
*No external servers • No data exfiltration • 100 % Office JS*

---

## What It Does

| Feature | Description |
|---------|-------------|
| **Verified-Sender Badging** | Compares the sender address to a curated whitelist of reputable corporate domains, your own organisation’s domain, and specific approved addresses.<br>• **Green badge** = verified<br>• **Red badge** = unverified |
| **Personal vs Business Detection** | Flags messages sent from consumer mailboxes (Gmail, Yahoo, etc.) with a caution badge to reinforce “business-only” communication policies. |
| **SPF / DKIM / DMARC Insight** | Parses raw headers and displays colour-coded **PASS / FAIL / N A** badges plus a concise summary—no need to open *Internet Headers*. |
| **Domain-Mismatch Warnings** | Cross-checks *From*, *Sender*, *Return-Path*, and DKIM-signing domains. Conflicts surface as a prominent ⚠️ alert. |
| **Link & Attachment Analysis** | • Decodes wrapped URLs (Safe Links, Proofpoint, Symantec, etc.)<br>• Counts total, unique, internal, and external domains<br>• Tallies attachments and badges the card accordingly |
| **Detailed Message Properties** | Shows conversation ID, message ID, subject, sender/recipient lists, timestamps, and more—each with a one-click *Copy to clipboard* button. |
| **Clean, Responsive UI** | Collapsible cards prevent clutter; long fields truncate with ellipsis + tooltip; theme toggle delivers dark-mode parity across Outlook desktop, web, and Mac. |

---

## How It Helps

* **Reduces phishing risk** by translating complex authentication data into plain-language pass/fail cues.  
* **Accelerates investigations**—instant, copy-ready identifiers for SOC and help-desk teams.  
* **Strengthens security awareness**—users learn *why* a message is safe or suspicious every time they read mail.  
* **Requires zero infrastructure**—everything runs inside the client; deployment is just an XML-manifest push.  
* **Delivers consistency**—same experience across all Outlook hosts, ensuring policy parity company-wide.

---

## Installation

### Freemium Plugin (Side-load)

> **Manifest URL**  
> ```
> https://phawandjian.github.io/identify-email-addin/VerifiedSenderAddin.xml
> ```

---

#### macOS (New & Legacy Outlook)

1. Open **Outlook for Mac**.  
2. In the **Home** ribbon, click **••• More options**.  
3. Choose **Get Add-ins**.  
4. Select **My add-ins** (left column).  
5. Under **Custom add-ins**, click **➕ Add a custom add-in → Add from URL…**  
   – Paste the manifest URL above and press **OK**.  
6. Confirm **Install/Continue** and wait for the green check-mark.  
7. *(Optional)* First launch & pin:  
   * Open any email → click the **Apps** (puzzle-piece) icon → choose **Verified Sender**.  
   * In the task-pane, click the **pin** icon to keep it docked.  

> **Tip:** If the add-in button doesn’t appear, quit and reopen Outlook once.

---

#### Windows  
*(Classic desktop & “New” Outlook)*

1. Open **Outlook for Windows**.  
2. Classic desktop: **File → Manage Add-ins** (opens a browser tab).  
   *New Outlook*: **Home → Get Add-ins**.  
3. In the Add-ins page, click **My add-ins**.  
4. Under **Custom add-ins**, choose **➕ Add a custom add-in → Add from URL…**  
   – Paste the manifest URL and confirm.  
5. Accept the security prompt → **Install** → wait for the green check-mark.  
6. **Restart Outlook** (ensures the ribbon button loads).  
7. *(Optional)* First launch & pin:  
   * Open any message → **Apps** icon → **Verified Sender**.  
   * Click the **pin** in the task-pane to keep it docked.

---

### Troubleshooting

* **Add-in doesn’t appear after install**  
  * Quit and reopen Outlook.  
  * Ensure you’re signed into the same Microsoft 365 mailbox in both Outlook and the browser tab.  
  * Verify the manifest URL is reachable (paste it in a browser—should download the XML).

* **Task pane closes every time**  
  * Click the **pin** button inside the pane so it stays docked.

* **“Custom add-ins” option greyed out**  
  * Your admin may have disabled sideloading. Ask IT to enable **Add from URL** or deploy the add-in centrally via the Microsoft 365 admin center.

---
Un-install / Remove the IdentifyEmail (“Verified Sender”) add-in

Tip: If your IT department deployed the add-in centrally, you will not see a Remove option—ask your Microsoft 365 admin to withdraw the deployment from the admin center.

macOS (both “New” & “Legacy” Outlook)
	1.	Open Outlook for Mac.
	2.	In the Home ribbon choose … More options ➜ Get Add-ins.
	3.	In the Add-ins window select My add-ins on the left.
	4.	Scroll to Custom add-ins. Locate Verified Sender.
	5.	Click the ⋯ (three-dot) menu on its tile and choose Remove → Remove again to confirm.  ￼
	6.	Close the Add-ins window.
	7.	If the task-pane pin is still visible in messages, quit and reopen Outlook once to clear the cached button.

Windows – Classic desktop (Microsoft 365 Current Channel)
	1.	In Outlook choose File ➜ Manage Add-ins (opens My Add-ins in your browser).
	2.	Under My add-ins find Verified Sender.
	3.	Click the ⋯ menu on the tile → Remove → Remove to confirm.  ￼
	4.	Return to Outlook and restart the app so the ribbon refreshes.

Windows – New Outlook (web-tech)
	1.	From the Home ribbon click Get Add-ins.
	2.	Choose My add-ins on the left.
	3.	Under Custom add-ins locate Verified Sender, open the ⋯ menu and pick Remove → Remove.  ￼
	4.	Close the dialog; the task-pane icon disappears immediately (restart if it lingers).
---

### License

MIT License

Copyright (c) 2025 Panos Hawandjian

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the “Software”), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
