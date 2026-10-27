# LinkedIn Experience Copier (existing Chrome session)

This project uses a Chrome extension with the LinkedIn session you already have open and a local Python service. It does **not** launch a separate browser or ask you to log in again.

## What it captures

- Profile name
- First Experience, with Company and Job Title in separate columns immediately after it
- Canonical LinkedIn profile URL
- Profile location shown near the name, kept separate from the current Experience location (`Experience's Geo`)
- The displayed time of the latest visible Activity entry (for example, `6mo • Edited`); follower counts and post text are not stored as activity
- The **whole visible Experience section** (including multiple roles where shown)
- It does not click or expand “See more”.

The extension waits for the Experience heading and its text to appear and stabilize without scrolling for that section, including in tabs opened in the background. To read Activity, it may briefly scroll to the visible Activity card and then restores the original page position. It works while you switch to other tabs, displays an on-page status badge, retries incomplete captures, and serializes capture submissions to the Python service. The LinkedIn profile URL is the primary identifier; a missing profile name does not block saving. If the browser cannot locate the section, Python tries a fallback extraction from the page HTML. Python deduplicates by profile URL within each worksheet and writes to `profiles.xlsx` by default. Bulk batches write to the selected worksheet, defaulting to `Profiles`; individual captures also go to `Profiles`.

## QA reporting

QA Status applies the parsed first-experience and verified Open to Work rules: Dead Contact, missing tenure, PDC 2/3, and Open to Work are Disqualified; PDC 1 is Qualified PDC 1 with “No Fatal Info Found”; PDC 2/3 comments show current tenure. QA Comments stay blank when no rule matches. Open to Work is marked Yes only when a badge or explicit intro-card label is detected; otherwise it is Not Available.

## Setup (Windows)

1. Extract the ZIP to a normal folder.
2. Open the folder in VS Code.
3. Double-click `start_service.bat` or run it in the VS Code terminal. Keep the terminal open.
4. In Chrome, open `chrome://extensions`.
5. Turn on **Developer mode**.
6. Click **Load unpacked** and select the project's `new ext` folder.
7. Pin **LinkedIn Experience Copier - New** to the toolbar.
8. Open LinkedIn in your existing Chrome profile and make sure you are logged in.
9. Click the extension and press **Start**.
10. Open **Bulk Opener in new tab** from the extension popup to paste profiles, choose an existing worksheet (Profiles is the default) or explicitly create a new one, set the tab delay and choose new or reused tabs. The first profile opens normally; later tabs open in the background without changing the focused profile. Duplicate checks apply only within the destination worksheet.
11. Browse individual LinkedIn profile pages normally. The extension attempts to capture each profile while Start is enabled.
12. Press **Stop** when finished. The workbook is saved as `profiles.xlsx` in the project folder.

## Notes

- Only individual URLs under `linkedin.com/in/` are captured. Company pages and non-LinkedIn tabs are ignored.
- The active profile is attempted immediately when Start is clicked. Existing inactive profile tabs are enabled; tabs you navigate to while capture is on are captured in the background. Chrome may throttle background tabs, and LinkedIn may not render below-the-fold sections until they are naturally loaded. Because this version does not scroll, a profile whose Experience section is not present in the DOM may need a retry or manual revisit.
- A small badge appears on individual profile pages with statuses such as waiting, sending, saved, retrying, or failed. If a capture fails after automatic retries, revisit the profile or click Start again on that profile. Check the extension console and Python terminal for errors.
- No “See more” button is clicked. Only content rendered/available in the Experience section is collected.
- A 90% success rate is a target, not a guarantee until tested against a representative batch of profiles and real network conditions. No auto-scroll is performed; therefore, content LinkedIn has not loaded into the page cannot be captured until LinkedIn renders it.
- The Python service listens only on `127.0.0.1`. The extension and service share a local token in their config files; keep the project private.
- Use the tool responsibly and follow LinkedIn's terms and applicable privacy rules.



