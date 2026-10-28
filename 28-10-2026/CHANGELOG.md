# 2.6.37
- Reorder report columns, restore Secondary Reason, and add stable timestamp profile IDs with collision suffixes.

# 2.6.36
- Flag an unclaimed first-Experience company page through existing keyword screening and QA comments when DOM evidence is sufficient.

# 2.6.35
- Add the fixed screening keyword list in the Python backend and merge it with the extension's editable keyword list.

# 2.6.34
- Remove Open to Work detection and its delayed rescan/diagnostic logic. Keep the existing output field neutral (Not Available) for schema compatibility.

# 2.6.33
- Add a fallback for the visible Open to work status card in the bounded profile intro; keep it separate from generic page text.

# 2.6.32
- Align Open to Work detection with the QA Validator reference, require profile-owner context, log structured diagnostics, and rescan mutations that occur during capture.

# 2.6.31
- Port the QA Validator Open to Work detector: inspect image/frame marker attributes, then match the owner name in the photo label or nearby profile header.

# 2.6.30
- Detect the explicit #OpenToWork marker in the profile owner photo alt/accessibility label.
- Bump the unpacked extension version to identify this detection fix.

# 2.6.29
- Add optional per-profile Target Company validation across the captured Experience history.
- Record Company Check results and matching Experience entry numbers; apply the approved QA outcomes.

# 2.6.28
- Detect Open to Work from the visible profile card and compact `#OPENTOWORK` labels.
- Add a separate Experience's Geo column, derived from the first Experience entry and backfilled from existing Experience text.

# Changelog

## 1.2.0
- Removed automatic scrolling to the Experience section.
- Added a small status badge on individual profile pages (waiting, sending, saved, retrying, failed).
- Added bounded retries for missing Experience content and temporary service errors.
- Serialized capture submissions in the extension service worker to reduce overlapping requests.
- Made profile name optional; the normalized LinkedIn profile URL is the primary identifier.
- Kept existing Experience-section selection and Python fallback logic.
- Documented that content LinkedIn has not rendered cannot be captured without scrolling.

- Added optional row-aligned Report’s Job Title input to Bulk Opener, exact normalized comparison against the first Experience job title, JT Match Yes/No output, and report column migration/order after QA Comments.
- Open to Work detection now checks explicit visible badge, photo-frame, icon, and intro-card signals across LinkedIn profile layouts.
- Removed the Secondary Reason report column and output; migrated existing sheets to 14.4-point default row height with text wrapping disabled.
- Auto Mode toggle changes now persist immediately; profile and Activity scrolling is gated behind Auto Mode.
- Expanded Open to Work checks across visible badge/photo-frame attributes near the primary profile heading.
- Auto scrolling now requires both the current per-profile Auto Mode queue record and the saved toggle; turning Auto Mode off cancels pending scrolls.
## Unreleased
- Auto Mode now scrolls the profile’s actual scrollable ancestor and reports a successful scroll only when the scroll position changes.
- Simplified Auto Mode scrolling: every profile content script reads the saved toggle, scrolls one viewport exactly one second after load, and waits for that action before capture.
- Persisted Auto Mode on each bulk record and used that per-tab flag to open, scroll, and advance each profile consistently across worker restarts.
- Auto Mode scroll is now injected per completed profile tab; its one-second timer and repeat-until-Experience loop run in the page, independent of content-script mode lookup and worker suspension.
- Moved Auto Mode’s one-second scroll timer into each profile tab so it fires even when the extension background worker is suspended.
- Auto Mode executes a direct page scroll one second after each profile finishes loading, targeting Experience and falling back to a viewport scroll.
- Auto Mode now scrolls the active profile downward while searching, so lazy-loaded Experience sections are brought into view before capture timeout.
- Added an Auto Mode switch to Bulk Opener. It opens one active batch tab at a time, scrolls to Experience, waits for confirmed Excel export, closes the batch tab, and continues; failed captures are recorded before continuing, and failed Excel saves leave the tab open.
- Restricted Experience keyword screening to the first role, including its description, while retaining whole-profile keyword screening for the other supported fields.
- Tightened Open to Work detection to explicit accessibility labels on the profile intro badge/image; arbitrary page text and generic view markers no longer qualify.
- Repeated URLs with different emails now open and capture separately; the later email is retained on its own row with QA Status Disqualified and Primary Reason Duplicated Link.
- Kept each profile in Sending until its own Excel row is confirmed written; stale queued messages now resolve after the writer flushes.
- Prevented one slow capture upload from blocking the rest, used background timers for hidden-tab retries, and logged separate capture and Excel flush times.
- Reduced Excel flush batching and removed redundant full-workbook scans after writes; improved Open to Work badge detection in clickable profile-photo frames.
- Activity now stores only the latest visible entry time (such as `6mo • Edited`), without post text or follower counts.
- Improved profile location capture by reading the intro text before About and the location line before Contact info, including one-word city names.
- Improved profile location capture for locations shown without a comma.
- Kept existing workbook metadata when a later profile capture cannot read that field.
- Removed the broken Upload Excel input feature and its service endpoint.
- Reduced bulk tab delay to a 3–60 second range, with a 3 second default.
- Updated setup instructions to load the `new ext` folder.

- Moved Bulk Opener into its own tab with separate link/email inputs, preview, opening options, queue controls, and a live queue table.
- Added named worksheet creation for each bulk batch; batch captures write to that sheet in the selected workbook.

- Bulk batches now wait for each profile capture to finish before opening the next tab, with a timeout and visible per-profile queue status.

- Bulk batches can target an existing worksheet, use Profiles by default, or create a new worksheet explicitly. Duplicate checks are scoped to the selected worksheet.
- Bulk profile tabs now open on the selected timer, with later tabs staying in the background and focus remaining on the first profile; LinkedIn 404 redirects are recorded as invalid without stopping the queue.
- Added explicit QA outcomes for Dead Contact, PDC 1/2/3, missing tenure, and verified Open to Work; tightened Open to Work badge detection.
- Added Company and Job Title columns after First Experience and migrate existing worksheets while preserving First Experience.
- Refined current-employer parsing to reduce false Dual Employment flags and avoid writing Full-time as a Job Title; PDC 2/3 comments now show Current Tenure, and empty QA comments stay blank.
- Reduced per-profile Experience wait and DOM scan cost, added a text-based Experience fallback, bounded HTML fallback payloads, and added one automatic reload before recording Unable to Capture.
- Capture uploads now serialize one POST per profile; removed per-profile Excel status polling, automatic profile resubmission, and automatic tab reloads. Failed captures are marked once, the bulk page no longer polls service status every three seconds, and Activity metadata waits were shortened.

- Successful captures now report Saved immediately after journal acceptance, matching the earlier app while the background Excel writer finishes the workbook flush.

- Bulk link and email inputs remain editable during an active batch so the next list can be prepared; batch settings stay locked until the current batch finishes.

- Experience detection now waits up to 18 seconds and performs one additional DOM-only retry before the single capture submission; it no longer exits early when later profile sections appear.

- Hidden LinkedIn tabs now pause Experience detection and resume when visible, matching the existing extension workflow so lazy-loaded Experience is captured only after LinkedIn renders it.

- Restored the working extension’s 100 ms Experience checks and 500 ms broad heading scan cadence while the profile tab is visible.

- Keyword screening now uses only the first Experience item (including its description). Bulk links starting with http/https each keep their email index, even when wrapped lines continue the preceding URL.
- Dual employment now compares company names for every Present role; 404 links are Disqualified as Invalid JT Link; repeated links with different batch emails create a separate Internal Duplicate row for the later email.
