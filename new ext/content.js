(() => {
  if (window.__liExperienceCopierInstalled) return;
  window.__liExperienceCopierInstalled = true;

  let enabled = false;
  let autoMode = false;
  let queuedAutoMode = false;
  let initialAutoScrollDone = false;
  let autoScrollTimer = null;
  let lastCapturedUrl = "";
  let lastSeenUrl = location.href;
  let urlSeenAt = Date.now();   // when this tab first showed the current profile
  let failedUrl = "";
  let timer = null;
  let inProgress = false;
  let pausedHidden = false;

  // Last Experience section we looked at. Used to avoid saving the PREVIOUS
  // profile's content under the NEW profile's URL while LinkedIn swaps the page.
  let lastSectionUrl = "";
  let lastSectionText = "";
  let headingEl = null;
  let lastBroadScan = 0;
  let lastLocateScan = 0;
  let lastTextFallback = 0;
  let emptyUrl = "";   // no automatic retry after the one capture attempt

  // The whole-page HTML fallback is up to 1.5 MB per attempt (sent twice when a profile has no Experience).
  // The text fallback above already covers the cases it used to help with. Set true to re-enable.
  const SEND_PAGE_HTML_FALLBACK = false;

  const EXPERIENCE_WAIT_MS = 18000; // wait for Experience to load before one capture submission
  const EXPERIENCE_RETRY_WAIT_MS = 6000; // final DOM-only retry; no extra Excel request
  const POLL_MS = 100;              // match the working extension; detection runs only while visible
  const STABLE_MS = 400;            // section text must stay unchanged this long
  const STALE_GUARD_MS = 2500;      // after navigation, ignore content identical to the previous profile
  const START_DELAY_MS = 150;       // delay before a capture starts (waiting happens inside the capture)
  const URL_POLL_MS = 250;          // SPA URL-change detection
  const MIN_EXTRA_CHARS = 6;        // a card needs at least this much text beyond its heading (so tiny cards count)
  // Headings of OTHER profile sections. The Experience container must never include these.
  const SECTION_HEADINGS = new Set(["about", "activity", "highlights", "analytics", "featured", "education",
    "skills", "languages", "interests", "projects", "courses", "recommendations", "honors & awards",
    "licenses & certifications", "volunteering", "publications", "causes", "patents", "test scores",
    "organizations", "people also viewed", "people you may know", "explore premium profiles",
    "more profiles for you"]);

  // Shown as a BLINKING red banner when the Experience could not be captured properly.
  const ALERT = {
    empty: "Couldn't capture the Experience - please capture manually",
    too_broad: "Captured extra profile text (trimmed) - please check this profile manually",
    unparsed: "Couldn't read the first experience - please capture manually",
    failed: "Couldn't save this profile - please capture manually"
  };

  const PROFILE_RE = /^https:\/\/(?:[a-z]{1,3}\.)?linkedin\.com\/in\/[^/?#]+/i;
  const cleanUrl = () => location.href.split("?")[0].split("#")[0].replace(/\/+$/, "");
  const isProfile = () => PROFILE_RE.test(location.href);
  function visibleProfileMetadata() {
    const main = document.querySelector("main") || document.body;
    const text = (main.innerText || main.textContent || "");
    const lines = text.split(/\n+/).map(x => x.trim()).filter(Boolean);
    const h1 = main.querySelector("h1");

    // Restrict headline/location/status checks to the profile intro card, not the
    // entire page (which can contain generic recommendations mentioning Open to Work).
    let intro = h1;
    for (let i = 0; intro && i < 9 && intro.parentElement; i++) {
      const candidate = intro.parentElement;
      const t = (candidate.innerText || candidate.textContent || "").trim();
      if (t.length > 0 && t.length < 4000 && candidate.querySelector("h1")) intro = candidate;
      else break;
    }
    const introText = (intro?.innerText || intro?.textContent || "").trim();
    let geo = "";
    const isGeoNoise = value => !value || /^[·•,]+$/.test(value.trim()) ||
      /^(?:location|contact info|open to work|about|activity|connections?|followers?)$/i.test(value.trim()) ||
      /\bconnections?\b|\bfollowers?\b|\bopen to work\b|\b[123](?:st|nd|rd)\b/i.test(value);
    if (h1) {
      const name = (h1.innerText || h1.textContent || "").trim().replace(/\s+/g, " ");
      let n = h1.parentElement;
      for (let depth = 0; n && depth < 16 && n !== main; depth++, n = n.parentElement) {
        const candidateLines = (n.innerText || n.textContent || "").split(/\n+/).map(x => x.trim()).filter(Boolean);
        const idx = candidateLines.findIndex(x => {
          const line = x.replace(/\s+/g, " ");
          return line === name || line.startsWith(name + " ");
        });
        if (idx < 0) continue;
        const nearby = candidateLines.slice(idx + 1, idx + 15);
        const contactIndex = nearby.findIndex(x => /contact info/i.test(x));
        if (contactIndex >= 0) {
          const contactLine = nearby[contactIndex].split(/\s*[·•]\s*contact info/i)[0].trim();
          let beforeContact = contactLine && !isGeoNoise(contactLine) ? contactLine : "";
          for (let j = contactIndex - 1; !beforeContact && j >= 0; j--) {
            if (!isGeoNoise(nearby[j])) beforeContact = nearby[j];
          }
          if (beforeContact) { geo = beforeContact; break; }
        }
        geo = nearby.find(x => /,|\b(?:area|region|metropolitan)\b/i.test(x) && !isGeoNoise(x)) || "";
        if (geo) break;
      }
    }
    // LinkedIn can place the location outside the h1's immediate wrapper. In the
    // intro card it sits immediately before Contact info, so use that relationship
    // as a fallback even for a single-word city with no comma or "Area" suffix.
    if (!geo) {
      const aboutIndex = lines.findIndex(x => /^about$/i.test(x));
      const introLines = lines.slice(0, aboutIndex >= 0 ? aboutIndex : Math.min(lines.length, 50));
      const contactIndex = introLines.findIndex(x => /contact info/i.test(x));
      if (contactIndex >= 0) {
        const contactLine = introLines[contactIndex].split(/\s*[·•]\s*contact info/i)[0].trim();
        if (!isGeoNoise(contactLine)) geo = contactLine;
        for (let i = contactIndex - 1; !geo && i >= 0; i--) {
          if (!isGeoNoise(introLines[i])) geo = introLines[i];
        }
      }
      if (!geo && h1) {
        const name = (h1.innerText || h1.textContent || "").trim().replace(/\s+/g, " ");
        const nameIndex = introLines.findIndex(x => x.replace(/\s+/g, " ") === name);
        if (nameIndex >= 0) {
          geo = introLines.slice(nameIndex + 1, nameIndex + 15)
            .find(x => /,|\b(?:area|region|metropolitan)\b/i.test(x) && !isGeoNoise(x)) || "";
        }
      }
    }
    const connectionMatch = text.match(/(?:^|\n)\s*([\d,.]+\+?)\s+connections?\b/i) ||
      introText.match(/\b([\d,.]+\+?)\s+connections?\b/i);
    const followerMatch = text.match(/(?:^|\n)\s*([\d,.]+\+?)\s+followers?\b/i) ||
      introText.match(/\b([\d,.]+\+?)\s+followers?\b/i);

    // LinkedIn changes badge markup between layouts. Search the primary profile card
    // around its h1 for explicit badge/frame tokens, not recommendation cards below.
    const badgeRoots = [];
    for (let node = h1, depth = 0; node && node !== main && depth < 12; node = node.parentElement, depth++) {
      const nodeText = (node.innerText || node.textContent || "").trim();
      if (nodeText.length > 5000) break;
      badgeRoots.push(node);
    }
    const normalizeBadgeValue = value => String(value || "").toLowerCase()
      .replace(/opentowork/g, "open to work")
      .replace(/[_-]+/g, " ").replace(/[^a-z0-9# ]+/g, " ").replace(/\s+/g, " ").trim();
    const explicitOpenWorkLabel = value => /^(?:#\s*)?open to work(?:\s+(?:(?:profile )?(?:photo|picture) )?(?:badge|frame|indicator))?$/.test(normalizeBadgeValue(value));
    const openWorkMarker = value => {
      const normalized = normalizeBadgeValue(value);
      return /(?:^|\s)#?open to work(?:$|\s)/.test(normalized) &&
        /\b(?:badge|frame|photo|picture|profile|indicator)\b/.test(normalized);
    };
    const openWorkNode = badgeRoots.some(root => {
      const nodes = [root, ...root.querySelectorAll(
        '[aria-label], [title], img[alt], [data-test-icon], [data-view-name], [data-testid], [data-test-id], [class]'
      )];
      return nodes.some(el => {
        const labels = [el.getAttribute?.('aria-label'), el.getAttribute?.('title'), el.getAttribute?.('alt')];
        if (labels.some(explicitOpenWorkLabel)) return true;
        const markers = [el.getAttribute?.('data-test-icon'), el.getAttribute?.('data-view-name'),
          el.getAttribute?.('data-testid'), el.getAttribute?.('data-test-id'), el.getAttribute?.('class')];
        return markers.some(value => explicitOpenWorkLabel(value) || openWorkMarker(value));
      });
    });
    const introHasOpenWorkBadge = badgeRoots.some(root =>
      (root.innerText || root.textContent || "").split(/\n+/).some(line => explicitOpenWorkLabel(line.trim())));
    // LinkedIn may expose Open to Work as a card below the name instead of labeling the photo frame.
    // Match the explicit audience line in the profile's rendered text without scanning every DOM node.
    const openWorkCard = lines.some(line =>
      /^open to work\s*[·•|]\s*(?:everyone on linkedin|all linkedin members|recruiters only)\b/i.test(line));
    const openToWork = (openWorkNode || introHasOpenWorkBadge || openWorkCard) ? "Yes" : "Not Available";

    // Read the actual Activity section text. LinkedIn layouts vary, and the post
    // or comment card is not consistently an <article> or a data-urn element.
    // The section heading anchors the scan; follower counts and tabs are discarded.
    let activity = "";
    const activityHeading = [...main.querySelectorAll("h2,h3,[role=heading]")]
      .find(e => /^activity(?:\s*\d*)?$/i.test((e.innerText||e.textContent||"").trim()));
    if (activityHeading) {
      let activityRoot = null;
      for (let node = activityHeading.parentElement, depth = 0; node && node !== main && node !== document.body && depth < 12; node = node.parentElement, depth++) {
        const text = (node.innerText || node.textContent || "").trim();
        if (!text || text.length > 18000) break;
        const otherSection = [...node.querySelectorAll("h2,h3,[role=heading]")]
          .some(e => e !== activityHeading && isOtherHeadingText(e.innerText||e.textContent||""));
        if (otherSection) break;
        activityRoot = node;
      }
      if (activityRoot) {
        const sectionLines = (activityRoot.innerText || activityRoot.textContent || "")
          .split(/\n+/).map(x => x.trim()).filter(Boolean);
        const headingIndex = sectionLines.findIndex(x => /^activity\b/i.test(x));
        const lines = sectionLines.slice(headingIndex >= 0 ? headingIndex + 1 : 0)
          .filter(x => !/\b[\d,.]+\+?\s+followers?\b/i.test(x))
          .filter(x => !/^(activity(?:\s*\d*)?|posts?|comments?|images?|documents?|reactions?|show all|see all activity)$/i.test(x))
          .filter(x => !/^(like|comment|repost|send|see translation|see more|show more|follow)$/i.test(x));
        const timePattern = /(?:^|[·•]\s*)(?:just now|\d+\s*(?:seconds?|secs?|minutes?|mins?|hours?|hrs?|days?|weeks?|months?|mos?|years?|yrs?|s|m|h|d|w|mo|y)(?:\s*ago)?)(?=$|[\s·•])|^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}(?:,\s*\d{4})?$/i;
        const timeIndex = lines.findIndex(x => timePattern.test(x));
        if (timeIndex >= 0) {
          const timeText = lines[timeIndex];
          const age = timeText.match(/(?:just now|\d+\s*(?:seconds?|secs?|minutes?|mins?|hours?|hrs?|days?|weeks?|months?|mos?|years?|yrs?|mo|yr|s|m|h|d|w|y)(?:\s*ago)?|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}(?:,\s*\d{4})?)/i)?.[0]?.trim() || timeText;
          const edited = /\bedited\b/i.test(timeText) || /^edited$/i.test(lines[timeIndex + 1] || "") || /^edited$/i.test(lines[timeIndex - 1] || "");
          activity = `${age}${edited ? " • Edited" : ""}`;
        }
      }
    }
    const headline = lines.slice(1,4).join(" · ");
    let about = "";
    const aboutHeading = [...main.querySelectorAll("h2,h3,[role=heading]")].find(e => /^about$/i.test((e.innerText||e.textContent||"").trim()));
    if (aboutHeading) about = aboutHeading.parentElement?.innerText || "";
    return {Geo:geo, Connections:connectionMatch?.[1]||"", Followers:followerMatch?.[1]||"",
      Activity:activity, "Open to Work":openToWork, Headline:headline, About:about};
  }
  async function captureProfileMetadata() {
    // Metadata reads must not move the page unless the user enabled Auto Mode.
    if (!autoMode) return visibleProfileMetadata();
    const originalY = window.scrollY || window.pageYOffset || 0;
    try {
      // LinkedIn often lazy-renders Activity below the fold. Scroll in bounded steps,
      // stopping as soon as the Activity heading is present, then allow its card to settle.
      // A hidden tab never renders lazily, so scrolling it cannot make Activity appear: one look is enough.
      const maxSteps = document.hidden ? 1 : 8;
      for (let i = 0; i < maxSteps; i++) {
        const main = document.querySelector("main") || document.body;
        const activityHeading = [...main.querySelectorAll("h2,h3,[role=heading]")]
          .find(e => /^activity(?:\s*\d*)?$/i.test((e.innerText||e.textContent||"").trim()));
        if (activityHeading) {
          activityHeading.scrollIntoView({behavior:"instant", block:"center"});
          await sleep(200);
          break;
        }
        window.scrollBy(0, Math.max(500, Math.floor(window.innerHeight * 0.75)));
        await sleep(120);
      }
      return visibleProfileMetadata();
    } finally {
      window.scrollTo(0, originalY);
    }
  }

  // Chrome throttles timers in hidden tabs (about one tick per second, and about one per minute after a few
  // minutes in the background), which stretches every wait below. The extension's background worker is not
  // throttled, so a hidden tab borrows its clock.
  const pageSleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const extensionAlive = () => Boolean(chrome.runtime?.id);
  const sleep = ms => (document.hidden && extensionAlive())
    ? chrome.runtime.sendMessage({ type: "LI_COPIER_SLEEP", ms }).catch(() => pageSleep(ms))
    : pageSleep(ms);

  // A tiny page-local status badge. It does not scroll the page or steal focus.
  let statusHost = null;
  let statusText = null;
  function ensureStatusBadge() {
    if (statusHost?.isConnected) return;
    statusHost = document.createElement("div");
    statusHost.id = "__liExperienceCopierStatusHost";
    statusHost.style.cssText = "position:fixed;top:12px;right:12px;z-index:2147483647;pointer-events:none;";
    const shadow = statusHost.attachShadow({ mode: "open" });
    shadow.innerHTML = `
      <style>
        .badge{font:600 12px/1.3 Arial,sans-serif;display:flex;align-items:center;gap:7px;padding:8px 11px;border-radius:999px;background:#202124;color:#fff;box-shadow:0 2px 10px #0003;max-width:260px}
        .dot{width:8px;height:8px;border-radius:50%;background:#fbbc04;flex:none}
        .text{white-space:normal;overflow-wrap:anywhere}
        .saved .dot{background:#34a853}.working .dot{background:#fbbc04}.error .dot{background:#ea4335}.off .dot{background:#9aa0a6}
        .alert{background:#c5221f;font-size:14px;max-width:340px;box-shadow:0 0 0 3px #ff8a80,0 4px 18px #0006;animation:liBlink .9s ease-in-out infinite}
        .alert .dot{background:#fff}
        @keyframes liBlink{0%,100%{opacity:1}50%{opacity:.25}}
      </style>
      <div class="badge working"><span class="dot"></span><span class="text">Waiting for Experience…</span></div>`;
    statusHost.badge = shadow.querySelector(".badge");
    statusText = shadow.querySelector(".text");
    (document.documentElement || document.body).appendChild(statusHost);
  }
  function setStatus(text, kind = "working") {
    if (!isProfile()) return;
    ensureStatusBadge();
    statusHost.badge.className = `badge ${kind}`;
    statusText.textContent = text;
  }
  function hideStatus() {
    if (statusHost) statusHost.style.display = "none";
  }
  function showStatus() {
    if (statusHost) statusHost.style.display = "block";
  }

  // textContent does not force a page layout (innerText does), so it is used for
  // all the repeated checks. innerText is read once, at the end, for clean line breaks.
  function headingCandidates(broad, label = "experience") {
    const out = [];
    // Normal case: a real heading tag. Broad case: ANY element whose whole text is
    // "Experience" (newer LinkedIn layouts use <p>/<span>/<div> instead of <h2>).
    for (const el of document.querySelectorAll(broad ? "*" : "h1,h2,h3,[role='heading']")) {
      if (broad && el.childElementCount) continue;
      if ((el.textContent || "").trim().toLowerCase() === label) out.push(el);
    }
    return out;
  }

  function isOtherHeadingText(text) {
    const t = (text || "").trim().toLowerCase();
    return t.length > 0 && t.length <= 40 && SECTION_HEADINGS.has(t);
  }
  function containsOtherHeading(node) {
    for (const el of node.querySelectorAll("h1,h2,h3,[role='heading']")) {
      if (isOtherHeadingText(el.textContent)) return true;
    }
    for (const el of node.querySelectorAll("*")) {
      if (el.childElementCount) continue;
      if (isOtherHeadingText(el.textContent)) return true;
    }
    return false;
  }

  // Finds the Experience CARD only: the smallest container with real content after the heading,
  // widened while it still holds no other section (Activity, Education, Skills...).
  // It stops at <main>, so the whole profile page can never be returned.
  function sectionFrom(heading) {
    const headLen = (heading.textContent || "").trim().length;
    let start = null;
    let node = heading;
    for (let depth = 0; node && depth < 12; depth++, node = node.parentElement) {
      if (node.tagName === "MAIN" || node === document.body) break;
      const len = (node.textContent || "").trim().length;
      if (len >= 25000) break;
      if (len >= headLen + MIN_EXTRA_CHARS) { start = node; break; }
    }
    // Nothing yet, or the first container with content already mixes in other sections:
    // not ready / not safe, so report "not found" instead of returning too much.
    if (!start || containsOtherHeading(start)) return null;
    let best = start;
    for (let up = 0; up < 6; up++) {
      const parent = best.parentElement;
      if (!parent || parent.tagName === "MAIN" || parent === document.body) break;
      if ((parent.textContent || "").length >= 25000) break;
      if (containsOtherHeading(parent)) break;
      best = parent;
    }
    return best;
  }

  function fallbackExperienceSection() {
    const main = document.querySelector("main");
    if (!main) return null;
    const lines = (main.innerText || main.textContent || "").split(/\n+/).map(x => x.trim()).filter(Boolean);
    const start = lines.findIndex(line => /^experience(?:\s+\d+)?$/i.test(line));
    if (start < 0) return null;
    let end = lines.length;
    for (let i = start + 1; i < lines.length; i++) {
      if (SECTION_HEADINGS.has(lines[i].toLowerCase())) { end = i; break; }
    }
    const text = lines.slice(start, end).join("\n");
    if (text.length < MIN_EXTRA_CHARS) return null;
    return {tagName:"section", textContent:text, innerText:text, innerHTML:""};
  }

  function locateExperienceSection() {
    const now = Date.now();
    if (now - lastLocateScan >= 500) {
      lastLocateScan = now;
      if (headingEl?.isConnected) {
        const cached = sectionFrom(headingEl);
        if (cached) return cached;
        headingEl = null;
      }
      for (const el of headingCandidates(false)) {
        const found = sectionFrom(el);
        if (found) { headingEl = el; return found; }
      }
      if (now - lastBroadScan >= 500) {  // match the working extension for newly rendered Experience headings
        lastBroadScan = now;
        for (const el of headingCandidates(true)) {
          const found = sectionFrom(el);
          if (found) { headingEl = el; return found; }
        }
      }
    }
    if (now - lastTextFallback >= 1500) {
      lastTextFallback = now;
      const fallback = fallbackExperienceSection();
      if (fallback) return fallback;
    }
    return null;
  }

  function waitForPageLoad(timeoutMs = 6000) {
    return new Promise(resolve => {
      let timer = null;
      const finish = () => {
        window.removeEventListener("load", finish);
        if (timer !== null) clearTimeout(timer);
        resolve();
      };
      window.addEventListener("load", finish, {once:true});
      if (document.readyState === "complete") finish();
      else timer = setTimeout(finish, timeoutMs);
    });
  }

  function scrollProfileDown(amount = Math.max(700, Math.floor(window.innerHeight * 0.85))) {
    const main = document.querySelector("main,[role='main']") || document.body;
    const candidates = [];
    for (let node = main; node && node !== document.documentElement; node = node.parentElement) {
      const style = window.getComputedStyle(node);
      if (/(auto|scroll|overlay|hidden)/i.test(style.overflowY) && node.scrollHeight > node.clientHeight + 8) {
        candidates.push(node);
      }
    }
    const root = document.scrollingElement;
    if (root && !candidates.includes(root)) candidates.push(root);
    for (const node of candidates) {
      const before = node.scrollTop;
      const maxTop = Math.max(0, node.scrollHeight - node.clientHeight);
      node.scrollTop = Math.min(maxTop, before + amount);
      if (node.scrollTop > before + 1) return true;
    }
    const beforeY = window.scrollY || window.pageYOffset || 0;
    window.scrollBy(0, amount);
    return (window.scrollY || window.pageYOffset || 0) > beforeY + 1;
  }

  function scheduleInitialAutoScroll(delay = 1000) {
    if (autoScrollTimer !== null) clearTimeout(autoScrollTimer);
    autoScrollTimer = null;
    initialAutoScrollDone = !autoMode;
    if (!autoMode || !enabled || !isProfile()) return;
    autoScrollTimer = setTimeout(() => {
      autoScrollTimer = null;
      if (!enabled || !autoMode || !isProfile()) return;
      const moved = scrollProfileDown();
      initialAutoScrollDone = true;
      setStatus(moved ? "Scrolled down to load Experience…" : "Profile is already at the bottom; checking Experience…", "working");
      console.info("[Experience Copier] Auto Mode one-second profile scroll", {moved, url:cleanUrl()});
    },delay);
  }

  async function waitForExperience(url, maxWaitMs = EXPERIENCE_WAIT_MS) {
    pausedHidden = false;
    const started = Date.now();
    const deadline = started + maxWaitMs;
    let lastText = null;
    let stableSince = 0;
    let best = null;
    let nextAutoScrollAt = started;
    let scrolledToSection = false;
    while (Date.now() < deadline) {
      if (!enabled || !isProfile() || cleanUrl() !== url) return null;
      // LinkedIn delays profile sections in background tabs. Wait until this tab is visible.
      if (document.hidden) { pausedHidden = true; return null; }
      const section = best?.isConnected ? best : locateExperienceSection();
      if (section) {
        best = section;
        if (autoMode && !scrolledToSection) {
          const target = typeof section.scrollIntoView === "function" ? section : headingEl;
          target?.scrollIntoView?.({behavior:"instant",block:"center"});
          scrolledToSection = true;
          await sleep(250);
        }
        const text = section.textContent || "";
        const now = Date.now();
        if (text !== lastText) {
          lastText = text;
          stableSince = now;
        }
        // Right after navigating, LinkedIn may still show the previous profile.
        const stale = lastSectionUrl && lastSectionUrl !== url &&
          text === lastSectionText && now - started < STALE_GUARD_MS;
        // A section can legitimately be tiny; we only require that it stops changing.
        if (!stale && now - stableSince >= STABLE_MS && (!autoMode || initialAutoScrollDone)) return section;
      } else if (autoMode && initialAutoScrollDone && Date.now() >= nextAutoScrollAt) {
        // LinkedIn may place the profile in an inner scroll container rather than
        // the window. Move the actual scrollable ancestor as the page lazy-loads.
        scrollProfileDown(Math.max(500,Math.floor(window.innerHeight * 0.7)));
        nextAutoScrollAt = Date.now() + 450;
      }
      await sleep(autoMode ? 180 : POLL_MS);
    }
    return best?.isConnected ? best : null;
  }

  let scheduleToken = 0;
  function cancelScheduledCapture() {
    scheduleToken++;
    if (typeof timer === "number") clearTimeout(timer);
    timer = null;
  }
  function scheduleCapture(delay = START_DELAY_MS, replace = false) {
    if (!enabled || !isProfile() || lastCapturedUrl === cleanUrl() || failedUrl === cleanUrl()) return;
    if (timer !== null && !replace) return;
    if (timer !== null) cancelScheduledCapture();
    const token = ++scheduleToken;
    const run = () => {
      if (token !== scheduleToken) return;
      timer = null;
      capture();
    };
    if (delay <= 0) {
      timer = "background";
      queueMicrotask(run);
    } else if (document.hidden && extensionAlive()) {
      timer = "background";
      chrome.runtime.sendMessage({ type: "LI_COPIER_SLEEP", ms: delay }).then(run).catch(() => {
        if (token === scheduleToken) timer = setTimeout(run, delay);
      });
    } else {
      timer = setTimeout(run, delay);
    }
  }

  async function capture() {
    if (!enabled || !isProfile() || inProgress) return;
    const url = cleanUrl();
    if (url === lastCapturedUrl || url === failedUrl) return;
    if (document.hidden) {
      setStatus("Waiting — open this tab to capture", "working");
      return;
    }
    inProgress = true;
    setStatus("Waiting for Experience section…", "working");
    try {
      let section = await waitForExperience(url);
      if (pausedHidden) {
        setStatus("Paused — open this tab to continue", "working");
        return;
      }
      if (!section && enabled && isProfile() && cleanUrl() === url) {
        setStatus("Experience has not loaded yet — waiting for page load, then retrying…", "working");
        await waitForPageLoad();
        if (document.hidden) pausedHidden = true;
        if (enabled && isProfile() && cleanUrl() === url && !document.hidden) {
          section = await waitForExperience(url, EXPERIENCE_RETRY_WAIT_MS);
        }
        if (pausedHidden) {
          setStatus("Paused — open this tab to continue", "working");
          return;
        }
      }
      if (!enabled || !isProfile() || cleanUrl() !== url) return;
      if (autoMode && section?.isConnected) {
        const target = typeof section.scrollIntoView === "function" ? section : headingEl;
        target?.scrollIntoView?.({behavior:"instant",block:"center"});
        await sleep(300);
      }
      const text = section ? (section.innerText || section.textContent || "").trim() : "";
      if (section) {
        lastSectionUrl = url;
        lastSectionText = section.textContent || "";
      }
      const metadata = await captureProfileMetadata();
      if (document.hidden) {
        pausedHidden = true;
        setStatus("Paused — open this tab to continue", "working");
        return;
      }
      const payload = {
        url,
        page_title: document.title || "",
        page_name: (document.querySelector("main h1")?.textContent || "").trim(),
        first_seen: urlSeenAt,
        experience_text: text,
        metadata,
        capture_method: text ? (autoMode ? "browser_dom_section_auto_scroll" : "browser_dom_section_no_scroll") : "python_html_fallback_no_scroll"
      };
      // The big HTML blobs are only a fallback for when no text was found.
      // Sending them every time made each capture many times heavier for nothing.
      if (!text) {
        const main = document.querySelector("main");
        payload.experience_html = section?.innerHTML || "";
        if (SEND_PAGE_HTML_FALLBACK) {
          const markup = main?.innerHTML || document.body.innerHTML || "";
          payload.page_html = markup.length > 1500000 ? markup.slice(0, 1500000) : markup;
        }
      }

      if (document.hidden) {
        pausedHidden = true;
        setStatus("Paused — open this tab to continue", "working");
        return;
      }
      setStatus("Sending capture to Excel…", "working");
      const response = await chrome.runtime.sendMessage({ type: "LI_COPIER_SUBMIT_CAPTURE", payload });
      const result = response?.result || {};
      const success = Boolean(response?.ok && result.ok);
      console.info("[Experience Copier] Experience section:", section ? `<${section.tagName.toLowerCase()}> ${text.length} chars` : "not found");
      const terminal = true;
      if (success && !text) {
        failedUrl = url;
        emptyUrl = url;
        setStatus(ALERT.empty, "alert");
      } else if (success && (result.quality || "ok") !== "ok") {
        lastCapturedUrl = url;
        failedUrl = "";
        emptyUrl = "";
        setStatus(ALERT[result.quality] || ALERT.unparsed, "alert");
        console.warn("[Experience Copier] Capture quality problem:", result.quality, result.message);
      } else if (success) {
        lastCapturedUrl = url;
        failedUrl = "";
        emptyUrl = "";
        setStatus(result.message || "Saved successfully", "saved");
        console.info("[Experience Copier] Capture accepted", result.record || url);
      } else {
        const error = result.error || response?.error || "Capture was not accepted";
        failedUrl = url;
        setStatus(ALERT.failed, "alert");
        console.warn("[Experience Copier] Capture failed; no automatic resubmission:", error);
      }
      chrome.runtime.sendMessage({
        type: "LI_COPIER_CAPTURE_RESULT",
        terminal,
        captureOk: success && Boolean(text),
        url,
        reason: success ? (result.message || "Experience text was not available") : (result.error || response?.error || "Capture failed")
      }).catch(() => {});
    } catch (e) {
      const error = String(e);
      failedUrl = url;
      setStatus(ALERT.failed, "alert");
      chrome.runtime.sendMessage({
        type: "LI_COPIER_CAPTURE_RESULT",
        terminal: true,
        captureOk: false,
        url,
        reason: error
      }).catch(() => {});
      console.warn("[Experience Copier] Capture error; no automatic resubmission:", error);
    } finally {
      inProgress = false;
      // If navigation happened while this profile was being processed, re-queue the current URL now.
      const currentUrl = cleanUrl();
      if (enabled && isProfile() && currentUrl && currentUrl !== url) {
        if (lastCapturedUrl !== currentUrl) lastCapturedUrl = "";
        if (failedUrl !== currentUrl) failedUrl = "";
        scheduleCapture(START_DELAY_MS, true);
      }
    }
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message?.type === "LI_COPIER_SET_ENABLED") {
      enabled = Boolean(message.enabled);
      if (message.autoMode !== undefined) autoMode = Boolean(message.autoMode);
      lastSeenUrl = location.href;
      if (!enabled) {
        cancelScheduledCapture();
        hideStatus();
      } else {
        showStatus();
        if (isProfile() && lastCapturedUrl !== cleanUrl()) scheduleCapture(START_DELAY_MS);
      }
      sendResponse?.({ ok: true, enabled });
    } else if (message?.type === "LI_COPIER_AUTO_SAVE_FAILED") {
      setStatus("Excel save is still pending — tab left open", "alert");
      sendResponse?.({ok:true});
    } else if (message?.type === "LI_COPIER_CAPTURE_NOW") {
      enabled = true;
      lastCapturedUrl = "";
      failedUrl = "";
      lastSectionUrl = "";
      lastSectionText = "";
      emptyUrl = "";
      showStatus();
      setStatus("Waiting for Experience section…", "working");
      scheduleCapture(0, true);
      sendResponse?.({ ok: true });
    }
    // Every reply above is synchronous, so the channel does not need to stay open.
  });

  chrome.storage.local.get({ captureEnabled: false, autoMode: false }).then(async result => {
    enabled = Boolean(result.captureEnabled);
    try {
      // Scroll only when both this queue item and the current user toggle authorize it.
      const options = await chrome.runtime.sendMessage({type:"LI_COPIER_GET_TAB_OPTIONS"});
      queuedAutoMode = Boolean(options?.autoMode);
      autoMode = queuedAutoMode && Boolean(result.autoMode);
    } catch (_) { queuedAutoMode = false; autoMode = false; }
    if (enabled && isProfile()) {
      showStatus();
      if (autoMode) scheduleInitialAutoScroll(1000);
      scheduleCapture(START_DELAY_MS);
    } else {
      hideStatus();
    }
  });
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    if (changes.autoMode) {
      autoMode = queuedAutoMode && Boolean(changes.autoMode.newValue);
      scheduleInitialAutoScroll(1000); // clears a pending scroll immediately when Auto Mode is turned off
    }
    if (changes.captureEnabled) {
      enabled = Boolean(changes.captureEnabled.newValue);
      if (!enabled) {
        cancelScheduledCapture();
        hideStatus();
      } else if (isProfile()) {
        showStatus();
        scheduleCapture(START_DELAY_MS);
      }
    }
  });

  // The moment a tab becomes visible, LinkedIn starts rendering it: capture then.
  document.addEventListener("visibilitychange", () => {
    if (document.hidden || !enabled || !isProfile()) return;
    const url = cleanUrl();
    if (lastCapturedUrl === url || emptyUrl === url || failedUrl === url) return;
    showStatus();
    scheduleCapture(START_DELAY_MS, true);
  });

  // LinkedIn is a single-page app. URL changes schedule capture even when this tab is inactive.
  const urlTimer = setInterval(() => {
    if (!extensionAlive()) {            // extension was reloaded: stop this orphaned copy
      clearInterval(urlTimer);
      observer.disconnect();
      return;
    }
    if (location.href !== lastSeenUrl) {
      lastSeenUrl = location.href;
      urlSeenAt = Date.now();
      lastCapturedUrl = "";
      failedUrl = "";
      headingEl = null;
      emptyUrl = "";
      initialAutoScrollDone = !autoMode;
      if (isProfile() && enabled) {
        showStatus();
        setStatus("Profile detected — waiting for Experience…", "working");
        chrome.runtime.sendMessage({type:"LI_COPIER_GET_TAB_OPTIONS"}).then(options =>
          chrome.storage.local.get({autoMode:false}).then(settings => {
            queuedAutoMode = Boolean(options?.autoMode);
            autoMode = queuedAutoMode && Boolean(settings.autoMode);
            scheduleInitialAutoScroll(1000);
            scheduleCapture(START_DELAY_MS, true);
          })
        ).catch(() => {
          queuedAutoMode = false;
          autoMode = false;
          scheduleInitialAutoScroll(1000);
          scheduleCapture(START_DELAY_MS, true);
        });
      } else {
        hideStatus();
      }
    }
  }, URL_POLL_MS);

  const observer = new MutationObserver(records => {
    if (!enabled || !isProfile() || inProgress || document.hidden) return;
    // Ignore our own status badge.
    if (!records.some(record => !statusHost || !statusHost.contains(record.target))) return;
    const url = cleanUrl();
    if (lastCapturedUrl === url || failedUrl === url || emptyUrl === url) return;
    scheduleCapture(250);
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
})();
