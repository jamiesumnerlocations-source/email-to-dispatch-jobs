/**
 * Email → Dispatch Jobs (Google Apps Script)
 *
 * Scans authorised Gmail senders for recent emails, extracts dispatch "moves"
 * (date, time, origin, destination, vehicle counts), and writes them to a
 * Google Sheet as one row per dispatchable job.
 *
 * Dedupe: SourceEmailId + Date + Time + Origin + Destination
 */

const CONFIG = {
  // Replace with your Sheet ID (do NOT commit real values)
  spreadsheetId: "YOUR_SPREADSHEET_ID",
  jobsSheetName: "Jobs",

  // Replace with your allowed sender list (do NOT commit real emails)
  authorisedSenders: [
    "authorised.sender1@example.com",
    "authorised.sender2@example.com",
  ],

  newerThanDays: 180,
  maxThreads: 30,
  countryBias: "UK",

  // If true, also create a DRAFT job when date found but time missing
  createDraftIfMissingTime: true,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Dispatch")
    .addItem("Scan authorised Gmail → Jobs", "scanGmailCreateJobs")
    .addToUi();
}

function scanGmailCreateJobs() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const jobsSheet = ss.getSheetByName(CONFIG.jobsSheetName);
  if (!jobsSheet) throw new Error(`Missing sheet: ${CONFIG.jobsSheetName}`);

  const headers = getHeaders_(jobsSheet);
  const existingJobIds = loadIds_(jobsSheet, headers, "JobID");
  const existingDedupe = loadDedupeKeys_(jobsSheet, headers);

  const query = buildQuery_();
  const threads = GmailApp.search(query, 0, CONFIG.maxThreads);

  for (const thread of threads) {
    const msg = thread.getMessages().pop();
    const sourceEmailId = msg.getId();

    const from = extractEmail_(msg.getFrom());
    if (!CONFIG.authorisedSenders.includes(from)) continue;

    const body = msg.getPlainBody() || "";
    const lines = body.split(/\r?\n/).map(normalizeLine_).filter(Boolean);

    const extracted = extractMoves_(lines);

    // If nothing extracted, create a placeholder DRAFT so you don't miss it.
    if (extracted.length === 0) {
      const jobId = generateId_(existingJobIds); existingJobIds.add(jobId);
      appendRow_(jobsSheet, headers, {
        JobID: jobId,
        CreatedAt: new Date(),
        SourceEmailId: sourceEmailId,
        Requester: from,
        VehicleType: "",
        Date: "",
        Time: "",
        Origin: "",
        Destination: "",
        DistanceKm: "",
        DurationMins: "",
        MapsUrl: "",
        Status: "DRAFT",
        Notes: "WARN:NO_DATE_TIME_FOUND",
      });
      continue;
    }

    for (const move of extracted) {
      const hasDate = Boolean(move.date);
      const hasTime = Boolean(move.time);

      if (!(hasDate && hasTime)) {
        if (!CONFIG.createDraftIfMissingTime) continue;
        if (!hasDate) continue; // don't create draft if date missing (too noisy)

        const key = dedupeKey_(sourceEmailId, move.date, move.time, move.origin, move.destination);
        if (existingDedupe.has(key)) continue;
        existingDedupe.add(key);

        const jobId = generateId_(existingJobIds); existingJobIds.add(jobId);
        appendRow_(jobsSheet, headers, {
          JobID: jobId,
          CreatedAt: new Date(),
          SourceEmailId: sourceEmailId,
          Requester: from,
          VehicleType: classifyVehicleType_(move.counts),
          Date: move.date,
          Time: move.time,
          Origin: move.origin,
          Destination: move.destination,
          DistanceKm: "",
          DurationMins: "",
          MapsUrl: "",
          Status: "DRAFT",
          Notes: "WARN:MISSING_TIME",
        });
        continue;
      }

      const key = dedupeKey_(sourceEmailId, move.date, move.time, move.origin, move.destination);
      if (existingDedupe.has(key)) continue;
      existingDedupe.add(key);

      const route = shouldRoute_(move.origin, move.destination)
        ? computeRoute_(move.origin, move.destination)
        : emptyRoute_();

      const jobId = generateId_(existingJobIds); existingJobIds.add(jobId);

      appendRow_(jobsSheet, headers, {
        JobID: jobId,
        CreatedAt: new Date(),
        SourceEmailId: sourceEmailId,
        Requester: from,
        VehicleType: classifyVehicleType_(move.counts),
        Date: move.date,
        Time: move.time,
        Origin: move.origin,
        Destination: move.destination,
        DistanceKm: route.distanceKm,
        DurationMins: route.durationMins,
        MapsUrl: route.mapsUrl,
        Status: "DRAFT",
        Notes: "",
      });
    }
  }
}

/* ===================== EXTRACTION ===================== */

function extractMoves_(lines) {
  const moves = [];
  let dateContext = "";
  let timeContext = "";
  let countsContext = { truck: 0, van: 0, car: 0 };

  let pending = newMove_();

  const flushIfMeaningful = () => {
    if (!pending.date) pending.date = dateContext;
    if (!pending.time) pending.time = timeContext;

    const pendingTotal = pending.counts.truck + pending.counts.van + pending.counts.car;
    const ctxTotal = countsContext.truck + countsContext.van + countsContext.car;
    if (pendingTotal === 0 && ctxTotal > 0) pending.counts = { ...countsContext };

    pending.origin = cleanPlace_(pending.origin);
    pending.destination = cleanPlace_(pending.destination);

    const hasAny = Boolean(pending.date || pending.time || pending.origin || pending.destination);
    if (hasAny) moves.push(pending);

    pending = newMove_();
  };

  for (const line of lines) {
    if (shouldIgnoreLine_(line)) continue;

    const wd = parseWeekdayDateHeader_(line);
    if (wd) {
      flushIfMeaningful();
      if (wd.date) dateContext = wd.date;
      if (wd.time) timeContext = wd.time;
      continue;
    }

    const bd = parseBareDate_(line);
    if (bd) {
      dateContext = bd;
      continue;
    }

    const t = detectTime_(line);
    if (t) timeContext = t;

    const vc = detectVehicleCounts_(line);
    countsContext = addCountsReturn_(countsContext, vc);

    if (/^from\s*:/i.test(line) && !/@/.test(line)) {
      pending.origin = stripLabel_(line);
      continue;
    }
    if (/^to\s*:/i.test(line) && !/@/.test(line)) {
      pending.destination = stripLabel_(line);
      continue;
    }

    if (/^collection address\s*:/i.test(line) || /^collection\s*:/i.test(line)) {
      const { time, place } = parseTimedPlace_(line);
      if (time) pending.time = time;
      pending.origin = place;
      continue;
    }
    if (/^drop off address\s*:/i.test(line) || /^drop off\s*:/i.test(line) || /^drop-off\s*:/i.test(line)) {
      const { time, place } = parseTimedPlace_(line);
      if (time) pending.time = time;
      pending.destination = place;
      continue;
    }

    const rt = extractInlineRoute_(line);
    if (rt.origin && rt.destination) {
      pending.origin = pending.origin || rt.origin;
      pending.destination = pending.destination || rt.destination;
      continue;
    }

    if (line === "--" || /^_{5,}$/.test(line)) {
      flushIfMeaningful();
      continue;
    }
  }

  flushIfMeaningful();
  return moves;
}

function newMove_() {
  return {
    date: "",
    time: "",
    origin: "",
    destination: "",
    counts: { truck: 0, van: 0, car: 0 },
  };
}

/* ===================== INLINE ROUTE/DATE/TIME ===================== */

function extractInlineRoute_(text) {
  const t = String(text || "").replace(/\s+/g, " ").trim();
  const patterns = [
    /\bmoved\s+from\s+(.+?)\s+to\s+(.+?)(?:\s+(?:at|on|by|for)\b|[?.!,]|$)/i,
    /\bfrom\s+(.+?)\s+to\s+(.+?)(?:\s+(?:at|on|by|for)\b|[?.!,]|$)/i,
  ];
  for (const rx of patterns) {
    const m = t.match(rx);
    if (m) return { origin: m[1].trim(), destination: m[2].trim() };
  }
  return { origin: "", destination: "" };
}

function parseWeekdayDateHeader_(line) {
  const m = String(line).match(
    /\b(mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b.*?\b(\d{1,2})(st|nd|rd|th)?\b.*?\b(jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\b/i
  );
  if (!m) return null;
  const dd = String(m[2]).padStart(2, "0");
  const mm = month_(m[4]);
  const date = mm ? `${dd}/${mm}/` : "";
  const time = detectTime_(line);
  return { date, time };
}

function parseBareDate_(line) {
  const m = String(line).match(
    /\b(0?[1-9]|[12]\d|3[01])(?:st|nd|rd|th)?(?:\s+of)?\s+(jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\b/i
  );
  if (!m) return "";
  const dd = String(m[1]).padStart(2, "0");
  const mm = month_(m[2]);
  return mm ? `${dd}/${mm}/` : "";
}

function detectTime_(text) {
  const s = String(text || "").toLowerCase();

  let m = s.match(/\b([01]?\d|2[0-3])[:. ]([0-5]\d)\b/);
  if (m) return String(m[1]).padStart(2, "0") + ":" + m[2];

  m = s.match(/\b(1[0-2]|0?[1-9])(?:[:. ]([0-5]\d))?\s?(am|pm)\b/);
  if (m) {
    let h = parseInt(m[1], 10);
    const mins = m[2] ? m[2] : "00";
    const ap = m[3];
    if (ap === "pm" && h !== 12) h += 12;
    if (ap === "am" && h === 12) h = 0;
    return String(h).padStart(2, "0") + ":" + mins;
  }

  return "";
}

/* ===================== VEHICLE COUNTS ===================== */

function detectVehicleCounts_(text) {
  const t = String(text || "").toLowerCase();
  const c = { truck: 0, van: 0, car: 0 };

  const trucks = t.match(/\b(\d+)\s*trucks?\b/);
  if (trucks) c.truck += parseInt(trucks[1], 10);
  if (/\btwo\s+trucks?\b/.test(t)) c.truck += 2;

  const vans = t.match(/\b(\d+)\s*vans?\b/);
  if (vans) c.van += parseInt(vans[1], 10);

  const rx = /\b(\d+)\s*x\s*([^,\n]+)/gi;
  let m;
  while ((m = rx.exec(t)) !== null) {
    const n = parseInt(m[1], 10);
    const label = m[2] || "";
    if (/\btruck\b/.test(label)) c.truck += n;
    else if (/\bvan\b|\bluton\b/.test(label)) c.van += n;
    else if (/\bcar\b|\b4x4\b/.test(label)) c.car += n;
  }

  if (/\blighting truck\b/.test(t) && c.truck === 0) c.truck += 1;
  if ((/\bluton\b/.test(t) || /\bvan\b/.test(t)) && c.van === 0) c.van += 1;

  return c;
}

function addCountsReturn_(a, b) {
  return { truck: a.truck + b.truck, van: a.van + b.van, car: a.car + b.car };
}

function classifyVehicleType_(c) {
  const types = [];
  if (c.truck) types.push("Truck");
  if (c.van) types.push("Van");
  if (c.car) types.push("Car");
  return types.length > 1 ? "Mixed" : (types[0] || "");
}

/* ===================== ROUTING ===================== */

function shouldRoute_(o, d) {
  if (!o || !d) return false;
  const bad = /(tbc|unknown|to be confirmed)/i;
  return !(bad.test(o) || bad.test(d));
}

function computeRoute_(origin, destination) {
  const o = biasUK_(cleanPlace_(origin));
  const d = biasUK_(cleanPlace_(destination));
  try {
    const res = Maps.newDirectionFinder()
      .setOrigin(o)
      .setDestination(d)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (!res?.routes?.length) return emptyRoute_();
    const leg = res.routes[0].legs?.[0];
    if (!leg?.distance?.value || !leg?.duration?.value) return emptyRoute_();

    return {
      distanceKm: Math.round((leg.distance.value / 1000) * 10) / 10,
      durationMins: Math.round(leg.duration.value / 60),
      mapsUrl: mapsUrl_(o, d),
    };
  } catch {
    return emptyRoute_();
  }
}

function emptyRoute_() {
  return { distanceKm: "", durationMins: "", mapsUrl: "" };
}

function mapsUrl_(o, d) {
  return `https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(o)}&destination=${encodeURIComponent(d)}&travelmode=driving`;
}

function biasUK_(place) {
  const p = String(place || "").trim();
  if (!p) return "";
  const hasPostcode = /\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b/i.test(p);
  if (hasPostcode || p.includes(",")) return p;
  return `${p}, ${CONFIG.countryBias}`;
}

/* ===================== CLEANING / NORMALIZATION ===================== */

function normalizeLine_(line) {
  let s = String(line || "").trim();
  if (!s) return "";

  s = s.replace(/^\*+/, "").replace(/\*+$/, "").trim();
  s = s.replace(/\s+/g, " ").trim();

  s = s.replace(
    /\*?(vehicles|from|to|collection|collection address|drop off|drop off address|drop-off|time)\*?\s*:/i,
    (m) => m.replace(/\*/g, "")
  );

  return s;
}

function shouldIgnoreLine_(line) {
  const l = line.toLowerCase();
  if (/^(subject|cc|bcc)\s*:\s*/i.test(line)) return true;
  if (l.includes("forwarded message")) return true;
  if (l.startsWith("----------")) return true;
  if (l.includes("http://") || l.includes("https://") || l.includes("www.")) return true;
  if (/^(from|to)\s*:\s*/i.test(line) && /@/.test(line)) return true;
  return false;
}

function cleanPlace_(p) {
  let s = String(p || "").trim();
  if (!s) return "";
  s = s.replace(/^(location|unit base|office)\s*-\s*/i, "").trim();
  s = s.replace(/\bw3w\s*:\s*\/\/\/[a-z.]+\b/ig, "").trim();
  s = s.replace(/^\*+/, "").trim();
  return s;
}

function stripLabel_(line) {
  return String(line).replace(/^[^:]+:\s*/i, "").trim();
}

function parseTimedPlace_(line) {
  const rest = stripLabel_(line);
  const time = detectTime_(rest) || "";
  const cleaned = rest.replace(/\b([01]?\d|2[0-3])[:.][0-5]\d\b.*?-\s*/i, "").trim();
  return { time, place: cleaned };
}

function month_(m) {
  const k = String(m || "").slice(0, 3).toLowerCase();
  return {
    jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
    jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
  }[k] || "";
}

/* ===================== DEDUPE + SHEET HELPERS ===================== */

function dedupeKey_(emailId, date, time, origin, dest) {
  return [
    String(emailId || "").trim(),
    String(date || "").trim(),
    String(time || "").trim(),
    String(origin || "").trim().toLowerCase(),
    String(dest || "").trim().toLowerCase(),
  ].join("|");
}

function loadDedupeKeys_(sheet, headers) {
  const idxEmail = headers.indexOf("SourceEmailId");
  const idxDate = headers.indexOf("Date");
  const idxTime = headers.indexOf("Time");
  const idxO = headers.indexOf("Origin");
  const idxD = headers.indexOf("Destination");

  const set = new Set();
  const values = sheet.getDataRange().getValues().slice(1);
  for (const r of values) {
    const key = [
      String(r[idxEmail] || "").trim(),
      String(r[idxDate] || "").trim(),
      String(r[idxTime] || "").trim(),
      String(r[idxO] || "").trim().toLowerCase(),
      String(r[idxD] || "").trim().toLowerCase(),
    ].join("|");
    if (key !== "||||") set.add(key);
  }
  return set;
}

function appendRow_(sheet, headers, obj) {
  sheet.appendRow(headers.map(h => (h in obj ? obj[h] : "")));
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}

function loadIds_(sheet, headers, colName) {
  const idx = headers.indexOf(colName);
  const set = new Set();
  if (idx < 0) return set;
  const values = sheet.getDataRange().getValues().slice(1);
  for (const row of values) {
    const v = String(row[idx] || "").trim();
    if (v) set.add(v);
  }
  return set;
}

function generateId_(set) {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  for (;;) {
    let id = "";
    for (let i = 0; i < 4; i++) id += chars[Math.floor(Math.random() * chars.length)];
    if (!set.has(id)) return id;
  }
}

function extractEmail_(fromField) {
  const m = String(fromField || "").match(/<([^>]+)>/);
  return (m ? m[1] : String(fromField || "")).trim().toLowerCase();
}

function buildQuery_() {
  return `(${CONFIG.authorisedSenders.map(a => `from:${a}`).join(" OR ")}) newer_than:${CONFIG.newerThanDays}d`;
}
