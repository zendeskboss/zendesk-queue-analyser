// =============================================================================
// Zendesk Queue Analyser — Google Apps Script
// Version: 1.1.0
// License: MIT
// Repository: https://github.com/zendeskboss/zendesk-queue-analyser
// =============================================================================

var PROPS = PropertiesService.getUserProperties();

var SHEET_NAMES = {
  CONFIG:    "⚙️ Setup",
  QUEUES:    "📋 Queues",
  WAITTIME:  "⏱️ Wait Times",
  QUEUESIZE: "📊 Queue Size",
  RAW:       "🗂️ Raw Events"
};

var CHANNELS = ["MESSAGING", "TALK", "SUPPORT"];

// ─── MENU ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎯 Queue Analyser")
    .addItem("🔧 Setup credentials", "showCredentialsSidebar")
    .addSeparator()
    .addItem("🔄 Load my queues", "loadQueues")
    .addSeparator()
    .addItem("▶️  Run analysis", "showAnalysisDialog")
    .addSeparator()
    .addItem("🗑️  Clear all results", "clearResults")
    .addItem("❓ Help & documentation", "showHelp")
    .addToUi();
}

// ─── SIDEBAR: CREDENTIALS ─────────────────────────────────────────────────────

function showCredentialsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("🔧 Setup Credentials")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveCredentials(subdomain, email, token) {
  subdomain = subdomain.trim().replace(/^https?:\/\//, "").replace(/\.zendesk\.com.*$/, "");
  email     = email.trim();
  token     = token.trim();

  if (!subdomain || !email || !token) {
    return { success: false, message: "All fields are required." };
  }

  var testUrl = "https://" + subdomain + ".zendesk.com/api/v2/queues";
  var options = buildRequestOptions(email, token);
  try {
    var resp = UrlFetchApp.fetch(testUrl, Object.assign({}, options, { muteHttpExceptions: true }));
    var code = resp.getResponseCode();
    if (code === 401) return { success: false, message: "Authentication failed. Check your email and API token." };
    if (code === 403) return { success: false, message: "Access denied. Make sure the account has admin access and omnichannel routing is enabled." };
    if (code === 404) return { success: false, message: "Subdomain not found. Double-check your Zendesk subdomain." };
    if (code !== 200) return { success: false, message: "Unexpected response (HTTP " + code + "). Check your details and try again." };
  } catch (e) {
    return { success: false, message: "Could not connect to Zendesk.\n\nDetail: " + e.message };
  }

  PROPS.setProperties({ zd_subdomain: subdomain, zd_email: email, zd_token: token });
  ensureSheets();
  return { success: true, message: "Connected to " + subdomain + ".zendesk.com ✓" };
}

function deleteCredentials() {
  PROPS.deleteProperty("zd_subdomain");
  PROPS.deleteProperty("zd_email");
  PROPS.deleteProperty("zd_token");
}

function getStoredCredentials() {
  var subdomain = PROPS.getProperty("zd_subdomain") || "";
  var email     = PROPS.getProperty("zd_email") || "";
  var hasToken  = !!PROPS.getProperty("zd_token");
  return { subdomain: subdomain, email: email, hasToken: hasToken };
}

// ─── SHEET BOOTSTRAP ──────────────────────────────────────────────────────────

function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_NAMES).forEach(function(key) {
    var name = SHEET_NAMES[key];
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  setupConfigSheet(ss.getSheetByName(SHEET_NAMES.CONFIG));
}

function setupConfigSheet(sheet) {
  sheet.clearContents();
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 320);

  sheet.getRange(1, 1, 1, 2).merge()
    .setValue("Zendesk Queue Analyser — Setup Guide")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#1F3864").setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  var instructions = [
    ["", ""],
    ["📋 How to use this tool", ""],
    ["Step 1", "Go to 🎯 Queue Analyser → 🔧 Setup credentials"],
    ["Step 2", "Enter your Zendesk subdomain, admin email & API token"],
    ["Step 3", "Go to 🎯 Queue Analyser → 🔄 Load my queues"],
    ["Step 4", "Go to 🎯 Queue Analyser → ▶️ Run analysis"],
    ["", ""],
    ["⚠️ Requirements", ""],
    ["Plan", "Zendesk Suite Professional, Enterprise, or Enterprise Plus"],
    ["Role", "Admin access on your Zendesk account"],
    ["Feature", "Omnichannel routing must be turned on"],
    ["Data", "Queue events are retained for the past 90 days only"],
    ["", ""],
    ["🔑 Where to find your API token", ""],
    ["In Zendesk", "Admin Center → Apps & integrations → Zendesk API → API tokens"],
    ["", ""],
    ["📞 Support & feedback", ""],
    ["GitHub", "https://github.com/YOUR_USERNAME/zendesk-queue-analyser"]
  ];

  sheet.getRange(2, 1, instructions.length, 2).setValues(instructions);

  [3, 9, 14, 17].forEach(function(r) {
    sheet.getRange(r, 1, 1, 2).merge().setBackground("#D9E1F2").setFontWeight("bold");
  });

  sheet.setFrozenRows(1);
}

// ─── LOAD QUEUES ──────────────────────────────────────────────────────────────

function loadQueues() {
  var creds = getCredentials();
  if (!creds) return;

  var url  = "https://" + creds.subdomain + ".zendesk.com/api/v2/queues";
  var resp = zdFetch(url, creds);
  if (!resp) return;

  var data   = JSON.parse(resp);
  var queues = data.queues || [];

  if (queues.length === 0) {
    SpreadsheetApp.getUi().alert("No queues found", "No active queues were found. Make sure omnichannel routing is on and at least one queue exists.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.QUEUES) || ss.insertSheet(SHEET_NAMES.QUEUES);
  sheet.clearContents();
  sheet.clearFormats();

  var headerRow = ["Queue ID", "Queue Name", "Description", "Priority", "Order", "Primary Groups", "Created", "Updated"];
  sheet.getRange(1, 1, 1, headerRow.length)
    .setValues([headerRow])
    .setBackground("#1F3864").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);

  [200, 200, 250, 80, 60, 200, 180, 180].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  var rows = queues.map(function(q) {
    var groups = (q.primary_groups && q.primary_groups.groups || []).map(function(g) { return g.name; }).join(", ");
    return [q.id, q.name, q.description || "", q.priority, q.order, groups, friendlyDate(q.created_at), friendlyDate(q.updated_at)];
  });

  sheet.getRange(2, 1, rows.length, headerRow.length).setValues(rows);
  for (var i = 0; i < rows.length; i++) {
    sheet.getRange(i + 2, 1, 1, headerRow.length).setBackground(i % 2 === 0 ? "#FFFFFF" : "#EEF2FF");
  }

  ss.setActiveSheet(sheet);
  ss.toast(queues.length + " queue(s) loaded. Now run ▶️ Run analysis.", "✅ Queues loaded", 6);
}

// ─── ANALYSIS DIALOG ──────────────────────────────────────────────────────────

function showAnalysisDialog() {
  var creds = getCredentials(true);
  if (!creds) return;
  var html = HtmlService.createHtmlOutputFromFile("AnalysisDialog").setWidth(480).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "▶️ Run Queue Analysis");
}

function getQueueListForDialog() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.QUEUES);
  if (!sheet || sheet.getLastRow() < 2) {
    return { error: "No queues loaded yet. Close this dialog and run 🎯 Queue Analyser → 🔄 Load my queues first." };
  }
  var data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var queues = data.filter(function(r) { return r[0]; }).map(function(r) { return { id: r[0], name: r[1] }; });
  return { queues: queues };
}

function runAnalysis(params) {
  var creds = getCredentials();
  if (!creds) return { error: "No credentials found. Please set up your credentials first." };

  try {
    var startTs  = new Date(params.startDate + "T00:00:00Z").getTime() / 1000;
    var endTs    = new Date(params.endDate   + "T23:59:59Z").getTime() / 1000;
    var daysDiff = (endTs - startTs) / 86400;

    if (daysDiff > 90) return { error: "Date range cannot exceed 90 days (Zendesk only retains 90 days of queue event data)." };
    if (daysDiff < 0)  return { error: "End date must be after start date." };

    var queueNameMap = buildQueueNameMap();
    var allEvents    = [];
    var queueIds     = params.queueIds;
    var channels     = params.channels;

    var fetchTargets = [];
    if (queueIds.length === 0) {
      channels.forEach(function(ch) { fetchTargets.push({ queueId: null, channel: ch }); });
    } else {
      queueIds.forEach(function(qid) {
        channels.forEach(function(ch) { fetchTargets.push({ queueId: qid, channel: ch }); });
      });
    }

    fetchTargets.forEach(function(target) {
      var events = fetchQueueEvents(creds, target.queueId, target.channel, startTs, endTs);
      allEvents  = allEvents.concat(events);
    });

    if (allEvents.length === 0) {
      return { error: "No queue events found for the selected period, queues, and channels. Try widening your date range or selecting different queues." };
    }

    writeRawEvents(allEvents, queueNameMap);

    var waitMetrics = calcWaitTimeMetrics(allEvents);
    writeWaitTimeSheet(waitMetrics, params, queueNameMap);

    var sizeMetrics = calcQueueSizeMetrics(allEvents);
    writeQueueSizeSheet(sizeMetrics, params, queueNameMap);

    return { success: true, eventCount: allEvents.length };

  } catch (e) {
    return { error: "An error occurred during analysis: " + e.message };
  }
}

// ─── QUEUE NAME HELPERS ───────────────────────────────────────────────────────

function buildQueueNameMap() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.QUEUES);
  var map   = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(function(row) { if (row[0]) map[row[0]] = row[1] || row[0]; });
  return map;
}

function queueLabel(queueId, queueNameMap) {
  return (queueNameMap && queueNameMap[queueId]) ? queueNameMap[queueId] : queueId;
}

// ─── DATA FETCHING ────────────────────────────────────────────────────────────

function fetchQueueEvents(creds, queueId, channel, startTs, endTs) {
  var baseUrl = "https://" + creds.subdomain + ".zendesk.com/api/v2/queue_events";
  var params  = ["channel=" + channel, "start_time=" + startTs, "end_time=" + endTs, "page_size=100"];
  if (queueId) params.push("queue_id=" + encodeURIComponent(queueId));

  var url       = baseUrl + "?" + params.join("&");
  var allEvents = [];
  var pageCount = 0;
  var maxPages  = 500;

  while (url && pageCount < maxPages) {
    var resp = zdFetch(url, creds);
    if (!resp) break;
    var data = JSON.parse(resp);
    allEvents = allEvents.concat(data.queue_data || []);
    pageCount++;
    url = (data.has_more && data.next_page_url) ? data.next_page_url : null;
  }

  return allEvents;
}

// ─── METRICS: WAIT TIMES ──────────────────────────────────────────────────────

function calcWaitTimeMetrics(events) {
  // OUTBOUND events with both time fields = a ticket left the queue; we can measure how long it waited
  var eligible = events.filter(function(e) {
    return e.event_type === "OUTBOUND" && e.start_queue_time && e.end_queue_time;
  });

  eligible.forEach(function(e) {
    e._waitSecs = (new Date(e.end_queue_time) - new Date(e.start_queue_time)) / 1000;
    e._date     = e.event_occurred_at ? e.event_occurred_at.substring(0, 10) : "";
  });

  // Count unique tickets entering queues per day (INBOUND events)
  var inbound = events.filter(function(e) { return e.event_type === "INBOUND"; });
  inbound.forEach(function(e) { e._date = e.event_occurred_at ? e.event_occurred_at.substring(0, 10) : ""; });

  var ticketsPerDay = {};
  inbound.forEach(function(e) {
    if (!e._date) return;
    if (!ticketsPerDay[e._date]) ticketsPerDay[e._date] = {};
    if (e.ticket_id) ticketsPerDay[e._date][e.ticket_id] = true;
  });

  var ticketCountByDay = {};
  Object.keys(ticketsPerDay).forEach(function(d) {
    ticketCountByDay[d] = Object.keys(ticketsPerDay[d]).length;
  });

  return {
    all:              aggregateWaitTimes(eligible),
    byQueue:          groupBy(eligible, "queue_id", aggregateWaitTimes),
    byChannel:        groupBy(eligible, "channel",  aggregateWaitTimes),
    byDay:            groupBy(eligible, "_date",    aggregateWaitTimes),
    ticketCountByDay: ticketCountByDay,
    eligible:         eligible
  };
}

function aggregateWaitTimes(events) {
  if (!events || events.length === 0) return { count: 0, avg: 0, max: 0, min: 0 };
  var waits = events.map(function(e) { return e._waitSecs; });
  var sum   = waits.reduce(function(a, b) { return a + b; }, 0);
  return {
    count: events.length,
    avg:   sum / events.length,
    max:   Math.max.apply(null, waits),
    min:   Math.min.apply(null, waits)
  };
}

// ─── METRICS: QUEUE SIZE ──────────────────────────────────────────────────────

function calcQueueSizeMetrics(events) {
  // current_channel_queue_size is a snapshot of queue depth at the moment of each event.
  // Use ALL events that carry the field — but exclude size=0 from the minimum
  // (a size of 0 just means the queue briefly emptied, not a meaningful floor for ops purposes).
  var sizeEvents = events.filter(function(e) {
    return e.current_channel_queue_size !== undefined && e.current_channel_queue_size !== null;
  });
  sizeEvents.forEach(function(e) {
    e._date = e.event_occurred_at ? e.event_occurred_at.substring(0, 10) : "";
  });

  // hourly_channel_queue_size_seconds must be SUMMED per hour-bucket first, then divided by 3600.
  // Per the Zendesk docs, each event contributes a positive (INBOUND) or negative (OUTBOUND)
  // number of seconds to the running total for that hour. Summing them gives the net
  // queue-size-seconds for the hour. Dividing by 3600 gives the hourly average queue size.
  // We must NOT average the raw per-event values — that cancels out positives and negatives.
  var hourBuckets = {}; // key: "YYYY-MM-DD HH" → { date, queueId, channel, totalSecs }
  events.forEach(function(e) {
    if (e.hourly_channel_queue_size_seconds === undefined || e.hourly_channel_queue_size_seconds === null) return;
    if (!e.event_occurred_at) return;

    // Bucket by hour
    var hourKey = e.event_occurred_at.substring(0, 13); // "YYYY-MM-DDTHH"
    var dateKey = e.event_occurred_at.substring(0, 10);

    // We need separate buckets per queue and channel to support grouping
    var bucketKey = (e.queue_id || "all") + "|" + (e.channel || "all") + "|" + hourKey;

    if (!hourBuckets[bucketKey]) {
      hourBuckets[bucketKey] = {
        date:      dateKey,
        queue_id:  e.queue_id  || "Unknown",
        channel:   e.channel   || "Unknown",
        totalSecs: 0
      };
    }
    hourBuckets[bucketKey].totalSecs += e.hourly_channel_queue_size_seconds;
  });

  // Convert each bucket to a hourly avg, keeping only non-negative results
  // (a negative total just means more left than entered in that hour — treat as 0)
  var hourlyRows = Object.keys(hourBuckets).map(function(k) {
    var b    = hourBuckets[k];
    var secs = Math.max(0, b.totalSecs); // floor at 0
    return {
      date:       b.date,
      queue_id:   b.queue_id,
      channel:    b.channel,
      hourlyAvg:  secs / 3600
    };
  });

  return {
    all:       aggregateQueueSize(sizeEvents, hourlyRows),
    byQueue:   groupBySize(sizeEvents, hourlyRows, "queue_id"),
    byChannel: groupBySize(sizeEvents, hourlyRows, "channel"),
    byDay:     groupBySize(sizeEvents, hourlyRows, "date")
  };
}

function aggregateQueueSize(sizeEvents, hourlyRows) {
  var result = { count: sizeEvents.length, min: null, max: null, avgHourly: null };

  if (sizeEvents.length > 0) {
    var sizes = sizeEvents.map(function(e) { return e.current_channel_queue_size; });
    var nonZero = sizes.filter(function(s) { return s > 0; });
    result.max = Math.max.apply(null, sizes);
    result.min = nonZero.length > 0 ? Math.min.apply(null, nonZero) : 0;
  }

  if (hourlyRows.length > 0) {
    // Daily average = sum of hourly avgs for the day / 24
    // Overall average = sum of all hourly avgs / total hours
    var sum = hourlyRows.reduce(function(a, r) { return a + r.hourlyAvg; }, 0);
    result.avgHourly = sum / hourlyRows.length;
  }

  return result;
}

function groupBySize(sizeEvents, hourlyRows, key) {
  var groups = {}, hGroups = {};

  sizeEvents.forEach(function(e) {
    var k = e[key] || "Unknown";
    if (!groups[k]) groups[k] = [];
    groups[k].push(e);
  });

  hourlyRows.forEach(function(r) {
    var k = r[key] || "Unknown";
    if (!hGroups[k]) hGroups[k] = [];
    hGroups[k].push(r);
  });

  var allKeys = Object.keys(groups);
  Object.keys(hGroups).forEach(function(k) { if (allKeys.indexOf(k) < 0) allKeys.push(k); });

  var result = {};
  allKeys.forEach(function(k) {
    result[k] = aggregateQueueSize(groups[k] || [], hGroups[k] || []);
  });
  return result;
}

// ─── SHEET WRITERS ────────────────────────────────────────────────────────────

function writeWaitTimeSheet(metrics, params, queueNameMap) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.WAITTIME) || ss.insertSheet(SHEET_NAMES.WAITTIME);
  sheet.clearContents();
  sheet.clearFormats();
  sheet.clearNotes();
  sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });

  var row = 1;
  writeSheetTitle(sheet, "⏱️ Wait Times", params, row);
  row += 3;

  // ── Plain-English summary ──
  var overall      = metrics.all;
  var totalTickets = Object.keys(metrics.ticketCountByDay).reduce(function(sum, d) {
    return sum + metrics.ticketCountByDay[d];
  }, 0);

  row = writeSectionHeader(sheet, "📌 Summary — What you need to know", row, "#FFF3CD", "#856404");
  var summaryLines = [
    ["What is a wait time?",        "How long a ticket sat in a queue before an agent accepted it. Lower is better."],
    ["What is an event?",           "Every time a ticket enters or leaves a queue it creates an event. One ticket can have multiple events if it moves between queues. Wait time is only measured when a ticket leaves (e.g. accepted by an agent)."],
    ["Tickets entering queues",     totalTickets + " unique tickets entered queues in this period"],
    ["Overall average wait",        overall.count > 0 ? formatDuration(overall.avg) + " per ticket" : "No data"],
    ["Longest any ticket waited",   overall.count > 0 ? formatDuration(overall.max) : "No data"]
  ];
  row = writeSummaryTable(sheet, summaryLines, row);
  row++;

  // ── By queue ──
  if (Object.keys(metrics.byQueue).length > 0) {
    row = writeSectionHeader(sheet, "By Queue", row, "#D9E1F2", "#1F3864");
    var qRows = Object.keys(metrics.byQueue).map(function(qid) {
      var m = metrics.byQueue[qid];
      return [queueLabel(qid, queueNameMap), m.count, formatDuration(m.avg), formatDuration(m.max), formatDuration(m.min)];
    });
    row = writeTable(sheet, ["Queue", "Tickets Processed", "Avg Wait", "Max Wait", "Min Wait"], qRows, row, "#D9E1F2");
    row++;
  }

  // ── By channel ──
  if (Object.keys(metrics.byChannel).length > 0) {
    row = writeSectionHeader(sheet, "By Channel", row, "#D9E1F2", "#1F3864");
    var chRows = Object.keys(metrics.byChannel).map(function(ch) {
      var m = metrics.byChannel[ch];
      return [ch, m.count, formatDuration(m.avg), formatDuration(m.max), formatDuration(m.min)];
    });
    row = writeTable(sheet, ["Channel", "Tickets Processed", "Avg Wait", "Max Wait", "Min Wait"], chRows, row, "#D9E1F2");
    row++;
  }

  // ── Daily breakdown + chart ──
  var dayKeys = Object.keys(metrics.byDay).sort();
  if (dayKeys.length > 0) {
    row = writeSectionHeader(sheet, "Daily Breakdown", row, "#D9E1F2", "#1F3864");
    var chartDataStartRow = row;

    sheet.getRange(row, 1, 1, 4)
      .setValues([["Date", "Tickets Entering Queue", "Avg Wait (mins)", "Max Wait (mins)"]])
      .setBackground("#D9E1F2").setFontWeight("bold");
    row++;

    var dayDataRows = dayKeys.map(function(d) {
      var m       = metrics.byDay[d];
      var tickets = metrics.ticketCountByDay[d] || 0;
      return [
        friendlyDate(d),
        tickets,
        m.avg > 0 ? Math.round(m.avg / 60 * 10) / 10 : 0,
        m.max > 0 ? Math.round(m.max / 60 * 10) / 10 : 0
      ];
    });

    sheet.getRange(row, 1, dayDataRows.length, 4).setValues(dayDataRows);
    for (var i = 0; i < dayDataRows.length; i++) {
      sheet.getRange(row + i, 1, 1, 4).setBackground(i % 2 === 0 ? "#FFFFFF" : "#F5F7FF");
    }
    row += dayDataRows.length;

    if (dayDataRows.length > 1) {
      addWaitTimeChart(sheet, sheet.getRange(chartDataStartRow, 1, dayDataRows.length + 1, 4), chartDataStartRow);
    }
    row++;
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 180);
  [3, 4, 5].forEach(function(c) { sheet.setColumnWidth(c, 140); });

  ss.setActiveSheet(sheet);
}

function writeQueueSizeSheet(metrics, params, queueNameMap) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.QUEUESIZE) || ss.insertSheet(SHEET_NAMES.QUEUESIZE);
  sheet.clearContents();
  sheet.clearFormats();
  sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });

  var row = 1;
  writeSheetTitle(sheet, "📊 Queue Size", params, row);
  row += 3;

  // Detect whether the selected channels can actually provide queue size data.
  // Zendesk only populates current_channel_queue_size and hourly_channel_queue_size_seconds
  // for MESSAGING and TALK. SUPPORT tickets are handled differently and these fields
  // are always 0 for Support channel events.
  var hasNonZeroSize    = metrics.all.max !== null && metrics.all.max > 0;
  var onlySupportQueried = params.channels.length > 0 &&
    params.channels.every(function(ch) { return ch === "SUPPORT"; });

  // ── Plain-English summary ──
  var overall = metrics.all;
  row = writeSectionHeader(sheet, "📌 Summary — What you need to know", row, "#FFF3CD", "#856404");

  if (!hasNonZeroSize) {
    // Show a clear notice rather than misleading zeros
    var noDataReason = onlySupportQueried
      ? "Queue size data is not available for the Support channel. Zendesk only tracks queue size metrics for Messaging and Talk. Try re-running the analysis with Messaging or Talk selected."
      : "No queue size data was returned for this period. This may be because all selected channels are Support-only, or because no tickets were waiting in queue during this time.";

    var summaryLines = [
      ["What is queue size?",  "The number of tickets sitting in a queue waiting for an agent at any moment."],
      ["⚠️ No queue size data", noDataReason]
    ];
    row = writeSummaryTable(sheet, summaryLines, row);
    sheet.getRange(row - 1 - summaryLines.length + 2, 1, 1, 5)
      .setBackground("#FDE8E8").setFontColor("#c0392b");
    sheet.setColumnWidth(1, 220);
    [2, 3, 4].forEach(function(c) { sheet.setColumnWidth(c, 160); });
    return;
  }

  var summaryLines = [
    ["What is queue size?",  "The number of tickets sitting in a queue waiting for an agent at any moment. A growing queue means demand is outpacing capacity."],
    ["Why does it matter?",  "Peaks show when your team is most stretched. Use this alongside wait times to make staffing and routing decisions."],
    ["Peak queue size",      overall.max + " tickets waiting at once"],
    ["Average queue size",   overall.avgHourly !== null ? overall.avgHourly.toFixed(1) + " tickets on average (hourly)" : "No data"]
  ];
  row = writeSummaryTable(sheet, summaryLines, row);
  row++;

  // ── By queue ──
  if (Object.keys(metrics.byQueue).length > 0) {
    row = writeSectionHeader(sheet, "By Queue", row, "#D5E8D4", "#2d7a4f");
    var qRows = Object.keys(metrics.byQueue).map(function(qid) {
      var m = metrics.byQueue[qid];
      return [queueLabel(qid, queueNameMap), m.min, m.max, m.avgHourly !== null ? m.avgHourly.toFixed(1) : "N/A"];
    });
    row = writeTable(sheet, ["Queue", "Min Size", "Max Size", "Avg Size (Hourly)"], qRows, row, "#D5E8D4");
    row++;
  }

  // ── By channel ──
  if (Object.keys(metrics.byChannel).length > 0) {
    row = writeSectionHeader(sheet, "By Channel", row, "#D5E8D4", "#2d7a4f");
    var chRows = Object.keys(metrics.byChannel).map(function(ch) {
      var m = metrics.byChannel[ch];
      return [ch, m.min, m.max, m.avgHourly !== null ? m.avgHourly.toFixed(1) : "N/A"];
    });
    row = writeTable(sheet, ["Channel", "Min Size", "Max Size", "Avg Size (Hourly)"], chRows, row, "#D5E8D4");
    row++;
  }

  // ── Daily breakdown + chart ──
  var dayKeys = Object.keys(metrics.byDay).sort();
  if (dayKeys.length > 0) {
    row = writeSectionHeader(sheet, "Daily Breakdown", row, "#D5E8D4", "#2d7a4f");
    var chartDataStartRow = row;

    sheet.getRange(row, 1, 1, 3)
      .setValues([["Date", "Max Queue Size", "Avg Queue Size (Hourly)"]])
      .setBackground("#D5E8D4").setFontWeight("bold");
    row++;

    var dayDataRows = dayKeys.map(function(d) {
      var m = metrics.byDay[d];
      return [
        friendlyDate(d),
        m.max !== null ? m.max : 0,
        m.avgHourly !== null ? Math.round(m.avgHourly * 10) / 10 : 0
      ];
    });

    sheet.getRange(row, 1, dayDataRows.length, 3).setValues(dayDataRows);
    for (var i = 0; i < dayDataRows.length; i++) {
      sheet.getRange(row + i, 1, 1, 3).setBackground(i % 2 === 0 ? "#FFFFFF" : "#F5F7FF");
    }
    row += dayDataRows.length;

    if (dayDataRows.length > 1) {
      addQueueSizeChart(sheet, sheet.getRange(chartDataStartRow, 1, dayDataRows.length + 1, 3), chartDataStartRow);
    }
  }

  sheet.setColumnWidth(1, 220);
  [2, 3, 4].forEach(function(c) { sheet.setColumnWidth(c, 160); });
}

function writeRawEvents(events, queueNameMap) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.RAW) || ss.insertSheet(SHEET_NAMES.RAW);
  sheet.clearContents();
  sheet.clearFormats();

  var headers = [
    "Event ID", "Queue Name", "Queue ID", "Channel", "Event Type",
    "Inbound Reason", "Outbound Reason", "Ticket ID",
    "Event Occurred At", "Start Queue Time", "End Queue Time",
    "Wait Time (mins)", "Current Queue Size", "Hourly Queue Size Secs"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground("#1F3864").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);

  if (events.length > 0) {
    var rows = events.map(function(e) {
      var waitMins = (e.event_type === "OUTBOUND" && e.start_queue_time && e.end_queue_time)
        ? Math.round((new Date(e.end_queue_time) - new Date(e.start_queue_time)) / 1000 / 60 * 10) / 10
        : "";
      return [
        e.queue_event_id || "",
        queueLabel(e.queue_id, queueNameMap),
        e.queue_id || "",
        e.channel || "",
        e.event_type || "",
        e.inbound_reason || "",
        e.outbound_reason || "",
        // Prefix ticket_id with apostrophe to force plain text — prevents Sheets
        // from auto-formatting large integers as date serial numbers
        e.ticket_id ? "'" + e.ticket_id : "",
        friendlyDate(e.event_occurred_at),
        friendlyDate(e.start_queue_time),
        friendlyDate(e.end_queue_time),
        waitMins,
        e.current_channel_queue_size !== undefined ? e.current_channel_queue_size : "",
        e.hourly_channel_queue_size_seconds !== undefined ? e.hourly_channel_queue_size_seconds : ""
      ];
    });

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // Force ticket ID column (col 8) to plain text format so Sheets never
    // interprets numeric IDs as date serials
    sheet.getRange(2, 8, rows.length, 1).setNumberFormat("@");

    for (var i = 0; i < rows.length; i++) {
      sheet.getRange(i + 2, 1, 1, headers.length).setBackground(i % 2 === 0 ? "#FFFFFF" : "#F8F9FF");
    }
  }

  [120, 180, 140, 100, 150, 140, 150, 90, 170, 170, 170, 120, 130, 150].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
}

// ─── SHEET HELPERS ────────────────────────────────────────────────────────────

function writeSheetTitle(sheet, title, params, row) {
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue(title)
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#1F3864").setFontColor("#FFFFFF")
    .setHorizontalAlignment("left");

  var subtitle = "Period: " + friendlyDate(params.startDate) + " to " + friendlyDate(params.endDate) +
    "  |  Channels: " + params.channels.join(", ") +
    "  |  Queues: " + (params.queueIds.length === 0 ? "All" : params.queueIds.length + " selected");

  sheet.getRange(row + 1, 1, 1, 5).merge()
    .setValue(subtitle)
    .setFontSize(10).setFontColor("#666666").setFontStyle("italic");

  sheet.getRange(row + 2, 1, 1, 5).merge()
    .setValue("Generated: " + friendlyDate(new Date().toISOString()))
    .setFontSize(9).setFontColor("#999999");
}

function writeSectionHeader(sheet, label, row, bgColor, fgColor) {
  sheet.getRange(row, 1, 1, 5).merge()
    .setValue(label)
    .setFontWeight("bold").setFontSize(11)
    .setBackground(bgColor || "#D9E1F2")
    .setFontColor(fgColor || "#1F3864");
  return row + 1;
}

function writeSummaryTable(sheet, data, row) {
  data.forEach(function(pair, i) {
    sheet.getRange(row + i, 1).setValue(pair[0]).setFontWeight("bold");
    sheet.getRange(row + i, 2, 1, 4).merge().setValue(pair[1]).setWrap(true);
    sheet.getRange(row + i, 1, 1, 5).setBackground(i % 2 === 0 ? "#FAFAFA" : "#FFFFFF");
  });
  return row + data.length + 1;
}

function writeTable(sheet, headers, rows, row, headerBg) {
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers]).setBackground(headerBg || "#D9E1F2").setFontWeight("bold");
  row++;

  if (rows.length > 0) {
    sheet.getRange(row, 1, rows.length, headers.length).setValues(rows);
    for (var i = 0; i < rows.length; i++) {
      sheet.getRange(row + i, 1, 1, headers.length).setBackground(i % 2 === 0 ? "#FFFFFF" : "#F5F7FF");
    }
    row += rows.length;
  } else {
    sheet.getRange(row, 1, 1, headers.length).merge()
      .setValue("No data available for this grouping.")
      .setFontStyle("italic").setFontColor("#999999");
    row++;
  }
  return row;
}

// ─── CHARTS ───────────────────────────────────────────────────────────────────

function addWaitTimeChart(sheet, dataRange, anchorRow) {
  try {
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dataRange)
      .setPosition(anchorRow, 6, 0, 0)
      .setOption("title", "Daily Wait Times")
      .setOption("width", 580)
      .setOption("height", 320)
      .setOption("series", {
        0: { type: "bars",  labelInLegend: "Tickets entering queue", targetAxisIndex: 1, color: "#BDD7EE" },
        1: { type: "line",  labelInLegend: "Avg wait (mins)",        targetAxisIndex: 0, color: "#2E75B6", lineWidth: 3 },
        2: { type: "line",  labelInLegend: "Max wait (mins)",        targetAxisIndex: 0, color: "#C55A11", lineWidth: 2, lineDashStyle: "4 4" }
      })
      .setOption("vAxes", {
        0: { title: "Wait time (mins)", minValue: 0 },
        1: { title: "Tickets" }
      })
      .setOption("curveType", "function")
      .setOption("legend", { position: "bottom" })
      .build();
    sheet.insertChart(chart);
  } catch(e) {
    Logger.log("Wait time chart failed: " + e.message);
  }
}

function addQueueSizeChart(sheet, dataRange, anchorRow) {
  try {
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dataRange)
      .setPosition(anchorRow, 5, 0, 0)
      .setOption("title", "Daily Queue Size")
      .setOption("width", 580)
      .setOption("height", 320)
      .setOption("series", {
        0: { type: "bars", labelInLegend: "Max queue size",          color: "#F4B183" },
        1: { type: "line", labelInLegend: "Avg queue size (hourly)", color: "#2E75B6", lineWidth: 3 }
      })
      .setOption("vAxes", { 0: { title: "Tickets in queue", minValue: 0 } })
      .setOption("curveType", "function")
      .setOption("legend", { position: "bottom" })
      .build();
    sheet.insertChart(chart);
  } catch(e) {
    Logger.log("Queue size chart failed: " + e.message);
  }
}

// ─── CLEAR / HELP ─────────────────────────────────────────────────────────────

function clearResults() {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.alert("Clear all results?",
    "This will clear the Wait Times, Queue Size, Queues, and Raw Events sheets. Your credentials and Setup sheet will not be affected.\n\nContinue?",
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_NAMES.WAITTIME, SHEET_NAMES.QUEUESIZE, SHEET_NAMES.QUEUES, SHEET_NAMES.RAW].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) { sheet.clearContents(); sheet.clearFormats(); }
  });
  ss.toast("All results cleared.", "✅ Done", 3);
}

function showHelp() {
  SpreadsheetApp.getUi().alert(
    "❓ Zendesk Queue Analyser — Help",
    "WHAT THIS TOOL DOES\n" +
    "Analyses wait times and queue sizes for your Zendesk Omnichannel Routing queues.\n\n" +
    "HOW TO USE IT\n" +
    "1. Setup credentials — subdomain, admin email, API token.\n" +
    "2. Load queues — fetches your active queue list.\n" +
    "3. Run analysis — pick a date range (max 90 days), queues, and channels.\n\n" +
    "UNDERSTANDING EVENTS\n" +
    "Every time a ticket enters or leaves a queue, it creates an event. One ticket can create several events if it moves between queues. Wait time is only calculated for OUTBOUND events — when a ticket left the queue (e.g. accepted by an agent).\n\n" +
    "SHEETS EXPLAINED\n" +
    "⚙️ Setup: getting started guide\n" +
    "📋 Queues: your queue list\n" +
    "⏱️ Wait Times: how long tickets waited before being picked up\n" +
    "📊 Queue Size: how many tickets were in queue at any time\n" +
    "🗂️ Raw Events: event-level data for custom analysis\n\n" +
    "MORE HELP\n" +
    "https://github.com/YOUR_USERNAME/zendesk-queue-analyser",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─── UTILITIES ────────────────────────────────────────────────────────────────

function getCredentials(silent) {
  var subdomain = PROPS.getProperty("zd_subdomain");
  var email     = PROPS.getProperty("zd_email");
  var token     = PROPS.getProperty("zd_token");
  if (!subdomain || !email || !token) {
    if (!silent) SpreadsheetApp.getUi().alert("No credentials found", "Go to 🎯 Queue Analyser → 🔧 Setup credentials first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
  return { subdomain: subdomain, email: email, token: token };
}

function buildRequestOptions(email, token) {
  return {
    method: "GET",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(email + "/token:" + token),
      "Content-Type":  "application/json"
    },
    muteHttpExceptions: true
  };
}

function zdFetch(url, creds) {
  var options = buildRequestOptions(creds.email, creds.token);
  var resp;
  try {
    resp = UrlFetchApp.fetch(url, options);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Network error", "Could not connect to Zendesk.\n\n" + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
  var code = resp.getResponseCode();
  if (code === 200) return resp.getContentText();

  var errMessages = {
    401: "Authentication failed. Your API token may have expired. Please re-enter your credentials.",
    403: "Access denied. Check you have admin access and omnichannel routing is enabled.",
    404: "Endpoint not found. Check your subdomain is correct.",
    429: "Rate limit hit. Zendesk allows 30 requests per minute for queue endpoints. Wait a minute and try again."
  };
  var msg = errMessages[code] || "Unexpected error (HTTP " + code + "):\n" + resp.getContentText().substring(0, 300);
  SpreadsheetApp.getUi().alert("Zendesk API Error", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return null;
}

function formatDuration(secs) {
  if (!secs && secs !== 0) return "N/A";
  secs = Math.round(secs);
  if (secs < 60)   return secs + "s";
  if (secs < 3600) return Math.floor(secs / 60) + "m " + (secs % 60) + "s";
  return Math.floor(secs / 3600) + "h " + Math.floor((secs % 3600) / 60) + "m";
}

/**
 * Converts any ISO date string or YYYY-MM-DD to "2 April 2026" (UTC).
 * Safe to call with null/empty — returns "".
 */
function friendlyDate(isoString) {
  if (!isoString) return "";
  var d = new Date(isoString);
  if (isNaN(d.getTime())) return String(isoString);
  var months = ["January","February","March","April","May","June",
                "July","August","September","October","November","December"];
  return d.getUTCDate() + " " + months[d.getUTCMonth()] + " " + d.getUTCFullYear();
}

function groupBy(arr, key, aggregateFn) {
  var groups = {};
  arr.forEach(function(item) {
    var k = item[key] || "Unknown";
    if (!groups[k]) groups[k] = [];
    groups[k].push(item);
  });
  var result = {};
  Object.keys(groups).forEach(function(k) { result[k] = aggregateFn(groups[k]); });
  return result;
}
