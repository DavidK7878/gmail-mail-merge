/**
 * Open & click tracking.
 *
 * The script must be deployed as a Web App (Execute as: Me, Access: Anyone)
 * and the resulting URL pasted into Settings → WebAppUrl.
 *
 *   ?a=o&t=<trackingId>            → open pixel hit
 *   ?a=c&t=<trackingId>&u=<url>    → click, then redirect to <url>
 */

/** Rewrite links and append the open pixel. */
function injectTracking_(html, trackingId, settings) {
  var base = settings.WebAppUrl;
  if (!base) return html;

  if (settings.TrackClicks) {
    html = html.replace(/href="(https?:\/\/[^"]+)"/gi, function (match, dest) {
      if (dest.indexOf(base) === 0) return match;            // already tracked
      if (/unsubscribe/i.test(dest)) return match;           // leave unsubscribe links pristine
      return 'href="' + base + '?a=c&t=' + trackingId + '&u=' + encodeURIComponent(dest) + '"';
    });
  }

  if (settings.TrackOpens) {
    html += '<img src="' + base + '?a=o&t=' + trackingId +
      '" width="1" height="1" alt="" style="width:1px;height:1px;border:0;">';
  }
  return html;
}

/** Web App entry point — every pixel load and tracked click lands here. */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.a || '';
  var trackingId = p.t || '';
  var dest = p.u || '';

  try {
    if (trackingId && (action === 'o' || action === 'c')) {
      recordTrackingEvent_(action, trackingId);
    }
  } catch (err) {
    // Never let a logging failure break the redirect for a real human clicking a link.
  }

  if (action === 'c' && /^https?:\/\//i.test(dest)) {
    var safe = dest.replace(/"/g, '%22').replace(/</g, '%3C').replace(/>/g, '%3E');
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head>' +
      '<meta http-equiv="refresh" content="0; url=' + safe + '">' +
      '</head><body>' +
      '<script>window.top.location.href = ' + JSON.stringify(dest) + ';</script>' +
      '<p>Redirecting… <a href="' + safe + '">continue</a></p>' +
      '</body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Pixel (or malformed) request: an empty response is fine — the hit is already logged,
  // and the img is 1×1 so nothing visible renders in the recipient's client.
  return HtmlService.createHtmlOutput('');
}

/** Find the row by TrackingId and bump its counters. Lock-guarded: doGet runs concurrently. */
function recordTrackingEvent_(action, trackingId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getContactsSheet_();
    var map = colMap_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var ids = sheet.getRange(2, map['trackingid'], lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === trackingId) {
        var row = i + 2;
        var email = String(cell_(sheet, map, row, 'Email').getValue());
        if (action === 'o') {
          var opens = Number(cell_(sheet, map, row, 'Opens').getValue()) || 0;
          cell_(sheet, map, row, 'Opens').setValue(opens + 1);
          cell_(sheet, map, row, 'LastOpenAt').setValue(new Date());
          logEvent('OPEN', email, 'Open #' + (opens + 1));
        } else {
          var clicks = Number(cell_(sheet, map, row, 'Clicks').getValue()) || 0;
          cell_(sheet, map, row, 'Clicks').setValue(clicks + 1);
          logEvent('CLICK', email, 'Click #' + (clicks + 1));
        }
        return;
      }
    }
  } finally {
    lock.releaseLock();
  }
}
