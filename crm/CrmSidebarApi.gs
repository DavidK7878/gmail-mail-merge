/**
 * Server-side functions called from CrmSidebar.html via google.script.run.
 */

function crmGetSidebarData() {
  var out = { contacts: 0, unclassified: 0, review: 0, needsReply: 0, dormant: 0, importRows: 0, importNew: 0, importPushed: 0 };
  try {
    var master = crmSheet_('Master');
    var map = crmColMap_(master, CRM_MASTER_COLUMNS);
    var lastRow = master.getLastRow();
    if (lastRow >= 2) {
      var data = master.getRange(2, 1, lastRow - 1, master.getLastColumn()).getValues();
      data.forEach(function (r) {
        if (!String(r[map['email'] - 1]).trim()) return;
        out.contacts++;
        if (!String(r[map['category'] - 1]).trim()) out.unclassified++;
        if (crmBool_(r[map['reviewneeded'] - 1])) out.review++;
        if (crmBool_(r[map['needsreply'] - 1])) out.needsReply++;
        if (String(r[map['stage'] - 1]) === 'Dormant') out.dormant++;
      });
    }
    var imp = crmSheet_('Import');
    var imap = crmColMap_(imp, ['Email']);
    if (imp.getLastRow() >= 2) {
      var idata = imp.getRange(2, 1, imp.getLastRow() - 1, imp.getLastColumn()).getValues();
      idata.forEach(function (r) {
        if (!String(r[imap['email'] - 1]).trim()) return;
        out.importRows++;
        if (imap['dedupestatus'] && String(r[imap['dedupestatus'] - 1]) === 'NEW') {
          if (imap['pushedat'] && r[imap['pushedat'] - 1]) out.importPushed++; else out.importNew++;
        }
      });
    }
  } catch (e) {
    out.error = String(e.message || e);
  }
  var scan = crmScanStatus_();
  var cls = crmClassifyStatus_();
  var settings = null;
  try { settings = crmSettings_(); } catch (e) { /* not set up yet */ }
  return {
    counts: out,
    scan: scan,
    classify: cls,
    hasApiKey: !!crmGetApiKey_(),
    model: settings ? settings.Model : '',
    hasCompanyDescription: !!(settings && settings.CompanyDescription),
    scanSince: settings ? Utilities.formatDate(settings.ScanSince, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    me: crmMyEmail_()
  };
}

function crmSidebarStartScan(full) { return crmStartScan(!!full); }
function crmSidebarStopScan() { return crmStopScan(); }
function crmSidebarClassify(mode) { return crmStartClassify_(mode === 'all' ? 'all' : 'unclassified'); }
function crmSidebarStopClassify() { return crmStopClassify(); }
function crmSidebarRecompute() { return crmRecomputePipeline(); }
function crmSidebarSync() { return crmSyncFromMailMerge(); }
function crmSidebarCheckImport() { return crmCheckImport(); }
function crmSidebarPush(includeSameCompany) { return crmPushToMailMerge(includeSameCompany ? ['NEW', 'SAME_COMPANY'] : ['NEW']); }
function crmSidebarImportDrive(ref) { return crmImportFromDrive(ref); }
function crmSidebarSetApiKey(key) { crmSetApiKey(key); return 'API key saved.'; }
function crmSidebarPreview(n) { return crmPreviewClassification(n || 5); }
function crmSidebarAudit() { return crmAudit(); }
