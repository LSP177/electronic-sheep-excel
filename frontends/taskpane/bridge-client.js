/* DataHelper task pane · R15
 * One controller for the four tabs: 帮我整理 / 清理工具 / 处理工具 / 小票.
 *
 * Safety model (unchanged from R13/R14, enforced by the host, not by this page):
 *   scan (captures range) -> prepare (locks plan) -> confirm (native Excel dialog) -> execute.
 * This page only chains those commands; the write is authorised solely by the
 * native dialog. VBA owns batch safety semantics. The page gives guidance and
 * C# validates message shapes; neither replaces the final VBA guard.
 * The dev simulator runs only in a normal browser with ?dev_sim=1.
 */
(function () {
  'use strict';

  var isNative = !!(window.chrome && window.chrome.webview && typeof window.chrome.webview.postMessage === 'function');
  var isSim = !isNative && /[?&]dev_sim=1\b/.test(window.location.search || '');

  /* ------------------------------------------------------------------ *
   * 1. Action registry — the only place the page describes an action.  *
   * ------------------------------------------------------------------ */
  // rank: 0 Unmerge | 1 cell cleanup | 2 BlankRows | 3 Dedup | 4 masking. alone: never batched.
  var ACTIONS = {
    Unmerge:        { label: '拆开合并的格子',       unit: '处', rank: 0 },
    Space:          { label: '清掉多余的空格',       unit: '格', rank: 1 },
    Unit:           { label: '文字数字改成数字',     unit: '格', rank: 1 },
    Date:           { label: '统一日期写法',         unit: '格', rank: 1 },
    Mark:           { label: '重复的行标颜色',       unit: '行', rank: 1 },
    BlankRows:      { label: '删除空行',             unit: '行', rank: 2 },
    Dedup:          { label: '删除重复的行',         unit: '行', rank: 3 },
    Transform:      { label: '打码',                 unit: '格', rank: 4 },
    Split:          { label: '数字和单位拆成两列',   unit: '格', alone: true },
    ClearTemporary: { label: '去掉插件留下的标记',   unit: '处', alone: true },
    Delivery:       { label: '交表前收拾',           unit: '处', alone: true },
    Lookup:         { label: '跨表补全',             unit: '格', alone: true },
    Append:         { label: '多表纵向追加',         unit: '行', alone: true },
    Unpivot:        { label: '宽表转长表',           unit: '行', alone: true },
    SplitSheet:     { label: '按列拆表',             unit: '行', alone: true },
    Reconcile:      { label: '两表对账',             unit: '处', alone: true }
  };
  var MASK = { Phone: '手机号打码', ID: '身份证号打码', Name: '姓名打码' };

  function baseCode(canon) {
    var c = String(canon || '');
    var i = c.indexOf(':');
    return i > 0 ? c.slice(0, i) : c;
  }
  function labelOf(canon) {
    var c = String(canon || '');
    if (c.indexOf('Transform:') === 0) return MASK[c.slice(10)] || '打码';
    if (c.indexOf('Lookup:') === 0) {
      var lp = c.split(':');
      if (lp.length === 8) {
        var rCnt = lp[3].split(',').length;
        return '跨表补全（从 ' + lp[1] + ' 取 ' + rCnt + ' 列 → ' + (lp[5] === 'AUTO' ? '右侧新增列' : lp[5] + ' 列') + '）';
      }
    }
    if (c.indexOf('Append:') === 0) {
      var ap = c.split(':');
      if (ap.length >= 6) {
        var sCnt = ap[1].split(',').filter(Boolean).length;
        return '多表纵向追加（' + sCnt + ' 张表 → 新表「' + (ap[5] || '多表追加汇总') + '」）';
      }
    }
    if (c.indexOf('Unpivot:') === 0) {
      var up = c.split(':');
      if (up.length === 8) {
        var vCnt = up[2].split(',').filter(Boolean).length;
        return '宽表转长表（展开 ' + vCnt + ' 列 → 新表「' + (up[7] || '宽表转长表明细') + '」）';
      }
    }
    if (c.indexOf('SplitSheet:') === 0) {
      var sp = c.split(':');
      if (sp.length === 5) {
        return '按列拆表（按 ' + sp[1] + ' 列拆分分发' + (sp[4] ? ' · 前缀「' + sp[4] + '」' : '') + '）';
      }
    }
    if (c === 'Unit:Force') return ACTIONS.Unit.label + '（同列多单位也转）';
    var a = ACTIONS[baseCode(c)];
    return a ? a.label : c;
  }
  function unitOf(canon) {
    var a = ACTIONS[baseCode(canon)];
    return a ? a.unit : '处';
  }
  // Client-side copy of the host rule, only to explain problems in plain words early.
  function batchProblem(list) {
    if (!list.length) return '请先选一项要处理的。';
    if (list.length > 6) return '一次最多做 6 项。';
    var ids = list.map(function (a) { return a.id; });
    if (list.length > 1 && ids.some(function (id) { return ACTIONS[id] && ACTIONS[id].alone; })) return '这一项要单独做，不能和其他项一起。';
    if (ids.indexOf('Mark') >= 0 && ids.indexOf('Dedup') >= 0) return '重复的行只能选「标颜色」或「删掉」其中一个。';
    if (ids.indexOf('Transform') >= 0 && (ids.indexOf('BlankRows') >= 0 || ids.indexOf('Dedup') >= 0)) return '打码不能和删除行一起做，请分开做。';
    for (var i = 1; i < ids.length; i++) {
      if ((ACTIONS[ids[i]].rank || 0) < (ACTIONS[ids[i - 1]].rank || 0)) return '操作顺序不对，请刷新窗格后再试。';
    }
    return '';
  }

  /* ------------------------------------------------------------------ *
   * 2. Host channel — request/response by id; pushes are 'invalidated'. *
   * ------------------------------------------------------------------ */
  var seq = 0;
  var pending = {};
  var NO_TIMEOUT = { confirm: 1, execute: 1 };

  function call(command, extra) {
    return new Promise(function (resolve, reject) {
      if (!isNative && !isSim) { reject(new Error('请在 Excel 的电子羊表格窗格里使用。')); return; }
      seq += 1;
      var req = { v: 1, id: 'ui_' + Date.now().toString(36) + '_' + seq, command: command };
      if (extra && extra.actions) req.actions = extra.actions;
      if (extra && extra.options) req.options = extra.options;
      if (extra && extra.path) req.path = extra.path;
      if (extra && Object.prototype.hasOwnProperty.call(extra, "license")) req.license = extra.license;
      var entry = { resolve: resolve, reject: reject, timer: null };
      if (!NO_TIMEOUT[command]) {
        entry.timer = setTimeout(function () {
          delete pending[req.id];
          reject(new Error('Excel 没有响应，请稍后再试。'));
        }, 60000);
      }
      pending[req.id] = entry;
      if (isNative) window.chrome.webview.postMessage(req); else Sim.dispatch(req);
    });
  }

  function onHostMessage(msg) {
    if (!msg) return;
    if (msg.command === 'execute' && msg.result && msg.result[9]) paintLicense(msg.result[9]);
    if (msg.type === 'invalidated') { onStale(String(msg.reason || '')); return; }
    var p = msg.id ? pending[msg.id] : null;
    if (!p) return; // replies to host-triggered test requests are not ours
    delete pending[msg.id];
    if (p.timer) clearTimeout(p.timer);
    if (msg.ok) { p.resolve(msg); return; }
    var e = new Error(msg.error || '没有完成');
    e.reply = msg;
    p.reject(e);
  }

  var ERRORS = [
    [/CANCELLED/, '已取消，表格没有被修改。'],
    [/BUSY/, '上一步还没结束，请稍等一下再点。'],
    [/NOT_READY/, '窗格还没准备好，请稍等再试。'],
    [/Expired scan session/i, '检查结果已经超过 5 分钟，请重新检查后再做。'],
    [/Locked range was removed/i, '前面一步已经把要处理的行全部删掉了，后面的步骤没有可处理的内容，已停下。'],
    [/No scan session/i, '请先检查一下表格。'],
    [/Stale session|snapshot changed|Range changed|Sheet changed/i, '工作表、数据或锁定选区已变化，请重新检查；本次未获授权执行。'],
    [/Business sheet required|Worksheet required/i, '请先切到一张普通的数据表（不是图表，也不是体检报告）。'],
    [/Range required/i, '请先在表格里选中单元格。'],
    [/One contiguous target required/i, '请只选一块连续的区域。'],
    [/Write plan limited to 100000/i, '范围超过 10 万格，请缩小范围分批处理。'],
    [/Reference worksheet not found/i, '找不到选的那张表，请重新选择。'],
    [/targetWriteCol count must match|returnCol must specify/i, '要取的列和填入的列数量不一致，请核对勾选。'],
    [/sheets must specify 2 to 32|至少勾选 2/i, '请至少勾选 2 张需要追加合并的工作表。'],
    [/Invalid colMap/i, '手动列名对应写法不对，请按「原列名=统一列名」填写，多组用分号隔开。'],
    [/idCols and valCols cannot overlap/i, '「左侧固定列」和「要展开的列」不能重叠，请核对勾选。'],
    [/attrHeader and valHeader must differ/i, '展开后的「属性列名」和「数值列名」不能写成一样的名字。'],
    [/Invalid sheetPrefix/i, '分表名称前缀不能超过 12 个字，且不能包含特殊符号。'],
    [/1 to 6 actions required/i, '一次最多做 6 项。'],
    [/Batch order must be|Unmerge must be first/i, '操作顺序不对，请刷新窗格后再试。'],
    [/Mark and Dedup cannot be combined/i, '重复的行只能选「标颜色」或「删掉」其中一个。'],
    [/Masking cannot be combined/i, '打码不能和删除行一起做，请分开做。'],
    [/Standalone/i, '这一项要单独做，不能和其他项一起。'],
    [/Request replay|Request budget/i, '窗格请求次数太多了，请关闭窗格再打开。']
  ];
  function friendly(errOrText) {
    var t = String((errOrText && errOrText.message) || errOrText || '').replace(/^VBA \[-?\d+\]\s*/, '');
    for (var i = 0; i < ERRORS.length; i++) if (ERRORS[i][0].test(t)) return ERRORS[i][1];
    return t || '没有完成，原因未知。';
  }

  /* ------------------------------------------------------------------ *
   * 3. State                                                            *
   * ------------------------------------------------------------------ */
  var S = {
    tab: 'auto', busy: false, help: false,
    ctx: { wb: '', sheet: '', sel: '', rows: 0, cols: 0, version: '', error: '' },
    scan: null, stale: '',
    auto: { phase: 'start', choice: {}, dupMode: 'Mark', last: null, checking: false },
    clean: { open: null, range: 'table', unitMode: 'Unit', forceMixed: false, dupMode: 'Mark', results: {} },
    proc: {
      open: null, results: {}, meta: null, metaLoading: false,
      lk: { target: '', src: '', srcKey: '', ret: '', retCols: [], tgtKey: '', writeModeType: 'auto', write: '', writeCols: [], dupRule: 'abort', writeMode: 'static' },
      preview: null, previewKey: '', previewing: false,
      ap: { sheets: [], colMode: 'union', summaryRowRule: 'exclude', sourceTagMode: 'sheet', targetSheetName: '多表追加汇总', colMap: '' },
      apPreview: null, apPreviewKey: '', apPreviewing: false,
      up: { sheet: '', idCols: [], valCols: [], attrHeader: '属性类别', valHeader: '数值', emptyPolicy: 'skip_blank', summaryRule: 'exclude', targetSheetName: '宽表转长表明细' },
      upPreview: null, upPreviewKey: '', upPreviewing: false,
      sp: { sheet: '', splitCol: '', blankKeyRule: 'bucket', summaryRule: 'exclude', sheetPrefix: '' },
      spPreview: null, spPreviewKey: '', spPreviewing: false,
      recon: { sheet: '', col: '', header: true },
      maskMode: 'Phone'
    },
    flow: null,
    receipt: { entries: [], unseen: 0, viewDay: null, confirmClear: false },
    alert: null
  };

  function getPath(path) { return path.split('.').reduce(function (o, k) { return o == null ? o : o[k]; }, S); }
  function setPath(path, v) {
    var ks = path.split('.'), o = S;
    for (var i = 0; i < ks.length - 1; i++) o = o[ks[i]];
    o[ks[ks.length - 1]] = v;
  }

  /* ------------------------------------------------------------------ *
   * 4. Receipt store — best effort, per workbook and day.               *
   * ------------------------------------------------------------------ */
  function dayKey(d) {
    d = d || new Date();
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
  var Store = {
    prefix: 'dh-receipt|',
    key: function (wb, day) { return this.prefix + (wb || '未命名') + '|' + day; },
    load: function (wb, day) {
      try { var v = JSON.parse(window.localStorage.getItem(this.key(wb, day)) || '[]'); return Array.isArray(v) ? v : []; }
      catch (e) { return []; }
    },
    save: function (wb, day, list) {
      try { window.localStorage.setItem(this.key(wb, day), JSON.stringify(list.slice(-200))); } catch (e) { /* memory only */ }
    },
    days: function (wb) {
      var out = [];
      try {
        var p = this.key(wb, '');
        for (var i = 0; i < window.localStorage.length; i++) {
          var k = window.localStorage.key(i);
          if (k && k.indexOf(p) === 0) out.push(k.slice(p.length));
        }
      } catch (e) { /* ignore */ }
      return out.sort().reverse();
    }
  };
  var receiptWb = null;
  function loadReceipt() {
    var identity = S.ctx.identity || S.ctx.wb;
    if (receiptWb === identity) return;
    receiptWb = identity;
    S.receipt.entries = Store.load(receiptWb, dayKey());
    S.receipt.viewDay = null;
  }

  /* ------------------------------------------------------------------ *
   * 5. Reading host results                                             *
   * ------------------------------------------------------------------ */
  function applyContext(r) {
    S.ctx.wb = String(r[0] || ''); S.ctx.sheet = String(r[1] || ''); S.ctx.sel = String(r[2] || '');
    S.ctx.rows = Number(r[3]) || 0; S.ctx.cols = Number(r[4]) || 0;
    if (r[5]) S.ctx.version = String(r[5]);
    S.ctx.identity = String(r[6] || S.ctx.wb);
    S.ctx.error = '';
    loadReceipt();
  }
  function applyCandidate(r) {
    S.ctx.wb = String(r[0] || S.ctx.wb); S.ctx.sheet = String(r[1] || S.ctx.sheet);
    S.ctx.sel = String(r[2] || ''); S.ctx.rows = Number(r[3]) || 0; S.ctx.cols = Number(r[4]) || 0;
    S.ctx.identity = String(r[5] || S.ctx.wb);
    S.ctx.error = '';
    loadReceipt();
  }
  function applyScan(r) {
    S.scan = {
      wb: String(r[0] || ''), sheet: String(r[1] || ''), addr: String(r[2] || ''),
      dirty: Number(r[4]) || 0, numText: Number(r[5]) || 0, merges: Number(r[6]) || 0,
      mergedCells: Number(r[7]) || 0, blank: Number(r[8]) || 0, dup: Number(r[9]) || 0,
      formulas: Number(r[10]) || 0, risk: Number(r[11]) || 0, dates: Number(r[12]) || 0,
      protectedNums: Number(r[13]) || 0,
      hiddenCols: String(r[14] || ''),
      rows: S.ctx.rows
    };
    S.stale = '';
  }

  var ctxTimer = null, quietUntil = 0;
  function scheduleCtx() {
    clearTimeout(ctxTimer);
    ctxTimer = setTimeout(function () { if (!S.busy) refreshContext(); }, 450);
  }
  function refreshContext() {
    return call('context').then(function (m) { applyContext(m.result); render(); })
      .catch(function (e) { S.ctx.error = friendly(e); render(); });
  }

  function onStale(reason) {
    // Our own flows move the selection and write data; their echoes can arrive just after the reply.
    if (S.busy || Date.now() < quietUntil) return;
    if (S.scan) {
      if (/data changed|reference sheet/i.test(reason)) S.stale = '表格内容有变动';
      else if (/sheet|workbook/i.test(reason)) S.stale = '你切换了工作表或文件';
      else if (!S.stale) S.stale = 'selection';
    }
    scheduleCtx();
    render();
  }

  /* ------------------------------------------------------------------ *
   * 6. Flows                                                            *
   * ------------------------------------------------------------------ */
  function checkTable(preselect) {
    if (S.busy) return;
    S.busy = true; S.auto.checking = true; S.alert = null;
    render();
    return call('candidate').then(function (m) {
      applyCandidate(m.result);
      if (!m.result[2]) { S.scan = null; S.auto.phase = 'empty'; return null; }
      return call('scan').then(function (s) {
        applyScan(s.result);
        S.auto.phase = 'result';
        S.auto.choice = { Space: S.scan.dirty > 0, Unmerge: S.scan.merges > 0, Unit: false, Date: false, BlankRows: false, Dup: false };
        S.auto.dupMode = 'Mark';
        (preselect || []).forEach(function (k) { S.auto.choice[k] = true; });
      });
    }).catch(function (e) {
      showAlert('bad', friendly(e));
    }).then(function () {
      S.busy = false; S.auto.checking = false; quietUntil = Date.now() + 800; render();
    });
  }

  function autoActions() {
    var c = S.auto.choice, a = [];
    if (c.Unmerge) a.push({ id: 'Unmerge' });
    if (c.Space) a.push({ id: 'Space' });
    if (c.Unit) a.push({ id: 'Unit' });
    if (c.Date) a.push({ id: 'Date' });
    if (c.Dup && S.auto.dupMode === 'Mark') a.push({ id: 'Mark' });
    if (c.BlankRows) a.push({ id: 'BlankRows' });
    if (c.Dup && S.auto.dupMode === 'Dedup') a.push({ id: 'Dedup' });
    return a;
  }

  // plan: { actions, origin:'auto'|'clean'|'process', key, title, range:'table'|'selection', guard() -> '' | message }
  function runPlan(plan) {
    if (S.busy) return;
    var problem = batchProblem(plan.actions);
    if (problem) { showAlert('warn', problem); render(); return; }
    S.busy = true; S.alert = null;
    S.flow = { phase: 'prep', plan: plan, sheet: '', addr: '', locked: [] };
    render();

    var step = plan.range === 'table'
      ? call('candidate').then(function (m) {
          applyCandidate(m.result);
          if (!m.result[2]) throw new Error('这张表是空的，没有要处理的内容。');
        })
      : call('context').then(function (m) { applyContext(m.result); });

    step.then(function () {
      var msg = plan.guard ? plan.guard() : '';
      if (msg) throw new Error(msg);
      return call('scan');
    }).then(function (m) {
      applyScan(m.result);
      return call('prepare', { actions: plan.actions });
    }).then(function (m) {
      var r = m.result || [];
      S.flow.phase = 'confirm'; S.flow.sheet = String(r[4] || ''); S.flow.addr = String(r[5] || '');
      S.flow.locked = String(r[6] || '').split('|').filter(Boolean);
      render();
      return call('confirm');
    }).then(function () {
      S.flow.phase = 'exec'; render();
      return call('execute');
    }).then(function (m) {
      finishRun(plan, m);
    }, function (e) {
      if (e && e.reply && e.reply.command === 'execute') { finishRun(plan, e.reply); return; }
      var text = friendly(e);
      S.flow = null;
      if (/CANCELLED/.test(String(e && e.message))) toast(text); else showAlert('bad', text);
    }).then(function () {
      S.busy = false; S.flow = null; quietUntil = Date.now() + 800; render();
    });
  }

  // Rows deleted later in the same run (BlankRows "R5:7C1:9", Dedup "行12") move the cells
  // listed by earlier steps up. Re-map those addresses so they match the sheet the user now sees.
  function remapAfterDeletes(rows) {
    rows.forEach(function (del, i) {
      var code = baseCode(del.code), spans = [];
      if (code !== 'BlankRows' && code !== 'Dedup') return;
      (del.changeRows || []).forEach(function (r) {
        var a = String(r[0] || ''), m;
        if ((m = a.match(/^R(\d+):(\d+)C(\d+):(\d+)$/))) spans.push({ r1: +m[1], r2: +m[2], c1: +m[3], c2: +m[4] });
        else if ((m = a.match(/^行(\d+)$/)) && String(r[1] || '').indexOf('删除') === 0) spans.push({ r1: +m[1], r2: +m[1], c1: 1, c2: 16384 });
      });
      if (!spans.length) return;
      var complete = spans.length >= (del.changed || 0);
      rows.slice(0, i).forEach(function (x) {
        if (!complete) { x.moved = 'unsure'; return; }
        (x.changeRows || []).forEach(function (r) {
          var m = String(r[0] || '').match(/^([A-Z]{1,3})(\d+)$/);
          if (!m) return;
          if (!r.orig) r.orig = r[0]; // address at the time of the change
          var col = colIndex(m[1]), row = +m[2], shift = 0, gone = false;
          spans.forEach(function (sp) {
            if (col < sp.c1 || col > sp.c2) return;
            if (row >= sp.r1 && row <= sp.r2) gone = true;
            else if (row > sp.r2) shift += sp.r2 - sp.r1 + 1;
          });
          if (gone) { r[0] = (r.orig || r[0]) + '（这一行后来被删除）'; x.moved = x.moved || 'yes'; }
          else if (shift) { r[0] = m[1] + (row - shift); x.moved = x.moved || 'yes'; }
        });
      });
    });
    return rows;
  }
  function rowsOf(list) {
    return (Array.isArray(list) ? list : []).map(function (row) {
      var skips = Array.isArray(row[4]) ? row[4] : [0, []];
      var changes = Array.isArray(row[6]) ? row[6] : [0, []];
      var changeRows = Array.isArray(changes[1]) ? changes[1] : [];
      return {
        code: String(row[0] || ''), count: Number(row[1]) || 0,
        unit: String(row[5] || unitOf(String(row[0] || ''))),
        changed: Number(changes[0]) || 0,
        changeRows: changeRows.filter(function (r) { return Array.isArray(r) && r.length >= 1; }),
        summary: String(row[2] || ''), audit: String(row[3] || '')
      };
    });
  }

  function finishRun(plan, reply) {
    var r = Array.isArray(reply.result) ? reply.result : [];
    var tag = String(r[0] || '');
    var out = {
      t: new Date().toISOString(), title: plan.title, origin: plan.origin,
      sheet: (S.flow && S.flow.sheet) || S.ctx.sheet, addr: (S.flow && S.flow.addr) || '',
      ok: tag === 'COMPLETED', rows: [], backup: '', error: '', failedAt: '', notDone: []
    };
    var planned = (S.flow && S.flow.locked.length) ? S.flow.locked : plan.actions.map(function (a) { return a.id === 'Transform' ? 'Transform:' + a.options.mode : a.id; });
    if (tag === 'COMPLETED') {
      out.backup = String(r[2] || ''); out.rows = remapAfterDeletes(rowsOf(r[3]));
    } else if (tag === 'FAILED') {
      out.backup = String(r[2] || ''); out.rows = remapAfterDeletes(rowsOf(r[3]));
      out.error = friendly(String(r[5] || reply.error || ''));
      var phase = String(r[1] || '');
      out.failedAt = phase.indexOf('operation:') === 0 || phase.indexOf('audit-check:') === 0 ? phase.slice(phase.indexOf(':') + 1) : '';
      out.stage = phase;
    } else {
      out.error = friendly(reply.error || '');
    }
    out.skipped = (Array.isArray(r[8]) ? r[8] : []).map(function (s) { return { code: String(s[0] || ''), why: String(s[1] || '') }; });
    var done = out.rows.map(function (x) { return x.code; });
    out.notDone = planned.filter(function (c) { return done.indexOf(c) < 0; });

    // Receipt: every real write is recorded, success or not.
    if (out.rows.length || out.ok || tag === 'FAILED') {
      S.receipt.entries.push({
        t: out.t, title: out.title, sheet: out.sheet, addr: out.addr, ok: out.ok,
        rows: out.rows.map(function (x) { return { code: x.code, count: x.count, skips: x.skips, unit: x.unit, audit: x.audit, note: unitNote(x) }; }),
        failedAt: out.failedAt, error: out.ok ? '' : out.error, backup: out.backup,
        skipped: (out.skipped || []).map(function (s) { return { code: s.code, why: s.why }; })
      });
      Store.save(receiptWb, dayKey(), S.receipt.entries);
      if (S.tab !== 'receipt') S.receipt.unseen += 1;
    }

    if (S.scan) S.stale = '表格刚被修改过';
    if (plan.origin === 'auto') { S.auto.phase = 'done'; S.auto.last = out; }
    else if (plan.origin === 'clean') { S.clean.results[plan.key] = out; if (plan.key === 'Unit') S.clean.forceMixed = false; }
    else S.proc.results[plan.key] = out;
    S.flow = null;
    try { window.scrollTo(0, 0); } catch (e) { /* ignore */ }
  }

  /* ---- process-tool helpers ---- */
  function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function colIndex(l) { var n = 0; l = String(l || '').toUpperCase(); for (var i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64); return n; }
  function firstColOf(addr) { var m = /^\$?([A-Z]{1,3})/i.exec(String(addr || '')); return m ? m[1].toUpperCase() : ''; }
  function metaFor(name) {
    var list = S.proc.meta || [];
    for (var i = 0; i < list.length; i++) if (list[i][0] === name) return list[i];
    return null;
  }
  function colsOf(name) { var m = metaFor(name); return m && Array.isArray(m[3]) ? m[3] : []; }
  function colLabel(sheet, letter) {
    var cols = colsOf(sheet);
    for (var i = 0; i < cols.length; i++) if (cols[i][0] === letter) return cols[i][1] || (letter + ' 列');
    return letter + ' 列';
  }
  function writeColChoices(sheet) {
    var cols = colsOf(sheet).slice();
    var last = cols.length ? colIndex(cols[cols.length - 1][0]) : 0;
    for (var k = 1; k <= 6; k++) cols.push([colLetter(last + k), '空列（表头右边）', true]);
    return cols;
  }

  function ensureMeta() {
    if (S.proc.meta || S.proc.metaLoading || S.busy) return;
    S.proc.metaLoading = true; render();
    refreshContext().then(function () { return call('lookup_meta'); }).then(function (m) {
      S.proc.meta = (Array.isArray(m.result) ? m.result : []).filter(function (s) { return Array.isArray(s) && s[0]; });
      initLookupDefaults(true);
      initAppendDefaults();
      initUnpivotDefaults(true);
      initSplitSheetDefaults(true);
      initReconDefaults();
    }).catch(function (e) { showAlert('bad', friendly(e)); })
      .then(function () { S.proc.metaLoading = false; render(); });
  }
  function syncLookupColumns(lk) {
    var sc = colsOf(lk.src), avail = sc.filter(function (c) { return c[0] !== lk.srcKey; });
    var validSet = {};
    avail.forEach(function (c) { validSet[c[0]] = true; });
    var kept = (Array.isArray(lk.retCols) ? lk.retCols : []).filter(function (c) { return validSet[c]; });
    if (!kept.length && avail.length) kept = [avail[0][0]];
    if (kept.length > 16) kept = kept.slice(0, 16);
    lk.retCols = kept;
    lk.ret = kept.join(',');
    var wc = writeColChoices(lk.target).filter(function (c) { return c[0] !== lk.tgtKey; });
    var emptyStarts = wc.filter(function (c) { return c[2]; });
    var baseList = emptyStarts.length ? emptyStarts : wc;
    var wList = Array.isArray(lk.writeCols) ? lk.writeCols.slice(0, kept.length) : [];
    var used = {};
    for (var i = 0; i < kept.length; i++) {
      var cur = wList[i];
      if (!cur || cur === lk.tgtKey || used[cur]) {
        var pick = '';
        for (var j = 0; j < baseList.length; j++) {
          if (!used[baseList[j][0]] && baseList[j][0] !== lk.tgtKey) { pick = baseList[j][0]; break; }
        }
        if (!pick) {
          var tc = colsOf(lk.target);
          var lastIdx = tc.length ? colIndex(tc[tc.length - 1][0]) : 1;
          var candIdx = lastIdx + i + 1;
          if (colLetter(candIdx) === lk.tgtKey) candIdx += 1;
          pick = colLetter(candIdx);
        }
        cur = pick;
      }
      used[cur] = true;
      wList[i] = cur;
    }
    lk.writeCols = wList;
    if (!lk.writeModeType) lk.writeModeType = 'auto';
    lk.write = lk.writeModeType === 'auto' ? 'AUTO' : wList.join(',');
  }
  function autoAppendPreviewInfo(lk) {
    var tc = colsOf(lk.target), sc = colsOf(lk.src);
    var lastIdx = tc.length ? colIndex(tc[tc.length - 1][0]) : 0;
    var existing = {};
    tc.forEach(function (c) { if (c[1]) existing[String(c[1]).toLowerCase()] = true; });
    var items = [];
    (lk.retCols || []).forEach(function (rc, idx) {
      var newCol = colLetter(lastIdx + idx + 1);
      var srcEntry = sc.filter(function (x) { return x[0] === rc; })[0];
      var baseName = (srcEntry && srcEntry[1] && srcEntry[1] !== '(空白)') ? srcEntry[1] : ('补全_' + rc);
      var name = baseName;
      if (existing[name.toLowerCase()]) name = lk.src + '_' + baseName;
      var sfx = 2;
      while (existing[name.toLowerCase()]) { name = lk.src + '_' + baseName + '_' + sfx; sfx++; }
      existing[name.toLowerCase()] = true;
      items.push({ srcCol: rc, tgtCol: newCol, header: name });
    });
    return items;
  }
  function initLookupDefaults(resetSource) {
    var lk = S.proc.lk, meta = S.proc.meta || [];
    lk.target = S.ctx.sheet;
    if (resetSource || !metaFor(lk.src) || lk.src === lk.target) {
      var others = meta.filter(function (s) { return s[0] !== lk.target && Array.isArray(s[3]) && s[3].length; });
      lk.src = others.length ? others[0][0] : '';
    }
    var sc = colsOf(lk.src), tc = colsOf(lk.target);
    lk.srcKey = sc[0] ? sc[0][0] : '';
    lk.retCols = sc[1] ? [sc[1][0]] : (sc[0] ? [sc[0][0]] : []);
    var keyLabel = sc[0] ? sc[0][1] : '';
    var match = tc.filter(function (c) { return c[1] && c[1] === keyLabel; })[0];
    lk.tgtKey = match ? match[0] : (tc[0] ? tc[0][0] : 'A');
    if (resetSource || !lk.writeModeType) lk.writeModeType = 'auto';
    lk.writeCols = [];
    syncLookupColumns(lk);
    S.proc.preview = null; S.proc.previewKey = '';
  }
  function initAppendDefaults() {
    var ap = S.proc.ap, meta = S.proc.meta || [];
    var nonEmpty = meta.filter(function (m) { return (Number(m[2]) || 0) >= 2 && Array.isArray(m[3]) && m[3].length > 0; }).map(function (m) { return m[0]; });
    var allNames = meta.map(function (m) { return m[0]; });
    var validSet = {};
    allNames.forEach(function (n) { validSet[n] = true; });
    var kept = (Array.isArray(ap.sheets) ? ap.sheets : []).filter(function (n) { return validSet[n]; });
    if (kept.length < 2) {
      kept = nonEmpty.length >= 2 ? nonEmpty.slice(0, 32) : allNames.slice(0, Math.min(32, allNames.length));
    }
    ap.sheets = kept;
    if (!ap.colMode) ap.colMode = 'union';
    if (!ap.summaryRowRule) ap.summaryRowRule = 'exclude';
    if (!ap.sourceTagMode) ap.sourceTagMode = 'sheet';
    if (!ap.targetSheetName) ap.targetSheetName = '多表追加汇总';
    if (ap.colMap == null) ap.colMap = '';
    S.proc.apPreview = null; S.proc.apPreviewKey = '';
  }
  function appendOptions() {
    var ap = S.proc.ap;
    var sheets = Array.isArray(ap.sheets) ? ap.sheets.filter(Boolean) : [];
    if (sheets.length < 2) throw new Error('请至少勾选 2 张需要合并的工作表。');
    if (sheets.length > 32) throw new Error('单次最多同时合并 32 张工作表。');
    var tgtName = String(ap.targetSheetName || '多表追加汇总').trim() || '多表追加汇总';
    if (tgtName.length > 31 || /[:|\\/?*\[\],]/.test(tgtName)) throw new Error('新建汇总表名称不能超过 31 个字，且不能包含特殊符号。');
    var rawMap = String(ap.colMap || '').replace(/；/g, ';').replace(/＝/g, '=').trim();
    if (rawMap) {
      var parts = rawMap.split(';');
      for (var i = 0; i < parts.length; i++) {
        var p = parts[i].trim();
        if (!p) continue;
        var eq = p.indexOf('=');
        if (eq <= 0 || eq >= p.length - 1 || p.indexOf('=', eq + 1) >= 0) {
          throw new Error('手动列名对应写法不对：「' + p + '」，请按「旧列名=新列名」格式填写。');
        }
      }
    }
    return {
      sheets: sheets.join(','),
      colMode: ap.colMode || 'union',
      summaryRowRule: ap.summaryRowRule || 'exclude',
      sourceTagMode: ap.sourceTagMode || 'sheet',
      targetSheetName: tgtName,
      colMap: rawMap
    };
  }
  function previewAppend() {
    if (S.busy) return;
    var opts;
    try { opts = appendOptions(); } catch (e) { showAlert('warn', e.message); render(); return; }
    S.busy = true; S.proc.apPreviewing = true; S.alert = null; render();
    call('append_preview', { options: opts }).then(function (m) {
      var r = m.result || [];
      S.proc.apPreview = {
        sheetCount: Number(r[0]) || 0,
        totalRows: Number(r[1]) || 0,
        netRows: Number(r[2]) || 0,
        summaryRows: Number(r[3]) || 0,
        colCount: Number(r[4]) || 0,
        unionOnlyCols: Number(r[5]) || 0,
        sourceErrors: Number(r[6]) || 0,
        targetSheet: String(r[7] || ''),
        colMode: String(r[8] || 'union'),
        summaryRowRule: String(r[9] || 'exclude'),
        sourceTagMode: String(r[10] || 'sheet'),
        canExecute: !!r[11],
        block: String(r[12] || ''),
        headers: Array.isArray(r[13]) ? r[13] : [],
        summarySamples: Array.isArray(r[14]) ? r[14] : [],
        sheetBreakdown: Array.isArray(r[15]) ? r[15] : [],
        anomalies: Array.isArray(r[16]) ? r[16] : []
      };
      S.proc.apPreviewKey = JSON.stringify(opts);
    }).catch(function (e) { showAlert('bad', friendly(e)); })
      .then(function () { S.busy = false; S.proc.apPreviewing = false; render(); });
  }
  function isSummaryHeaderName(label) {
    var s = String(label || '').replace(/[\s\u3000\u00A0:：]/g, '').toLowerCase();
    return ['合计', '总计', '小计', '行合计', '列合计', '本月合计', '本季合计', '全年合计', '本年合计', '汇总', '累计数', '累计', 'total', 'subtotal', 'grandtotal', 'rowtotal', 'summary'].indexOf(s) >= 0;
  }
  function initUnpivotDefaults(forceReset) {
    var up = S.proc.up, curSheet = S.ctx.sheet || ((S.proc.meta && S.proc.meta[0]) ? S.proc.meta[0][0] : '');
    var cols = colsOf(curSheet);
    if (forceReset || up.sheet !== curSheet || !up.idCols || !up.idCols.length) {
      up.sheet = curSheet;
      if (cols.length >= 2) {
        up.idCols = [cols[0][0]];
        var vList = [];
        for (var i = 1; i < cols.length && vList.length < 64; i++) {
          if (!isSummaryHeaderName(cols[i][1])) vList.push(cols[i][0]);
        }
        if (!vList.length) vList = [cols[1][0]];
        up.valCols = vList;
      } else if (cols.length === 1) {
        up.idCols = [cols[0][0]];
        up.valCols = [];
      } else {
        up.idCols = [];
        up.valCols = [];
      }
    }
    if (!up.attrHeader) up.attrHeader = '属性类别';
    if (!up.valHeader) up.valHeader = '数值';
    if (!up.emptyPolicy) up.emptyPolicy = 'skip_blank';
    if (!up.summaryRule) up.summaryRule = 'exclude';
    if (!up.targetSheetName) up.targetSheetName = '宽表转长表明细';
    S.proc.upPreview = null; S.proc.upPreviewKey = '';
  }
  function unpivotOptions() {
    var up = S.proc.up;
    var idCols = Array.isArray(up.idCols) ? up.idCols.filter(Boolean) : [];
    var valCols = Array.isArray(up.valCols) ? up.valCols.filter(Boolean) : [];
    if (!idCols.length) throw new Error('请至少勾选 1 个左侧固定维度列（如姓名、编号）。');
    if (idCols.length > 16) throw new Error('固定维度列最多支持 16 列。');
    if (!valCols.length) throw new Error('请至少勾选 1 个需要展开的横向指标列（如 1月、2月、3月）。');
    if (valCols.length > 64) throw new Error('单次最多展开 64 个横向指标列。');
    var idMap = {};
    idCols.forEach(function (c) { idMap[c] = true; });
    for (var i = 0; i < valCols.length; i++) {
      if (idMap[valCols[i]]) throw new Error('「固定维度列」与「要展开的列」不能包含同一列（' + valCols[i] + ' 列）。');
    }
    var attrH = String(up.attrHeader || '属性类别').trim() || '属性类别';
    var valH = String(up.valHeader || '数值').trim() || '数值';
    if (attrH.toLowerCase() === valH.toLowerCase()) throw new Error('展开后的「属性列名」和「数值列名」不能同名。');
    var tgtName = String(up.targetSheetName || '宽表转长表明细').trim() || '宽表转长表明细';
    if (tgtName.length > 31 || /[:|\\/?*\[\],]/.test(tgtName)) throw new Error('新建工作表名称不能超过 31 个字，且不能含特殊符号。');
    return {
      idCols: idCols.join(','),
      valCols: valCols.join(','),
      attrHeader: attrH,
      valHeader: valH,
      emptyPolicy: up.emptyPolicy || 'skip_blank',
      summaryRule: up.summaryRule || 'exclude',
      targetSheetName: tgtName
    };
  }
  function previewUnpivot() {
    if (S.busy) return;
    var opts;
    try { opts = unpivotOptions(); } catch (e) { showAlert('warn', e.message); render(); return; }
    S.busy = true; S.proc.upPreviewing = true; S.alert = null; render();
    call('unpivot_preview', { options: opts }).then(function (m) {
      var r = m.result || [];
      S.proc.upPreview = {
        sheetName: String(r[0] || ''),
        totalSourceRows: Number(r[1]) || 0,
        netOutputRows: Number(r[2]) || 0,
        idColCount: Number(r[3]) || 0,
        valColCount: Number(r[4]) || 0,
        summaryRowCount: Number(r[5]) || 0,
        summaryColCount: Number(r[6]) || 0,
        skippedBlanks: Number(r[7]) || 0,
        sourceErrors: Number(r[8]) || 0,
        targetSheet: String(r[9] || ''),
        attrHeader: String(r[10] || '属性类别'),
        valHeader: String(r[11] || '数值'),
        emptyPolicy: String(r[12] || 'skip_blank'),
        summaryRule: String(r[13] || 'exclude'),
        canExecute: !!r[14],
        block: String(r[15] || ''),
        samples: Array.isArray(r[16]) ? r[16] : [],
        summarySamples: Array.isArray(r[17]) ? r[17] : [],
        anomalies: Array.isArray(r[18]) ? r[18] : []
      };
      S.proc.upPreviewKey = JSON.stringify(opts);
    }).catch(function (e) { showAlert('bad', friendly(e)); })
      .then(function () { S.busy = false; S.proc.upPreviewing = false; render(); });
  }
  function initSplitSheetDefaults(forceReset) {
    var sp = S.proc.sp, curSheet = S.ctx.sheet || ((S.proc.meta && S.proc.meta[0]) ? S.proc.meta[0][0] : '');
    var cols = colsOf(curSheet);
    if (forceReset || sp.sheet !== curSheet || !sp.splitCol) {
      sp.sheet = curSheet;
      sp.splitCol = cols[0] ? cols[0][0] : 'A';
    }
    if (!sp.blankKeyRule) sp.blankKeyRule = 'bucket';
    if (!sp.summaryRule) sp.summaryRule = 'exclude';
    if (sp.sheetPrefix == null) sp.sheetPrefix = '';
    S.proc.spPreview = null; S.proc.spPreviewKey = '';
  }
  function splitSheetOptions() {
    var sp = S.proc.sp;
    var col = String(sp.splitCol || '').trim().toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(col)) throw new Error('请选择要按哪一列拆分工作表。');
    var pfx = String(sp.sheetPrefix || '').trim();
    if (pfx.length > 12 || /[:|\\/?*\[\],]/.test(pfx)) throw new Error('分表名称前缀不能超过 12 个字，且不能包含特殊符号。');
    return {
      splitCol: col,
      blankKeyRule: sp.blankKeyRule || 'bucket',
      summaryRule: sp.summaryRule || 'exclude',
      sheetPrefix: pfx
    };
  }
  function previewSplitSheet() {
    if (S.busy) return;
    var opts;
    try { opts = splitSheetOptions(); } catch (e) { showAlert('warn', e.message); render(); return; }
    S.busy = true; S.proc.spPreviewing = true; S.alert = null; render();
    call('split_sheet_preview', { options: opts }).then(function (m) {
      var r = m.result || [];
      S.proc.spPreview = {
        sheetName: String(r[0] || ''),
        splitCol: String(r[1] || ''),
        splitHeader: String(r[2] || ''),
        totalRows: Number(r[3]) || 0,
        netRows: Number(r[4]) || 0,
        groupCount: Number(r[5]) || 0,
        blankRows: Number(r[6]) || 0,
        summaryRows: Number(r[7]) || 0,
        sourceErrors: Number(r[8]) || 0,
        colCount: Number(r[9]) || 0,
        blankKeyRule: String(r[10] || 'bucket'),
        summaryRule: String(r[11] || 'exclude'),
        sheetPrefix: String(r[12] || ''),
        canExecute: !!r[13],
        block: String(r[14] || ''),
        groups: Array.isArray(r[15]) ? r[15] : [],
        summarySamples: Array.isArray(r[16]) ? r[16] : [],
        anomalies: Array.isArray(r[17]) ? r[17] : []
      };
      S.proc.spPreviewKey = JSON.stringify(opts);
    }).catch(function (e) { showAlert('bad', friendly(e)); })
      .then(function () { S.busy = false; S.proc.spPreviewing = false; render(); });
  }
  function pickSplitColumn() {
    if (S.busy) return;
    call('context').then(function (m) {
      applyContext(m.result);
      var l = firstColOf(S.ctx.sel);
      if (!l) { showAlert('warn', '没读到选中的列。'); return; }
      S.proc.sp.sheet = S.ctx.sheet;
      S.proc.sp.splitCol = l;
      S.proc.spPreview = null; S.proc.spPreviewKey = '';
      toast('已选 ' + l + ' 列作为拆分列');
    }).catch(function (e) { showAlert('bad', friendly(e)); }).then(render);
  }
  function initReconDefaults() {
    var meta = S.proc.meta || [];
    var others = meta.filter(function (s) { return s[0] !== S.ctx.sheet; });
    S.proc.recon.sheet = others.length ? others[0][0] : '';
    var c = colsOf(S.proc.recon.sheet);
    S.proc.recon.col = c[0] ? c[0][0] : 'A';
  }
  function lookupOptions() {
    var lk = S.proc.lk;
    if (!lk.src) throw new Error('请选一张来源表。');
    var retCols = Array.isArray(lk.retCols) && lk.retCols.length ? lk.retCols : (lk.ret ? String(lk.ret).split(',') : []);
    if (!retCols.length) throw new Error('请至少勾选 1 个要取过来的列。');
    if (retCols.length > 16) throw new Error('单次最多选择 16 个要取过来的列。');
    var seenRet = {};
    retCols.forEach(function (c) {
      if (!/^[A-Z]{1,3}$/.test(c)) throw new Error('要取的列选择不完整。');
      if (c === lk.srcKey) throw new Error('「按哪列对上」和「要取的列」不能包含同一列（' + c + ' 列）。');
      if (seenRet[c]) throw new Error('要取的列不能重复（' + c + ' 列）。');
      seenRet[c] = true;
    });
    if (!/^[A-Z]{1,3}$/.test(lk.srcKey) || !/^[A-Z]{1,3}$/.test(lk.tgtKey)) throw new Error('匹配键列的选择不完整。');
    var targetWriteSpec = 'AUTO';
    if (lk.writeModeType === 'explicit') {
      var wCols = Array.isArray(lk.writeCols) && lk.writeCols.length === retCols.length ? lk.writeCols : (lk.write && lk.write !== 'AUTO' ? String(lk.write).split(',') : []);
      if (wCols.length !== retCols.length) throw new Error('请为每个要取的列分别指定填入的目标列。');
      var seenW = {};
      wCols.forEach(function (c) {
        if (!/^[A-Z]{1,3}$/.test(c)) throw new Error('填入的目标列选择不完整。');
        if (c === lk.tgtKey) throw new Error('当前表里「按哪列对上」和「填到哪列」不能是同一列（' + c + ' 列）。');
        if (seenW[c]) throw new Error('填入的目标列不能重复（' + c + ' 列），不同字段请填到不同列。');
        seenW[c] = true;
      });
      targetWriteSpec = wCols.join(',');
    }
    return { sourceSheet: lk.src, sourceKeyCol: lk.srcKey, returnCol: retCols.join(','), targetKeyCol: lk.tgtKey, targetWriteCol: targetWriteSpec, dupRule: lk.dupRule, writeMode: lk.writeMode };
  }
  // The engine's block reason names the rule codes; show the option labels the page offers.
  function plainBlock(text) {
    return text.replace(/“取首次出现行 \(first\)”/g, '「用第一次出现的那条」').replace(/“跳过重复键 \(skip\)”/g, '「这些编号都不填」')
      .replace(/显式选择/g, '在「更多设置 · 编号重复时」里选');
  }
  function previewLookup() {
    if (S.busy) return;
    var opts;
    try { opts = lookupOptions(); } catch (e) { showAlert('warn', e.message); render(); return; }
    S.busy = true; S.proc.previewing = true; S.alert = null; render();
    call('lookup_preview', { options: opts }).then(function (m) {
      var r = m.result || [];
      S.proc.preview = {
        total: Number(r[5]) || 0, matched: Number(r[6]) || 0, unmatched: Number(r[7]) || 0,
        blankKeys: Number(r[8]) || 0, dupKeys: Number(r[9]) || 0, conflicts: Number(r[10]) || 0,
        typeDiff: Number(r[11]) || 0, protectedCells: Number(r[12]) || 0,
        writeSummary: String(r[14] || ''),
        canExecute: !!r[16], block: plainBlock(String(r[17] || '')),
        samples: Array.isArray(r[18]) ? r[18] : [], anomalies: Array.isArray(r[19]) ? r[19] : [],
        sourceErrors: Number(r[20]) || 0
      };
      S.proc.previewKey = JSON.stringify(opts);
    }).catch(function (e) { showAlert('bad', friendly(e)); })
      .then(function () { S.busy = false; S.proc.previewing = false; render(); });
  }
  function pickColumn(field) {
    if (S.busy) return;
    call('context').then(function (m) {
      applyContext(m.result);
      if (S.ctx.sheet !== S.proc.lk.target) { showAlert('warn', '请在「' + S.proc.lk.target + '」这张表里选中一列再点。'); return; }
      var l = firstColOf(S.ctx.sel);
      if (!l) { showAlert('warn', '没读到选中的列。'); return; }
      if (field && field.indexOf('write_') === 0) {
        var idx = Number(field.slice(6)) || 0;
        S.proc.lk.writeCols[idx] = l;
        S.proc.lk.write = S.proc.lk.writeCols.join(',');
      } else {
        S.proc.lk[field] = l;
        syncLookupColumns(S.proc.lk);
      }
      S.proc.preview = null; S.proc.previewKey = '';
      toast('已选 ' + l + ' 列');
    }).catch(function (e) { showAlert('bad', friendly(e)); }).then(render);
  }

  /* ------------------------------------------------------------------ *
   * 7. Rendering                                                        *
   * ------------------------------------------------------------------ */
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function num(n) { return Number(n || 0).toLocaleString('zh-CN'); }
  function hhmm(iso) { var d = new Date(iso); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
  function $(id) { return document.getElementById(id); }

  var ICON = {
    check: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#1E6B52" stroke-width="2.4" aria-hidden="true"><path d="M5 12.5l4 4L19 7"/></svg>',
    chev: '<svg class="chev" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>',
    search: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.4" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg>',
    okBig: '<svg width="40" height="40" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="11" fill="#1E6B52"/><path d="M7 12.5l3.2 3L17 9" fill="none" stroke="#fff" stroke-width="2.2"/></svg>',
    badBig: '<svg width="40" height="40" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="11" fill="#A8261B"/><path d="M12 6.8v6.4M12 16.6v.4" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/></svg>',
    pick: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M5 3l14 8-6 2-2 6z"/></svg>',
    table: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16"/></svg>',
    down: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#8E897F" stroke-width="2" aria-hidden="true"><path d="M12 4v16M6 14l6 6 6-6"/></svg>',
    shield: '<svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><rect x="3" y="4" width="18" height="15" rx="2"/><path d="M3 8h18"/><path d="M8.5 13.5l2.5 2.5 4.5-4.5"/></svg>'
  };

  function render() {
    renderTabs(); renderCtx(); renderAlert(); renderHelp();
    ['auto', 'clean', 'process', 'receipt'].forEach(function (t) { $('view-' + t).classList.toggle('hidden', S.tab !== t); });
    if (S.tab === 'auto') renderAuto();
    else if (S.tab === 'clean') renderClean();
    else if (S.tab === 'process') renderProcess();
    else renderReceipt();
    renderOverlay();
  }

  function renderTabs() {
    var tabs = document.querySelectorAll('.tab');
    for (var i = 0; i < tabs.length; i++) tabs[i].setAttribute('aria-pressed', String(tabs[i].getAttribute('data-tab') === S.tab));
    var dot = $('receiptDot');
    dot.textContent = S.receipt.unseen > 9 ? '9+' : String(S.receipt.unseen);
    dot.classList.toggle('hidden', !S.receipt.unseen);
  }

  function renderCtx() {
    var el = $('ctx'), c = S.ctx, h = ICON.table;
    if (c.error) { el.innerHTML = h + '<span class="grow">' + esc(c.error) + '</span><button class="link-btn" data-act="ctx-refresh">刷新</button>'; return; }
    if (!c.sheet) { el.innerHTML = h + '<span class="grow muted">正在读取当前表…</span>'; return; }
    if (S.tab === 'receipt') { el.innerHTML = h + '<span class="grow"><b>' + esc(c.wb || '当前文件') + '</b> · 今天</span>'; return; }
    if (S.tab === 'process') { el.innerHTML = h + '<span class="grow">当前表：<b>' + esc(c.sheet) + '</b></span>'; return; }
    if (S.tab === 'clean') {
      var where = S.clean.range === 'table'
        ? (S.scan && S.scan.sheet === c.sheet ? '数据区域 ' + esc(S.scan.addr) : '整张表的数据')
        : '选中 <b>' + esc(c.sel || '—') + '</b>';
      el.innerHTML = h + '<span class="grow"><b>' + esc(c.sheet) + '</b> · ' + where + '</span>' +
        '<span class="mini-seg" role="group" aria-label="处理范围">' +
        '<button data-act="clean-range" data-v="table" aria-pressed="' + (S.clean.range === 'table') + '">整张表</button>' +
        '<button data-act="clean-range" data-v="selection" aria-pressed="' + (S.clean.range === 'selection') + '">选中的</button></span>';
      return;
    }
    var info = S.scan && S.scan.sheet === c.sheet && S.auto.phase !== 'start'
      ? ' · ' + esc(S.scan.addr) + (S.scan.rows ? ' · ' + num(S.scan.rows) + ' 行' : '')
      : '';
    el.innerHTML = h + '<span class="grow"><b>' + esc(c.sheet) + '</b>' + info + '</span>';
  }

  function showAlert(kind, text) { S.alert = { kind: kind, text: text }; }
  function renderAlert() {
    var a = S.alert;
    $('alert').innerHTML = a ? '<div class="banner ' + a.kind + '"><span class="grow">' + esc(a.text) + '</span><button class="x" data-act="alert-close" aria-label="关闭">×</button></div>' : '';
  }
  var toastTimer = null;
  function toast(text) {
    var t = $('toast'); t.textContent = text; t.classList.remove('hidden');
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.classList.add('hidden'); }, 2600);
  }

  function renderHelp() {
    var el = $('helpCard');
    $('helpBtn').setAttribute('aria-expanded', String(S.help));
    el.classList.toggle('hidden', !S.help);
    if (!S.help) return;
    el.innerHTML =
      '<div><b>帮我整理</b>：不知道从哪下手就用它。检查 → 勾选 → 在 Excel 弹窗里确认一次 → 完成。</div>' +
      '<div><b>清理工具</b>：知道要做什么时，点开一项单独处理。</div>' +
      '<div><b>处理工具</b>：跨表补全、两表对账、打码、交表前收拾。</div>' +
      '<div><b>小票</b>：每次修改都会记下来，全部做完复制一次。</div>' +
      '<div class="muted tiny">每次修改前都会自动备份当天的原文件；文件不会自动保存；只有在 Excel 弹窗里点「是」才会真的修改。</div>';
  }

  /* ---- 帮我整理 ---- */
  var AUTO_ITEMS = [
    { key: 'Space', stat: 'dirty', safe: true, title: '多余的空格', unit: '格', desc: '文字前后藏着空格，查找和匹配会对不上。' },
    { key: 'Unmerge', stat: 'merges', safe: true, title: '合并的格子', unit: '处', desc: '拆开后每格补上原来的值，才能正常排序、筛选。' },
    { key: 'Unit', stat: 'numText', title: '写成文字的数字', unit: '格', desc: '像「20件」「1,200」这种，改成真正的数字才能求和。单位会去掉，编号前面的 0 会保留。' },
    { key: 'Date', stat: 'dates', title: '日期写法统一', unit: '格', desc: '仅统计现有规则认准的日期常量（不含公式、合并格、超级表表头）。无法识别的原样保留；候选数不保证等于修改数。' },
    { key: 'BlankRows', stat: 'blank', title: '空行', unit: '行', desc: '删掉整行都是空的行，下面的行会往上移。', risky: true },
    { key: 'Dup', stat: 'dup', title: '一模一样的行', unit: '行', desc: '保留第一次出现的那一行。', options: [['Mark', '只标颜色'], ['Dedup', '删掉多余的']], needsVisible: true }
  ];

  function renderAuto() {
    var el = $('view-auto'), a = S.auto;
    if (a.phase === 'done' && a.last) { el.innerHTML = autoDoneHtml(a.last); return; }
    if (a.phase === 'empty') {
      el.innerHTML = '<h1>这张表是空的</h1><p class="lead">切到有数据的表，再点检查。</p>' + checkButton('重新检查');
      return;
    }
    if (a.phase === 'result' && S.scan) { el.innerHTML = autoResultHtml(); return; }
    el.innerHTML =
      '<div style="display:flex;flex-direction:column;gap:8px;padding-top:12px">' +
      '<h1 style="font-size:25px">先检查一下这张表</h1>' +
      '<p class="lead">我会找出多余的空格、合并的格子、空行、重复的行这类问题。检查只看不改。</p></div>' +
      '<div class="steps3">' +
      '<div class="step"><span class="n">1</span><b>我来检查</b><span>几秒钟</span></div>' +
      '<div class="step"><span class="n">2</span><b>你点确定</b><span>Excel 会问你一次</span></div>' +
      '<div class="step"><span class="n">3</span><b>整理好</b><span>先自动备份</span></div></div>' +
      '<div style="flex:1;min-height:40px"></div>' +
      checkButton('检查我的表格') +
      '<p class="foot-note">会自动框出这张表的数据区域。只想处理一部分？去「清理工具」选「选中的」。</p>';
  }
  function checkButton(text) {
    return '<button class="btn btn-primary btn-hero" data-act="check"' + (S.busy ? ' disabled' : '') + '>' +
      (S.auto.checking ? '<span class="spin"></span>正在检查…' : ICON.search + esc(text)) + '</button>';
  }

  function autoResultHtml() {
    var s = S.scan, c = S.auto.choice, found = 0, safeFound = 0, clean = [];
    AUTO_ITEMS.forEach(function (it) {
      if (!it.stat) return;
      if (s[it.stat] > 0) { found++; if (it.safe) safeFound++; } else clean.push(it.title);
    });
    var html = '';
    if (S.stale) {
      var why = S.stale === 'selection' ? '选区已变，旧计划已失效；再次处理会重新框选整表、检查并要求确认' : S.stale + '，下面的数字可能不准了';
      html += '<div class="banner ' + (S.stale === 'selection' ? 'info' : 'warn') + '"><span class="grow">' + esc(why) + '</span>' +
        (S.stale === 'selection' ? '' : '<button class="link-btn" data-act="check">重新检查</button>') + '</div>';
    }
    html += '<div style="display:flex;flex-direction:column;gap:6px;padding:0 4px">' +
      '<h1>' + (found ? '发现 ' + found + ' 类问题' + (safeFound ? '，' + safeFound + ' 类可以直接交给我' : '') : '挺干净的，没发现常见问题') + '</h1>' +
      '<p class="lead">' + (safeFound ? '推荐的已经替你勾好。「要你决定」的会改变数据，看一眼再勾。' : (found ? '下面这些会改变数据，看一眼再决定要不要处理。' : '日期候选使用明确日期规则计数，不会猜测歧义日期。')) + '</p>' +
      (S.stale && S.stale !== 'selection' ? '' : '<button class="link-btn" style="align-self:flex-start" data-act="check">重新检查</button>') + '</div>';

    var safe = AUTO_ITEMS.filter(function (it) { return it.safe && s[it.stat] > 0; });
    if (safe.length) {
      html += '<div class="choice-wrap"><div class="sec-title ok">推荐处理 · 不改变数据内容</div>' + safe.map(choiceHtml).join('') + '</div>';
    }
    var decide = AUTO_ITEMS.filter(function (it) { return !it.safe && (!it.stat || s[it.stat] > 0); });
    html += '<div class="choice-wrap"><div class="sec-title warn">要你决定 · 默认不改，勾上就一起处理</div>' + decide.map(choiceHtml).join('') + '</div>';
    if (clean.length) html += '<div class="okline">' + ICON.check + '<span>没问题：' + esc(clean.join('、')) + '</span></div>';
    if (s.formulas > 0) html += '<div class="okline">' + ICON.check + '<span>表里有 ' + num(s.formulas) + ' 个公式，整理时不会改动它们。</span></div>';

    var acts = autoActions();
    var isDefault = acts.length && !c.Unit && !c.Date && !c.BlankRows && !c.Dup && c.Space === (s.dirty > 0) && c.Unmerge === (s.merges > 0);
    var deletes = acts.filter(function (x) { return x.id === 'BlankRows' || x.id === 'Dedup'; }).length;
    html += '<div class="sticky-foot">' +
      '<button class="btn btn-primary btn-hero" data-act="auto-run"' + (acts.length && !S.busy ? '' : ' disabled') + '>' +
      (acts.length ? (isDefault ? '按推荐方案整理' : '按我的选择整理') + '（' + acts.length + ' 项）' : '先勾上要处理的项') + '</button>' +
      '<p class="foot-note">Excel 会弹窗请你确认一次 · 修改前自动备份' + (deletes ? '<br>删除行放在最后做，基于清理后的数据判断' : '') + '</p></div>';
    return html;
  }

  function choiceHtml(it) {
    var s = S.scan, on = !!S.auto.choice[it.key];
    var count = it.stat ? num(s[it.stat]) + ' ' + it.unit : '';
    var h = '<div class="choice-wrap" style="gap:0">' +
      '<label class="choice' + (on ? ' on' : '') + (on && !it.safe ? ' risky' : '') + '">' +
      '<input type="checkbox" data-bind="auto.choice.' + it.key + '"' + (on ? ' checked' : '') + '>' +
      '<span class="body"><span class="head"><span>' + esc(it.title) + '</span><span class="cnt">' + count + '</span></span>' +
      '<span class="desc">' + esc(it.key === 'Dup' && it.stat && s && s[it.stat] ? '保留第一次出现的那一行；这里会删掉 ' + num(s[it.stat]) + ' 行。' : it.desc) + '</span>' +
      (it.risky ? '<span class="tag warn">会删除整行</span>' : '') + '</span></label>';
    if (it.options && on) {
      h += '<div class="opts" style="margin:8px 0 0 32px">' + it.options.map(function (o) {
        return '<label class="opt"><input type="radio" name="autoDup" data-bind="auto.dupMode" value="' + o[0] + '"' + (S.auto.dupMode === o[0] ? ' checked' : '') + '>' + esc(o[1]) + '</label>';
      }).join('') + '</div>';
    }
    return h + '</div>';
  }

  // Number conversion reports which columns were kept and which units were dropped.
  function unitNote(x) {
    var code = baseCode(x.code), pick = null;
    if (code === 'Unit') pick = /同列单位不统一[^；]*|本次去掉了单位：[^；]*/g;
    else if (code === 'Date') pick = /保留 \d+ 处：月\/日\/年[^；]*/g;
    if (!pick) return '';
    return (String(x.summary || '').match(pick) || []).join('；');
  }
  function unitNoteHtml(x) {
    var n = unitNote(x);
    return n ? '<div class="tiny muted">' + esc(n) + '</div>' : '';
  }
  function resultRowsHtml(out) {
    var rows = out.rows.map(function (x) {
      var detail = '';
      if (x.changed) {
        var list = (x.changeRows || []).map(function (r) {
          var addr = String(r[0] || ''), reason = String(r[2] || ''); addr = addr.replace(/^R(\d+):(\d+)C\d+:\d+$/, function (m, r1, r2) { return r1 === r2 ? '第 ' + r1 + ' 行' : '第 ' + r1 + '–' + r2 + ' 行'; });
          var jump = '';
          var m = reason.match(/首次出现在行(\d+)/);
          if (m) jump = ' <button class="link-btn" data-act="jump" data-text="A' + m[1] + '">跳到行' + m[1] + '</button>';
          return '<div class="skip-item">' + esc(addr) + jump + '</div>';
        });
        detail = '<details class="skip"><summary>改了 ' + num(x.changed) + ' 处（点开看明细' +
          (x.moved === 'yes' ? '；位置已按后面删掉的行换算' : x.moved === 'unsure' ? '；位置是删行之前的' : '') + '）</summary>' +
          '<div class="skip-list">' + (list.slice(0, 80).join('') || '明细见本地审计记录') +
          (list.length > 80 ? '<div class="skip-item">…</div>' : '') + '</div></details>';
      }
      return '<div class="row"><span>' + esc(labelOf(x.code)) + unitNoteHtml(x) + detail + '</span><span class="v">' + num(x.count) + ' ' + (x.unit || unitOf(x.code)) + '</span></div>';
    });
    out.notDone.forEach(function (c) {
      rows.push('<div class="row bad"><span>' + esc(labelOf(c)) + '</span><span class="v">' + (c === out.failedAt ? '在这里停下' : '没做') + '</span></div>');
    });
    return rows.length ? '<div class="card rows">' + rows.join('') + '</div>' : '';
  }
  function skipNoteHtml(out) {
    var list = out.skipped || [];
    if (!list.length) return '';
    var items = list.map(function (s) { return '<div class="skip-item">' + esc(labelOf(s.code)) + '：' + esc(s.why) + '</div>'; }).join('');
    return '<div class="banner warn"><span class="grow">这几步没法做，已经跳过：' + list.length + ' 项' +
      '<details class="skip" style="margin-top:4px"><summary>为什么</summary><div class="skip-list">' + items + '</div></details></span></div>';
  }
  function backupHtml(out) {
    if (!out.backup) return '';
    return '<div class="card kv"><span class="grow"><div class="tiny muted">' + (out.ok ? '修改前的备份已存好' : '可以用这份备份恢复') + '</div></span>' +
      '<button class="link-btn" data-act="open-path" data-text="' + esc(out.backup) + '">打开文件夹</button>' +
      '<button class="link-btn" data-act="copy" data-text="' + esc(out.backup) + '">复制路径</button></div>';
  }
  function autoDoneHtml(out) {
    var skipN = (out.skipped || []).length;
    var html = '<div class="res-head">' + (out.ok ? ICON.okBig : ICON.badBig) + '<div><h1>' +
      (out.ok ? ('整理好了' + (skipN ? '（跳过 ' + skipN + ' 项）' : '')) : (out.rows.length ? '只完成了一部分' : '没有改动')) + '</h1>' +
      '<div class="sub">' + esc(out.sheet) + (out.addr ? ' · ' + esc(out.addr) : '') + ' · 文件还没保存</div></div></div>';
    if (!out.ok) html += '<div class="banner bad"><span class="grow">' + esc(out.error) + (out.rows.length ? '<br>已完成的部分不会自动撤回，需要的话用下面的备份恢复。' : '') + '</span></div>';
    html += skipNoteHtml(out) + resultRowsHtml(out) + backupHtml(out);
    if (out.ok && S.scan) {
      var keptN = (S.scan.protectedNums || 0) + (S.scan.formulas || 0);
      if (keptN > 0) html += '<p class="foot-note">另有 ' + num(keptN) + ' 处按规则没动（编号前导零 / 公式 / 合并格），不影响使用。</p>';
    }
    html += '<div class="okline">' + ICON.check + '<span>已记进今天的小票，全部做完再去复制。</span></div>';
    var s = S.scan, left = [];
    if (out.ok && s) {
      AUTO_ITEMS.forEach(function (it) {
        if (it.stat && s[it.stat] > 0 && !S.auto.choice[it.key]) left.push(it);
      });
    }
    if (left.length) {
      html += '<p class="foot-note" style="padding:0 4px">这次没处理：' + esc(left.map(function (it) { return it.title + '（' + num(s[it.stat]) + ' ' + it.unit + '）'; }).join('、')) +
        ' <button class="link-btn" data-act="check-again" data-keys="' + left.map(function (it) { return it.key; }).join(',') + '">也处理掉</button></p>';
    }
    html += '<div style="flex:1"></div><div class="btn-row"><button class="btn btn-ghost" data-act="go-tab" data-v="receipt">看小票</button>' +
      '<button class="btn btn-ghost" data-act="check">再检查一次</button></div>';
    return html;
  }

  /* ---- 清理工具 ---- */
  var CLEAN_TOOLS = [
    { key: 'Space', title: '清掉多余的空格', stat: 'dirty', unit: '格', desc: '清掉文字前后和中间多余的空格、看不见的字符。英文单词之间的空格、编号前面的 0 和公式都不动。表里有合并格时要先拆开。' },
    { key: 'Unmerge', title: '拆开合并的格子', stat: 'merges', unit: '处', desc: '拆开上下方向、单列的合并格，并把原来的值填进每一格。横向的合并会跳过。' },
    { key: 'Unit', title: '数字和单位分开', stat: 'numText', unit: '格', desc: '把「20件」「1,200」这类写成文字的数字变成真正的数字。',
      options: [['Unit', '原地改成数字'], ['Split', '单位拆到右边新列']], bind: 'clean.unitMode' },
    { key: 'Date', title: '统一日期写法', stat: 'dates', unit: '格', desc: '能认准的日期文字改成真正的日期；认不准的原样保留，不会乱猜。' },
    { key: 'BlankRows', title: '删除空行', stat: 'blank', unit: '行', desc: '删掉整行都是空的行，下面的行会往上移。' },
    { key: 'Dup', title: '处理重复的行', stat: 'dup', unit: '行', desc: '按整行内容判断，保留第一次出现的那一行；含公式或备注的重复行会保留。删除前请先清除筛选。',
      options: [['Mark', '只标颜色'], ['Dedup', '删掉多余的']], bind: 'clean.dupMode' },
    { key: 'ClearTemporary', title: '去掉插件留下的标记', desc: '清掉电子羊表格生成的体检报告和标出的颜色。你自己的表和原来的底色不动。作用于整个文件。' }
  ];
  function cleanActions(key) {
    if (key === 'Unit') return [{ id: S.clean.unitMode === 'Unit' && S.clean.forceMixed ? 'Unit:Force' : S.clean.unitMode }];
    if (key === 'Dup') return [{ id: S.clean.dupMode }];
    return [{ id: key }];
  }
  function cleanRange(key) {
    if (key === 'ClearTemporary') return 'table';
    if (key === 'Unit' && S.clean.unitMode === 'Split') return 'selection';
    return S.clean.range;
  }

  function renderClean() {
    var el = $('view-clean'), s = S.scan && S.scan.sheet === S.ctx.sheet ? S.scan : null;
    var html = '<p class="lead" style="padding:0 4px">点一项展开，单独处理。' +
      (s ? '数字来自上次检查' + (S.stale ? '（之后表格有变动，可能不准）' : '') + '。' : '还没检查过，') +
      (s ? '' : '<button class="link-btn" data-act="clean-scan">先检查一下</button>') + '</p>';
    // 清理/处理工具是工具箱：不管表里有没有问题，所有工具都在，数量只是提示
    html += '<div class="acc">' + CLEAN_TOOLS.map(function (t) {
      var open = S.clean.open === t.key;
      var pill = t.stat && s ? '<span class="pill' + (s[t.stat] ? '' : ' zero') + '">' + num(s[t.stat]) + '</span>' : '';
      var h = '<div class="acc-item' + (open ? ' open' : '') + '" id="clean-' + t.key + '">' +
        '<button class="acc-head" data-act="acc" data-scope="clean" data-v="' + t.key + '" aria-expanded="' + open + '">' +
        '<span class="t"><b>' + esc(t.title) + '</b></span>' + pill + ICON.chev + '</button>';
      if (open) h += '<div class="acc-body">' + cleanBody(t, s) + '</div>';
      return h + '</div>';
    }).join('') + '</div>';
    html += '<p class="foot-note">想一次处理好几项？用 <button class="link-btn" data-act="go-tab" data-v="auto">帮我整理</button>，它会排好顺序一起做。</p>';
    el.innerHTML = html;
  }
  function cleanBody(t, s) {
    var h = '<p>' + esc(t.desc) + '</p>';
    if (t.options) {
      var cur = getPath(t.bind);
      h += '<div class="opts">' + t.options.map(function (o) {
        return '<label class="opt"><input type="radio" name="opt-' + t.key + '" data-bind="' + t.bind + '" value="' + o[0] + '"' + (cur === o[0] ? ' checked' : '') + '>' + esc(o[1]) + '</label>';
      }).join('') + '</div>';
    }
    if (t.key === 'Unit' && S.clean.unitMode === 'Split') h += '<p class="tiny">拆到新列要先在 Excel 里<b>只选中这一列</b>，右边会插入一列放单位。</p>';
    if (t.key === 'Unit' && S.clean.unitMode === 'Unit') {
      // One-run override of the mixed-unit guard; reset after every run (finishRun).
      h += '<label class="opt" style="margin-top:6px"><input type="checkbox" data-bind="clean.forceMixed"' + (S.clean.forceMixed ? ' checked' : '') + '>同列有多个单位时也转换</label>' +
        (S.clean.forceMixed ? '<p class="tiny" style="color:var(--bad)">⚠ 会去掉单位文字只剩数字（件/台/箱 等信息会丢），只影响这一次。</p>'
          : '<p class="tiny">默认：一列里单位不统一时整列不动，避免把单位吞掉。</p>');
    }
    if (t.key === 'Dup' && S.clean.dupMode === 'Dedup') h += '<p class="tiny" style="color:var(--warn)">会删除整行，下面的行往上移。</p>';
    var range = cleanRange(t.key);
    h += '<p class="tiny">范围：' + (t.key === 'ClearTemporary' ? '整个文件' : range === 'table' ? '整张表的数据（自动框选）' : '你在 Excel 里选中的区域' + (S.ctx.sel ? '（现在是 ' + esc(S.ctx.sel) + '）' : '')) + '</p>';
    var res = S.clean.results[t.key];
    if (res) h += inlineResultHtml(res);
    var count = (t.stat && s && s[t.stat] && range === 'table') ? '这 ' + num(s[t.stat]) + ' ' + t.unit : '';
    h += '<button class="btn btn-primary" data-act="clean-run" data-v="' + t.key + '"' + (S.busy ? ' disabled' : '') + '>' + (res ? '再做一次' : '处理' + count) + '</button>' +
      '<span class="foot-note">Excel 会弹窗确认 · 修改前自动备份</span>';
    return h;
  }
  // Only suggest a bounded range from actual scan metadata; never guess row 50.
  // Preserve the header so an explicit hasHeader option is applied exactly once.
  function suggestedSelectionRange(selection, scanAddress) {
    var col = String(selection || '').match(/^\$?([A-Z]{1,3})/);
    var bounds = String(scanAddress || '').match(/^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/);
    if (!col || !bounds) return '';
    var first = Number(bounds[2]), last = Number(bounds[4] || bounds[2]);
    if (first < 1 || last < first || last > 1048576) return '';
    function columnIndex(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
    var chosen = columnIndex(col[1]);
    if (chosen < columnIndex(bounds[1]) || chosen > columnIndex(bounds[3] || bounds[1])) return '';
    return col[1] + first + ':' + col[1] + last;
  }

  function inlineResultHtml(out) {
    var h = '<div class="inline-result ' + (out.ok ? 'ok' : 'bad') + '"><b>' +
      (out.ok ? '做好了 · ' + hhmm(out.t) : (out.rows.length ? '只完成了一部分' : '没有改动')) + '</b>';
    if (out.ok) h += '<span>' + out.rows.map(function (x) { return esc(labelOf(x.code)) + ' ' + num(x.count) + ' ' + (x.unit || unitOf(x.code)) + '（' + num(x.count) + ' ' + (x.unit || unitOf(x.code)) + '）'; }).join('；') + (out.rows.length ? '' : '没有需要改的') + '</span>' +
      out.rows.map(unitNoteHtml).join('');
    else {
      h += '<span>' + esc(out.error) + '</span>';
      // 选区类错误：给「帮我选好」，自动选中当前列的数据区
      if (/选区|单列|数据区|太大/.test(out.error || '')) {
        var help = suggestedSelectionRange(S.ctx.sel, S.scan && S.scan.addr);
        if (help) h += '<button class="link-btn" data-act="jump" data-text="' + esc(help) + '">帮我选好（' + esc(help) + '）</button>';
      }
    }
    if (out.backup) h += '<span class="path">备份：' + esc(out.backup) + '</span>';
    h += '</div>';
    return h;
  }

  /* ---- 处理工具 ---- */
  var PROC_TOOLS = [
    { key: 'lookup', title: '跨表补全', sub: '按编号把另一张表的数据填过来' },
    { key: 'append', title: '多表纵向追加', sub: '把多张结构相近的表合成一张明细表' },
    { key: 'unpivot', title: '宽表转长表', sub: '把横着排的月份/地区列转成一维明细表' },
    { key: 'splitsheet', title: '按列拆表', sub: '按部门/地区把一张总表拆成多张独立分表' },
    { key: 'recon', title: '两表对账', sub: '找出两张表对不上的地方' },
    { key: 'mask', title: '隐藏敏感信息', sub: '手机号、身份证号、姓名打码' },
    { key: 'delivery', title: '交表前收拾', sub: '把表整理成可以发给别人的样子' }
  ];
  function renderProcess() {
    var el = $('view-process');
    el.innerHTML = '<div class="acc">' + PROC_TOOLS.map(function (t) {
      var open = S.proc.open === t.key;
      var h = '<div class="acc-item' + (open ? ' open' : '') + '" id="proc-' + t.key + '">' +
        '<button class="acc-head" data-act="acc" data-scope="proc" data-v="' + t.key + '" aria-expanded="' + open + '">' +
        '<span class="t"><b>' + esc(t.title) + '</b><span>' + esc(t.sub) + '</span></span>' + ICON.chev + '</button>';
      if (open) h += '<div class="acc-body">' + procBody(t.key) + '</div>';
      return h + '</div>';
    }).join('') + '</div>';
  }
  function selectHtml(bind, choices, cur) {
    return '<select data-bind="' + bind + '">' + choices.map(function (c) {
      return '<option value="' + esc(c[0]) + '"' + (c[0] === cur ? ' selected' : '') + '>' + esc(c[1]) + '</option>';
    }).join('') + '</select>';
  }
  function colChoices(cols) { return cols.map(function (c) { return [c[0], c[0] + ' · ' + (c[1] || '（无表头）')]; }); }
  function selNowHtml(hint) {
    return '<div class="selnow">' + ICON.table + '<span class="grow">' + esc(hint) + '：<b>' + esc(S.ctx.sel || '—') + '</b>' +
      (S.ctx.sheet ? '（' + esc(S.ctx.sheet) + '）' : '') + '</span><button class="link-btn" data-act="ctx-refresh">刷新</button></div>';
  }
  function procBody(key) {
    var res = S.proc.results[key], h = '';
    if (key === 'lookup' || key === 'append' || key === 'unpivot' || key === 'splitsheet' || key === 'recon') {
      if (!S.proc.meta) return '<p><span class="spin"></span> 正在读取各张表的列名…</p>';
    }
    if (key === 'lookup') {
      var lk = S.proc.lk, meta = S.proc.meta;
      var srcSheets = meta.filter(function (m) { return m[0] !== lk.target; });
      if (!srcSheets.length) return '<p>这个文件里只有「' + esc(lk.target) + '」一张表，没有可以取数据的另一张表。</p>';
      syncLookupColumns(lk);
      var srcCols = colsOf(lk.src), availRetCols = srcCols.filter(function (c) { return c[0] !== lk.srcKey; });
      var retSet = {};
      (lk.retCols || []).forEach(function (c) { retSet[c] = true; });
      if (S.ctx.sheet && S.ctx.sheet !== lk.target) h += '<div class="banner warn"><span class="grow">你现在在「' + esc(S.ctx.sheet) + '」。补全会填到当前表，请切回「' + esc(lk.target) + '」再点开始，或</span><button class="link-btn" data-act="lk-retarget">改成填到当前表</button></div>';
      h += '<div class="inner"><div class="inner-title"><span class="num-badge">1</span>从这张表取</div>' +
        '<label class="field">表' + selectHtml('proc.lk.src', srcSheets.map(function (m) { return [m[0], m[0]]; }), lk.src) + '</label>' +
        '<label class="field">按哪列对上' + selectHtml('proc.lk.srcKey', colChoices(srcCols), lk.srcKey) + '</label>' +
        '<div class="ret-head"><span>要取哪些列（已选 <b>' + (lk.retCols ? lk.retCols.length : 0) + '</b> 列）</span>' +
        '<span class="ret-actions"><button type="button" class="link-btn" data-act="lk-ret-all">全选其余列</button>' +
        (lk.retCols && lk.retCols.length > 1 ? '<button type="button" class="link-btn" data-act="lk-ret-first">只选第1列</button>' : '') + '</span></div>';
      if (!availRetCols.length) {
        h += '<div class="banner warn"><span class="grow">来源表只有 1 列，没有可取出的其他字段列。</span></div>';
      } else {
        h += '<div class="ret-grid">' + availRetCols.map(function (c) {
          var on = !!retSet[c[0]];
          return '<label class="ret-chk' + (on ? ' on' : '') + '"><input type="checkbox" data-act="lk-ret-toggle" data-v="' + esc(c[0]) + '"' + (on ? ' checked' : '') + '><span>' + esc(c[0] + ' · ' + (c[1] || '（无表头）')) + '</span></label>';
        }).join('') + '</div>';
      }
      h += '</div>' +
        '<div style="display:flex;justify-content:center">' + ICON.down + '</div>' +
        '<div class="inner"><div class="inner-title"><span class="num-badge">2</span>填到「' + esc(lk.target) + '」</div>' +
        '<div class="field pick"><span>按哪列对上</span>' + selectHtml('proc.lk.tgtKey', colChoices(colsOf(lk.target)).concat(hasCol(lk.target, lk.tgtKey) ? [] : [[lk.tgtKey, lk.tgtKey + ' 列']]), lk.tgtKey) +
        '<button class="pick-btn" data-act="lk-pick" data-v="tgtKey" aria-label="用 Excel 里选中的列" title="先在表里选中一列再点">' + ICON.pick + '</button></div>' +
        '<div class="opts"><label class="opt"><input type="radio" name="lkWriteModeType" data-bind="proc.lk.writeModeType" value="auto"' + (lk.writeModeType !== 'explicit' ? ' checked' : '') + '>在表右侧新增列（推荐）</label>' +
        '<label class="opt"><input type="radio" name="lkWriteModeType" data-bind="proc.lk.writeModeType" value="explicit"' + (lk.writeModeType === 'explicit' ? ' checked' : '') + '>填入指定列</label></div>';
      var autoInfo = autoAppendPreviewInfo(lk);
      if (lk.writeModeType !== 'explicit') {
        var autoDesc = autoInfo.map(function (x) { return x.tgtCol + ' 列（' + x.header + '）'; }).join('、');
        h += '<div class="auto-note">将在「' + esc(lk.target) + '」最右侧安全新增 <b>' + autoInfo.length + '</b> 列：' + esc(autoDesc || '—') + '。若右侧已有数据、合并格或超级表会自动中止保护。</div>';
      } else {
        var wcChoices = colChoices(writeColChoices(lk.target));
        h += '<div class="map-list">' + (lk.retCols || []).map(function (rc, idx) {
          var curW = (lk.writeCols && lk.writeCols[idx]) || 'B';
          var listForRow = wcChoices.concat(wcChoices.some(function (c) { return c[0] === curW; }) ? [] : [[curW, curW + ' 列']]);
          return '<div class="map-row"><span class="map-src" title="来源 ' + esc(rc + ' · ' + colLabel(lk.src, rc)) + '">来源 ' + esc(rc + ' · ' + colLabel(lk.src, rc)) + ' →</span>' +
            selectHtml('proc.lk.writeCol_' + idx, listForRow, curW) +
            '<button class="pick-btn" style="width:36px;height:36px" data-act="lk-pick" data-v="write_' + idx + '" aria-label="用 Excel 里选中的列" title="先在表里选中一列再点">' + ICON.pick + '</button></div>';
        }).join('') + '</div>';
      }
      h += '</div>';
      var retLabels = (lk.retCols || []).map(function (rc) { return colLabel(lk.src, rc); }).join('、');
      var tgtDesc = lk.writeModeType !== 'explicit'
        ? ('最右侧新增 ' + (lk.retCols ? lk.retCols.length : 1) + ' 列（' + autoInfo.map(function (x) { return x.tgtCol; }).join('、') + ' 列）')
        : ((lk.writeCols || []).join('、') + ' 列');
      h += '<p class="sentence">按<b>' + esc(colLabel(lk.target, lk.tgtKey)) + '</b>去「' + esc(lk.src) + '」里找，把<b>' + esc(retLabels) + '</b>填进「' + esc(lk.target) + '」的 <b>' + esc(tgtDesc) + '</b>。</p>';
      h += '<details class="more"' + (S.proc.preview && S.proc.preview.dupKeys ? ' open' : '') + '><summary>更多设置：重复编号、写入方式</summary><div class="inner">' +
        '<label class="field">编号重复时' + selectHtml('proc.lk.dupRule', [['abort', '先停下，让我处理（推荐）'], ['first', '用第一次出现的那条'], ['skip', '这些编号都不填']], lk.dupRule) + '</label>' +
        '<label class="field">写入方式' + selectHtml('proc.lk.writeMode', [['static', '固定的值（推荐）'], ['formula', '公式，来源变了会跟着变']], lk.writeMode) + '</label></div></details>';
      var pv = S.proc.preview, fresh = false;
      try { fresh = pv && S.proc.previewKey === JSON.stringify(lookupOptions()); } catch (e) { fresh = false; }
      if (pv && fresh) {
        h += '<div class="inner"><div class="chips"><span class="chip ok">能填上 ' + num(pv.matched) + ' 格</span>' +
          (pv.unmatched ? '<span class="chip warn">没找到 ' + num(pv.unmatched) + ' 行</span>' : '') +
          (pv.sourceErrors ? '<span class="chip warn">来源含错误值 ' + num(pv.sourceErrors) + ' 格（将跳过）</span>' : '') +
          (pv.dupKeys ? '<span class="chip warn">来源编号重复 ' + num(pv.dupKeys) + '</span>' : '') +
          (pv.typeDiff ? '<span class="chip warn">格式不一致 ' + num(pv.typeDiff) + '</span>' : '') +
          (pv.protectedCells ? '<span class="chip">已有内容不覆盖 ' + num(pv.protectedCells) + ' 格</span>' : '') +
          (pv.blankKeys ? '<span class="chip">编号为空 ' + num(pv.blankKeys) + ' 行</span>' : '') + '</div>';
        if (pv.samples.length) {
          h += '<table class="tbl"><tr><th>位置</th><th>编号</th><th>会填入</th></tr>' + pv.samples.slice(0, 4).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>';
          }).join('') + '</table>';
        }
        if (pv.anomalies.length) {
          h += '<table class="tbl"><tr><th>位置</th><th>情况</th></tr>' + pv.anomalies.slice(0, 5).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + (r[2] ? '：' + esc(r[2]) : '') + '</td></tr>';
          }).join('') + '</table>';
        }
        if (!pv.canExecute || pv.block) h += '<div class="banner warn"><span class="grow">现在还不能写入：' + esc(pv.block || '请检查列的选择。') + '</span></div>';
        h += '</div>';
      }
      if (res) h += inlineResultHtml(res);
      var canRun = pv && fresh && pv.canExecute && !pv.block;
      h += canRun
        ? '<button class="btn btn-primary" data-act="lk-run"' + (S.busy ? ' disabled' : '') + '>没问题，开始补全（' + num(pv.matched) + ' 格）</button>'
        : '<button class="btn btn-dark" data-act="lk-preview"' + (S.busy ? ' disabled' : '') + '>' + (S.proc.previewing ? '<span class="spin"></span>正在预览…' : '先预览一下') + '</button>';
      h += '<span class="foot-note">预览只看不改 · 开始补全时 Excel 会弹窗确认</span>';
      return h;
    }
    if (key === 'append') {
      var ap = S.proc.ap, metaAll = S.proc.meta || [];
      if (metaAll.length < 2) return '<p>这个文件里只有 ' + metaAll.length + ' 张数据表，至少要有 2 张表才能纵向追加汇总。</p>';
      var selSet = {};
      (ap.sheets || []).forEach(function (n) { selSet[n] = true; });
      h += '<div class="inner"><div class="inner-title"><span class="num-badge">1</span><span>勾选要合并哪些表（已选 <b>' + (ap.sheets ? ap.sheets.length : 0) + '</b> 张）</span></div>' +
        '<div class="ret-head"><span>按勾选顺序从上往下追加，原表不会改动</span>' +
        '<span class="ret-actions"><button type="button" class="link-btn" data-act="ap-sheet-all">全选有数据的表</button>' +
        (ap.sheets && ap.sheets.length > 2 ? '<button type="button" class="link-btn" data-act="ap-sheet-first2">只选前2张</button>' : '') + '</span></div>' +
        '<div class="ret-grid">' + metaAll.map(function (m) {
          var sName = m[0], rCnt = Math.max(0, (Number(m[2]) || 0) - 1), cCnt = Array.isArray(m[3]) ? m[3].length : 0;
          var on = !!selSet[sName];
          return '<label class="ret-chk' + (on ? ' on' : '') + '"><input type="checkbox" data-act="ap-sheet-toggle" data-v="' + esc(sName) + '"' + (on ? ' checked' : '') + '>' +
            '<span title="' + esc(sName + '（' + rCnt + ' 行 · ' + cCnt + ' 列）') + '">' + esc(sName) + ' <small style="color:var(--ink-2)">(' + rCnt + '行)</small></span></label>';
        }).join('') + '</div></div>';
      h += '<div style="display:flex;justify-content:center">' + ICON.down + '</div>';
      h += '<div class="inner"><div class="inner-title"><span class="num-badge">2</span>列名对齐与小计行处理</div>' +
        '<label class="field">列名对齐' + selectHtml('proc.ap.colMode', [['union', '保留全部列，缺失格留空（推荐）'], ['common', '只保留每张表都有的同名列']], ap.colMode) + '</label>' +
        '<label class="field">小计/合计' + selectHtml('proc.ap.summaryRowRule', [['exclude', '识别并排除，不混入明细（推荐）'], ['keep', '全部保留，当成普通行追加']], ap.summaryRowRule) + '</label>' +
        '<label class="field">来源标记' + selectHtml('proc.ap.sourceTagMode', [['sheet', '首列加上「来源工作表」（推荐）'], ['sheet_row', '首列加上「表名!行号」'], ['none', '不加来源标记列']], ap.sourceTagMode) + '</label></div>';
      h += '<p class="sentence">把 <b>' + esc((ap.sheets || []).join('、') || '所选工作表') + '</b> 按表头对齐纵向合并到新工作表「<b>' + esc(ap.targetSheetName || '多表追加汇总') + '</b>」' +
        (ap.summaryRowRule !== 'keep' ? '，自动剔除底部或中间的小计/合计行并记入小票' : '') + '。</p>';
      h += '<p class="tiny">不同名列不会自动合并；例如「产品编码」与「商品编号」默认保留为两列。需要合并时请展开下面的同义列名设置。</p>';
      h += '<details class="more"><summary>更多设置：新建表名、手动指定同义列名</summary><div class="inner">' +
        '<label class="field">新表名称<input type="text" class="txt-in" data-bind="proc.ap.targetSheetName" value="' + esc(ap.targetSheetName || '多表追加汇总') + '" maxlength="26" placeholder="多表追加汇总"></label>' +
        '<label class="field">同义列名<input type="text" class="txt-in" data-bind="proc.ap.colMap" value="' + esc(ap.colMap || '') + '" placeholder="例：产品编码=商品编号;单价=售价"></label>' +
        '<span class="tiny muted">插件默认只自动对齐完全同名的列（忽略首尾空格和大小写），绝不乱猜近义词。若两张表同一列叫法不同，可在此按「旧列名=统一列名」显式指定。</span></div></details>';
      var apv = S.proc.apPreview, apFresh = false;
      try { apFresh = apv && S.proc.apPreviewKey === JSON.stringify(appendOptions()); } catch (e) { apFresh = false; }
      if (apv && apFresh) {
        h += '<div class="inner"><div class="chips">' +
          '<span class="chip ok">可追加 ' + num(apv.netRows) + ' 行明细</span>' +
          '<span class="chip">合并后 ' + num(apv.colCount) + ' 列' + (apv.sourceTagMode !== 'none' ? '（另加来源列）' : '') + '</span>' +
          (apv.summaryRows ? '<span class="chip warn">识别出小计/合计 ' + num(apv.summaryRows) + ' 行（' + (apv.summaryRowRule === 'exclude' ? '将排除' : '将保留') + '）</span>' : '<span class="chip">无小计/合计行干扰</span>') +
          (apv.unionOnlyCols ? '<span class="chip warn">仅部分表有的列 ' + num(apv.unionOnlyCols) + ' 个</span>' : '') +
          (apv.sourceErrors ? '<span class="chip warn">来源含错误值 ' + num(apv.sourceErrors) + ' 格（将跳过）</span>' : '') +
          '<span class="chip">输出至新表「' + esc(apv.targetSheet) + '」</span></div>';
        if (apv.summarySamples && apv.summarySamples.length) {
          h += '<div class="auto-note" style="background:var(--warn-soft);border-color:#E8C98F;color:var(--warn-ink)">' +
            '<b>请确认识别出的小计/合计行（共 ' + num(apv.summaryRows) + ' 行）：</b>' +
            (apv.summaryRowRule === 'exclude' ? '当前设为「排除」，以下行不会混入汇总明细，并会记入审计日志；如不想排除可在上方改为「全部保留」。' : '当前设为「全部保留」，以下行将一并追加。') + '</div>';
          h += '<table class="tbl"><tr><th>位置</th><th>行首标识</th><th>识别原因</th></tr>' + apv.summarySamples.slice(0, 6).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>';
          }).join('') + '</table>';
        }
        if (apv.headers && apv.headers.length) {
          h += '<table class="tbl"><tr><th>对齐后列名</th><th>出现表数</th></tr>' + apv.headers.slice(0, 8).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + num(r[1]) + ' / ' + num(r[2]) + ' 张表</td></tr>';
          }).join('') + '</table>';
        }
        if (apv.anomalies && apv.anomalies.length) {
          h += '<table class="tbl"><tr><th>位置/字段</th><th>情况</th></tr>' + apv.anomalies.slice(0, 6).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + (r[2] ? '：' + esc(r[2]) : '') + '</td></tr>';
          }).join('') + '</table>';
        }
        if (!apv.canExecute || apv.block) h += '<div class="banner warn"><span class="grow">现在还不能执行追加：' + esc(apv.block || '请检查勾选的工作表与表头。') + '</span></div>';
        h += '</div>';
      }
      if (res) h += inlineResultHtml(res);
      var canRunAp = apv && apFresh && apv.canExecute && !apv.block;
      h += canRunAp
        ? '<button class="btn btn-primary" data-act="ap-run"' + (S.busy ? ' disabled' : '') + '>没问题，开始追加汇总（' + num(apv.netRows) + ' 行）</button>'
        : '<button class="btn btn-dark" data-act="ap-preview"' + (S.busy ? ' disabled' : '') + '>' + (S.proc.apPreviewing ? '<span class="spin"></span>正在预览表头与小计行…' : '先预览一下') + '</button>';
      h += '<span class="foot-note">预览只看不改 · 开始追加时会自动新建工作表，原表不改动</span>';
      return h;
    }
    if (key === 'unpivot') {
      var up = S.proc.up;
      if (S.ctx.sheet && up.sheet !== S.ctx.sheet) initUnpivotDefaults(true);
      var uCols = colsOf(S.ctx.sheet);
      if (uCols.length < 2) return '<p>当前表「' + esc(S.ctx.sheet) + '」至少要有 2 列（至少 1 列固定维度 + 1 列要展开的指标）才能做宽表转长表。</p>';
      var idSet = {}, valSet = {};
      (up.idCols || []).forEach(function (c) { idSet[c] = true; });
      (up.valCols || []).forEach(function (c) { valSet[c] = true; });
      h += '<div class="inner"><div class="inner-title"><span class="num-badge">1</span><span>勾选左侧固定不变的列（已选 <b>' + (up.idCols ? up.idCols.length : 0) + '</b> 列）</span></div>' +
        '<div class="ret-head"><span>像姓名、编号、部门这种每行都要保留的列</span>' +
        '<span class="ret-actions"><button type="button" class="link-btn" data-act="up-id-first1">只选第1列</button>' +
        (uCols.length >= 3 ? '<button type="button" class="link-btn" data-act="up-id-first2">选前2列</button>' : '') + '</span></div>' +
        '<div class="ret-grid">' + uCols.map(function (c) {
          var on = !!idSet[c[0]];
          return '<label class="ret-chk' + (on ? ' on' : '') + '"><input type="checkbox" data-act="up-id-toggle" data-v="' + esc(c[0]) + '"' + (on ? ' checked' : '') + '>' +
            '<span>' + esc(c[0] + ' · ' + (c[1] || '（无表头）')) + '</span></label>';
        }).join('') + '</div></div>';
      h += '<div style="display:flex;justify-content:center">' + ICON.down + '</div>';
      var availValCols = uCols.filter(function (c) { return !idSet[c[0]]; });
      h += '<div class="inner"><div class="inner-title"><span class="num-badge">2</span><span>勾选要往下展开的列（已选 <b>' + (up.valCols ? up.valCols.length : 0) + '</b> 列）</span></div>' +
        '<div class="ret-head"><span>像 1月、2月、3月或各地区这种横着排的列</span>' +
        '<span class="ret-actions"><button type="button" class="link-btn" data-act="up-val-rest">全选其余业务列</button></span></div>';
      if (!availValCols.length) {
        h += '<div class="banner warn"><span class="grow">请先在上方少勾几个固定列，留出需要展开的列。</span></div>';
      } else {
        h += '<div class="ret-grid">' + availValCols.map(function (c) {
          var on = !!valSet[c[0]], isSumCol = isSummaryHeaderName(c[1]);
          return '<label class="ret-chk' + (on ? ' on' : '') + '"><input type="checkbox" data-act="up-val-toggle" data-v="' + esc(c[0]) + '"' + (on ? ' checked' : '') + '>' +
            '<span>' + esc(c[0] + ' · ' + (c[1] || '（无表头）')) + (isSumCol ? ' <small style="color:var(--warn)">(汇总列)</small>' : '') + '</span></label>';
        }).join('') + '</div>';
      }
      h += '<label class="field">空单元格' + selectHtml('proc.up.emptyPolicy', [['skip_blank', '跳过空白格，不生成空行（推荐）'], ['skip_blank_zero', '空白格和等于 0 的格都跳过'], ['keep_all', '全部保留（含空白格）']], up.emptyPolicy) + '</label>' +
        '<label class="field">小计/合计' + selectHtml('proc.up.summaryRule', [['exclude', '自动排除合计行与合计列（推荐）'], ['keep', '全部保留，当成普通数据展开']], up.summaryRule) + '</label></div>';
      h += '<p class="sentence">保留「' + esc(S.ctx.sheet) + '」的 <b>' + esc((up.idCols || []).join('、') || '—') + ' 列</b>，把 <b>' + esc((up.valCols || []).join('、') || '—') + ' 列</b> 展开成「<b>' + esc(up.attrHeader || '属性类别') + '</b>」和「<b>' + esc(up.valHeader || '数值') + '</b>」两列，写入新表「<b>' + esc(up.targetSheetName || '宽表转长表明细') + '</b>」。</p>';
      h += '<details class="more"><summary>更多设置：自定义展开后的列名、新表名称</summary><div class="inner">' +
        '<label class="field">属性列名<input type="text" class="txt-in" data-bind="proc.up.attrHeader" value="' + esc(up.attrHeader || '属性类别') + '" maxlength="32" placeholder="例：月份 / 地区"></label>' +
        '<label class="field">数值列名<input type="text" class="txt-in" data-bind="proc.up.valHeader" value="' + esc(up.valHeader || '数值') + '" maxlength="32" placeholder="例：销量 / 金额"></label>' +
        '<label class="field">新表名称<input type="text" class="txt-in" data-bind="proc.up.targetSheetName" value="' + esc(up.targetSheetName || '宽表转长表明细') + '" maxlength="26" placeholder="宽表转长表明细"></label></div></details>';
      var upv = S.proc.upPreview, upFresh = false;
      try { upFresh = upv && S.proc.upPreviewKey === JSON.stringify(unpivotOptions()); } catch (e) { upFresh = false; }
      if (upv && upFresh) {
        var totalSumHits = (upv.summaryRowCount || 0) + (upv.summaryColCount || 0);
        h += '<div class="inner"><div class="chips">' +
          '<span class="chip ok">可生成 ' + num(upv.netOutputRows) + ' 行明细</span>' +
          '<span class="chip">输出 ' + num(upv.idColCount + 2) + ' 列</span>' +
          (totalSumHits ? '<span class="chip warn">识别出合计行/列 ' + num(totalSumHits) + ' 处（' + (upv.summaryRule === 'exclude' ? '将排除' : '将保留') + '）</span>' : '') +
          (upv.skippedBlanks ? '<span class="chip">过滤空值/零值 ' + num(upv.skippedBlanks) + ' 格</span>' : '') +
          (upv.sourceErrors ? '<span class="chip warn">来源含错误值 ' + num(upv.sourceErrors) + ' 格（将跳过）</span>' : '') +
          '<span class="chip">输出至新表「' + esc(upv.targetSheet) + '」</span></div>';
        if (upv.summarySamples && upv.summarySamples.length) {
          h += '<div class="auto-note" style="background:var(--warn-soft);border-color:#E8C98F;color:var(--warn-ink)">' +
            '<b>请确认识别出的合计行/列（共 ' + num(totalSumHits) + ' 处）：</b>' +
            (upv.summaryRule === 'exclude' ? '当前设为「自动排除」，以下汇总行/列不会混入展开后的明细。' : '当前设为「全部保留」，以下汇总行/列将一并展开。') + '</div>';
          h += '<table class="tbl"><tr><th>位置</th><th>名称</th><th>识别原因</th></tr>' + upv.summarySamples.slice(0, 5).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>';
          }).join('') + '</table>';
        }
        if (upv.samples && upv.samples.length) {
          h += '<table class="tbl"><tr><th>固定列内容</th><th>' + esc(upv.attrHeader) + '</th><th>' + esc(upv.valHeader) + '</th></tr>' + upv.samples.slice(0, 5).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>';
          }).join('') + '</table>';
        }
        if (upv.anomalies && upv.anomalies.length) {
          h += '<table class="tbl"><tr><th>位置</th><th>情况</th></tr>' + upv.anomalies.slice(0, 5).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + (r[2] ? '：' + esc(r[2]) : '') + '</td></tr>';
          }).join('') + '</table>';
        }
        if (!upv.canExecute || upv.block) h += '<div class="banner warn"><span class="grow">现在还不能执行转换：' + esc(upv.block || '请检查勾选的列。') + '</span></div>';
        h += '</div>';
      }
      if (res) h += inlineResultHtml(res);
      var canRunUp = upv && upFresh && upv.canExecute && !upv.block;
      h += canRunUp
        ? '<button class="btn btn-primary" data-act="up-run"' + (S.busy ? ' disabled' : '') + '>没问题，开始转换（生成 ' + num(upv.netOutputRows) + ' 行）</button>'
        : '<button class="btn btn-dark" data-act="up-preview"' + (S.busy ? ' disabled' : '') + '>' + (S.proc.upPreviewing ? '<span class="spin"></span>正在预览转换结果…' : '先预览一下') + '</button>';
      h += '<span class="foot-note">预览只看不改 · 开始转换时会自动新建工作表，原宽表不改动</span>';
      return h;
    }
    if (key === 'splitsheet') {
      var sp = S.proc.sp;
      if (S.ctx.sheet && sp.sheet !== S.ctx.sheet) initSplitSheetDefaults(true);
      var spCols = colsOf(S.ctx.sheet);
      if (!spCols.length) return '<p>当前表「' + esc(S.ctx.sheet) + '」是空的，没有可以拆分的列。</p>';
      var spChoices = colChoices(spCols).concat(hasCol(S.ctx.sheet, sp.splitCol) ? [] : [[sp.splitCol || 'A', (sp.splitCol || 'A') + ' 列']]);
      h += '<div class="inner"><div class="inner-title"><span class="num-badge">1</span><span>按哪一列拆分「' + esc(S.ctx.sheet) + '」</span></div>' +
        '<div class="field pick"><span>拆分依据列</span>' + selectHtml('proc.sp.splitCol', spChoices, sp.splitCol) +
        '<button class="pick-btn" data-act="sp-pick" aria-label="用 Excel 里选中的列" title="先在表里选中一列再点">' + ICON.pick + '</button></div>' +
        '<label class="field">遇到空值时' + selectHtml('proc.sp.blankKeyRule', [['bucket', '集中放进「（未填写）」分表（推荐）'], ['skip', '跳过该行，不放进任何分表'], ['abort', '先停下报错，让我补齐再拆']], sp.blankKeyRule) + '</label>' +
        '<label class="field">小计/合计' + selectHtml('proc.sp.summaryRule', [['exclude', '自动排除合计行，不生成「合计」分表（推荐）'], ['keep', '全部保留，当成普通行拆分']], sp.summaryRule) + '</label></div>';
      h += '<p class="sentence">按「' + esc(S.ctx.sheet) + '」的 <b>' + esc(colLabel(S.ctx.sheet, sp.splitCol)) + '</b> 不同取值，为每个类别自动新建一张带表头的工作表' +
        (sp.sheetPrefix ? '（表名前缀「<b>' + esc(sp.sheetPrefix) + '</b>」）' : '') + '，原总表保持不动。</p>';
      h += '<details class="more"><summary>更多设置：给拆出来的分表名统一加前缀</summary><div class="inner">' +
        '<label class="field">分表前缀<input type="text" class="txt-in" data-bind="proc.sp.sheetPrefix" value="' + esc(sp.sheetPrefix || '') + '" maxlength="12" placeholder="例：分表_ / 2026_（可留空）"></label>' +
        '<span class="tiny muted">若工作簿里已有同名工作表，会自动加上 _2、_3 后缀保护，绝不覆盖已有工作表；单次最多拆出 50 张分表。</span></div></details>';
      var spv = S.proc.spPreview, spFresh = false;
      try { spFresh = spv && S.proc.spPreviewKey === JSON.stringify(splitSheetOptions()); } catch (e) { spFresh = false; }
      if (spv && spFresh) {
        h += '<div class="inner"><div class="chips">' +
          '<span class="chip ok">将新建 ' + num(spv.groupCount) + ' 张分表</span>' +
          '<span class="chip ok">共分发 ' + num(spv.netRows) + ' 行明细</span>' +
          (spv.summaryRows ? '<span class="chip warn">识别出小计/合计 ' + num(spv.summaryRows) + ' 行（' + (spv.summaryRule === 'exclude' ? '将排除' : '将保留') + '）</span>' : '') +
          (spv.blankRows ? '<span class="chip warn">拆分列为空 ' + num(spv.blankRows) + ' 行</span>' : '') +
          (spv.sourceErrors ? '<span class="chip warn">来源含错误值 ' + num(spv.sourceErrors) + ' 格（将跳过）</span>' : '') + '</div>';
        if (spv.summarySamples && spv.summarySamples.length) {
          h += '<div class="auto-note" style="background:var(--warn-soft);border-color:#E8C98F;color:var(--warn-ink)">' +
            '<b>请确认识别出的小计/合计行（共 ' + num(spv.summaryRows) + ' 行）：</b>' +
            (spv.summaryRule === 'exclude' ? '当前设为「自动排除」，不会误生成名为「合计」的分表。' : '当前设为「全部保留」，以下行将一并拆分。') + '</div>';
          h += '<table class="tbl"><tr><th>位置</th><th>行首标识</th><th>识别原因</th></tr>' + spv.summarySamples.slice(0, 5).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>';
          }).join('') + '</table>';
        }
        if (spv.groups && spv.groups.length) {
          h += '<table class="tbl"><tr><th>分类值</th><th>新建分表名</th><th>明细行数</th></tr>' + spv.groups.slice(0, 8).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + num(r[2]) + ' 行</td></tr>';
          }).join('') + '</table>';
        }
        if (spv.anomalies && spv.anomalies.length) {
          h += '<table class="tbl"><tr><th>位置</th><th>情况</th></tr>' + spv.anomalies.slice(0, 5).map(function (r) {
            return '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + (r[2] ? '：' + esc(r[2]) : '') + '</td></tr>';
          }).join('') + '</table>';
        }
        if (!spv.canExecute || spv.block) h += '<div class="banner warn"><span class="grow">现在还不能执行拆表：' + esc(spv.block || '请检查拆分列选择。') + '</span></div>';
        h += '</div>';
      }
      if (res) h += inlineResultHtml(res);
      var canRunSp = spv && spFresh && spv.canExecute && !spv.block;
      h += canRunSp
        ? '<button class="btn btn-primary" data-act="sp-run"' + (S.busy ? ' disabled' : '') + '>没问题，开始拆表（新建 ' + num(spv.groupCount) + ' 张表 · ' + num(spv.netRows) + ' 行）</button>'
        : '<button class="btn btn-dark" data-act="sp-preview"' + (S.busy ? ' disabled' : '') + '>' + (S.proc.spPreviewing ? '<span class="spin"></span>正在预览分表清单…' : '先预览一下') + '</button>';
      h += '<span class="foot-note">预览只看不改 · 开始拆表时会自动新建分表，原总表不改动</span>';
      return h;
    }
    if (key === 'recon') {
      var rc = S.proc.recon, others = S.proc.meta.filter(function (m) { return m[0] !== S.ctx.sheet; });
      if (!others.length) return '<p>这个文件里只有一张表，没法对账。</p>';
      h += '<p>把当前表选中的那一列，和另一张表的一列逐个对比，对不上的会标出颜色（可以用「去掉插件留下的标记」清掉），不改数据。</p>';
      h += selNowHtml('先在当前表选中要对的那一列，现在选中');
      h += '<div class="inner"><div class="inner-title">和哪一列对</div>' +
        '<label class="field">表' + selectHtml('proc.recon.sheet', others.map(function (m) { return [m[0], m[0]]; }), rc.sheet) + '</label>' +
        '<label class="field">列' + selectHtml('proc.recon.col', colChoices(colsOf(rc.sheet)), rc.col) + '</label>' +
        '<label class="field" style="grid-template-columns:auto 1fr"><input type="checkbox" data-bind="proc.recon.header"' + (rc.header ? ' checked' : '') + ' style="width:18px;height:18px;accent-color:var(--accent)">第一行是表头，不参与对比</label></div>';
      if (res) h += inlineResultHtml(res);
      h += '<button class="btn btn-primary" data-act="recon-run"' + (S.busy ? ' disabled' : '') + '>开始对账</button><span class="foot-note">Excel 会弹窗确认 · 修改前自动备份</span>';
      return h;
    }
    if (key === 'mask') {
      h += '<p>把选中区域里的敏感信息打上星号，表头和公式跳过。注意：当天的备份里仍然是原文。</p>' + selNowHtml('先选中要打码的列，现在选中');
      h += '<div class="opts three">' + [['Phone', '手机号'], ['ID', '身份证号'], ['Name', '姓名']].map(function (o) {
        return '<label class="opt"><input type="radio" name="mask" data-bind="proc.maskMode" value="' + o[0] + '"' + (S.proc.maskMode === o[0] ? ' checked' : '') + '>' + o[1] + '</label>';
      }).join('') + '</div><p class="tiny">例：138****5678 · 110101********1234 · 张*</p>';
      if (res) h += inlineResultHtml(res);
      h += '<button class="btn btn-primary" data-act="mask-run"' + (S.busy ? ' disabled' : '') + '>给选中的区域打码</button><span class="foot-note">Excel 会弹窗确认 · 修改前自动备份</span>';
      return h;
    }
    h += '<p>文件里所有可见的数据表：回到 A1、缩放 100%、冻结首行、打开网格线。不改任何数据。</p>';
    if (res) h += inlineResultHtml(res);
    h += '<button class="btn btn-primary" data-act="delivery-run"' + (S.busy ? ' disabled' : '') + '>收拾整个文件</button><span class="foot-note">Excel 会弹窗确认 · 修改前自动备份</span>';
    return h;
  }
  function hasCol(sheet, l) { return colsOf(sheet).some(function (c) { return c[0] === l; }); }
  function hasWriteCol(lk) { return writeColChoices(lk.target).some(function (c) { return c[0] === lk.write; }); }

  /* ---- 小票 ---- */
  function receiptSummary(entries) {
    var byCode = {}, skipped = {}, runs = 0, cells = 0;
    entries.forEach(function (e) {
      runs += 1;
      e.rows.forEach(function (x) {
        var k = labelOf(x.code);
        if (!byCode[k]) byCode[k] = { count: 0, unit: x.unit || unitOf(x.code), units: [], kept: [] };
        byCode[k].count += x.count;
        // Number conversion trace: union of dropped units, and columns kept because units were mixed.
        var gone = String(x.note || '').match(/本次去掉了单位：([^；]*)/);
        if (gone) gone[1].split('/').forEach(function (u) { if (u && byCode[k].units.indexOf(u) < 0) byCode[k].units.push(u); });
        var kept = String(x.note || '').match(/同列单位不统一，整列保留：[^；]*|保留 \d+ 处：月\/日\/年[^；]*/);
        if (kept && byCode[k].kept.indexOf(kept[0]) < 0) byCode[k].kept.push(kept[0]);
      });
      (e.skipped || []).forEach(function (s) {
        var k = labelOf(s.code);
        if (!skipped[k]) skipped[k] = s.why;
      });
    });
    var keys = Object.keys(byCode);
    keys.forEach(function (k) { });
    return { byCode: byCode, skipped: skipped, runs: runs, keys: keys };
  }
  function receiptSummaryText(entries, day) {
    var s = receiptSummary(entries);
    var lines = ['整理小票（汇总）· 《' + (S.ctx.wb || '当前文件') + '》 · ' + day, '————————————————'];
    lines.push('共执行 ' + s.runs + ' 次，合计改动：');
    s.keys.forEach(function (k) {
      lines.push('  ' + k + '：' + num(s.byCode[k].count) + ' ' + s.byCode[k].unit);
      if (s.byCode[k].units.length) lines.push('    本次去掉了单位：' + s.byCode[k].units.join('/'));
      s.byCode[k].kept.forEach(function (n) { lines.push('    ' + n); });
    });
    var sk = Object.keys(s.skipped);
    if (sk.length) {
      lines.push('没有处理的（原因）：');
      sk.forEach(function (k) { lines.push('  ' + k + '：' + s.skipped[k]); });
    }
    lines.push('————————————————');
    lines.push('每次修改前都已自动备份；不支持逐步撤销，文件需自己保存。需要明细请在小票页查看当天记录。');
    return lines.join('\n');
  }
  function receiptText(entries, day) {
    var lines = ['整理小票 · 《' + (S.ctx.wb || '当前文件') + '》 · ' + day, '————————————————'];
    entries.forEach(function (e) {
      lines.push(hhmm(e.t) + ' ' + e.title + ' · ' + e.sheet + (e.addr ? ' ' + e.addr : '') + (e.ok ? '' : '（没有全部完成）'));
      e.rows.forEach(function (x) {
        lines.push('  ' + labelOf(x.code) + '：' + num(x.count) + ' ' + (x.unit || unitOf(x.code)) + '（' + num(x.count) + ' ' + (x.unit || unitOf(x.code)) + '）');
        if (x.note) lines.push('    ' + x.note);
      });
      if (!e.ok) lines.push('  停在：' + (e.failedAt ? labelOf(e.failedAt) : '执行前') + '，' + e.error);
      (e.skipped || []).forEach(function (s) { lines.push('  跳过：' + labelOf(s.code) + '（' + s.why + '）'); });
    });
    lines.push('————————————————');
    lines.push('共 ' + entries.length + ' 次执行记录。破坏性操作前核验当日基线，后续复用；不支持逐步撤销。文件需自己保存。小票不是完整审计凭证。');
    return lines.join('\n');
  }
  function renderReceipt() {
    var el = $('view-receipt'), today = dayKey();
    var day = S.receipt.viewDay || today;
    var entries = day === today ? S.receipt.entries : Store.load(receiptWb, day);
    var html = '<div style="display:flex;flex-direction:column;gap:4px;padding:0 4px"><h1>' + (day === today ? '今天的小票' : esc(day) + ' 的小票') + '</h1>' +
      '<p class="lead">' + (day === today ? '今天在这个文件上做的每一次修改都记在这里。全部做完，复制一次发给别人。' : '以前的记录，只能查看和复制。 <button class="link-btn" data-act="day" data-v="">回到今天</button>') + '</p></div>';
    if (receiptWb !== S.ctx.wb && Store.days(S.ctx.wb).length) html += '<div class="banner info"><span>旧版按文件名保存的小票仍在本地；因无法核实同名文件归属，未自动合并。历史操作请以本地审计为准。</span></div>';
    if (entries.length) {
      html += '<button class="acc-head" data-act="toggle-receipt-detail" aria-expanded="' + (S.receipt.open ? 'true' : 'false') + '">' +
        '<span class="t"><b>整理小票明细</b></span><span class="pill">' + entries.length + ' 次</span>' + ICON.chev + '</button>';
      html += '<div class="acc-item' + (S.receipt.open ? ' open' : '') + '" style="' + (S.receipt.open ? '' : 'display:none') + '">';
    }
    html += '<div class="paper"><div class="ttl"><b>整理小票</b><span>' + esc(S.ctx.wb || '当前文件') + ' · ' + esc(day) + '</span></div><hr>';
    if (!entries.length) {
      html += '<div class="empty">还没有记录。用「帮我整理」或工具做完一次修改，这里就会自动记下。</div>';
    } else {
      html += entries.map(function (e) {
        var lines = '<div class="entry"><b>' + hhmm(e.t) + ' ' + esc(e.title) + '</b><span>' + esc(e.sheet) + (e.addr ? ' ' + esc(e.addr) : '') + '</span>';
        lines += e.rows.map(function (x) {
          return '<span class="li"><span>' + esc(labelOf(x.code)) + (x.note ? '<br><small>' + esc(x.note) + '</small>' : '') + '</span><span>' + num(x.count) + ' ' + (x.unit || unitOf(x.code)) + '</span></span>';
        }).join('');
        if (!e.ok) lines += '<span class="li bad"><span>停在：' + esc(e.failedAt ? labelOf(e.failedAt) : '执行前') + '</span><span>未完成</span></span>';
        return lines + '</div>';
      }).join('');
      html += '<hr><span>共 ' + entries.length + ' 次修改，每次修改前都已自动备份</span>';
    }
    html += '</div>';
    if (entries.length) html += '</div>';
    if (entries.length) {
      html += '<button class="btn btn-dark" style="height:52px;font-size:16px" data-act="copy-receipt">复制小票（汇总）</button>' +
        '<button class="link-btn" style="align-self:center" data-act="copy-receipt-detail">复制完整明细</button>';
      if (day === today) html += S.receipt.confirmClear
        ? '<div class="btn-row"><button class="btn btn-ghost" data-act="clear-receipt" style="color:var(--bad)">确定清空今天的小票</button><button class="btn btn-ghost" data-act="clear-cancel">算了</button></div>'
        : '<button class="link-btn" style="align-self:center" data-act="clear-receipt">清空今天的小票</button>';
    }
    var backups = [];
    entries.forEach(function (e) { if (e.backup && backups.indexOf(e.backup) < 0) backups.push(e.backup); });
    if (backups.length) {
      html += '<div class="choice-wrap"><div class="sec-title">备份文件<small>修改前自动保存的当天原文件</small></div><div class="list">' +
        backups.map(function (b) {
          return '<div class="list-row"><span class="grow path">' + esc(b) + '</span><button class="link-btn" data-act="open-path" data-text="' + esc(b) + '">打开文件夹</button><button class="link-btn" data-act="copy" data-text="' + esc(b) + '">复制路径</button></div>';
        }).join('') + '</div></div>';
    }
    var days = Store.days(receiptWb).filter(function (d) { return d !== today; });
    html += '<div class="choice-wrap"><div class="sec-title">其他</div><div class="list">';
    days.slice(0, 7).forEach(function (d) {
      html += '<button class="list-row" data-act="day" data-v="' + esc(d) + '"><span class="grow">' + esc(d) + ' 的小票</span>' + ICON.chev.replace('class="chev"', 'class="chev" style="transform:rotate(-90deg)"') + '</button>';
    });
    html += '<button class="list-row" data-act="help"><span class="grow">使用说明</span></button>' +
      '<div class="list-row"><span class="grow">关于电子羊表格</span><span class="muted tiny">' + esc(S.ctx.version || '—') + '</span></div></div></div>';
    el.innerHTML = html;
  }

  /* ---- overlay ---- */
  function renderOverlay() {
    var el = $('overlay'), f = S.flow;
    el.classList.toggle('hidden', !f);
    document.body.style.overflow = f ? 'hidden' : '';
    if (!f) { el.innerHTML = ''; return; }
    if (f.phase === 'prep') {
      el.innerHTML = '<div class="ov-icon"><span class="spin" style="width:28px;height:28px;border-width:3px"></span></div><h2>正在准备…</h2><p>读取表格，锁定这次要做的事。</p>';
      return;
    }
    if (f.phase === 'exec') {
      el.innerHTML = '<div class="ov-icon"><span class="spin" style="width:28px;height:28px;border-width:3px"></span></div><h2>正在修改…</h2><p>请先不要操作 Excel，做完会自动显示结果。</p>';
      return;
    }
    el.innerHTML = '<div class="ov-icon">' + ICON.shield + '</div>' +
      '<div style="display:flex;flex-direction:column;gap:8px"><h2>请在 Excel 弹出的窗口里确认</h2>' +
      '<p>点「是」才会修改表格，点「否」什么都不会改。<br>没看到窗口？它可能被挡在后面，点一下任务栏里的 Excel。</p></div>' +
      '<div class="ov-card"><span class="lbl">窗口里会写着</span>' +
      '<div style="display:flex;justify-content:space-between;gap:8px"><span class="muted">改哪里</span><b>' + esc(f.sheet) + ' ' + esc(f.addr) + '</b></div>' +
      '<span class="lbl">按顺序做</span><ol>' + f.locked.map(function (c) { return '<li>' + esc(labelOf(c)) + '</li>'; }).join('') + '</ol>' +
      '<span class="tiny muted">确定后先备份当天的原文件，再开始修改。</span></div>';
  }

  /* ------------------------------------------------------------------ *
   * 8. Events                                                           *
   * ------------------------------------------------------------------ */
  function copyText(text) {
    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); } catch (e) { /* ignore */ }
      document.body.removeChild(ta);
    }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).catch(fallback);
      else fallback();
    } catch (e) { fallback(); }
    toast('已复制，去微信或邮件里粘贴就行');
  }

  function onClick(ev) {
    var b = ev.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act'), v = b.getAttribute('data-v');
    switch (act) {
      case 'tab':
        S.tab = v; S.alert = null; S.help = false;
        if (v === 'receipt') S.receipt.unseen = 0; S.receipt.open = false;
        if (v === 'process' && S.proc.open && (S.proc.open === 'lookup' || S.proc.open === 'append' || S.proc.open === 'unpivot' || S.proc.open === 'splitsheet' || S.proc.open === 'recon')) ensureMeta();
        break;
      case 'go-tab': S.tab = v; S.alert = null; if (v === 'receipt') S.receipt.unseen = 0; S.receipt.open = false; break;
      case 'help': S.help = !S.help; break;
      case 'alert-close': S.alert = null; break;
      case 'ctx-refresh': refreshContext(); return;
      case 'check': checkTable(); return;
      case 'check-again': S.tab = 'auto'; checkTable((b.getAttribute('data-keys') || '').split(',').filter(Boolean)); return;
      case 'clean-scan': checkTable(); return;
      case 'auto-run':
        runPlan({ actions: autoActions(), origin: 'auto', key: 'auto', title: '帮我整理', range: 'table' });
        return;
      case 'clean-range':
        S.clean.range = v;
        if (v === 'selection') refreshContext();
        break;
      case 'acc': {
        var scope = b.getAttribute('data-scope');
        var o = scope === 'clean' ? S.clean : S.proc;
        o.open = o.open === v ? null : v;
        if (scope === 'proc' && o.open && (v === 'lookup' || v === 'append' || v === 'unpivot' || v === 'splitsheet' || v === 'recon')) ensureMeta();
        if (scope === 'proc' && o.open && (v === 'mask' || v === 'recon')) refreshContext();
        if (scope === 'clean' && o.open && S.clean.range === 'selection') refreshContext();
        render();
        var item = $((scope === 'clean' ? 'clean-' : 'proc-') + v);
        if (item && o.open) { try { item.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch (e) { /* ignore */ } }
        return;
      }
      case 'clean-run': {
        var t = CLEAN_TOOLS.filter(function (x) { return x.key === v; })[0];
        var range = cleanRange(v);
        runPlan({
          actions: cleanActions(v), origin: 'clean', key: v, title: t.title, range: range,
          guard: function () {
            if (v === 'Unit' && S.clean.unitMode === 'Split' && S.ctx.cols !== 1) return '拆到新列要先在 Excel 里只选中一列（现在选中的是 ' + (S.ctx.sel || '—') + '）。';
            return '';
          }
        });
        return;
      }
      case 'sp-preview': previewSplitSheet(); return;
      case 'sp-pick': pickSplitColumn(); return;
      case 'sp-run': {
        var spOpts;
        try { spOpts = splitSheetOptions(); } catch (e) { showAlert('warn', e.message); break; }
        var curSpSheet = S.ctx.sheet;
        runPlan({
          actions: [{ id: 'SplitSheet', options: spOpts }], origin: 'process', key: 'splitsheet', title: '按列拆表', range: 'table',
          guard: function () { return S.ctx.sheet !== curSpSheet ? '你刚刚切换了工作表，请切回「' + curSpSheet + '」再点开始。' : ''; }
        });
        return;
      }
      case 'up-preview': previewUnpivot(); return;
      case 'up-id-toggle': {
        var up0 = S.proc.up, idList = (up0.idCols || []).slice(), idPos = idList.indexOf(v);
        var uOrder = colsOf(S.ctx.sheet).map(function (c) { return c[0]; });
        if (idPos >= 0) {
          if (idList.length <= 1) { toast('至少保留 1 个左侧固定列'); render(); return; }
          idList.splice(idPos, 1);
        } else {
          if (idList.length >= 16) { showAlert('warn', '固定列最多选 16 列。'); render(); return; }
          idList.push(v);
          idList.sort(function (a, b) { return uOrder.indexOf(a) - uOrder.indexOf(b); });
          up0.valCols = (up0.valCols || []).filter(function (c) { return c !== v; });
        }
        up0.idCols = idList;
        if (!up0.valCols || !up0.valCols.length) {
          var rem = colsOf(S.ctx.sheet).filter(function (c) { return idList.indexOf(c[0]) < 0 && !isSummaryHeaderName(c[1]); }).map(function (c) { return c[0]; });
          up0.valCols = rem;
        }
        S.proc.upPreview = null; S.proc.upPreviewKey = '';
        break;
      }
      case 'up-val-toggle': {
        var up1 = S.proc.up, vList = (up1.valCols || []).slice(), vPos = vList.indexOf(v);
        var uOrder2 = colsOf(S.ctx.sheet).map(function (c) { return c[0]; });
        if (vPos >= 0) {
          if (vList.length <= 1) { toast('至少保留 1 个要展开的列'); render(); return; }
          vList.splice(vPos, 1);
        } else {
          if (vList.length >= 64) { showAlert('warn', '单次最多展开 64 列。'); render(); return; }
          vList.push(v);
          vList.sort(function (a, b) { return uOrder2.indexOf(a) - uOrder2.indexOf(b); });
        }
        up1.valCols = vList;
        S.proc.upPreview = null; S.proc.upPreviewKey = '';
        break;
      }
      case 'up-id-first1': {
        var cAll1 = colsOf(S.ctx.sheet);
        if (cAll1.length) {
          S.proc.up.idCols = [cAll1[0][0]];
          S.proc.up.valCols = cAll1.slice(1).filter(function (c) { return !isSummaryHeaderName(c[1]); }).map(function (c) { return c[0]; });
          S.proc.upPreview = null; S.proc.upPreviewKey = '';
        }
        break;
      }
      case 'up-id-first2': {
        var cAll2 = colsOf(S.ctx.sheet);
        if (cAll2.length >= 2) {
          S.proc.up.idCols = [cAll2[0][0], cAll2[1][0]];
          S.proc.up.valCols = cAll2.slice(2).filter(function (c) { return !isSummaryHeaderName(c[1]); }).map(function (c) { return c[0]; });
          S.proc.upPreview = null; S.proc.upPreviewKey = '';
        }
        break;
      }
      case 'up-val-rest': {
        var idSetNow = {};
        (S.proc.up.idCols || []).forEach(function (c) { idSetNow[c] = true; });
        var restCols = colsOf(S.ctx.sheet).filter(function (c) { return !idSetNow[c[0]] && !isSummaryHeaderName(c[1]); }).map(function (c) { return c[0]; });
        if (!restCols.length) restCols = colsOf(S.ctx.sheet).filter(function (c) { return !idSetNow[c[0]]; }).map(function (c) { return c[0]; });
        S.proc.up.valCols = restCols.slice(0, 64);
        S.proc.upPreview = null; S.proc.upPreviewKey = '';
        break;
      }
      case 'up-run': {
        var upOpts;
        try { upOpts = unpivotOptions(); } catch (e) { showAlert('warn', e.message); break; }
        var curS = S.ctx.sheet;
        runPlan({
          actions: [{ id: 'Unpivot', options: upOpts }], origin: 'process', key: 'unpivot', title: '宽表转长表', range: 'table',
          guard: function () { return S.ctx.sheet !== curS ? '你刚刚切换了工作表，请切回「' + curS + '」再点开始。' : ''; }
        });
        return;
      }
      case 'ap-preview': previewAppend(); return;
      case 'ap-sheet-toggle': {
        var ap0 = S.proc.ap, sList = (ap0.sheets || []).slice(), sPos = sList.indexOf(v);
        if (sPos >= 0) {
          if (sList.length <= 2) { toast('至少保留 2 张要合并的工作表'); render(); return; }
          sList.splice(sPos, 1);
        } else {
          if (sList.length >= 32) { showAlert('warn', '单次最多同时合并 32 张工作表。'); render(); return; }
          var mOrder = (S.proc.meta || []).map(function (m) { return m[0]; });
          sList.push(v);
          sList.sort(function (a, b) { return mOrder.indexOf(a) - mOrder.indexOf(b); });
        }
        ap0.sheets = sList;
        S.proc.apPreview = null; S.proc.apPreviewKey = '';
        break;
      }
      case 'ap-sheet-all': {
        var metaList = S.proc.meta || [];
        var nonE = metaList.filter(function (m) { return (Number(m[2]) || 0) >= 2 && Array.isArray(m[3]) && m[3].length > 0; }).map(function (m) { return m[0]; });
        S.proc.ap.sheets = (nonE.length >= 2 ? nonE : metaList.map(function (m) { return m[0]; })).slice(0, 32);
        S.proc.apPreview = null; S.proc.apPreviewKey = '';
        break;
      }
      case 'ap-sheet-first2': {
        var metaList2 = S.proc.meta || [];
        S.proc.ap.sheets = metaList2.slice(0, 2).map(function (m) { return m[0]; });
        S.proc.apPreview = null; S.proc.apPreviewKey = '';
        break;
      }
      case 'ap-run': {
        var apOpts;
        try { apOpts = appendOptions(); } catch (e) { showAlert('warn', e.message); break; }
        runPlan({
          actions: [{ id: 'Append', options: apOpts }], origin: 'process', key: 'append', title: '多表纵向追加', range: 'table'
        });
        return;
      }
      case 'lk-preview': previewLookup(); return;
      case 'lk-pick': pickColumn(v); return;
      case 'lk-retarget': initLookupDefaults(false); break;
      case 'lk-ret-toggle': {
        var lk0 = S.proc.lk, list0 = (lk0.retCols || []).slice(), pos = list0.indexOf(v);
        if (pos >= 0) {
          if (list0.length <= 1) { toast('至少保留 1 个要取过来的列'); render(); return; }
          list0.splice(pos, 1);
        } else {
          if (list0.length >= 16) { showAlert('warn', '单次最多同时补全 16 列。'); render(); return; }
          var order = colsOf(lk0.src).map(function (c) { return c[0]; });
          list0.push(v);
          list0.sort(function (a, b) { return order.indexOf(a) - order.indexOf(b); });
        }
        lk0.retCols = list0;
        syncLookupColumns(lk0);
        S.proc.preview = null; S.proc.previewKey = '';
        break;
      }
      case 'lk-ret-all': {
        var lkAll = S.proc.lk;
        lkAll.retCols = colsOf(lkAll.src).filter(function (c) { return c[0] !== lkAll.srcKey; }).slice(0, 16).map(function (c) { return c[0]; });
        syncLookupColumns(lkAll);
        S.proc.preview = null; S.proc.previewKey = '';
        break;
      }
      case 'lk-ret-first': {
        var lkFirst = S.proc.lk, avail0 = colsOf(lkFirst.src).filter(function (c) { return c[0] !== lkFirst.srcKey; });
        lkFirst.retCols = avail0.length ? [avail0[0][0]] : [];
        syncLookupColumns(lkFirst);
        S.proc.preview = null; S.proc.previewKey = '';
        break;
      }
      case 'lk-run': {
        var opts;
        try { opts = lookupOptions(); } catch (e) { showAlert('warn', e.message); break; }
        var target = S.proc.lk.target;
        runPlan({
          actions: [{ id: 'Lookup', options: opts }], origin: 'process', key: 'lookup', title: '跨表补全', range: 'table',
          guard: function () { return S.ctx.sheet !== target ? '你现在在「' + S.ctx.sheet + '」，补全会填到当前表。请切回「' + target + '」再点开始。' : ''; }
        });
        return;
      }
      case 'recon-run': {
        var rc = S.proc.recon;
        runPlan({
          actions: [{ id: 'Reconcile', options: { otherSheet: rc.sheet, otherCol: rc.col, hasHeader: rc.header ? 'true' : 'false' } }],
          origin: 'process', key: 'recon', title: '两表对账', range: 'selection',
          guard: function () { return S.ctx.cols !== 1 ? '请先在当前表只选中要对的那一列（现在选中的是 ' + (S.ctx.sel || '—') + '）。' : ''; }
        });
        return;
      }
      case 'mask-run':
        runPlan({ actions: [{ id: 'Transform', options: { mode: S.proc.maskMode } }], origin: 'process', key: 'mask', title: MASK[S.proc.maskMode], range: 'selection' });
        return;
      case 'delivery-run':
        runPlan({ actions: [{ id: 'Delivery' }], origin: 'process', key: 'delivery', title: '交表前收拾', range: 'table' });
        return;
      case 'copy': copyText(b.getAttribute('data-text') || ''); return;
      case 'jump': {
        var a = b.getAttribute('data-text') || '';
        call('select-range', { path: a }).then(function () { }, function (e) { toast(friendly(e && e.message || '跳转失败')); });
        return;
      }
      case 'open-path': {
        var pth = b.getAttribute('data-text') || '';
        if (!pth) { toast('没有可打开的路径'); return; }
        call('open-path', { path: pth }).then(function () { toast('已在文件夹里定位到备份'); }, function (e) { toast(friendly(e && e.message || '打开失败')); });
        return;
      }
      case 'copy-receipt': {
        var day = S.receipt.viewDay || dayKey();
        var list = day === dayKey() ? S.receipt.entries : Store.load(receiptWb, day);
        copyText(receiptSummaryText(list, day));
        return;
      }
      case 'toggle-receipt-detail': S.receipt.open = !S.receipt.open; render(); return;
      case 'copy-receipt-detail': {
        var d2 = S.receipt.viewDay || dayKey();
        var l2 = d2 === dayKey() ? S.receipt.entries : Store.load(receiptWb, d2);
        copyText(receiptText(l2, d2));
        return;
      }
      case 'clear-receipt':
        if (!S.receipt.confirmClear) { S.receipt.confirmClear = true; break; }
        S.receipt.entries = []; S.receipt.confirmClear = false;
        Store.save(receiptWb, dayKey(), []);
        toast('今天的小票已清空');
        break;
      case 'clear-cancel': S.receipt.confirmClear = false; break;
      case 'day': S.receipt.viewDay = v || null; break;
      default: return;
    }
    render();
  }

  function onChange(ev) {
    var el = ev.target, path = el.getAttribute && el.getAttribute('data-bind');
    if (!path) return;
    var val = el.type === 'checkbox' ? el.checked : el.value;
    if (path.indexOf('proc.lk.writeCol_') === 0) {
      var wIdx = Number(path.slice('proc.lk.writeCol_'.length)) || 0;
      S.proc.lk.writeCols[wIdx] = val;
      S.proc.lk.write = S.proc.lk.writeCols.join(',');
      S.proc.preview = null; S.proc.previewKey = '';
      render();
      return;
    }
    setPath(path, val);
    if (path.indexOf('proc.lk.') === 0) {
      if (path === 'proc.lk.src') {
        var sc = colsOf(val);
        S.proc.lk.srcKey = sc[0] ? sc[0][0] : '';
        S.proc.lk.retCols = sc[1] ? [sc[1][0]] : (sc[0] ? [sc[0][0]] : []);
      }
      syncLookupColumns(S.proc.lk);
      S.proc.preview = null; S.proc.previewKey = '';
    }
    if (path.indexOf('proc.ap.') === 0) { S.proc.apPreview = null; S.proc.apPreviewKey = ''; }
    if (path.indexOf('proc.up.') === 0) { S.proc.upPreview = null; S.proc.upPreviewKey = ''; }
    if (path.indexOf('proc.sp.') === 0) { S.proc.spPreview = null; S.proc.spPreviewKey = ''; }
    if (path === 'proc.recon.sheet') { var c = colsOf(val); S.proc.recon.col = c[0] ? c[0][0] : 'A'; }
    render();
  }

  /* ------------------------------------------------------------------ *
   * 9. Dev simulator: bridge-sim.js (loaded only with ?dev_sim=1, never   *
   *    packaged). It receives the internals it needs through these hooks. *
   * ------------------------------------------------------------------ */
  var Sim = null;
  if (isSim) {
    if (typeof window.DHSimFactory === 'function') {
      Sim = window.DHSimFactory({ ACTIONS: ACTIONS, baseCode: baseCode, onHostMessage: onHostMessage, S: S, dayKey: dayKey,
        Store: Store, num: num, $: $, render: render,
        resetReceipt: function () { receiptWb = null; loadReceipt(); } });
    } else {
      isSim = false;
    }
  }

  // Presentation only: the shared VBA entry checks and records every plan.
  var licenseRefreshing = false;
  function paintLicense(state) {
    var text = '开源免费版 (MIT)';
    if (Array.isArray(state) && state[0] === 'LICENSED' && state[6] && state[6] !== 'OSS') {
      text = '已解锁（' + String(state[6]) + '）';
    }
    document.getElementById('licenseStatus').textContent = text;
    document.getElementById('licenseDialogStatus').textContent = text;
  }
  function refreshLicense() {
    if (!isNative || licenseRefreshing || S.busy) return Promise.resolve();
    licenseRefreshing = true;
    return call('license-status').then(function (reply) { paintLicense(reply.result); }, function () { paintLicense(null); })
      .then(function () { licenseRefreshing = false; });
  }
  document.getElementById('licenseOpen').addEventListener('click', function () {
    if (S.busy) return;
    document.getElementById('licenseMessage').textContent = '';
    document.getElementById('licenseDialog').showModal();
    refreshLicense();
  });
  document.getElementById('licenseClose').addEventListener('click', function () { document.getElementById('licenseDialog').close(); });
  document.getElementById('licenseDialog').addEventListener('close', function () { document.getElementById('licenseInput').value = ''; });
  document.getElementById('licenseSubmit').addEventListener('click', function () {
    var input = document.getElementById('licenseInput'), button = this;
    var card = input.value; input.value = ''; button.disabled = true;
    call('license-activate', { license: card }).then(function (reply) {
      document.getElementById('licenseMessage').textContent = String(reply.result[2] || '');
      return refreshLicense();
    }, function (error) {
      document.getElementById('licenseMessage').textContent = friendly(error);
    }).then(function () { button.disabled = false; });
  });
  window.addEventListener('focus', function () { refreshLicense(); });
  document.addEventListener('visibilitychange', function () { if (!document.hidden) refreshLicense(); });

  /* ------------------------------------------------------------------ *
   * 10. Start                                                           *
   * ------------------------------------------------------------------ */
  document.addEventListener('click', onClick);
  document.addEventListener('change', onChange);
  if (isNative) window.chrome.webview.addEventListener('message', function (e) { onHostMessage(e.data); });
  if (isSim) {
    Sim.reset();
    document.body.classList.add('sim');
    $('simbar').classList.remove('hidden');
    $('simbar').addEventListener('click', Sim.onBar);
    Sim.bar();
    window.__dh = { S: S, Sim: Sim, render: render };
  }
  render();
  if (isNative || isSim) { refreshContext(); if (isNative) setTimeout(refreshLicense, 500); }
  else { showAlert('warn', '这个页面要在 Excel 的电子羊表格窗格里使用。'); render(); }
})();
