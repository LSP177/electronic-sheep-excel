
window.DHSimFactory = function(api) {
  return {
    reset: function() {},
    onBar: function() {},
    bar: function() {},
    dispatch: function(req) {
      setTimeout(function() {
        if (req.command === 'context' || req.command === 'candidate') {
          api.onHostMessage({ v: 1, id: req.id, command: req.command, ok: true, result: ['TestBook.xlsx', 'Orders', 'A1:C5', 5, 3, '1.0.0-patch2'] });
        } else if (req.command === 'scan') {
          api.onHostMessage({ v: 1, id: req.id, command: 'scan', ok: true, result: ['TestBook.xlsx', 'Orders', 'A1:C5', 0, 0, 0, 0, 0, 0, 0, 0, 0] });
        } else if (req.command === 'lookup_meta') {
          api.onHostMessage({ v: 1, id: req.id, command: 'lookup_meta', ok: true, result: [
            ['Orders', 'A1:C5', 5, [['A', '商品编号'], ['B', '数量'], ['C', '备注']]],
            ['Products', 'A1:E6', 6, [['A', '商品编号'], ['B', '商品名称'], ['C', '单价'], ['D', '备注'], ['E', '上架日期']]]
          ]});
        } else if (req.command === 'lookup_preview') {
          var o = req.options || {};
          api.onHostMessage({ v: 1, id: req.id, command: 'lookup_preview', ok: true, result: [
            'TestBook.xlsx', 'Orders', o.sourceSheet, o.sourceKeyCol, o.returnCol,
            4, 7, 1, 0, 0, 0, 1, 0, o.targetKeyCol, 'D(商品名称),E(单价),F(Products_备注),G(上架日期)',
            o.dupRule + '/' + o.writeMode, true, '',
            [['D..G:2', '0012', '商品名称=无线鼠标 | 单价=89.5 | 备注=首发 | 上架日期=2026-03-01', 'Products!行2']],
            [['E3', '来源含错误值（将跳过）', '来源表「Products」C3（单价）为错误值 #N/A，保护跳过不写入错误'], ['A4', '类型差异（文本/数字）', '键值「12」在两表中存储类型不同']],
            1
          ]});
        } else if (req.command === 'prepare') {
          var act = req.actions[0];
          var o = act.options;
          var canon = 'Lookup:' + o.sourceSheet + ':' + o.sourceKeyCol + ':' + o.returnCol + ':' + o.targetKeyCol + ':' + o.targetWriteCol + ':' + o.dupRule + ':' + o.writeMode;
          api.onHostMessage({ v: 1, id: req.id, command: 'prepare', ok: true, result: ['PREPARED_BOUND_WRITE_PLAN', 'SESS-P2', 'C:\\TestBook.xlsx', 'Orders', 'Orders', 'A1:C5', canon] });
        } else if (req.command === 'confirm') {
          api.onHostMessage({ v: 1, id: req.id, command: 'confirm', ok: true, result: 'CONFIRMED_BOUND_WRITE_PLAN' });
        } else if (req.command === 'execute') {
          api.onHostMessage({ v: 1, id: req.id, command: 'execute', ok: true, result: [
            'COMPLETED', 'done', 'C:\\Backup\\TestBook_当日基线.xlsx',
            [['Lookup:Products:A:B,C,D,E:A:AUTO:abort:static', 7, '跨表补全 7 格（4 列，右侧新增列）', 'C:\\Audit\\lookup.tsv', [2, [['E3', '来源错误值(#N/A)保护跳过'], ['A4', '类型差异跳过']]]]]
          ]});
        }
      }, 10);
    }
  };
};
