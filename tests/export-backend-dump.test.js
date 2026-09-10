// ตัวดึงชีตดิบฝั่ง Apps Script (dumpSheetForExport) — หัวใจของชุดข้อมูลที่ไม่มี getter เดิม
// รันจริงด้วยการ stub Utilities/SpreadsheetApp ของ Apps Script ขึ้นมา
// จุดที่กันไว้:
//  1) Date object จากชีตต้องถูกอ่านที่โซนเวลาไทย ไม่ใช่ UTC (ไม่งั้นวันเพี้ยนไป 1 วัน)
//  2) แถวว่างท้ายชีตต้องไม่ติดมาในไฟล์
//  3) กรองช่วงวันที่/ไลน์ได้จริง และคอลัมน์ที่สั่งตัดต้องหายไปจาก payload
const { backend, grabFn, assert } = require('./_export-extract');

// ── stub Apps Script เท่าที่ dumpSheetForExport ใช้ ────────────────────
function makeEnv(sheets) {
  const env = {};
  env.Utilities = {
    formatDate: function(date, tz, fmt) {
      assert.strictEqual(tz, 'Asia/Bangkok', 'ต้องจัดรูปแบบวันที่ที่โซนเวลาไทยเสมอ');
      // จำลอง Asia/Bangkok = UTC+7
      const t = new Date(date.getTime() + 7 * 3600000);
      const p = function(n) { return ('0' + n).slice(-2); };
      const y = t.getUTCFullYear(), mo = p(t.getUTCMonth() + 1), d = p(t.getUTCDate());
      const hh = p(t.getUTCHours()), mi = p(t.getUTCMinutes()), ss = p(t.getUTCSeconds());
      if (fmt === 'yyyy-MM-dd') return y + '-' + mo + '-' + d;
      if (fmt === 'yyyy-MM') return y + '-' + mo;
      return y + '-' + mo + '-' + d + ' ' + hh + ':' + mi + ':' + ss;
    }
  };
  env.SpreadsheetApp = {
    getActiveSpreadsheet: function() {
      return {
        getSheetByName: function(name) {
          if (!sheets[name]) return null;
          return { getDataRange: function() { return { getValues: function() { return sheets[name]; } }; } };
        }
      };
    }
  };
  return env;
}

function loadDump(sheets) {
  const env = makeEnv(sheets);
  const src = [
    grabFn('normalizeHeaderName', backend, 0),
    grabFn('exportDateKey', backend, 0),
    grabFn('exportInDateRange', backend, 0),
    grabFn('normalizeExportDateInput', backend, 0),
    grabFn('dumpSheetForExport', backend, 0),
    'return { dumpSheetForExport: dumpSheetForExport, exportDateKey: exportDateKey, exportInDateRange: exportInDateRange };'
  ].join('\n');
  return new Function('Utilities', 'SpreadsheetApp', src)(env.Utilities, env.SpreadsheetApp);
}

// ── exportDateKey อ่านได้ทุกรูปแบบที่ชีตเก็บจริง ───────────────────────
const m = loadDump({});
assert.strictEqual(m.exportDateKey('2026-09-10 14:26:47'), '2026-09-10', 'รูปแบบที่ processTransaction เขียนลงชีต');
assert.strictEqual(m.exportDateKey('2026-09-10'), '2026-09-10');
assert.strictEqual(m.exportDateKey('10/9/2026'), '2026-09-10', 'd/M/yyyy จากไฟล์นำเข้าเก่า');
assert.strictEqual(m.exportDateKey(''), '', 'ค่าว่างต้องได้ค่าว่าง');
assert.strictEqual(m.exportDateKey(null), '');
assert.strictEqual(m.exportDateKey(new Date('2026-08-31T17:00:00.000Z')), '2026-09-01',
  '31 ส.ค. 17:00Z คือ 1 ก.ย. เวลาไทย — ต้องนับเป็นเดือนกันยายน (เคสเดียวกับที่ Dashboard เคยพลาด)');
assert.strictEqual(m.exportDateKey(new Date('invalid')), '', 'Date ที่ใช้ไม่ได้ต้องได้ค่าว่าง ไม่ใช่ NaN');

// ── ช่วงวันที่ ───────────────────────────────────────────────────────
assert.strictEqual(m.exportInDateRange('2026-09-05', '2026-09-01', '2026-09-10'), true);
assert.strictEqual(m.exportInDateRange('2026-09-01', '2026-09-01', '2026-09-10'), true, 'วันเริ่มต้องรวม');
assert.strictEqual(m.exportInDateRange('2026-09-10 23:59:59', '2026-09-01', '2026-09-10'), true, 'วันสิ้นสุดต้องรวม');
assert.strictEqual(m.exportInDateRange('2026-08-31', '2026-09-01', '2026-09-10'), false);
assert.strictEqual(m.exportInDateRange('', '2026-09-01', '2026-09-10'), true, 'แถวไม่มีวันที่ต้องไม่ถูกตัดทิ้ง');
assert.strictEqual(m.exportInDateRange('1999-01-01', '', ''), true, 'ไม่ตั้งช่วง = เอาทุกแถว');

// ── ดึงชีตจริง ───────────────────────────────────────────────────────
const sheets = {
  PRHeaders: [
    ['pr_id', 'status', 'created_by', 'created_at', 'line', 'total_amount'],
    ['PR-001', 'APPROVED', 'somchai', '2026-09-02 09:00:00', 'H9', 12000],
    ['PR-002', 'PENDING', 'somsak', '2026-09-09 10:00:00', 'Lug&Screw', 3400],
    ['PR-003', 'REJECTED', 'somchai', '2026-08-20 08:00:00', 'H9', 900],
    ['', '', '', '', '', ''],
    [null, null, null, null, null, null]
  ],
  MiscExpenseLog: [
    ['Date Time', 'User', 'Expense ID', 'Action Type'],
    [new Date('2026-09-04T03:00:00.000Z'), 'admin', 'EX-1', 'edit'],
    ['2026-09-06 11:00:00', 'admin', 'EX-2', 'delete']
  ]
};
const d = loadDump(sheets);

// ไม่กรอง = ได้ทุกแถวที่มีข้อมูล และแถวว่างต้องหลุดออกไป
const all = d.dumpSheetForExport('PRHeaders', {});
assert.deepStrictEqual(all.headers, ['pr_id', 'status', 'created_by', 'created_at', 'line', 'total_amount']);
assert.strictEqual(all.rows.length, 3, 'แถวว่างท้ายชีตต้องไม่ติดมา (ชีตจริงมักมีแถวว่างค้างอยู่)');

// กรองช่วงวันที่
const sept = d.dumpSheetForExport('PRHeaders', { dateColumn: 'created_at', from: '2026-09-01', to: '2026-09-10' });
assert.strictEqual(sept.rows.length, 2, 'ใบเดือนสิงหาคมต้องไม่ติดมา');
assert.deepStrictEqual(sept.rows.map(function(r) { return r[0]; }), ['PR-001', 'PR-002']);

// กรองไลน์
const h9 = d.dumpSheetForExport('PRHeaders', { lineColumn: 'line', line: 'H9' });
assert.deepStrictEqual(h9.rows.map(function(r) { return r[0]; }), ['PR-001', 'PR-003']);
// ชื่อไลน์ต้องเทียบแบบไม่แคร์ตัวพิมพ์ (ชีตจริงพิมพ์ไม่เหมือนกันทุกแถว)
assert.strictEqual(d.dumpSheetForExport('PRHeaders', { lineColumn: 'line', line: 'h9' }).rows.length, 2);

// ตัดคอลัมน์ที่ห้ามหลุด — ต้องหายจาก headers และจากทุกแถว
const noMoney = d.dumpSheetForExport('PRHeaders', { dropColumns: ['total_amount'] });
assert.strictEqual(noMoney.headers.indexOf('total_amount'), -1, 'คอลัมน์ที่สั่งตัดต้องไม่อยู่ในหัวตาราง');
assert.strictEqual(noMoney.rows[0].length, 5, 'คอลัมน์ที่สั่งตัดต้องหายจากแถวด้วย ไม่ใช่แค่หัว');
assert.strictEqual(noMoney.rows[0].indexOf(12000), -1, 'ค่าที่ถูกตัดต้องไม่หลุดออกไปกับ payload');

// Date object ในชีตต้องถูกแปลงเป็นข้อความเวลาไทย ไม่ใช่ UTC ISO
const mlog = d.dumpSheetForExport('MiscExpenseLog', {});
assert.strictEqual(typeof mlog.rows[0][0], 'string', 'Date object ต้องถูกแปลงเป็นข้อความก่อนส่งกลับ');
assert.strictEqual(mlog.rows[0][0], '2026-09-04 10:00:00', '03:00Z = 10:00 เวลาไทย');

// เรียงใหม่สุดขึ้นก่อน (ชีต audit ต้องอ่านของล่าสุดได้ทันที)
const newest = d.dumpSheetForExport('MiscExpenseLog', { newestFirst: true });
assert.strictEqual(newest.rows[0][2], 'EX-2', 'newestFirst ต้องกลับลำดับให้');

// จำกัดจำนวนแถว
assert.strictEqual(d.dumpSheetForExport('PRHeaders', { limit: 2 }).rows.length, 2);

// ชีตที่ยังไม่มีในสเปรดชีต ต้องตอบว่า missing ไม่ใช่โยน error ทั้งงาน
const gone = d.dumpSheetForExport('ชีตที่ไม่มีอยู่', {});
assert.strictEqual(gone.missing, true, 'ชีตที่ไม่มีต้องตอบ missing (ระบบใหม่ยังไม่มีชีต audit บางตัว)');
assert.deepStrictEqual(gone.rows, []);

// ── exportPrBundle ต้องจับคู่ 3 ชีตด้วย pr_id ที่เหลือจากการกรองหัวใบ ──
const prSrc = grabFn('exportPrBundle', backend, 0);
assert(/filterByPrId/.test(prSrc), 'PRLines/PRAudit ต้องถูกกรองตาม pr_id ของหัวใบที่เหลือ');
// บรรทัดในใบ PR ไม่มีคอลัมน์วันที่ของตัวเอง ถ้าไปกรองด้วยวันจะได้ใบที่หัวมีแต่ไม่มีรายการ
assert(/dumpSheetForExport\(SPARE_APP_CONFIG\.prLinesSheetName, \{\}\)/.test(prSrc),
  'PRLines ต้องถูกดึงมาทั้งชีตแล้วค่อยกรองด้วย pr_id ไม่ใช่กรองด้วยวันที่');
assert(/dumpSheetForExport\(SPARE_APP_CONFIG\.prAuditSheetName, \{\}\)/.test(prSrc),
  'PRAudit ต้องถูกกรองด้วย pr_id เช่นกัน');

// ── exportManifest ต้องนับแบบเบา และแยก "ยังไม่มีชีต" ออกจาก "0 แถว" ──
const manifestSrc = grabFn('exportManifest', backend, 0);
assert(/getLastRow\(\)/.test(manifestSrc), 'ต้องนับด้วย getLastRow ไม่ใช่อ่านค่าทั้งชีต');
assert(!/getDataRange\(\)\.getValues\(\)/.test(manifestSrc), 'ห้ามอ่านทุกค่าเพียงเพื่อจะนับแถว');
assert(/counts\[name\] = -1/.test(manifestSrc), 'ชีตที่ยังไม่มีต้องเป็น -1 เพื่อแยกจากชีตที่มีแต่ว่างเปล่า');

console.log('export-backend-dump: OK');
