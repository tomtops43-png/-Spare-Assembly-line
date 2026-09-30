// รันฟังก์ชันสถานะเครื่องจักรของ Backend.gs ของจริง บนชีตจำลองในหน่วยความจำ
// stub เฉพาะบริการของ Apps Script (SpreadsheetApp/Utilities/LockService) + ด่านสิทธิ์ผู้ใช้
// นาฬิกาเลื่อนได้ (clock.now) เพื่อทดสอบการนับชั่วโมงเครื่องหยุด
//
// ไฟล์นี้ชื่อไม่ลงท้าย .test.js จึงไม่ถูก run-all.js หยิบไปรันเป็นเทสต์เอง
const { backend, grabFn, grabVar, grabScalar } = require('./_export-extract');

function makeSheet(rows) {
  const data = rows || [];
  return {
    rows: data,
    getLastRow: function() { return data.length; },
    getLastColumn: function() { return data.reduce(function(m, r) { return Math.max(m, r.length); }, 0); },
    appendRow: function(row) { data.push(row.slice()); },
    getDataRange: function() {
      return { getValues: function() { return data.map(function(r) { return r.slice(); }); } };
    },
    getRange: function(row, col, numRows, numCols) {
      return {
        setValues: function(values) {
          for (let r = 0; r < (numRows || 1); r += 1) {
            data[row - 1 + r] = data[row - 1 + r] || [];
            for (let c = 0; c < (numCols || 1); c += 1) data[row - 1 + r][col - 1 + c] = values[r][c];
          }
        },
        setValue: function(v) { data[row - 1] = data[row - 1] || []; data[row - 1][col - 1] = v; }
      };
    }
  };
}

// users: { token: { username, role, line, viewOnly } }
function makeMachineStatusBackend(opts) {
  opts = opts || {};
  const clock = { now: opts.now || Date.parse('2026-10-01T01:00:00Z') };
  const sheets = {};
  Object.keys(opts.sheets || {}).forEach(function(name) { sheets[name] = makeSheet(opts.sheets[name]); });
  const users = opts.users || {};
  let uuid = 0;

  class FakeDate extends Date {
    constructor(...args) { if (args.length) super(...args); else super(clock.now); }
    static now() { return clock.now; }
    // ค่าจากชีตเป็น Date ของจริง (คนละคลาสกับ FakeDate) — ใน Apps Script เป็นคลาสเดียวกัน
    static [Symbol.hasInstance](v) { return Object.prototype.toString.call(v) === '[object Date]'; }
  }

  const env = {
    SPARE_APP_CONFIG: { machinesSheetName: 'Machines', machineStatusSheetName: 'MachineStatus', machineStatusLogSheetName: 'MachineStatusLog' },
    SpreadsheetApp: { getActiveSpreadsheet: function() { return {}; } },
    getOrCreateSheet: function(ss, name) { sheets[name] = sheets[name] || makeSheet([]); return sheets[name]; },
    LockService: { getScriptLock: function() { return { waitLock: function() {}, releaseLock: function() {} }; } },
    Utilities: {
      getUuid: function() { uuid += 1; return 'uuid-' + String(uuid).padStart(8, '0'); },
      formatDate: function(date, tz, fmt) {
        const t = new Date(date.getTime() + 7 * 3600000);
        const p = function(n) { return ('0' + n).slice(-2); };
        const y = t.getUTCFullYear(), mo = p(t.getUTCMonth() + 1), d = p(t.getUTCDate());
        const hh = p(t.getUTCHours()), mi = p(t.getUTCMinutes()), ss = p(t.getUTCSeconds());
        if (fmt === 'yyyyMMdd-HHmmss') return '' + y + mo + d + '-' + hh + mi + ss;
        return y + '-' + mo + '-' + d + ' ' + hh + ':' + mi + ':' + ss;
      }
    },
    session: function(payload) {
      const u = users[payload && payload.authToken];
      if (!u) throw new Error('กรุณาเข้าสู่ระบบ');
      return u;
    }
  };

  const src = [
    grabScalar('MACHINE_STATUS_MAX_PARTS', backend),
    grabVar('MACHINE_HEADERS', backend),
    grabVar('MACHINE_STATUS_HEADERS', backend),
    grabVar('MACHINE_STATUS_LOG_HEADERS', backend),
    grabVar('MACHINE_STATUS_KEYS', backend),
    grabVar('MACHINE_REASON_KEYS', backend),
    grabVar('MACHINE_IMPACT_KEYS', backend),
    'function requirePermission(payload) { return session(payload); }',
    'function requireWarehouseWriter(payload) { var u = session(payload); if (u.viewOnly) throw new Error("บัญชีนี้เป็นสิทธิ์ดูอย่างเดียว จึงบันทึกข้อมูลส่วนนี้ไม่ได้"); return u; }',
    grabFn('normalizeRole', backend, 0),
    grabFn('getMachinesSheet', backend, 0),
    grabFn('rowToMachine', backend, 0),
    grabFn('getMachines', backend, 0),
    grabFn('getMachineStatusSheet', backend, 0),
    grabFn('getMachineStatusLogSheet', backend, 0),
    grabFn('ensureMachineStatusHeaders', backend, 0),
    grabFn('machineStatusNow', backend, 0),
    grabFn('machineStatusTimeText', backend, 0),
    grabFn('machineStatusParseTime', backend, 0),
    grabFn('machineStatusParseParts', backend, 0),
    grabFn('sanitizeMachineStatusParts', backend, 0),
    grabFn('machineStatusParseImpact', backend, 0),
    grabFn('sanitizeMachineImpact', backend, 0),
    grabFn('machineStatusRowToObject', backend, 0),
    grabFn('machineStatusLogRowToObject', backend, 0),
    grabFn('machineStatusUserCanEditLine', backend, 0),
    grabFn('getMachineStatusBoard', backend, 0),
    grabFn('updateMachineStatus', backend, 0),
    grabFn('updateMachineStatusUnlocked', backend, 0),
    grabFn('getMachineStatusHistory', backend, 0),
    'return { getMachineStatusBoard: getMachineStatusBoard, updateMachineStatus: updateMachineStatus, getMachineStatusHistory: getMachineStatusHistory };'
  ].join('\n');
  const api = new Function('SPARE_APP_CONFIG', 'SpreadsheetApp', 'getOrCreateSheet', 'LockService', 'Utilities', 'session', 'Date', src)(
    env.SPARE_APP_CONFIG, env.SpreadsheetApp, env.getOrCreateSheet, env.LockService, env.Utilities, env.session, FakeDate
  );
  api.clock = clock;
  api.sheets = sheets;
  api.advanceMinutes = function(min) { clock.now += min * 60000; };
  return api;
}

module.exports = { makeMachineStatusBackend };
