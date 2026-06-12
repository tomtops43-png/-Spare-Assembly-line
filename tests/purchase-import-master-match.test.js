const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const start = html.indexOf('    function normalizePurchaseImportMatchText');
const end = html.indexOf('    function loadPurchaseImportMasterRows', start);
assert(start >= 0 && end > start);
const context = { console, Number, String, Array, Object, Math };
context.refreshPurchaseImportRow = function(row) {
  row.qty_ordered = Number(row.qty_ordered || 0);
  row.review_required = !String(row.part_name || '').trim() || !String(row.model || '').trim() || !(row.qty_ordered > 0);
};
vm.createContext(context);
vm.runInContext(html.slice(start, end), context);
const rows = [
  { line:'H9', part_name:'SOLENOID VALVE', brand:'', model:'', qty_ordered:2, unit_price:1544, remark:'Imported from PR PDF', include:false, review_required:true },
  { line:'H9', part_name:'AIR CYLINDER', brand:'', model:'', qty_ordered:3, unit_price:4085, remark:'Imported from PR PDF', include:false, review_required:true },
  { line:'H9', part_name:'PROXIMITY', brand:'', model:'', qty_ordered:10, unit_price:490, remark:'Imported from PR PDF', include:false, review_required:true }
];
const master = [
  { line:'H9', name:'Solenoid Valve', brand:'SMC', model:'SY5120-5LZ', unit_price:1544, no:'H9-2' },
  { line:'H9', name:'Solenoid Valve', brand:'SMC', model:'SY3340-5LZ', unit_price:500, no:'H9-1' },
  { line:'H9', name:'Air Cylinder', brand:'SMC', model:'CQ2B40-50D', unit_price:4085, no:'H9-3' },
  { line:'H9', name:'Air Cylinder', brand:'SMC', model:'CDQ2B32-25D', unit_price:2418, no:'H9-4' },
  { line:'H9', name:'Proximity', brand:'OMRON', model:'E2E-X5ME1', unit_price:490, no:'H9-5' },
  { line:'H9', name:'Proximity', brand:'KEYENCE', model:'EV-108M', unit_price:490, no:'H9-6' }
];
const enriched = context.enrichPurchaseImportRowsFromMaster(rows, master);
assert.strictEqual(enriched, 2);
assert.strictEqual(rows[0].brand, 'SMC');
assert.strictEqual(rows[0].model, 'SY5120-5LZ');
assert.strictEqual(rows[0].part_id, 'H9-2');
assert.strictEqual(rows[0].review_required, false);
assert.strictEqual(rows[1].model, 'CQ2B40-50D');
assert.strictEqual(rows[1].include, true);
assert.strictEqual(rows[2].model, ''); // ambiguous same name + same price: do not guess
assert.strictEqual(rows[2].review_required, true);
console.log('Purchase import Item Master matching checks passed');
