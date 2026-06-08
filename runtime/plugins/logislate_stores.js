'use strict';

// Drop this file into runtime/plugins/ on the USB agent.
// Requires: npm install xlsx exceljs  (run once in usb-root/)

const fs   = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(__dirname, '..', '..', 'bridge', 'config.json');

function loadConfig() {
  return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
}

function safeDir() {
  const cfg = loadConfig();
  return (cfg.runtime || {}).safe_directory || '';
}

function validatePath(target, base) {
  const rel = path.relative(base, target);
  return !rel.startsWith('..') && !path.isAbsolute(rel);
}

function formatToShortDate(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return String(val).split(' ')[0];
  return d.toISOString().split('T')[0];
}

function compareStrings(a, b) {
  a = String(a || '').toLowerCase().trim();
  b = String(b || '').toLowerCase().trim();
  if (!a || !b) return 0;
  if (a === b) return 1;
  const map = new Map();
  for (let i = 0; i < a.length - 1; i++) {
    const bi = a.substring(i, i + 2);
    map.set(bi, (map.get(bi) || 0) + 1);
  }
  let intersect = 0;
  for (let i = 0; i < b.length - 1; i++) {
    const bi = b.substring(i, i + 2);
    if (map.get(bi) > 0) { map.set(bi, map.get(bi) - 1); intersect++; }
  }
  return (2.0 * intersect) / (a.length + b.length - 2);
}

module.exports = {
  name: 'logislate_stores',
  description:
    'Process a store-based logistics manifest Excel file from the safe directory. ' +
    'Deduplicates entries, fuzzy-groups deliveries by Sending_Store and IBTs by Receiving_Store, ' +
    'splits output into per-store tabs with short/long-distance and IBT sections, ' +
    'and saves the result to an output subfolder inside safe_directory. ' +
    'Use input_file="auto" to process the first .xlsx found.',
  parameters: {
    type: 'object',
    properties: {
      input_file: {
        type: 'string',
        description:
          'Filename of the .xlsx manifest inside safe_directory, or "auto" to pick the first .xlsx found.',
      },
      output_folder: {
        type: 'string',
        description:
          'Subfolder name inside safe_directory where the output file will be written. Defaults to "output".',
      },
    },
    required: ['input_file'],
  },

  async run({ input_file, output_folder = 'output' }) {
    let XLSX, ExcelJS;
    try { XLSX = require('xlsx'); }
    catch { return 'Error: "xlsx" package not installed. Run: npm install xlsx'; }
    try { ExcelJS = require('exceljs'); }
    catch { return 'Error: "exceljs" package not installed. Run: npm install exceljs'; }

    const safe = safeDir();
    if (!safe) return 'Error: safe_directory is not set in bridge/config.json.';
    if (!fs.existsSync(safe)) return `Error: safe_directory does not exist: ${safe}`;

    // ── resolve input file ────────────────────────────────────────────────────
    let inputPath;
    if (!input_file || input_file === 'auto') {
      const found = fs.readdirSync(safe).filter(f => /\.xlsx?$/i.test(f));
      if (!found.length) return `Error: No .xlsx files found in safe_directory: ${safe}`;
      inputPath = path.join(safe, found[0]);
    } else {
      inputPath = path.join(safe, input_file);
    }
    if (!validatePath(inputPath, safe)) return 'Error: input_file path escapes safe_directory.';
    if (!fs.existsSync(inputPath)) return `Error: File not found: ${inputPath}`;

    // ── resolve output folder ─────────────────────────────────────────────────
    const outDir = path.join(safe, output_folder);
    if (!validatePath(outDir, safe)) return 'Error: output_folder path escapes safe_directory.';
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    // ── read workbook ─────────────────────────────────────────────────────────
    let rawData;
    try {
      const buf   = fs.readFileSync(inputPath);
      const wb_in = XLSX.read(buf, { type: 'buffer', cellDates: true });
      rawData     = XLSX.utils.sheet_to_json(wb_in.Sheets[wb_in.SheetNames[0]]);
    } catch (err) {
      return `Error reading file: ${err.message}`;
    }
    if (!rawData.length) return 'Error: Spreadsheet has no data rows.';

    // ── step 1: clean & deduplicate ───────────────────────────────────────────
    const seen = new Set();
    const cleanedData = [];
    rawData.forEach(row => {
      const refSource = row.Shipper_Ref || row.External_ID || '';
      const cleanID   = String(refSource).replace(/\.0$/, '').trim();
      if (cleanID && !seen.has(cleanID)) {
        seen.add(cleanID);
        const isIBT         = cleanID.startsWith('46');
        const sendingStore  = String(row.Sending_Store  || 'Unknown').trim();
        const receivingStore= String(row.Receiving_Store || 'Unknown').trim();
        cleanedData.push({
          ...row,
          _internalID:    cleanID,
          _isIBT:         isIBT,
          _shortDate:     formatToShortDate(row.Booking_Date),
          _tripID:        String(row.ID || ''),
          _dist:          parseFloat(row.Delivery_Distance || 0),
          _weight:        parseFloat(row.Delivery_Weight || 0),
          _customer:      String(row.Customer || ''),
          _address:       String(row.To_Location || ''),
          _sendingStore:  sendingStore,
          _receivingStore:receivingStore,
          // deliveries tab under sending store; IBTs tab under receiving store
          _storeKey:      isIBT ? receivingStore : sendingStore,
        });
      }
    });
    cleanedData.sort((a, b) => a._storeKey.localeCompare(b._storeKey));

    // ── step 2: group ─────────────────────────────────────────────────────────
    const processedRefs  = new Set();
    const groupedResults = [];
    for (let i = 0; i < cleanedData.length; i++) {
      const p = cleanedData[i];
      if (processedRefs.has(p._internalID)) continue;
      const group = {
        ...p,
        subInvoices: [],
        subTrips:    [],
        totalW:      p._weight,
        reason:      'No Grouping done',
        confValue:   1.0,
      };
      processedRefs.add(p._internalID);

      for (let j = i + 1; j < cleanedData.length; j++) {
        const c = cleanedData[j];
        if (processedRefs.has(c._internalID)) continue;
        let isMatch = false, matchType = '';
        if (p._shortDate === c._shortDate) {
          if (!p._isIBT && !c._isIBT) {
            if (p._sendingStore && p._sendingStore === c._sendingStore) {
              const cSim  = compareStrings(p._customer, c._customer);
              const aSim  = compareStrings(p._address,  c._address);
              const dDiff = Math.abs(p._dist - c._dist);
              if      (cSim > 0.88 && aSim > 0.85) { isMatch = true; matchType = 'Name & Address'; }
              else if (cSim > 0.88 && dDiff <= 10)  { isMatch = true; matchType = 'Name & Distance (Unable to match Addresses)'; }
            }
          } else if (p._isIBT && c._isIBT) {
            if (p._sendingStore === c._sendingStore && p._receivingStore === c._receivingStore) {
              isMatch = true; matchType = 'match on Sending and Receiving StoreNames';
            }
          }
        }
        if (isMatch) {
          group.subInvoices.push(c._internalID);
          group.subTrips.push(c._tripID);
          group.totalW      += c._weight;
          group.reason       = `Grouped via ${matchType}`;
          group.confValue    = 0.95;
          processedRefs.add(c._internalID);
        }
      }
      groupedResults.push(group);
    }

    // ── step 3: build Excel output ────────────────────────────────────────────
    const wb_out   = new ExcelJS.Workbook();
    const rawSheet = wb_out.addWorksheet('raw data');
    const allHeaders = Object.keys(rawData[0]);
    rawSheet.addRow(allHeaders);
    cleanedData.forEach(r => rawSheet.addRow(allHeaders.map(h => r[h])));

    const storeKeys = [...new Set(groupedResults.map(g => g._storeKey || 'Unknown'))];
    for (const storeKey of storeKeys) {
      const sheet = wb_out.addWorksheet(
        storeKey.substring(0, 31).replace(/[\[\]\*\?\/\\]/g, '')
      );
      sheet.columns = [
        { header: 'Booking_Date',      key: 'h_date',   width: 15 },
        { header: 'Shipper_Ref',       key: 'h_ref',    width: 15 },
        { header: 'ID',                key: 'h_id',     width: 15 },
        { header: 'Customer',          key: 'h_cust',   width: 25 },
        { header: 'To_Location',       key: 'h_addr',   width: 35 },
        { header: 'Sending_Store',     key: 'h_send',   width: 18 },
        { header: 'Receiving_Store',   key: 'h_recv',   width: 18 },
        { header: 'Delivery_Distance', key: 'h_dist',   width: 15 },
        { header: 'Delivery_Weight',   key: 'h_weight', width: 15 },
        { header: 'Grouped_invoice',   key: 'h_ginv',   width: 30 },
        { header: 'Grouped_trip',      key: 'h_gtrip',  width: 30 },
        { header: 'Confidence',        key: 'h_conf',   width: 12 },
        { header: 'Reasoning',         key: 'h_reas',   width: 50 },
      ];

      const coData = groupedResults.filter(g => (g._storeKey || 'Unknown') === storeKey);
      const d_le   = coData.filter(r => !r._isIBT && r._dist <= 150);
      const d_gt   = coData.filter(r => !r._isIBT && r._dist >  150);
      const i_le   = coData.filter(r =>  r._isIBT && r._dist <= 150);
      const i_gt   = coData.filter(r =>  r._isIBT && r._dist >  150);

      const mapRow = r => ({
        h_date:   r._shortDate,
        h_ref:    r._internalID,
        h_id:     r._tripID,
        h_cust:   r._customer,
        h_addr:   r._address,
        h_send:   r._sendingStore  || '',
        h_recv:   r._receivingStore || '',
        h_dist:   r._dist,
        h_weight: r.totalW,
        h_ginv:   r.subInvoices.length ? r.subInvoices.join(' // ') : '',
        h_gtrip:  r.subTrips.length    ? r.subTrips.join(' // ')    : '',
        h_conf:   (r.confValue * 100) + '%',
        h_reas:   r.reason,
      });

      [d_le, d_gt, [{}], i_le, i_gt].forEach(subset =>
        subset.forEach(r => {
          if (r._internalID) sheet.addRow(mapRow(r));
          else sheet.addRow({});
        })
      );
      sheet.getRow(1).font = { bold: true };
    }

    const timestamp  = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outputName = `Logistics_Stores_${timestamp}.xlsx`;
    const outputPath = path.join(outDir, outputName);

    try {
      const buf = await wb_out.xlsx.writeBuffer();
      fs.writeFileSync(outputPath, buf);
    } catch (err) {
      return `Error writing output file: ${err.message}`;
    }

    return (
      `Done. Processed ${cleanedData.length} unique records into ${groupedResults.length} grouped entries ` +
      `across ${storeKeys.length} store tab(s). Output saved to: ${outputPath}`
    );
  },
};
