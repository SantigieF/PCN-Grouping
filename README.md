/**
 * EU Poison Centre Notification (PCN) — Mode 1 / Conservative grouping screen (v1.2 diagnostics).
 *
 * This is a preliminary regulatory screening aid, NOT an Annex VIII validation,
 * IUCLID submission, PCN authorisation, or UFI allocation tool.
 *
 * Expected source structure:
 *   row 1: raw-material headings (one per four-column block; may be merged)
 *   row 2: Actual concentration | Concentration range | Min | Max
 *   column A: product identifier; each subsequent row is one product
 *
 * The script does NOT modify the source. Every raw-material heading is treated
 * as a component key regardless of whether the raw material is a substance
 * or a mixture-in-mixture (MiM). No type classification is requested.
 *
 * This revision adds source-cell diagnostics and reads native percentage formatting.
 * A component reference sheet allows OPTIONAL common IDs for synonymous
 * raw-material headings and a YES/NO major-concern flag for range validation.
 * All group-submission decisions remain preliminary pending PCN review.
 */
const PCN_MODE1 = {
  SOURCE_SHEET: 'Formulations',       // CHANGE to your source tab name
  MATERIAL_HEADER_ROW: 1,           // row with raw material names
  SUBHEADER_ROW: 2,                 // row with Actual / Range / Min / Max
  FIRST_DATA_ROW: 3,
  PRODUCT_ID_COLUMN: 1,             // 1 = column A
  FIRST_MATERIAL_COLUMN: 2,         // 2 = column B; unique stable raw-material headings
  INPUT_PERCENT_MODE: 'AUTO',        // AUTO: detect Google Sheets percentage formatting;
                                     // NUMBER: 12 means 12%; FRACTION: 0.12 means 12%.
  VALUES_ARE_FRACTIONS: false,      // Legacy setting; ignored unless INPUT_PERCENT_MODE is removed.
  ACTUAL_TOTAL_TOLERANCE: 1,        // diagnostic tolerance in percentage points, NOT a legal PCN threshold
  READ_BATCH_ROWS: 150,
  WRITE_BATCH_ROWS: 2000,
  EPSILON: 1e-7
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PCN Mode 1')
    .addItem('1. Set up / update lookup sheets', 'pcnMode1Setup')
    .addItem('2. Run conservative grouping', 'pcnMode1Run')
    .addToUi();
}

/** Run first to create optional component IDs / major-concern and product CLP lookups. */
function pcnMode1Setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = pcnMode1Source_(ss);
  const materials = pcnMode1Materials_(source);
  const ids = pcnMode1ProductIds_(source);
  pcnMode1EnsureRegister_(ss, materials.map(x => x.name));
  pcnMode1EnsureClassifications_(ss, ids);
  ss.toast('Lookup sheets ready. Substance/MiM type is NOT required. Fill major-concern YES/NO and final-mixture CLP classifications.', 'PCN Mode 1', 10);
}

/** Run again after completing the two reference sheets. */
function pcnMode1Run() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = pcnMode1Source_(ss);
  const materials = pcnMode1Materials_(source);
  pcnMode1EnsureRegister_(ss, materials.map(x => x.name));
  pcnMode1EnsureClassifications_(ss, pcnMode1ProductIds_(source));
  const register = pcnMode1ReadRegister_(ss);
  const classes = pcnMode1ReadClasses_(ss);
  const issues = [];
  const products = [];
  const seenIds = new Set();
  const lastColumn = source.getLastColumn();
  const lastRow = source.getLastRow();
  // Read source formatting ONCE. For percentage-formatted numeric cells, Sheets
  // getValues() returns fractions (e.g. displayed 12% is numeric 0.12).
  // All rows of a material column should use the same number format.
  const percentFormats = source.getRange(PCN_MODE1.FIRST_DATA_ROW, 1, 1, lastColumn).getNumberFormats()[0];
  const diagnosticRows = [['Product ID', 'Sheet row', 'Raw material', 'Actual cell', 'Raw actual', 'Actual % interpreted', 'Min cell', 'Reported min %', 'Max cell', 'Reported max %', 'Width pp', 'Permitted width pp', 'Applicable table', 'Status', 'Explanation', 'Source number format']];
  const totalRows = [['Product ID', 'Source row', 'Sum actual %', 'Sum lower bounds %', 'Sum upper bounds %', 'Difference from 100 pp', 'Raw-material count', 'Largest component contributions', 'Interpretation']];
  let widthFlagCount = 0;
  let totalFlagCount = 0;

  for (let rowStart = PCN_MODE1.FIRST_DATA_ROW; rowStart <= lastRow; rowStart += PCN_MODE1.READ_BATCH_ROWS) {
    const count = Math.min(PCN_MODE1.READ_BATCH_ROWS, lastRow - rowStart + 1);
    const raw = source.getRange(rowStart, 1, count, lastColumn).getValues();
    for (let j = 0; j < raw.length; j++) {
      const values = raw[j];
      const sheetRow = rowStart + j;
      const suppliedId = String(values[PCN_MODE1.PRODUCT_ID_COLUMN - 1] == null ? '' : values[PCN_MODE1.PRODUCT_ID_COLUMN - 1]).trim();
      const nonblank = materials.some(m => ['actual', 'min', 'max'].some(f => m[f] != null && values[m[f]] !== '' && values[m[f]] != null));
      if (!suppliedId && !nonblank) continue;
      const productId = suppliedId || 'MISSING_ID_ROW_' + sheetRow;
      const ownIssues = [];
      const addIssue = (component, code, detail) => {
        const entry = { productId, sheetRow, component, code, detail };
        ownIssues.push(entry);
        issues.push(entry);
      };
      if (!suppliedId) addIssue('', 'MISSING_PRODUCT_ID', 'Product ID is blank.');
      if (seenIds.has(productId)) addIssue('', 'DUPLICATE_PRODUCT_ID', 'Product ID appears in multiple source rows.');
      seenIds.add(productId);
      const clp = classes.get(productId) || { health: '', physical: '' };
      if (!clp.health || !clp.physical) addIssue('', 'MISSING_FINAL_CLP', 'Enter both final-mixture health and physical classifications in PCN_Product_Classification; use NONE explicitly if none.');

      const components = [];
      const componentIds = new Set();
      let totalActual = 0;
      let totalLow = 0;
      let totalHigh = 0;
      for (const m of materials) {
        const actual = pcnMode1Percent_(values[m.actual], percentFormats[m.actual]);
        const low = pcnMode1Percent_(values[m.min], percentFormats[m.min]);
        const high = pcnMode1Percent_(values[m.max], percentFormats[m.max]);
        if (actual == null && low == null && high == null) continue;
        if (actual == null) {
          addIssue(m.name, 'MISSING_ACTUAL', 'Minimum/maximum supplied but actual concentration is blank.');
          continue;
        }
        if (!Number.isFinite(actual) || actual < 0 || actual > 100) {
          addIssue(m.name, 'INVALID_ACTUAL', 'Actual concentration must be within 0–100%.');
          continue;
        }
        if (actual === 0) {
          if ((low != null && low !== 0) || (high != null && high !== 0)) {
            addIssue(m.name, 'ZERO_WITH_NONZERO_RANGE', 'Actual is zero but the range is not zero; verify component presence.');
          }
          continue;  // Absent raw materials do not form part of the fingerprint.
        }
        totalActual += actual;
        const info = register.get(pcnMode1NormaliseName_(m.name)) || {};
        // Identity is based on a stable raw-material heading or an OPTIONAL common
        // ID. We never need to know whether the component is a substance or MiM.
        const componentId = 'RM:' + pcnMode1NormaliseName_(info.id || m.name);
        const concern = pcnMode1Concern_(info.concern);
        if (concern == null) addIssue(m.name, 'MAJOR_CONCERN_UNVERIFIED', 'Enter YES or NO in PCN_Component_Register to validate Annex VIII Table 1/2 widths. Grouping remains subject to review.');
        if (componentIds.has(componentId)) addIssue(m.name, 'DUPLICATE_CANONICAL_COMPONENT', 'Two raw-material columns resolve to the same component identity; check whether quantities need aggregation.');
        componentIds.add(componentId);
        const actualA1 = pcnMode1ColumnLabel_(m.actual + 1) + sheetRow;
        const minA1 = pcnMode1ColumnLabel_(m.min + 1) + sheetRow;
        const maxA1 = pcnMode1ColumnLabel_(m.max + 1) + sheetRow;
        const reportedWidth = Number.isFinite(low) && Number.isFinite(high) ? high - low : null;
        const permittedWidth = concern == null || !Number.isFinite(low) ? null : pcnMode1Width_(low, concern);
        let rangeProblem = '';
        let explanation = '';
        if (low == null || high == null) {
          rangeProblem = 'MISSING_REPORTED_RANGE';
          explanation = 'Actual concentration is positive but minimum or maximum is blank.';
        } else if (![low, high].every(Number.isFinite) || low < 0 || high > 100 || low > high || actual < low - PCN_MODE1.EPSILON || actual > high + PCN_MODE1.EPSILON) {
          rangeProblem = 'INVALID_RANGE';
          explanation = 'Required: 0 <= min <= actual <= max <= 100, using consistent percentage units.';
        } else if (low === 0) {
          rangeProblem = 'ZERO_LOWER_BOUND_REVIEW';
          explanation = 'Zero lower limit on a positive component requires review under applicable PCN concentration rules.';
        } else if (concern != null && high - low > permittedWidth + PCN_MODE1.EPSILON) {
          widthFlagCount++;
          rangeProblem = 'ANNEX_VIII_WIDTH';
          explanation = 'Range ' + pcnMode1Fmt_(low) + '–' + pcnMode1Fmt_(high) + '% has width ' + pcnMode1Fmt_(high - low) +
            ' percentage points; the ' + (concern ? 'Table 1' : 'Table 2') + ' limit is ' + permittedWidth +
            ' because the lower bound is in ' + pcnMode1Band_(low, concern) + '. If a higher band is crossed, the lower maximum width still applies.';
        }
        if (rangeProblem) {
          addIssue(m.name, rangeProblem, explanation + ' Source: ' + minA1 + ' and ' + maxA1 + '.');
          diagnosticRows.push([productId, sheetRow, m.name, actualA1, String(values[m.actual]), pcnMode1Fmt_(actual),
            minA1, pcnMode1Fmt_(low), maxA1, pcnMode1Fmt_(high), pcnMode1Fmt_(reportedWidth),
            pcnMode1Fmt_(permittedWidth), concern == null ? 'UNVERIFIED' : concern ? 'Table 1' : 'Table 2', rangeProblem, explanation, percentFormats[m.actual] || '']);
        }
        if (Number.isFinite(low)) totalLow += low;
        if (Number.isFinite(high)) totalHigh += high;
        components.push({ material: m.name, id: componentId, low, high, actual, concern });
      }
      if (Math.abs(totalActual - 100) > PCN_MODE1.ACTUAL_TOTAL_TOLERANCE) {
        totalFlagCount++;
        const topContributors = components.slice().sort((a, b) => b.actual - a.actual).slice(0, 10)
          .map(c => c.material + '=' + pcnMode1Fmt_(c.actual) + '%').join('; ');
        addIssue('', 'ACTUAL_TOTAL_REVIEW', 'Actual concentrations total ' + pcnMode1Fmt_(totalActual) +
          '% (lower-range sum ' + pcnMode1Fmt_(totalLow) + '%, upper-range sum ' + pcnMode1Fmt_(totalHigh) +
          '%). Check raw values, percentage formatting, duplicated columns and completeness; a sum of reported UPPER bounds above 100% is not itself an error.');
        totalRows.push([productId, sheetRow, pcnMode1Fmt_(totalActual), pcnMode1Fmt_(totalLow), pcnMode1Fmt_(totalHigh),
          pcnMode1Fmt_(totalActual - 100), components.length, topContributors,
          'Only ACTUAL sums are compared with 100%. If this is near 1%, check fraction-vs-percent input mode.']);
      }
      // The classification key is kept separate so different known CLP classes never cluster.
      const classKey = pcnMode1ClassKey_(clp.health) + '|' + pcnMode1ClassKey_(clp.physical);
      const componentSet = components.map(c => c.id).sort();
      const bands = components.map(c => c.id + '@' + pcnMode1Fmt_(c.low) + ':' + pcnMode1Fmt_(c.high)).sort();
      const groupKey = JSON.stringify([classKey, bands]);
      const cohortKey = JSON.stringify([classKey, componentSet]);
      const bandsOnlyKey = JSON.stringify(bands);
      const componentsOnlyKey = JSON.stringify(componentSet);
      products.push({ productId, sheetRow, components, issues: ownIssues, health: clp.health, physical: clp.physical, groupKey, cohortKey, bandsOnlyKey, componentsOnlyKey });
    }
  }

  const groupsByKey = new Map();
  const cohorts = new Map();
  const sameRangesIgnoringClp = new Map();
  const sameComponentsIgnoringClp = new Map();
  for (const product of products) {
    if (!groupsByKey.has(product.groupKey)) groupsByKey.set(product.groupKey, []);
    groupsByKey.get(product.groupKey).push(product);
    cohorts.set(product.cohortKey, (cohorts.get(product.cohortKey) || 0) + 1);
    sameRangesIgnoringClp.set(product.bandsOnlyKey, (sameRangesIgnoringClp.get(product.bandsOnlyKey) || 0) + 1);
    sameComponentsIgnoringClp.set(product.componentsOnlyKey, (sameComponentsIgnoringClp.get(product.componentsOnlyKey) || 0) + 1);
  }
  const groups = Array.from(groupsByKey, ([key, members]) => ({ key, members }))
    .sort((a, b) => b.members.length - a.members.length || a.key.localeCompare(b.key));
  let nextGroupId = 1;
  for (const group of groups) {
    group.id = group.members.length > 1 ? 'M1-' + String(nextGroupId++).padStart(4, '0') : '';
    group.status = group.members.some(x => x.issues.length) ? 'REVIEW_REQUIRED' : 'POTENTIAL_GROUP';
  }

  const summary = [['Group ID', 'Products', 'Assessment', 'Health CLP', 'Physical CLP', 'Source product IDs']];
  const members = [['Product ID', 'Source row', 'Candidate group', 'Group size', 'Assessment', 'Components', 'Same-composition cohort size', 'Issue codes / note']];
  const ranges = [['Candidate group', 'Component heading', 'Raw-material component key', 'Reported min %', 'Reported max %', 'Minimum actual %', 'Maximum actual %', 'Annex VIII table', 'Width check']];
  const exceptions = [['Product ID', 'Source row', 'Raw material', 'Code', 'Explanation']];
  for (const issue of issues) exceptions.push([issue.productId, issue.sheetRow, issue.component, issue.code, issue.detail]);

  for (const group of groups) {
    if (group.id) {
      summary.push([group.id, group.members.length, group.status, group.members[0].health, group.members[0].physical,
        group.members.map(m => m.productId).join(', ').slice(0, 40000)]);
      const first = group.members[0];
      for (const c of first.components.sort((a, b) => a.id.localeCompare(b.id))) {
        const all = group.members.flatMap(m => m.components.filter(z => z.id === c.id));
        const actuals = all.map(z => z.actual);
        const widthStatus = c.concern == null ? 'REVIEW: hazard status missing'
          : (c.low == null || c.high == null) ? 'REVIEW: missing bounds'
          : (c.high - c.low <= pcnMode1Width_(c.low, c.concern) + PCN_MODE1.EPSILON) ? 'Width OK'
          : 'REVIEW: width exceeds limit';
        ranges.push([group.id, c.material, c.id, c.low, c.high, Math.min(...actuals), Math.max(...actuals),
          c.concern === true ? 'Table 1' : c.concern === false ? 'Table 2' : 'Unknown', widthStatus]);
      }
    }
    for (const p of group.members) {
      let status = group.id ? group.status : (p.issues.length ? 'REVIEW_REQUIRED' : 'NO_EXACT_MATCH');
      const cohortCount = cohorts.get(p.cohortKey) || 1;
      let unmatchedReason = '';
      if (!group.id) {
        if ((sameRangesIgnoringClp.get(p.bandsOnlyKey) || 0) > 1) unmatchedReason = 'Same components and ranges exist, but final CLP health/physical classification differs or is missing';
        else if (cohortCount > 1) unmatchedReason = 'Same component set and CLP classification, but different reported ranges (potential Mode 2 candidate)';
        else if ((sameComponentsIgnoringClp.get(p.componentsOnlyKey) || 0) > 1) unmatchedReason = 'Same components exist, but classifications and/or reported ranges differ';
        else unmatchedReason = 'No other product has the same present component identities';
      }
      const message = [...new Set(p.issues.map(x => x.code))].concat(unmatchedReason ? [unmatchedReason] : []).join('; ');
      members.push([p.productId, p.sheetRow, group.id || '', group.members.length, status,
        p.components.length, cohortCount, message]);
    }
  }

  pcnMode1Write_(ss, 'PCN_Mode1_Summary', summary);
  pcnMode1Write_(ss, 'PCN_Mode1_Members', members);
  pcnMode1Write_(ss, 'PCN_Mode1_Ranges', ranges);
  pcnMode1Write_(ss, 'PCN_Mode1_Exceptions', exceptions);
  pcnMode1Write_(ss, 'PCN_Mode1_Range_Diagnostics', diagnosticRows);
  pcnMode1Write_(ss, 'PCN_Mode1_Total_Diagnostics', totalRows);
  ss.toast('Screened ' + products.length + ' products; ' + widthFlagCount + ' width warnings and ' + totalFlagCount + ' actual-total warnings. See PCN_Mode1_*_Diagnostics.', 'PCN Mode 1', 15);
}

function pcnMode1Source_(ss) {
  const sh = ss.getSheetByName(PCN_MODE1.SOURCE_SHEET);
  if (!sh) throw new Error('Source tab "' + PCN_MODE1.SOURCE_SHEET + '" not found. Update SOURCE_SHEET in PCN_MODE1 config.');
  if (sh.getLastRow() < PCN_MODE1.FIRST_DATA_ROW) throw new Error('Source sheet has no data rows.');
  return sh;
}

function pcnMode1Materials_(sh) {
  const lastCol = sh.getLastColumn();
  const top = sh.getRange(PCN_MODE1.MATERIAL_HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const sub = sh.getRange(PCN_MODE1.SUBHEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  let currentName = '';
  const blocks = new Map();
  for (let col = PCN_MODE1.FIRST_MATERIAL_COLUMN - 1; col < lastCol; col++) {
    // Merged material headings are nonblank only in the left-most cell.
    const label = String(top[col] || '').trim();
    if (label) currentName = label;
    if (!currentName) continue;
    const field = pcnMode1SubField_(sub[col]);
    if (!field) continue;
    if (!blocks.has(currentName)) blocks.set(currentName, { name: currentName, actual: null, min: null, max: null });
    const obj = blocks.get(currentName);
    if (obj[field] != null) throw new Error('Duplicate "' + field + '" subcolumn under material "' + currentName + '". Ensure material headings are unique.');
    obj[field] = col; // zero-based index for getValues arrays
  }
  const out = [...blocks.values()];
  const invalid = out.filter(b => b.actual == null || b.min == null || b.max == null);
  if (!out.length || invalid.length) throw new Error('Cannot recognise Actual/Min/Max subcolumns. Check header row config and names. Incomplete blocks: ' + invalid.slice(0, 8).map(x => x.name).join(', '));
  return out;
}

function pcnMode1SubField_(name) {
  const s = String(name || '').trim().toLowerCase().replace(/[_\-]+/g, ' ');
  if (/\b(min|minimum|lower)\b/.test(s)) return 'min';
  if (/\b(max|maximum|upper)\b/.test(s)) return 'max';
  if (/\b(actual|act|nominal)\b/.test(s) || (/\bconc(?:entration)?\b/.test(s) && !/\brange\b/.test(s))) return 'actual';
  return null; // concentration-range descriptive/text subcolumn is not needed for Mode 1
}

function pcnMode1ProductIds_(sh) {
  const last = sh.getLastRow();
  const data = sh.getRange(PCN_MODE1.FIRST_DATA_ROW, PCN_MODE1.PRODUCT_ID_COLUMN, last - PCN_MODE1.FIRST_DATA_ROW + 1, 1).getValues();
  return [...new Set(data.map(r => String(r[0] == null ? '' : r[0]).trim()).filter(Boolean))];
}

function pcnMode1EnsureRegister_(ss, names) {
  // Existing five-column registers from the first script are supported without
  // deleting any data; the old Substance/MiM column is ignored completely.
  const headers = ['Raw material heading', 'Common component ID (optional)', 'Major concern (YES/NO)', 'Notes'];
  let sh = ss.getSheetByName('PCN_Component_Register');
  if (!sh) {
    sh = ss.insertSheet('PCN_Component_Register');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const layout = pcnMode1RegisterLayout_(sh);
  const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  const known = new Set(data.map(r => pcnMode1NormaliseName_(r[0])));
  const add = names.filter(n => !known.has(pcnMode1NormaliseName_(n)))
    .map(n => [n, ...Array(layout.width - 1).fill('')]);
  pcnMode1Append_(sh, add, layout.width);
}

function pcnMode1RegisterLayout_(sh) {
  const width = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, width).getDisplayValues()[0]
    .map(x => String(x || '').trim().toUpperCase());
  if (headers[0] !== 'RAW MATERIAL HEADING') {
    throw new Error('PCN_Component_Register: expected "Raw material heading" in A1.');
  }
  // Legacy sheet: A heading, B ID, C Type, D major concern, E Notes.
  if (headers[2] && headers[2].indexOf('TYPE') === 0 && headers[3] && headers[3].indexOf('MAJOR CONCERN') === 0) {
    return { width: Math.max(5, width), idIndex: 1, concernIndex: 3 };
  }
  // Simplified sheet: A heading, B optional ID, C major concern, D Notes.
  if (headers[2] && headers[2].indexOf('MAJOR CONCERN') === 0) {
    return { width: Math.max(4, width), idIndex: 1, concernIndex: 2 };
  }
  throw new Error('PCN_Component_Register: expected the simplified 4-column register or the previous 5-column register.');
}

function pcnMode1EnsureClassifications_(ss, ids) {
  const headers = ['Product ID', 'Health CLP classification (or NONE)', 'Physical CLP classification (or NONE)'];
  const sh = pcnMode1EnsureSheet_(ss, 'PCN_Product_Classification', headers);
  const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  const known = new Set(data.map(r => String(r[0] == null ? '' : r[0]).trim()));
  const add = ids.filter(id => !known.has(id)).map(id => [id, '', '']);
  pcnMode1Append_(sh, add, headers.length);
}

function pcnMode1EnsureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (headers.some((h, i) => existing[i] !== h)) throw new Error('The headers in ' + name + ' were changed. Restore original headings to avoid reading the wrong columns.');
  }
  return sh;
}

function pcnMode1Append_(sh, rows, width) {
  for (let start = 0; start < rows.length; start += PCN_MODE1.WRITE_BATCH_ROWS) {
    const part = rows.slice(start, start + PCN_MODE1.WRITE_BATCH_ROWS);
    sh.getRange(sh.getLastRow() + 1, 1, part.length, width).setValues(part);
  }
}

function pcnMode1ReadRegister_(ss) {
  const sh = ss.getSheetByName('PCN_Component_Register');
  const out = new Map();
  const layout = pcnMode1RegisterLayout_(sh);
  if (sh.getLastRow() === 1) return out;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, layout.width).getValues();
  for (const r of rows) {
    const key = pcnMode1NormaliseName_(r[0]);
    if (key) out.set(key, { id: r[layout.idIndex], concern: r[layout.concernIndex] });
  }
  return out;
}

function pcnMode1ReadClasses_(ss) {
  const sh = ss.getSheetByName('PCN_Product_Classification');
  const out = new Map();
  if (sh.getLastRow() === 1) return out;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (const r of rows) {
    const id = String(r[0] == null ? '' : r[0]).trim();
    if (id) out.set(id, { health: String(r[1] || '').trim(), physical: String(r[2] || '').trim() });
  }
  return out;
}

function pcnMode1Percent_(v, numberFormat) {
  if (v === '' || v == null || v === '-') return null;
  const isTextWithPercent = typeof v === 'string' && /%\s*$/.test(v.trim());
  let n;
  if (typeof v === 'number') n = v;
  else {
    const str = String(v).trim().replace(/\s/g, '').replace('%', '').replace(',', '.');
    if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(str)) return NaN;
    n = Number(str);
  }
  // Text containing '%' (e.g. '12%') is ALREADY percent units.
  if (isTextWithPercent) return n;
  const mode = String(PCN_MODE1.INPUT_PERCENT_MODE || (PCN_MODE1.VALUES_ARE_FRACTIONS ? 'FRACTION' : 'NUMBER')).toUpperCase();
  if (mode === 'FRACTION') return n * 100;
  if (mode === 'NUMBER') return n;
  if (mode !== 'AUTO') throw new Error('INPUT_PERCENT_MODE must be AUTO, NUMBER or FRACTION.');
  // In Google Sheets, number format "0.00%" means underlying 0.12 represents 12%.
  return /%/.test(numberFormat || '') ? n * 100 : n;
}

function pcnMode1Band_(low, concern) {
  const lt = threshold => low < threshold - PCN_MODE1.EPSILON;
  if (concern) {
    if (lt(0.1)) return '0–<0.1%';
    if (lt(1)) return '0.1–<1%';
    if (lt(10)) return '1–<10%';
    if (lt(25)) return '10–<25%';
    return '25–<100%';
  }
  if (lt(1)) return '0–<1%';
  if (lt(10)) return '1–<10%';
  if (lt(25)) return '10–<25%';
  return '25–<100%';
}

function pcnMode1ColumnLabel_(col) {
  let result = '';
  while (col > 0) {
    col--;
    result = String.fromCharCode(65 + col % 26) + result;
    col = Math.floor(col / 26);
  }
  return result;
}

function pcnMode1Width_(min, isMajorConcern) {
  // Allow tiny Sheets formula floating-point deviations at exact band thresholds;
  // do not permit a materially lower value to be treated as the higher band.
  const lt = threshold => min < threshold - PCN_MODE1.EPSILON;
  if (isMajorConcern) {
    if (lt(0.1)) return 0.1;
    if (lt(1)) return 0.3;
    if (lt(10)) return 1;
    if (lt(25)) return 3;
    return 5;
  }
  if (lt(1)) return 1;
  if (lt(10)) return 3;
  if (lt(25)) return 10;
  return 20;
}

function pcnMode1Concern_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase();
  if (['YES', 'Y', 'TRUE', '1'].includes(s)) return true;
  if (['NO', 'N', 'FALSE', '0'].includes(s)) return false;
  return null;
}

function pcnMode1ClassKey_(v) {
  const s = String(v || '').trim().replace(/\s+/g, ' ').toUpperCase();
  if (['NONE', 'NOT CLASSIFIED', 'NO CLASSIFICATION', 'NOT APPLICABLE'].includes(s)) return 'NONE';
  // For hazard categories entered as a semicolon-separated set, ordering is irrelevant.
  return s ? s.split(';').map(x => x.trim()).filter(Boolean).sort().join('; ') : 'UNKNOWN';
}

function pcnMode1NormaliseName_(v) {
  return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
}

function pcnMode1Fmt_(v) {
  return v == null || !Number.isFinite(v) ? 'MISSING' : Number(v.toFixed(10)).toString();
}

function pcnMode1Write_(ss, name, content) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  const safe = content.map(row => row.map(v => typeof v === 'string' && /^=/.test(v) ? "'" + v : v));
  for (let start = 0; start < safe.length; start += PCN_MODE1.WRITE_BATCH_ROWS) {
    const part = safe.slice(start, start + PCN_MODE1.WRITE_BATCH_ROWS);
    sh.getRange(start + 1, 1, part.length, content[0].length).setValues(part);
  }
  sh.getRange(1, 1, 1, content[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
}
