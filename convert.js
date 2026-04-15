#!/usr/bin/env node
/**
 * Shopify (Matrixify) → WooCommerce Product Import CSV
 * GlossyLounge Migration Script
 *
 * Usage:
 *   node convert.js [input.xlsx] [output.csv] [--dry-run]
 *
 * Defaults:
 *   input  = ./input.xlsx
 *   output = ./woocommerce-import.csv
 *
 * Gift cards (handles giftcard / glossy-lounge-gift-cards or Type "Gift Card"):
 *   exported as simple (one variant) or variable with Attribute 1 = pa_denomination.
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ─── CLI args ────────────────────────────────────────────────────────────────
const args    = process.argv.slice(2).filter(a => !a.startsWith('--'));
const flags   = process.argv.slice(2).filter(a => a.startsWith('--'));
const DRY_RUN = flags.includes('--dry-run');

const INPUT_FILE      = args[0] || path.join(__dirname, 'input.xlsx');
const OUTPUT_FILE     = args[1] || path.join(__dirname, 'woocommerce-import.csv');
const SKU_LOG_FILE    = path.join(__dirname, 'sku-conflicts.log');
const COLLECTION_FILE = path.join(__dirname, 'collection-export.xlsx');

// ─── Colour data ─────────────────────────────────────────────────────────────
// slug → { name, hex }
const COLOUR_MAP = {
  'jet-black':          { name: 'Jet Black',         hex: '#383E4E' },
  'stone-grey':         { name: 'Stone Grey',         hex: '#939597' },
  'white-cloud':        { name: 'White Cloud',        hex: '#EDEDE6' },
  'leafy-khaki':        { name: 'Leafy Khaki',        hex: '#706F48' },
  'sea-sage-green':     { name: 'Sea Sage Green',     hex: '#96A480' },
  'natural-taupe':      { name: 'Natural Taupe',      hex: '#AF9483' },
  'sunset-red':         { name: 'Sunset Red',         hex: '#D1332C' },
  'cherry-red':         { name: 'Cherry Red',         hex: '#C14A52' },
  'violet-blue':        { name: 'Violet Blue',        hex: '#464F75' },
  'ocean-pale-blue':    { name: 'Ocean Pale Blue',    hex: '#9CC3DC' },
  'pastel-lavender':    { name: 'Pastel Lavender',    hex: '#B9B3D1' },
  'pastel-pink':        { name: 'Pastel Pink',        hex: '#F7C8C2' },
  'steel-blue':         { name: 'Steel Blue',         hex: '#999FBF' },
  'wheat':              { name: 'Wheat',              hex: '#CDBFBA' },
  'slate-green':        { name: 'Slate Green',        hex: '#81988E' },
  'pink':               { name: 'Pink',               hex: '#FBD8DB' },
  'lavender':           { name: 'Lavender',           hex: '#B7A0BB' },
  'white':              { name: 'White',              hex: '#FFFFFF' },
  'navy':               { name: 'Navy',               hex: '#293A65' },
  'navy-blue':          { name: 'Navy',               hex: '#293A65' },
  'powder-blue':        { name: 'Powder Blue',        hex: '#D8EFF5' },
  'muted-blue':         { name: 'Muted Blue',         hex: '#E8EDF0' },
  'warm-beige':         { name: 'Warm Beige',         hex: '#EDE3D4' },
  'nectar':             { name: 'Nectar',             hex: '#ECDACC' },
  'brown':              { name: 'Brown',              hex: '#6A493C' },
  'stone':              { name: 'Stone',              hex: '#BAAFA4' },
  'offwhite':           { name: 'Offwhite',           hex: '#F2EDE4' },
  'ivory':              { name: 'Ivory',              hex: '#FFF9E6' },
  'light-pink':         { name: 'Light Pink',         hex: '#F0D5D2' },
  'sky-blue':           { name: 'Sky Blue',           hex: '#B9CFE3' },
  'olive':              { name: 'Olive',              hex: '#7E9055' },
  'green':              { name: 'Green',              hex: '#777C63' },
  'mustard-yellow':     { name: 'Mustard Yellow',     hex: '#DCCC50' },
  'red':                { name: 'Red',                hex: '#E00030' },
  'butter-yellow':      { name: 'Butter Yellow',      hex: '#FFF3BC' },
  'blush-pink':         { name: 'Blush Pink',         hex: '#F0D5D3' },
  'rust-brown':         { name: 'Rust Brown',         hex: '#81553D' },
  'taupe':              { name: 'Taupe',              hex: '#967E6C' },
  'charcoal-grey':      { name: 'Charcoal Grey',      hex: '#636364' },
  'khaki':              { name: 'Leafy Khaki',        hex: '#706F48' },
  'pale-blue':          { name: 'Ocean Pale Blue',    hex: '#9CC3DC' },
  'pale-lavender':      { name: 'Pastel Lavender',    hex: '#B9B3D1' },
  'black':              { name: 'Jet Black',          hex: '#383E4E' },
  'grey':               { name: 'Stone Grey',         hex: '#939597' },
  'blue':               { name: 'Violet Blue',        hex: '#464F75' },
  'beige':              { name: 'Warm Beige',         hex: '#EDE3D4' },
  // Two-tone slugs
  'sky-blue-ivory':     { name: 'Sky Blue / Ivory',   hex: '#B9CFE3,#FFF9E6' },
  'white-navy':         { name: 'White / Navy',       hex: '#FFFFFF,#293A65' },
  'brown-offwhite':     { name: 'Brown / Offwhite',   hex: '#BD9F83,#F2EDE4' },
  'ivory-brown':        { name: 'Ivory / Brown',      hex: '#FFF9E6,#6A493C' },
  'light-pink-ivory':   { name: 'Light Pink / Ivory', hex: '#F0D5D2,#FFF9E6' },
  'brown-stone':        { name: 'Brown / Stone',      hex: '#6A493C,#BAAFA4' },
  'ivory-green':        { name: 'Ivory / Green',      hex: '#FFF9E6,#777C63' },
  'olive-mustard-yellow':{ name: 'Olive / Mustard Yellow', hex: '#7E9055,#DCCC50' },
  // New colours (hex values from Shopify metafields)
  'turkish-coffee':     { name: 'Turkish Coffee',    hex: '#403A36' },
  'dark-earth':         { name: 'Dark Earth',        hex: '#4d2e19' },
  'peppercorn':         { name: 'Peppercorn',        hex: '#bfbdbb' },
  'jester-red':         { name: 'Jester Red',        hex: '#751921' },
  'deep-blue':          { name: 'Deep Blue',         hex: '#3e4071' },
  'rio-red':            { name: 'Rio Red',            hex: '#751922' },
  'cashmere-rose':      { name: 'Cashmere Rose',     hex: '#CF859E' },
  'olive-green':        { name: 'Olive Green',       hex: '#676B4A' },
  'periwinkle-sky':     { name: 'Periwinkle Sky',    hex: '#5F6BB0' },
  'sage-green':         { name: 'Sea Sage Green',    hex: '#96A480' },
  'warm-taupe':         { name: 'Warm Taupe',        hex: '#BB9C8A' },
  'feminine-pink':      { name: 'Feminine Pink',     hex: '#C599C7' },
};

// hex → name lookup (for metafield-based resolution), all keys lowercased
const HEX_TO_NAME = {};
for (const [slug, { name, hex }] of Object.entries(COLOUR_MAP)) {
  const key = hex.toLowerCase();
  if (!HEX_TO_NAME[key]) HEX_TO_NAME[key] = name;
}

// name → { name, hex } lookup (for Option2 Colour resolution), keys lowercased/trimmed
const COLOUR_BY_NAME = {};
for (const [slug, entry] of Object.entries(COLOUR_MAP)) {
  const key = entry.name.toLowerCase();
  if (!COLOUR_BY_NAME[key]) COLOUR_BY_NAME[key] = entry;
}
// Alias: Shopify Option2 uses "Pale Lavender" but COLOUR_MAP canonical name is "Pastel Lavender"
if (!COLOUR_BY_NAME['pale lavender']) COLOUR_BY_NAME['pale lavender'] = COLOUR_BY_NAME['pastel lavender'];
// Additional hex values seen in Shopify metafields but not in COLOUR_MAP
const EXTRA_HEX = {
  // Old metafield hex values that may still appear in Shopify data — map to correct names
  // (These differ from the corrected COLOUR_MAP hex values but may exist in product metafields)
  '#6b585a': 'Peppercorn',
  '#574334': 'Dark Earth',
  '#98253b': 'Jester Red',
  '#45387f': 'Deep Blue',
  '#86293b': 'Rio Red',
  '#b29784': 'Natural Taupe',
  '#f494be': 'Pastel Pink',
  '#c599c7': 'Feminine Pink',
};
for (const [hex, name] of Object.entries(EXTRA_HEX)) {
  if (!HEX_TO_NAME[hex]) HEX_TO_NAME[hex] = name;
}

/**
 * Sorted colour slugs by length DESC so we always try the longest suffix first
 * (e.g. "olive-mustard-yellow" before "olive").
 */
const COLOUR_SLUGS_BY_LENGTH = Object.keys(COLOUR_MAP).sort((a, b) => b.length - a.length);

// ─── Helpers ─────────────────────────────────────────────────────────────────

function deriveProductGroup(handle) {
  for (const slug of COLOUR_SLUGS_BY_LENGTH) {
    const suffix = `-${slug}`;
    if (handle.endsWith(suffix)) {
      return handle.slice(0, handle.length - suffix.length);
    }
  }
  // Noor Stars collaboration: strip -ns, -ns-1, -ns-2 etc.
  const nsMatch = handle.match(/^(.+)-ns(-\d+)?$/);
  if (nsMatch) return nsMatch[1];
  return handle;
}

function deriveColourFromHandle(handle) {
  for (const slug of COLOUR_SLUGS_BY_LENGTH) {
    const suffix = `-${slug}`;
    if (handle.endsWith(suffix)) {
      return COLOUR_MAP[slug] || null;
    }
  }
  // Noor Stars collaboration products: -ns, -ns-1, -ns-2, etc.
  if (/-ns(-\d+)?$/.test(handle)) {
    return { name: 'Noor Stars', hex: '#C4A265' };
  }
  return null;
}

/** Resolve colour from Shopify's Colour/Color option value (Option2 in XLSX). */
function resolveColourFromOptionValue(optionValue) {
  if (!optionValue) return null;
  const key = String(optionValue).trim().toLowerCase();
  if (!key) return null;
  return COLOUR_BY_NAME[key] || null;
}

/** Look through all product_color metafields for an entry matching the handle. */
function extractColourFromMetafields(handle, metafields) {
  for (const val of metafields) {
    if (!val) continue;
    const pipeIdx = val.lastIndexOf('|');
    if (pipeIdx === -1) continue;
    const mHandle = val.slice(pipeIdx + 1).trim();
    if (mHandle === handle) {
      const hex = val.slice(0, pipeIdx).trim();
      const name = HEX_TO_NAME[hex.toLowerCase()] || null;
      return { hex, name };
    }
  }
  return null;
}

// ─── Collection → category mapping (from collection-export.xlsx) ─────────────

const JUNK_COLLECTION_HANDLE_PREFIX = ['spo-', 'test'];
const JUNK_COLLECTION_HANDLES = new Set([
  'ultimate-search-bestseller-collection-do-not-delete',
  'colour-test-collection',
]);

function isJunkCollection(handle) {
  const h = (handle || '').toLowerCase().trim();
  if (!h) return true;
  if (JUNK_COLLECTION_HANDLES.has(h)) return true;
  return JUNK_COLLECTION_HANDLE_PREFIX.some(p => h.startsWith(p));
}

/**
 * Read collection-export.xlsx (Smart Collections + Custom Collections sheets)
 * and build a Map: productHandle → Set<collectionTitle>.
 */
function loadCollectionLookup(filePath) {
  const lookup = new Map(); // productHandle → Set<title>

  if (!fs.existsSync(filePath)) {
    console.warn(`⚠️  Collection file not found: ${filePath} — all products will be "Uncategorized"`);
    return lookup;
  }

  const wb = XLSX.readFile(filePath);

  for (const sheetName of wb.SheetNames) {
    // Skip the export summary sheet
    if (sheetName.toLowerCase().includes('summary')) continue;

    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (rows.length < 2) continue;

    const headers = rows[0];
    const colIdx = {};
    headers.forEach((h, i) => { if (h) colIdx[h] = i; });

    const HANDLE_IDX      = colIdx['Handle'];
    const TITLE_IDX       = colIdx['Title'];
    const TOP_ROW_IDX     = colIdx['Top Row'];
    const PROD_HANDLE_IDX = colIdx['Product: Handle'];

    if (HANDLE_IDX === undefined || PROD_HANDLE_IDX === undefined) continue;

    let currentTitle = null;
    let currentHandle = null;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];

      // Detect new collection: Top Row is truthy or Handle changes
      const rowHandle = row[HANDLE_IDX];
      const isTopRow = TOP_ROW_IDX !== undefined &&
        (row[TOP_ROW_IDX] === true || row[TOP_ROW_IDX] === 1 ||
         row[TOP_ROW_IDX] === 'true' || row[TOP_ROW_IDX] === 'Y');

      if (rowHandle && (rowHandle !== currentHandle || isTopRow)) {
        currentHandle = rowHandle;
        currentTitle = TITLE_IDX !== undefined ? String(row[TITLE_IDX] || '') : '';

        // Filter junk collections by handle
        if (isJunkCollection(currentHandle)) {
          currentTitle = null; // mark as skip
        }
      }

      // Skip if this collection is filtered out
      if (!currentTitle) continue;

      // Map product handle to this collection title
      const prodHandle = row[PROD_HANDLE_IDX];
      if (!prodHandle) continue;

      if (!lookup.has(prodHandle)) lookup.set(prodHandle, new Set());
      lookup.get(prodHandle).add(currentTitle);
    }
  }

  console.log(`📁 Loaded ${lookup.size} product → collection mappings from: ${filePath}`);
  return lookup;
}

function buildCategories(handle, collectionLookup) {
  const titles = collectionLookup.get(handle);
  if (!titles || titles.size === 0) return 'Uncategorized';
  return [...titles].join(',');
}

/**
 * Excel auto-converts sizes like "1-2", "3-4", "6-12", "6- 12" to dates.
 * Fix: normalise spaces around dash, then wrap with Excel formula ="value".
 * Excel shows the raw string; WooCommerce CSV importer strips ="..." on import.
 *
 * Also filters junk Option1 values: colour words, gift card amounts, Default Title.
 */
const JUNK_SIZE = /^(default title|blue|green|red|duties|\$[\d,]+|aed[\d,.]+)/i;
const DATE_LIKE_SIZE = /^\d{1,2}-\d{1,2}$/; // 1-2, 2-3, 6-12, 9-12 etc.

function normaliseSizeLabel(size) {
  return size.trim().replace(/\s*-\s*/, '-');
}

function safeSizeValue(size) {
  if (!size) return size;
  const normalised = normaliseSizeLabel(size);
  if (DATE_LIKE_SIZE.test(normalised)) {
    return `="${normalised}"`;
  }
  return normalised;
}

/** Term slug for pa_size: lowercase, 2XL → 2xl, One Size → one-size. */
function sizeToTermSlug(size) {
  if (!size) return '';
  const s = normaliseSizeLabel(size);
  if (/^one\s*size$/i.test(s)) return 'one-size';
  if (/^2\s*xl$/i.test(s) || s === '2XL') return '2xl';
  return s.toLowerCase();
}

function isValidSize(size) {
  if (!size) return false;
  const s = String(size).trim();
  if (!s) return false;
  if (JUNK_SIZE.test(s)) return false;
  return true;
}

/** Shopify gift-card handles; Matrixify `Type` is usually "Gift Card". */
const GIFT_CARD_HANDLES = new Set(['giftcard', 'glossy-lounge-gift-cards']);

function isGiftCardProduct(topRow, handle, typeColIndex) {
  if (GIFT_CARD_HANDLES.has(handle)) return true;
  if (typeColIndex === undefined) return false;
  const t = String(topRow[typeColIndex] != null ? topRow[typeColIndex] : '').toLowerCase();
  return /\bgift\s*card\b/.test(t) || t === 'gift_card' || t === 'giftcard';
}

/**
 * If UTF-8 text was mis-decoded as Latin-1 (common in bad CSV/Excel round-trips),
 * the string shows mojibake (e.g. Ã© for é, â€™ for ', â€" for —).
 * Re-encode JS code units as Latin-1 bytes and decode as UTF-8 to recover.
 * Only runs when typical mojibake byte-pair patterns are detected so valid
 * Unicode text is never touched.
 */
const MOJIBAKE_HINT =
  /\u00c3[\u00a0-\u00bf\u0080-\u00bf]|\u00c2[\u00a0-\u00bf]|\u00e2\u20ac/;

function repairUtf8MojibakeIfPresent(str) {
  if (str == null || str === '') return str;
  const s = String(str);
  if (!MOJIBAKE_HINT.test(s)) return s;
  try {
    return Buffer.from(s, 'latin1').toString('utf8');
  } catch {
    return s;
  }
}

/** Strip newlines from any text value so it never breaks CSV row boundaries. */
function flattenText(val) {
  if (!val) return '';
  return String(val).replace(/\r?\n/g, ' ').replace(/\r/g, ' ').trim();
}

/** Clean HTML for WooCommerce: flatten to single line, strip junk Shopify editor artifacts. */
function cleanHTML(html) {
  if (!html) return '';
  return String(html)
    .replace(/\r?\n/g, '')
    .replace(/\r/g, '')
    .replace(/<meta[^>]*charset[^>]*>/gi, '')
    .replace(/\s*data-mce-fragment="1"/gi, '')
    .trim();
}

/** Escape a CSV field value: wrap in quotes, double any internal quotes. */
function csvField(val) {
  if (val === null || val === undefined) return '';
  const s = String(val);
  if (s.startsWith('="') && s.endsWith('"')) return s;
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function csvRow(fields) {
  return fields.map(csvField).join(',');
}

// ─── CSV header definition ────────────────────────────────────────────────────
const CSV_HEADERS = [
  'ID', 'Type', 'SKU', 'Name', 'Published',
  'Short description', 'Description', 'Date sale price starts', 'Date sale price ends',
  'Tax status', 'Tax class', 'In stock?', 'Stock', 'Low stock amount',
  'Backorders allowed?', 'Sold individually?', 'Manage stock?', 'Weight (kg)', 'Length (cm)',
  'Width (cm)', 'Height (cm)', 'Allow customer reviews?', 'Purchase note',
  'Sale price', 'Regular price', 'Categories', 'Tags', 'Shipping class',
  'Images', 'Download limit', 'Download expiry days', 'Parent', 'Grouped products',
  'Upsells', 'Cross-sells', 'External URL', 'Button text', 'Position',
  'Attribute 1 name', 'Attribute 1 value(s)', 'Attribute 1 visible', 'Attribute 1 global',
  'Meta: product_group', 'Meta: color_name', 'Meta: color_hex', 'Meta: complete_the_look',
  'Meta: maincat', 'Meta: subcat',
];

// ─── Main processing ──────────────────────────────────────────────────────────

function buildEmptyProduct() {
  return Object.fromEntries(CSV_HEADERS.map(h => [h, '']));
}

/**
 * Ensure SKU uniqueness within the export.
 * - If a SKU is seen for the first time, it is returned unchanged.
 * - If a SKU is re-used for the same handle + size, we keep the first row and
 *   drop the duplicates by returning an empty string (caller can skip row).
 * - If a SKU is re-used for a different product handle, we append a numerical
 *   suffix (-2, -3, ...) to make it unique and record a warning.
 */
function makeUniqueSku(baseSku, handle, sizeKey, seenSkus, skuConflicts, warnings) {
  if (!baseSku) return '';
  const key = `${baseSku}@@${handle}@@${sizeKey || ''}`;

  if (!seenSkus.has(baseSku)) {
    // First time we see this raw SKU.
    seenSkus.set(baseSku, { ownerKey: key, count: 1 });
    return baseSku;
  }

  const info = seenSkus.get(baseSku);

  if (info.ownerKey === key) {
    // Exact same product+size combination already used this SKU — treat this
    // as a duplicate row from the source and let the caller skip it by
    // returning an empty string.
    return '';
  }

  // SKU conflict between different products – generate a suffixed variant.
  info.count += 1;
  const newSku = `${baseSku}-${info.count}`;

  // Record the new SKU as seen, owned by this product+size combination.
  seenSkus.set(newSku, { ownerKey: key, count: 1 });

  const owners = skuConflicts.get(baseSku) || new Set();
  owners.add(info.ownerKey);
  owners.add(key);
  skuConflicts.set(baseSku, owners);

  warnings.push(`SKU conflict for "${baseSku}" → assigned new SKU "${newSku}" for handle "${handle}" size "${sizeKey || ''}"`);
  return newSku;
}

function processFile(inputPath) {
  console.log(`\n📂 Reading: ${inputPath}`);
  if (!fs.existsSync(inputPath)) {
    console.error(`❌ File not found: ${inputPath}`);
    process.exit(1);
  }

  const workbook = XLSX.readFile(inputPath);
  const collectionLookup = loadCollectionLookup(COLLECTION_FILE);
  const sheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('product')) || workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

  const headers = rawData[0];
  const rows    = rawData.slice(1);

  // Build column index map
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });

  // Identify all product_color metafield column indices
  const colorMetaCols = headers
    .map((h, i) => [h, i])
    .filter(([h]) => h && /my_fields\.product_color\d+/.test(h))
    .map(([, i]) => i);

  const CTL_COL       = col['Metafield: my_fields.complete_the_look [single_line_text_field]'];
  const HANDLE_COL    = col['Handle'];
  const TITLE_COL     = col['Title'];
  const BODY_COL      = col['Body HTML'];
  const TYPE_COL      = col['Type'];
  const TAGS_COL      = col['Tags'];
  const STATUS_COL    = col['Status'];
  const PUBLISHED_COL = col['Published'];
  const IMAGE_SRC_COL = col['Image Src'];
  const IMAGE_POS_COL = col['Image Position'];
  const OPT1_NAME_COL = col['Option1 Name'];
  const OPT1_VAL_COL  = col['Option1 Value'];
  const OPT2_NAME_COL = col['Option2 Name'];
  const OPT2_VAL_COL  = col['Option2 Value'];
  const VAR_SKU_COL   = col['Variant SKU'];
  const VAR_PRICE_COL = col['Variant Price'];
  const VAR_CMP_COL   = col['Variant Compare At Price'];
  const VAR_QTY_COL   = col['Variant Inventory Qty'];

  // ── Group all rows by Handle ──────────────────────────────────────────────
  const productMap = new Map(); // handle → rows[]
  for (const row of rows) {
    const handle = row[HANDLE_COL];
    if (!handle) continue;
    if (!productMap.has(handle)) productMap.set(handle, []);
    productMap.get(handle).push(row);
  }

  // ── Stats ─────────────────────────────────────────────────────────────────
  let totalProducts    = 0;
  let totalVariations  = 0;
  const warnings       = [];
  const outputRows     = [CSV_HEADERS.join(',')];

  // Track SKU usage to ensure uniqueness across all rows.
  const seenSkus       = new Map();     // baseSku or generatedSku → { ownerKey, count }
  const skuConflicts   = new Map();     // baseSku → Set<ownerKey>

  // ── Non-sellable product filters ─────────────────────────────────────────
  const SKIP_HANDLES  = new Set(['bundle-shirt-hoodie',
    'additional-fees', 'bryan-product-1', 'bryan-product-2']);
  const SKIP_PREFIXES = ['test', 'copy-of', 'duties-for-'];

  // ── Process each product ──────────────────────────────────────────────────
  for (const [handle, productRows] of productMap) {

    if (SKIP_HANDLES.has(handle) || SKIP_PREFIXES.some(p => handle.startsWith(p))) {
      warnings.push(`SKIPPED (non-sellable): ${handle}`);
      continue;
    }

    // The "top row" carries product-level data (Title, Body HTML, Status, etc.)
    // Matrixify marks the first row of each product with Top Row = true.
    const topRow = productRows.find(r => {
      const v = r[col['Top Row']];
      return v === true || v === 'true' || v === 'Y' || v === 'y' || v === 1;
    }) || productRows[0];

    const giftCard    = isGiftCardProduct(topRow, handle, TYPE_COL);
    const title       = repairUtf8MojibakeIfPresent(flattenText(topRow[TITLE_COL]));
    const bodyHTML    = repairUtf8MojibakeIfPresent(String(topRow[BODY_COL] || ''));
    const tags        = repairUtf8MojibakeIfPresent(flattenText(topRow[TAGS_COL]));
    const status    = (topRow[STATUS_COL] || '').toLowerCase();
    const published = status === 'active' ? '1' : '0';

    // ── Images (collect across all rows, sort by position) ────────────────
    const imageMap = new Map(); // position → src
    for (const row of productRows) {
      const src = row[IMAGE_SRC_COL];
      const pos = row[IMAGE_POS_COL];
      if (src && pos !== null && pos !== undefined) {
        // Keep only the first image per position slot
        if (!imageMap.has(pos)) imageMap.set(pos, src);
      }
    }
    const images = [...imageMap.entries()]
      .sort((a, b) => Number(a[0]) - Number(b[0]))
      .map(([, src]) => src)
      .join(',');

    // ── Collect variants (apparel sizes and/or gift-card denominations) ───
    // giftDenom: Shopify denomination label (Option1 Value), not used for apparel.
    const sizeVariants = []; // { size, giftDenom, sku, price, comparePrice, qty }
    const seenSizes      = new Set();

    for (const row of productRows) {
      const opt1name = (row[OPT1_NAME_COL] || '').toLowerCase();
      const opt1val  = row[OPT1_VAL_COL];
      const sku      = row[VAR_SKU_COL];

      if (!sku) continue; // image-only row

      const price        = row[VAR_PRICE_COL] !== null ? Number(row[VAR_PRICE_COL]) : null;
      const comparePrice = row[VAR_CMP_COL]  !== null ? Number(row[VAR_CMP_COL])  : null;
      const qty          = row[VAR_QTY_COL]  !== null ? Number(row[VAR_QTY_COL])  : 0;

      if (giftCard) {
        let denom = String(opt1val || '').trim();
        if (!denom || /^default title$/i.test(denom)) denom = '';
        const giftDenom =
          denom ||
          (price !== null && !Number.isNaN(price) ? String(price) : String(sku));
        const vKey = `gift:${sku}`;
        if (seenSizes.has(vKey)) continue;
        seenSizes.add(vKey);
        sizeVariants.push({ size: '', giftDenom, sku, price, comparePrice, qty });
        continue;
      }

      const isSize = opt1name === 'size';
      // Normalise and validate — skips junk values like "Default Title", "blue", gift amounts
      const rawSize = isSize ? String(opt1val || '').trim() : '';
      const size    = (rawSize && isValidSize(rawSize)) ? normaliseSizeLabel(rawSize) : '';

      const sizeKey = size || sku;
      if (seenSizes.has(sizeKey)) continue;
      seenSizes.add(sizeKey);

      sizeVariants.push({ size, giftDenom: '', sku, price, comparePrice, qty });
    }

    const hasSize =
      !giftCard && sizeVariants.some(v => v.size !== '');
    const useGiftVariations = giftCard && sizeVariants.length > 1;
    const isVariable        = hasSize || useGiftVariations;
    const productType       = isVariable ? 'variable' : 'simple';
    const totalStock    = sizeVariants.reduce((s, v) => s + (v.qty || 0), 0);

    // ── Pricing at product level ──────────────────────────────────────────
    const firstVariant  = sizeVariants[0] || {};
    let productRegular  = '';
    let productSale     = '';

    if (firstVariant.comparePrice && firstVariant.comparePrice > firstVariant.price) {
      // Compare-at is higher → regular = compare, sale = current price
      productRegular = String(firstVariant.comparePrice);
      productSale    = String(firstVariant.price);
    } else if (firstVariant.price !== null && firstVariant.price !== undefined) {
      productRegular = String(firstVariant.price);
    }

    // ── Variation attribute string (pa_size or pa_denomination for gift cards) ─
    // Keep Excel-safe wrapper ="1-2" etc. so opening in Excel does not corrupt
    // kids sizes; your post-import snippet in WooCommerce will clean this.
    const attr1Name = useGiftVariations ? 'pa_denomination' : hasSize ? 'pa_size' : '';
    const attr1Value = useGiftVariations
      ? sizeVariants.map(v => safeSizeValue(v.giftDenom)).join('|')
      : sizeVariants.filter(v => v.size).map(v => safeSizeValue(v.size)).join('|');

    // ── Colour / product_group ────────────────────────────────────────────
    const productGroup = deriveProductGroup(handle);

    // Collect all color metafield values from the top row
    const metaVals = colorMetaCols.map(i => topRow[i]).filter(Boolean);

    // Priority: Option2 Colour (Shopify admin) > handle suffix > metafield
    const opt2name = (topRow[OPT2_NAME_COL] || '').toLowerCase();
    const opt2val  = (opt2name === 'colour' || opt2name === 'color') ? topRow[OPT2_VAL_COL] : null;
    const colourFromOption = resolveColourFromOptionValue(opt2val);
    const colourFromHandle = deriveColourFromHandle(handle);
    const colourFromMeta   = extractColourFromMetafields(handle, metaVals);
    const colour           = colourFromOption || colourFromHandle || colourFromMeta;

    const colorName = colour ? (colour.name || '') : '';
    const colorHex  = colour ? (colour.hex  || '') : '';

    if (!colour && !giftCard) {
      warnings.push(`No colour resolved for: ${handle}`);
    }

    // ── Complete the look ─────────────────────────────────────────────────
    const ctlRaw = CTL_COL !== undefined ? (topRow[CTL_COL] || '') : '';
    const completeLook = repairUtf8MojibakeIfPresent(flattenText(ctlRaw));

    // ── Categories (from collection-export.xlsx lookup) ─────────────────
    const categories = buildCategories(handle, collectionLookup);

    // ── Extract maincat / subcat from tags ─────────────────────────────────
    const tagList = (tags || '').split(',').map(t => t.trim().toLowerCase());
    const maincatValues = tagList
      .filter(t => t.startsWith('maincat:'))
      .map(t => t.slice('maincat:'.length).trim())
      .filter(Boolean);
    const subcatValues = tagList
      .filter(t => t.startsWith('subcat:'))
      .map(t => t.slice('subcat:'.length).trim())
      .filter(Boolean);

    // ── Clean tags (remove SPO internals) ─────────────────────────────────
    const cleanTags = (tags || '')
      .split(',')
      .map(t => t.trim())
      .filter(t => !t.startsWith('spo-') && !t.startsWith('group:') && !t.startsWith('maincat:')
               && !t.startsWith('subcat:') && !t.startsWith('model') && t !== '')
      .join(', ');

    // ──────────────────────────────────────────────────────────────────────
    // Build the PRODUCT row
    // ──────────────────────────────────────────────────────────────────────
    const productSKU = handle; // use handle as unique product identifier

    const product = buildEmptyProduct();
    product['Type']               = productType;
    product['SKU']                = productSKU;
    product['Name']               = title;
    product['Published']          = published;
    product['Description']        = cleanHTML(bodyHTML);
    product['Tax status']         = 'taxable';
    product['In stock?']          = totalStock > 0 ? 1 : 0;
    product['Stock']              = totalStock;
    product['Manage stock?']      = 1;
    product['Backorders allowed?']= 0;
    product['Sale price']         = productSale;
    product['Regular price']      = productRegular;
    product['Categories']         = categories;
    product['Tags']               = cleanTags;
    product['Images']             = images;
    product['Attribute 1 name']   = attr1Name;
    product['Attribute 1 value(s)'] = attr1Name ? attr1Value : '';
    product['Attribute 1 visible'] = attr1Name ? 1 : '';
    product['Attribute 1 global']  = attr1Name ? 1 : '';
    product['Meta: product_group'] = productGroup;
    product['Meta: color_name']    = colorName;
    product['Meta: color_hex']     = colorHex;
    product['Meta: complete_the_look'] = completeLook;
    product['Meta: maincat']           = maincatValues.join('|');
    product['Meta: subcat']            = subcatValues.join('|');

    outputRows.push(csvRow(CSV_HEADERS.map(h => product[h])));
    totalProducts++;

    // ──────────────────────────────────────────────────────────────────────
    // Build VARIATION rows (only for variable products)
    // ──────────────────────────────────────────────────────────────────────
    if (isVariable) {
      for (const variant of sizeVariants) {
        const attrVal = useGiftVariations ? variant.giftDenom : variant.size;
        if (useGiftVariations && !variant.giftDenom) continue;

        let varRegular = '';
        let varSale    = '';

        if (variant.comparePrice && variant.comparePrice > variant.price) {
          varRegular = String(variant.comparePrice);
          varSale    = String(variant.price);
        } else if (variant.price !== null && variant.price !== undefined) {
          varRegular = String(variant.price);
        }

        const variation = buildEmptyProduct();
        variation['Type']             = 'variation';
        const baseVarSku = variant.sku ||
          `${productSKU}-${useGiftVariations ? variant.giftDenom : variant.size}`;
        const uniqKey = useGiftVariations ? variant.giftDenom : (variant.size || '');
        const uniqueVarSku            = makeUniqueSku(baseVarSku, handle, uniqKey, seenSkus, skuConflicts, warnings);
        if (!uniqueVarSku) {
          continue;
        }
        variation['SKU']              = uniqueVarSku;
        variation['Name']             = '';
        variation['Published']        = published;
        variation['In stock?']        = variant.qty > 0 ? 1 : 0;
        variation['Stock']            = variant.qty;
        variation['Manage stock?']    = 1;
        variation['Sale price']       = varSale;
        variation['Regular price']    = varRegular;
        variation['Parent']           = productSKU;
        variation['Attribute 1 name'] = attr1Name;
        variation['Attribute 1 value(s)'] = safeSizeValue(attrVal);

        outputRows.push(csvRow(CSV_HEADERS.map(h => variation[h])));
        totalVariations++;
      }
    }
  }

  return { outputRows, totalProducts, totalVariations, warnings, inputRowCount: rows.length, skuConflicts };
}

// ─── Dry-run preview ──────────────────────────────────────────────────────────
function printPreview(outputRows) {
  console.log('\n📋 DRY-RUN PREVIEW (first 3 products + their variations):\n');
  const headerLine = outputRows[0];
  const headers    = headerLine.split(',');
  const pick       = ['Type','SKU','Name','Regular price','Sale price','Stock',
                      'Categories','Attribute 1 value(s)','Parent',
                      'Meta: product_group','Meta: color_name','Meta: color_hex'];

  // Find first few product + variation blocks
  let productCount  = 0;
  let i             = 1;
  while (i < outputRows.length && productCount < 3) {
    // Parse the CSV row properly
    const fields = parseCSVLine(outputRows[i]);
    const get    = (colName) => {
      const idx = headers.indexOf(colName);
      return idx >= 0 ? (fields[idx] || '') : '';
    };

    if (get('Type') !== 'variation') {
      productCount++;
      console.log(`── Product #${productCount} ─────────────────────────────────`);
    } else {
      console.log(`   └─ Variation`);
    }

    for (const col of pick) {
      const v = get(col);
      if (v) console.log(`   ${col.padEnd(28)}: ${v.slice(0, 80)}`);
    }
    console.log('');
    i++;
  }
}

/** Minimal CSV line parser (handles quoted fields). */
function parseCSVLine(line) {
  const result = [];
  let current  = '';
  let inQuote  = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuote) {
      if (ch === '"') {
        if (line[i + 1] === '"') { current += '"'; i++; }
        else inQuote = false;
      } else {
        current += ch;
      }
    } else {
      if (ch === '"')      { inQuote = true; }
      else if (ch === ',') { result.push(current); current = ''; }
      else                 { current += ch; }
    }
  }
  result.push(current);
  return result;
}

// ─── Entry point ─────────────────────────────────────────────────────────────

const { outputRows, totalProducts, totalVariations, warnings, inputRowCount, skuConflicts } = processFile(INPUT_FILE);

if (DRY_RUN) {
  printPreview(outputRows);
} else {
  fs.writeFileSync(OUTPUT_FILE, '\uFEFF' + outputRows.join('\n'), 'utf8');
  console.log(`\n✅ Written: ${OUTPUT_FILE}`);
}

// Summary
console.log('\n═══════════════════════════════════');
console.log('  CONVERSION SUMMARY');
console.log('═══════════════════════════════════');
console.log(`  Input rows         : ${inputRowCount}`);
console.log(`  Products processed : ${totalProducts}`);
console.log(`  Variation rows     : ${totalVariations}`);
console.log(`  Total CSV rows     : ${totalProducts + totalVariations + 1} (incl. header)`);
console.log(`  Warnings           : ${warnings.length}`);
if (warnings.length > 0) {
  console.log('\n  WARNINGS:');
  warnings.slice(0, 20).forEach(w => console.log('  ⚠️  ' + w));
  if (warnings.length > 20) console.log(`  ... and ${warnings.length - 20} more`);
}
console.log('═══════════════════════════════════\n');

// SKU conflict report written to file for later inspection.
if (skuConflicts && skuConflicts.size > 0) {
  const lines = [];
  lines.push('Base SKU,Assigned To Handles/Sizes,Total Variants');
  for (const [baseSku, owners] of skuConflicts.entries()) {
    const ownerList = [...owners].map(k => {
      const [, handle, size] = k.split('@@');
      return `${handle}${size ? ` (size ${size})` : ''}`;
    });
    lines.push(`${baseSku},"${ownerList.join('; ')}",${ownerList.length}`);
  }
  try {
    fs.writeFileSync(SKU_LOG_FILE, lines.join('\n'), 'utf8');
    console.log(`SKU conflict report written to: ${SKU_LOG_FILE}`);
  } catch (err) {
    console.error('Failed to write SKU conflict report:', err.message || err);
  }
}
