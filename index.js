const fs = require('fs');
const path = require('path');
const readline = require('readline');
const XLSX = require('xlsx');
const qrcode = require('qrcode-terminal');
const { Client, LocalAuth } = require('whatsapp-web.js');
const { LoadUtils } = require('whatsapp-web.js/src/util/Injected/Utils');

function prompt(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(question, (ans) => { rl.close(); resolve(ans); }));
}

async function promptRequired(label, current) {
  const hint = current ? ` (Enter = "${current}")` : '';
  while (true) {
    const ans = (await prompt(`${label}${hint}: `)).trim();
    if (ans) return ans;
    if (current) return current;
    console.log('Value required — please type something.');
  }
}

const CONFIG_PATH = path.join(__dirname, 'config.json');
const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
const args = process.argv.slice(2);
const LIST_ONLY = args.includes('--list-groups');
const SETUP_ONLY = args.includes('--setup');
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const STATUS_COLUMN = 'WhatsApp Add Status';
const MIN_PHONE_DIGITS = 8;
const MAX_E164_DIGITS = 15;
const COUNTRY_CALLING_CODES = {
  malaysia: '60',
  my: '60',
  singapore: '65',
  sg: '65',
  china: '86',
  cn: '86',
  usa: '1',
  us: '1',
  america: '1',
  uk: '44',
  gb: '44',
  unitedkingdom: '44',
  england: '44',
  thailand: '66',
  th: '66',
  hongkong: '852',
  hk: '852',
  taiwan: '886',
  tw: '886',
  vietnam: '84',
  vn: '84',
};

function printBanner() {
  console.log(`
============================================================
 __        ______   ____            _     _
 \\ \\      / / ___| / ___|  __ _  __| | __| | ___ _ __
  \\ \\ /\\ / /\\___ \\| |  _  / _\` |/ _\` |/ _\` |/ _ \\ '__|
   \\ V  V /  ___) | |_| | (_| | (_| | (_| |  __/ |
    \\_/\\_/  |____/ \\____|\\__,_|\\__,_|\\__,_|\\___|_|

              WSadder - by Lay Yang
                 www.dkly.net
        https://github.com/osscv/WSadder
============================================================
`);
}

async function pickFromList(label, items, suggestIdx, formatter = (x) => x) {
  console.log(`\n${label}:`);
  items.forEach((it, i) => console.log(`  ${String(i + 1).padStart(2)}. ${formatter(it)}`));
  const hint = suggestIdx >= 0 ? ` (Enter = ${suggestIdx + 1})` : '';
  while (true) {
    const ans = (await prompt(`Pick [1-${items.length}]${hint}: `)).trim();
    if (ans === '' && suggestIdx >= 0) return items[suggestIdx];
    const n = Number(ans);
    if (Number.isInteger(n) && n >= 1 && n <= items.length) return items[n - 1];
    console.log('Invalid choice, try again.');
  }
}

function suggestIndex(headers, patterns) {
  for (const p of patterns) {
    const i = headers.findIndex((h) => p.test(h));
    if (i >= 0) return i;
  }
  return -1;
}

function excelPath(file) {
  return path.isAbsolute(file) ? file : path.join(__dirname, file);
}

function resolveCountryCode(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const digits = raw.replace(/\D/g, '');
  if (digits) return digits.replace(/^00/, '');
  return COUNTRY_CALLING_CODES[raw.toLowerCase().replace(/[^a-z]/g, '')] || '';
}

async function configureSourceFileAndColumns({ save = true } = {}) {
  const here = fs.readdirSync(__dirname).filter((f) => f.toLowerCase().endsWith('.xlsx'));
  let chosenFile;
  if (here.length === 0) {
    chosenFile = (await prompt('No .xlsx files found in this folder. Enter path to Excel file: ')).trim();
  } else {
    const currentIdx = cfg.excelFile ? here.indexOf(cfg.excelFile) : -1;
    chosenFile = await pickFromList('Excel files in this folder', here, currentIdx >= 0 ? currentIdx : 0);
  }
  if (!chosenFile) { console.error('No Excel file chosen.'); process.exit(1); }
  cfg.excelFile = chosenFile;

  const wb = XLSX.readFile(excelPath(chosenFile));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { defval: null, header: 1 });
  const headers = (grid[0] || []).map((h) => (h == null ? '' : String(h)));
  if (headers.length === 0) { console.error('No header row found in first sheet.'); process.exit(1); }

  const phoneIdx = suggestIndex(headers, [/whats/i, /phone|mobile/i]);
  cfg.phoneColumn = await pickFromList('Phone column', headers, phoneIdx);

  const nameIdx = suggestIndex(headers, [/full.*name.*(ic|passport)/i, /full.*name/i, /^name\b/i]);
  cfg.nameColumn = await pickFromList('Name column', headers, nameIdx);

  const ccDefault = cfg.defaultCountryCode || '60';
  const ccAns = (await prompt(`\nDefault country code or country name, e.g. 60, SG, Singapore, USA (Enter = ${ccDefault}): `)).trim();
  cfg.defaultCountryCode = resolveCountryCode(ccAns || ccDefault) || ccDefault;

  if (save) {
    saveConfig();
    console.log('\nSource file info saved:');
    console.log(`  excelFile:          ${cfg.excelFile}`);
    console.log(`  phoneColumn:        ${cfg.phoneColumn}`);
    console.log(`  nameColumn:         ${cfg.nameColumn}`);
    console.log(`  defaultCountryCode: ${cfg.defaultCountryCode}`);
  }
}

async function editEventName() {
  cfg.eventName = await promptRequired('\nEvent name', cfg.eventName);
  saveConfig();
  console.log(`Event name saved: ${cfg.eventName}`);
}

async function editEventDate() {
  cfg.eventDate = await promptRequired('\nEvent date (e.g. 4 June 2026)', cfg.eventDate);
  saveConfig();
  console.log(`Event date saved: ${cfg.eventDate}`);
}

async function runSetup() {
  console.log('=== Setup ===');

  // 1. Excel file
  const here = fs.readdirSync(__dirname).filter((f) => f.toLowerCase().endsWith('.xlsx'));
  let chosenFile;
  if (here.length === 0) {
    chosenFile = (await prompt('No .xlsx files found in this folder. Enter path to Excel file: ')).trim();
  } else {
    const currentIdx = cfg.excelFile ? here.indexOf(cfg.excelFile) : -1;
    chosenFile = await pickFromList('Excel files in this folder', here, currentIdx >= 0 ? currentIdx : 0);
  }
  if (!chosenFile) { console.error('No Excel file chosen.'); process.exit(1); }
  cfg.excelFile = chosenFile;

  // 2. Read column headers from first sheet
  const wb = XLSX.readFile(path.join(__dirname, chosenFile));
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { defval: null, header: 1 });
  const headers = (grid[0] || []).map((h) => (h == null ? '' : String(h)));
  if (headers.length === 0) { console.error('No header row found in first sheet.'); process.exit(1); }

  // 3. Phone column — prefer "WhatsApp", then "Phone/Mobile"
  const phoneIdx = suggestIndex(headers, [/whats/i, /phone|mobile/i]);
  cfg.phoneColumn = await pickFromList('Phone column', headers, phoneIdx);

  // 4. Name column — prefer "Full name (IC/Passport)", then "Full name", then "Name"
  const nameIdx = suggestIndex(headers, [/full.*name.*(ic|passport)/i, /full.*name/i, /^name\b/i]);
  cfg.nameColumn = await pickFromList('Name column', headers, nameIdx);

  // 5. Default country code
  const ccDefault = cfg.defaultCountryCode || '60';
  const ccAns = (await prompt(`\nDefault country code or country name, e.g. 60, SG, Singapore, USA (Enter = ${ccDefault}): `)).trim();
  cfg.defaultCountryCode = resolveCountryCode(ccAns || ccDefault) || ccDefault;

  // 6. Event name (required — no hardcoded default)
  cfg.eventName = await promptRequired('\nEvent name', cfg.eventName);

  // 7. Event date (required — no hardcoded default)
  cfg.eventDate = await promptRequired('Event date (e.g. 4 June 2026)', cfg.eventDate);

  saveConfig();
  console.log('\nConfig saved:');
  console.log(`  excelFile:          ${cfg.excelFile}`);
  console.log(`  phoneColumn:        ${cfg.phoneColumn}`);
  console.log(`  nameColumn:         ${cfg.nameColumn}`);
  console.log(`  defaultCountryCode: ${cfg.defaultCountryCode}`);
  console.log(`  eventName:          ${cfg.eventName}`);
  console.log(`  eventDate:          ${cfg.eventDate}`);
}

function normalizePhone(raw, cc = '60') {
  if (raw == null) return null;
  const s = String(raw).trim();
  const digits = s.replace(/\D/g, '');
  if (!digits) return null;
  // Already international (had a leading +) — trust it as-is.
  if (s.startsWith('+')) return digits;
  // Local form starting with 0 — replace 0 with default country code.
  if (digits.startsWith('0')) return cc + digits.slice(1);
  // Already has the default country code prefix.
  if (digits.startsWith(cc)) return digits;
  // Bare local number with no leading 0 — assume default country.
  return cc + digits;
}

function normalizePhoneForWhatsApp(raw, cc = '60') {
  if (raw == null) return null;
  const s = String(raw).trim();
  const defaultCountryCode = resolveCountryCode(cc);
  if (!s) return null;

  // Masked values such as "+6011-2373-XXX" are incomplete and cannot be used.
  if (/[A-Za-z]/.test(s)) return null;

  const digits = s.replace(/\D/g, '');
  if (!digits) return null;

  let normalized;
  if (/^\s*\+/.test(s)) {
    // International format: +60-1162383838, +1 (415) 555-0101, etc.
    normalized = digits;
  } else if (digits.startsWith('00') && digits.length > 2) {
    // Common international dialing prefix: 00601162383838 -> 601162383838.
    normalized = digits.slice(2);
  } else if (digits.startsWith('0')) {
    if (!defaultCountryCode) return null;
    // Local format: 01162383838 -> 601162383838 when defaultCountryCode is 60.
    normalized = defaultCountryCode + digits.slice(1);
  } else if (defaultCountryCode && digits.startsWith(defaultCountryCode)) {
    normalized = digits;
  } else {
    if (!defaultCountryCode) return null;
    // Bare local number with no leading 0.
    normalized = defaultCountryCode + digits;
  }

  if (normalized.length < MIN_PHONE_DIGITS || normalized.length > MAX_E164_DIGITS) return null;
  if (!/^[1-9]\d+$/.test(normalized)) return null;
  return normalized;
}

function ensureStatusColumn(ws) {
  const grid = XLSX.utils.sheet_to_json(ws, { defval: null, header: 1 });
  const headers = (grid[0] || []).map((h) => (h == null ? '' : String(h)));
  let statusCol = headers.findIndex((h) => h === STATUS_COLUMN);
  if (statusCol < 0) {
    statusCol = headers.length;
    XLSX.utils.sheet_add_aoa(ws, [[STATUS_COLUMN]], {
      origin: XLSX.utils.encode_cell({ r: 0, c: statusCol }),
    });
  }
  return statusCol;
}

function setExcelStatus(run, rowIndex, status) {
  if (!run || rowIndex == null) return;
  XLSX.utils.sheet_add_aoa(run.ws, [[status]], {
    origin: XLSX.utils.encode_cell({ r: rowIndex, c: run.statusCol }),
  });
}

function saveExcelRun(run) {
  XLSX.writeFile(run.wb, run.filePath);
}

function loadUsers() {
  const filePath = excelPath(cfg.excelFile);
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const statusCol = ensureStatusColumn(ws);
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  const seen = new Set();
  const users = [];
  const initialStatuses = [];
  for (const [idx, r] of rows.entries()) {
    const rowIndex = idx + 1; // zero-based worksheet row; data starts after header row
    const phone = normalizePhoneForWhatsApp(r[cfg.phoneColumn], cfg.defaultCountryCode);
    if (!phone) {
      initialStatuses.push({ rowIndex, status: 'missing phone - skipped' });
      continue;
    }
    if (seen.has(phone)) {
      initialStatuses.push({ rowIndex, status: 'duplicate phone - skipped' });
      continue;
    }
    seen.add(phone);
    const name = (r[cfg.nameColumn] ?? '').toString().trim();
    users.push({ name, phone, rowIndex });
  }
  return { wb, ws, sheetName, filePath, statusCol, users, initialStatuses };
}

function saveConfig() {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2));
}

let client = null;
let sessionLoss = null;

function throwIfSessionLost() {
  if (sessionLoss) throw sessionLoss;
}

class StartupAuth extends LocalAuth {
  async afterBrowserInitialized() {
    await super.afterBrowserInitialized();
    this.client.pupPage.on('requestfailed', (request) => {
      if (request.failure()?.errorText === 'net::ERR_NETWORK_CHANGED') {
        this.networkChanged = true;
      }
    });
  }
}

function createClient() {
  return new Client({
    authStrategy: new StartupAuth(),
    takeoverOnConflict: true,
    // Fail a hung page load after 60s (default 180s) so initClient can retry it.
    puppeteer: { headless: true, protocolTimeout: 60000, args: ['--no-sandbox', '--disable-setuid-sandbox'] },
    // Use the live WhatsApp Web page; the archived HTML may be missing or stale.
    webVersionCache: { type: 'none' },
  });
}

async function destroyClient(instance) {
  if (!instance) return;
  try {
    await instance.destroy();
  } catch (_) {
    // Failed launches may not have a fully initialized browser to destroy.
  }
}

async function withTimeout(promise, timeoutMs, label) {
  let timer;
  let progressTimer;
  const startedAt = Date.now();
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${Math.round(timeoutMs / 1000)}s`)), timeoutMs);
  });
  progressTimer = setInterval(() => {
    const elapsed = Date.now() - startedAt;
    const pct = Math.min(99, Math.round((elapsed / timeoutMs) * 100));
    process.stdout.write(`${label}... ${pct}% (${Math.round(elapsed / 1000)}s/${Math.round(timeoutMs / 1000)}s)\r`);
  }, 1000);
  try {
    const result = await Promise.race([promise, timeout]);
    process.stdout.write(`${label}... 100%\n`);
    return result;
  } finally {
    clearTimeout(timer);
    clearInterval(progressTimer);
  }
}

function unusedRegisterClientEvents(instance) {
  instance.on('qr', (qr) => {
    console.log('\nOpen WhatsApp on your phone â†’ Settings â†’ Linked devices â†’ Link a device, then scan:\n');
    qrcode.generate(qr, { small: true });
  });

  instance.on('loading_screen', (percent, message) => {
    process.stdout.write(`Loading WhatsApp Web... ${percent || 0}% ${message || ''}\r`);
  });

  instance.on('auth_failure', (m) => console.error('Auth failure:', m));
  instance.on('disconnected', (r) => console.log('Disconnected:', r));

  instance.on('authenticated', async () => {
    if (fastPathStarted || alreadyProceeded) return;
    fastPathStarted = true;

    console.log('\nAuthenticated. Fetching groups (fast path)...');

    // Give the page a few seconds to initialise, then start polling Store.
    await new Promise((r) => setTimeout(r, 8000));
    if (alreadyProceeded) return;

    const groups = await waitForGroups(50000); // 50 s max
    if (groups && groups.length > 0 && !alreadyProceeded) {
      alreadyProceeded = true;
      await proceedWithGroups(groups.map(wrapGroup));
      return;
    }

    if (!alreadyProceeded) {
      console.log('\nFast path didn\'t find groups â€” waiting for full sync...');
    }
  });

  instance.on('ready', async () => {
    if (alreadyProceeded) return;
    console.log('\nWhatsApp Web fully synced.');
    try {
      const chats = await withTimeout(client.getChats(), 20000, 'Loading chats');
      if (alreadyProceeded) return;
      alreadyProceeded = true;
      await proceedWithGroups(chats.filter((c) => c.isGroup));
    } catch (e) {
      if (alreadyProceeded) return;
      console.log(`\n${e.message}. Trying fast group extraction again...`);
      const groups = await waitForGroups(30000);
      if (groups && groups.length > 0 && !alreadyProceeded) {
        alreadyProceeded = true;
        await proceedWithGroups(groups.map(wrapGroup));
        return;
      }
      throw e;
    }
  });
}

function registerClientEvents(instance) {
  let authenticated = false;
  const canProceed = () => instance === client && !alreadyProceeded;
  const loseSession = (reason) => {
    sessionLoss ??= new Error(`WhatsApp session lost (${reason}). Restart WSadder and log in again if prompted.`);
  };

  async function loadGroupsAndProceed(timeoutMs, options) {
    if (!canProceed()) return;
    let ownsWorkflow = false;
    try {
      const groups = await waitForGroups(timeoutMs, options, instance);
      if (!canProceed()) return;
      alreadyProceeded = true;
      ownsWorkflow = true;
      throwIfSessionLost();
      await proceedWithGroups((groups || []).map(wrapGroup));
    } catch (e) {
      if (!ownsWorkflow && !canProceed()) return;
      alreadyProceeded = true;
      stopFullSyncProgress();
      console.error(`\n${(sessionLoss || e).message}`);
      await destroyClient(instance);
      process.exit(1);
    }
  }

  instance.on('qr', (qr) => {
    if (instance !== client || sessionLoss) return;
    if (authenticated) {
      loseSession('WhatsApp requested a new login');
      return;
    }
    console.log('\nOpen WhatsApp on your phone -> Settings -> Linked devices -> Link a device, then scan:\n');
    qrcode.generate(qr, { small: true });
  });

  instance.on('loading_screen', (percent, message) => {
    process.stdout.write(`Loading WhatsApp Web... ${percent || 0}% ${message || ''}\r`);
  });

  instance.on('auth_failure', (m) => console.error('Auth failure:', m));
  instance.on('disconnected', (reason) => {
    if (instance !== client) return;
    console.log('Disconnected:', reason);
    if (authenticated) loseSession(reason);
  });

  instance.on('authenticated', async () => {
    if (!canProceed()) return;
    authenticated = true;
    if (fastPathStarted) return;
    fastPathStarted = true;

    console.log('\nAuthenticated. Fetching groups directly from WhatsApp Web...');

    await sleep(8000);
    await loadGroupsAndProceed(180000, {
      minWaitMs: 12000,
      stableMs: 9000,
    });
  });

  instance.on('ready', async () => {
    if (!canProceed()) return;
    authenticated = true;
    stopFullSyncProgress(true);
    console.log('\nWhatsApp Web fully synced.');
    await loadGroupsAndProceed(30000, {
      minWaitMs: 3000,
      stableMs: 3000,
    });
  });
}

function isAuthStartupTimeout(message) {
  return /^auth timeout$|^Waiting failed: \d+ms exceeded$/i.test(message);
}

async function initClient(retries = 3) {
  console.log('Starting WhatsApp Web (may retry if startup times out)...');
  // Chrome aborts in-flight loads when any interface address changes (e.g. a
  // flapping VPN adapter under WSL mirrored networking). These are transient,
  // so allow extra quick retries for them without consuming the normal budget.
  let networkChangedRetries = 10;
  let sawNetworkChange = false;
  for (let i = 0; i < retries; i++) {
    sessionLoss = null;
    client = createClient();
    registerClientEvents(client);
    try {
      await client.initialize();
      return;
    } catch (e) {
      if (sessionLoss) return;
      const msg = e?.message || String(e);
      const navigationInterrupted = /net::ERR_NETWORK_CHANGED/.test(msg);
      sawNetworkChange ||= navigationInterrupted || !!client.authStrategy.networkChanged;
      stopFullSyncProgress();
      const failedClient = client;
      client = null;
      fastPathStarted = false;
      await destroyClient(failedClient);
      if (navigationInterrupted && networkChangedRetries-- > 0) {
        console.log('\nNetwork changed during load, retrying...');
        i--;
        await new Promise((r) => setTimeout(r, 1000));
        continue;
      }
      const retryable = isAuthStartupTimeout(msg)
        || /already running|destroyed|Target closed|Page closed|timed out|net::ERR_/i.test(msg);
      if (i < retries - 1 && retryable) {
        console.log(`\nWhatsApp Web failed to load (${msg.split('\n')[0]}), retrying (${i + 2}/${retries})...`);
        await new Promise((r) => setTimeout(r, 3000));
        continue;
      }
      if (sawNetworkChange && (isAuthStartupTimeout(msg) || navigationInterrupted)) {
        console.error('Chrome reported a network change while loading WhatsApp Web. Stabilize the network before retrying.');
      }
      throw e;
    }
  }
}

/*
client.on('qr', (qr) => {
  console.log('\nOpen WhatsApp on your phone → Settings → Linked devices → Link a device, then scan:\n');
  qrcode.generate(qr, { small: true });
});

*/

// ---------- Fast group extraction via Store (bypasses full sync) ----------

function extractFromStore() {
  try {
    const groupsById = {};

    function idToString(id) {
      if (!id) return null;
      if (typeof id === 'string') return id;
      if (id._serialized) return id._serialized;
      if (id.user && id.server) return `${id.user}@${id.server}`;
      if (typeof id.toString === 'function') {
        const text = id.toString();
        if (/@g\.us$/.test(text)) return text;
      }
      return null;
    }

    function collectionToArray(value) {
      if (!value) return [];
      if (Array.isArray(value)) return value;
      if (Array.isArray(value.models)) return value.models;
      if (Array.isArray(value._models)) return value._models;
      if (typeof value.getModelsArray === 'function') return value.getModelsArray();
      if (typeof value.getModels === 'function') return value.getModels();
      if (typeof value.serialize === 'function') {
        const serialized = value.serialize();
        return Array.isArray(serialized) ? serialized : [];
      }
      return [];
    }

    function participantListFrom(candidate) {
      const participants =
        candidate.groupMetadata?.participants ||
        candidate.participants ||
        candidate.__x_groupMetadata?.participants ||
        [];
      return collectionToArray(participants)
        .map((p) => {
          const id = idToString(p.id || p.wid || p);
          return id ? { id: { _serialized: id } } : null;
        })
        .filter(Boolean);
    }

    function addGroup(candidate) {
      if (!candidate) return;
      const serializedId = idToString(candidate.id || candidate.wid || candidate.groupId);
      const idLooksLikeGroup = serializedId && /@g\.us$/.test(serializedId);
      const isGroupFlag = typeof candidate.isGroup === 'function' ? candidate.isGroup() : candidate.isGroup;
      const idGroupFlag = typeof candidate.id?.isGroup === 'function' ? candidate.id.isGroup() : false;
      if (!serializedId || !(idLooksLikeGroup || isGroupFlag || idGroupFlag || candidate.groupMetadata)) return;

      const name = String(
        candidate.name ||
        candidate.formattedTitle ||
        candidate.__x_title ||
        candidate.__x_formattedTitle ||
        candidate.subject ||
        candidate.pushname ||
        candidate.groupMetadata?.subject ||
        candidate.groupMetadata?.name ||
        serializedId
      );

      groupsById[serializedId] = {
        name,
        id: { _serialized: serializedId },
        isGroup: true,
        participants: participantListFrom(candidate),
      };
    }

    // Current whatsapp-web.js versions use webpack modules, not window.Store.
    if (typeof window.require === 'function') {
      let collections = null;
      try { collections = window.require('WAWebCollections'); } catch (_) {}
      for (const chat of collectionToArray(collections?.Chat)) addGroup(chat);
      for (const contact of collectionToArray(collections?.Contact)) addGroup(contact);
      for (const meta of collectionToArray(collections?.GroupMetadata || collections?.WAWebGroupMetadataCollection)) {
        addGroup({
          id: meta.id,
          name: meta.subject || meta.name,
          groupMetadata: meta,
          isGroup: true,
        });
      }
    }

    // Legacy fallback for older WhatsApp Web builds.
    const S = window.Store;
    if (S) {
      for (const key of Object.keys(S)) {
        const val = S[key];
        if (!val || typeof val !== 'object') continue;

        const modelSources = [
          val.models,
          val._models,
          typeof val.getModels === 'function' ? val.getModels() : undefined,
          typeof val.getModelsArray === 'function' ? val.getModelsArray() : undefined,
          typeof val.filter === 'function' ? val.filter(() => true) : undefined,
        ];

        for (const arr of modelSources) {
          if (!Array.isArray(arr) || arr.length === 0) continue;
          for (const item of arr) addGroup(item);
        }
      }
    }

    const groups = Object.values(groupsById);
    return groups.length > 0 ? groups : null;
  } catch (_) {}
  return null;
}

async function waitForGroupsOld(timeoutMs) {
  const pollInterval = 3000;
  const deadline = Date.now() + timeoutMs;
  const startedAt = Date.now();

  // First, wait for window.Store to exist (up to 20 s).
  try {
    await client.pupPage.waitForFunction(
      () => typeof window.Store !== 'undefined' && window.Store !== null,
      { polling: 2000, timeout: Math.min(timeoutMs, 20000) },
    );
  } catch (_) { /* Store didn't appear — we'll still try the poll loop below */ }

  while (Date.now() < deadline) {
    try {
      const result = await client.pupPage.evaluate(extractFromStore);
      if (result && result.length > 0) {
        console.log(`\nGroups found: ${result.length}`);
        return result;
      }
    } catch (_) {}

    elapsed += pollInterval;
    process.stdout.write(`  syncing... (${Math.round(elapsed / 1000)}s)\r`);
    await new Promise((r) => setTimeout(r, pollInterval));
  }
  return null;
}

function normalizeGroupList(groups) {
  const byId = new Map();
  for (const group of groups || []) {
    const id = group?.id?._serialized;
    if (!id) continue;
    byId.set(id, {
      ...group,
      name: group.name || id,
      id: { _serialized: id },
      participants: Array.isArray(group.participants) ? group.participants : [],
    });
  }
  return [...byId.values()].sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
}

function groupSignature(groups) {
  return normalizeGroupList(groups)
    .map((g) => `${g.id._serialized}:${g.name}`)
    .join('|');
}

async function waitForGroups(timeoutMs, options = {}, instance = client) {
  const pollInterval = options.pollIntervalMs || 3000;
  const stableMs = options.stableMs ?? 6000;
  const minWaitMs = options.minWaitMs ?? 6000;
  const deadline = Date.now() + timeoutMs;
  const startedAt = Date.now();
  let bestGroups = [];
  let lastSignature = '';
  let stableSince = 0;
  const canKeepPolling = () => instance === client && !alreadyProceeded && !sessionLoss;

  while (canKeepPolling() && Date.now() < deadline) {
    try {
      const result = normalizeGroupList(await instance.pupPage.evaluate(extractFromStore));
      if (result.length > 0 && canKeepPolling()) {
        const signature = groupSignature(result);
        if (signature !== lastSignature) {
          lastSignature = signature;
          stableSince = Date.now();
        }
        if (result.length >= bestGroups.length) bestGroups = result;

        const elapsed = Date.now() - startedAt;
        if (elapsed >= minWaitMs && Date.now() - stableSince >= stableMs) {
          process.stdout.write(`  syncing groups... 100% - found ${bestGroups.length} groups\n`);
          return bestGroups;
        }
      }
    } catch (_) {}
    if (!canKeepPolling()) return null;

    const elapsed = Date.now() - startedAt;
    const pct = Math.min(99, Math.round((elapsed / timeoutMs) * 100));
    const found = bestGroups.length > 0 ? `, found ${bestGroups.length}` : '';
    process.stdout.write(`  syncing groups... ${pct}% (${Math.round(elapsed / 1000)}s/${Math.round(timeoutMs / 1000)}s${found})\r`);
    await sleep(pollInterval);
  }

  if (!canKeepPolling()) return null;

  if (bestGroups.length > 0) {
    process.stdout.write(`  syncing groups... 100% - found ${bestGroups.length} groups\n`);
    return bestGroups;
  }

  process.stdout.write('  syncing groups... 100% - no groups found\n');
  return null;
}

async function ensureWWebJS(timeoutMs = 30000) {
  const hasWWebJS = async () => {
    try {
      return await client.pupPage.evaluate(() => typeof window.WWebJS !== 'undefined');
    } catch (_) {
      return false;
    } finally {
      throwIfSessionLost();
    }
  };

  throwIfSessionLost();
  if (await hasWWebJS()) return;

  throwIfSessionLost();
  try {
    await client.pupPage.evaluate(LoadUtils);
  } catch (_) {
    // If WhatsApp Web is already injecting helpers, the poll below will catch it.
  }

  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    throwIfSessionLost();
    if (await hasWWebJS()) return;
    await sleep(500);
  }

  throwIfSessionLost();
  throw new Error(`WhatsApp helper injection timed out after ${Math.round(timeoutMs / 1000)}s`);
}

async function hydrateGroup(group) {
  const id = group?.id?._serialized;
  if (!id) return group;

  try {
    await ensureWWebJS(30000);
    throwIfSessionLost();
    const fullGroup = await withTimeout(client.getChatById(id), 30000, 'Loading selected group');
    throwIfSessionLost();
    if (fullGroup?.isGroup) return fullGroup;
  } catch (e) {
    throwIfSessionLost();
    console.warn(`Could not load full group details (${e.message}); using the loaded group summary.`);
  }

  return group;
}

function wrapGroup(raw) {
  return {
    get id() { return raw.id; },
    get name() { return raw.name; },
    get isGroup() { return true; },
    get participants() { return raw.participants || []; },
    addParticipants: async (ids) => {
      await ensureWWebJS(30000);
      throwIfSessionLost();
      const fullGroup = await client.getChatById(raw.id._serialized);
      throwIfSessionLost();
      return fullGroup.addParticipants(ids);
    },
  };
}

// whatsapp-web.js's GroupChat.getInviteCode() calls WAWebMexFetchGroupInviteCodeJob, but
// WhatsApp Web now only loads that module lazily with the "Invite via link" drawer, so the
// library call throws. Load the drawer's bundle first, then fetch the code.
async function fetchInviteCode(groupId) {
  await ensureWWebJS(30000);
  throwIfSessionLost();
  const codeRes = await client.pupPage.evaluate(async (chatId) => {
    const MODULE = 'WAWebMexFetchGroupInviteCodeJob';
    const loaded = () => Boolean(window.require('__debug')?.modulesMap?.[MODULE]);
    if (!loaded()) {
      await window.require('JSResourceForInteraction')('WAWebGroupInviteLinkDrawer.react').load();
    }
    try {
      return await window.require(MODULE).fetchMexGroupInviteCode(chatId);
    } catch (err) {
      if (err.name === 'ServerStatusCodeError') return undefined;
      throw err;
    }
  }, groupId);
  const code = codeRes?.code ?? codeRes;
  if (typeof code !== 'string' || !code.trim()) {
    throw new Error('WhatsApp returned no invite code (are you an admin of this group?)');
  }
  return code.trim();
}

// ---------- processUser ----------

async function processUser(u, group, inviteLink, existingIds, log, progress) {
  throwIfSessionLost();
  const tag = progress ? `[${progress.n}/${progress.total}] ` : '';
  const waId = `${u.phone}@c.us`;
  const label = `${tag}${u.name || '(no name)'} <+${u.phone}>`;
  if (progress) console.log(`${tag}→ processing ${u.name || '(no name)'} <+${u.phone}>...`);

  if (existingIds.has(waId)) {
    console.log(`= already in group: ${label}`);
    log?.skipped.push(u);
    return 'already in group';
  }
  if (cfg.dryRun) {
    console.log(`[dry-run] would add: ${label}`);
    return 'dry-run';
  }

  let registered;
  try {
    registered = await client.isRegisteredUser(waId);
  } catch (e) {
    throwIfSessionLost();
    console.error(`! WhatsApp lookup failed for ${label}: ${e.message}`);
    log?.failed.push({ ...u, error: `WhatsApp lookup failed: ${e.message}` });
    return `failed: WhatsApp lookup failed (${e.message})`;
  }
  throwIfSessionLost();
  if (!registered) {
    console.log(`x not on WhatsApp: ${label}`);
    log?.notOnWhatsApp.push(u);
    return 'not on WhatsApp';
  }

  let addedOk = false;
  let statusCode = null;
  try {
    const res = await group.addParticipants([waId]);
    const entry = res && (res[waId] || res);
    statusCode = entry && (entry.code ?? entry.status);
    addedOk = statusCode === 200;
  } catch (e) {
    throwIfSessionLost();
    console.error(`! error adding ${label}: ${e.message}`);
  }

  throwIfSessionLost();
  if (addedOk) {
    console.log(`+ added: ${label}`);
    log?.added.push(u);
    existingIds.add(waId);
    return 'added';
  }
  if (statusCode === 409) {
    console.log(`= already in group: ${label}`);
    log?.skipped.push(u);
    existingIds.add(waId);
    return 'already in group';
  }
  if (!inviteLink && cfg.inviteMessage.includes('{link}')) {
    console.error(`! not added and no invite link to DM (code=${statusCode}): ${label}`);
    log?.failed.push({ ...u, code: statusCode, error: 'no invite link' });
    return statusCode ? `failed (add code=${statusCode}; no invite link to DM)` : 'failed: no invite link to DM';
  }
  try {
    const msg = cfg.inviteMessage
      .replaceAll('{name}', u.name || '')
      .replaceAll('{link}', inviteLink || '')
      .replaceAll('{eventName}', cfg.eventName || '')
      .replaceAll('{eventDate}', cfg.eventDate || '');
    await ensureWWebJS(30000);
    throwIfSessionLost();
    await client.sendMessage(waId, msg);
    throwIfSessionLost();
    console.log(`~ DM invited (couldn't add directly, code=${statusCode}): ${label}`);
    log?.invited.push(u);
    return statusCode ? `invited via DM (add code=${statusCode})` : 'invited via DM';
  } catch (e) {
    throwIfSessionLost();
    console.error(`! DM invite failed for ${label}: ${e.message}`);
    log?.failed.push({ ...u, code: statusCode, error: e.message });
    return statusCode ? `failed (add code=${statusCode}; DM: ${e.message})` : `failed: ${e.message}`;
  }
}

// ---------- interactive prompts ----------

async function chooseNextAction() {
  console.log('\nWhat next?');
  console.log('  1. Start now — add everyone from the Excel');
  console.log('  2. Test first — try one name + phone you type in');
  console.log('  3. Reselect the group');
  console.log('  4. Edit Event Name');
  console.log('  5. Edit Event Date');
  console.log('  6. Edit Source File & Info');
  console.log('  7. Exit');
  while (true) {
    const ans = (await prompt('Pick [1-7]: ')).trim();
    if (ans === '1') return 'start';
    if (ans === '2') return 'test';
    if (ans === '3') return 'reselect-group';
    if (ans === '4') return 'edit-event-name';
    if (ans === '5') return 'edit-event-date';
    if (ans === '6') return 'edit-source-info';
    if (ans === '7' || ans === '') return 'exit';
    console.log('Invalid choice, try again.');
  }
}

async function pickAndSaveGroup(groups, currentGroup = null) {
  if (groups.length === 0) {
    console.log('No groups found on this account.');
    return null;
  }
  console.log('\nYour groups:');
  groups.forEach((g, i) => {
    const currentMarker = currentGroup?.id?._serialized === g.id._serialized ? '*' : ' ';
    console.log(`${currentMarker} ${String(i + 1).padStart(2)}. ${g.name}    [id: ${g.id._serialized}]`);
  });
  if (currentGroup) {
    console.log('\n* current configured group');
  }

  while (true) {
    const keepHint = currentGroup ? ', K = keep current' : '';
    const ans = (await prompt(`\nPick a group [1-${groups.length}]${keepHint}: `)).trim();
    if (ans === '') {
      console.log(currentGroup ? 'Type a group number, or K to keep the current group.' : 'Type a group number.');
      continue;
    }
    if (currentGroup && /^k(eep)?$/i.test(ans)) return null;

    const n = Number(ans);
    if (Number.isInteger(n) && n >= 1 && n <= groups.length) {
      const pick = groups[n - 1];
      cfg.groupName = pick.name;
      cfg.groupId = pick.id._serialized;
      saveConfig();
      console.log(`\nSaved to config.json:\n  groupName: ${cfg.groupName}\n  groupId:   ${cfg.groupId}`);
      return pick;
    }
    console.log('Invalid choice, try again.');
  }
}

// ---------- core logic — runs once groups are available ----------

async function proceedWithGroups(groups) {
  // Resolve target group from config first.
  const interactiveSession = LIST_ONLY || SETUP_ONLY;
  let group = null;
  if (cfg.groupId) {
    group = groups.find((g) => g.id._serialized === cfg.groupId) || null;
  }
  if (!group && cfg.groupName) {
    group = groups.find((g) => g.name === cfg.groupName) || null;
    if (group) {
      cfg.groupId = group.id._serialized;
      saveConfig();
    }
  }
  if (!group && !interactiveSession && cfg.groupId) {
    group = wrapGroup({
      name: cfg.groupName || cfg.groupId,
      id: { _serialized: cfg.groupId },
      participants: [],
    });
  }

  // Interactive pick when --list-groups, --setup, or no group is configured yet.
  const needsPick = interactiveSession || !group;
  if (needsPick) {
    const picked = await pickAndSaveGroup(groups, group);
    throwIfSessionLost();
    if (picked) group = picked;
    if (!group) {
      console.log('No selection — group config unchanged.');
      await client.destroy();
      process.exit(0);
    }
  }

  let action = 'start';
  let inviteLink = null;
  while (true) {
    group = await hydrateGroup(group);
    throwIfSessionLost();
    console.log(`\nTarget group: ${group.name}  [${group.id._serialized}]`);

    inviteLink = null;
    try {
      const code = await fetchInviteCode(group.id._serialized);
      throwIfSessionLost();
      inviteLink = `https://chat.whatsapp.com/${code}`;
      console.log(`Invite link: ${inviteLink}`);
    } catch (e) {
      throwIfSessionLost();
      console.warn(`Could not fetch invite link (${e.message}).`);
      if (cfg.inviteMessage?.includes('{link}')) {
        console.warn('People who cannot be added directly will NOT be DMed, because the message would have no link.');
      }
    }

    // After interactive pick (--setup / --list-groups), ask what to do next.
    if (!interactiveSession) break;

    action = await chooseNextAction();
    throwIfSessionLost();
    if (action === 'exit') {
      console.log('Exiting. Run `npm start` later when you are ready.');
      await client.destroy();
      process.exit(0);
    }

    if (action === 'reselect-group') {
      const picked = await pickAndSaveGroup(groups, group);
      if (picked) group = picked;
      continue;
    }

    if (action === 'edit-event-name') {
      await editEventName();
      continue;
    }

    if (action === 'edit-event-date') {
      await editEventDate();
      continue;
    }

    if (action === 'edit-source-info') {
      await configureSourceFileAndColumns();
      continue;
    }

    break;
  }

  const existingIds = new Set((group.participants || []).map((p) => p.id?._serialized).filter(Boolean));

  if (action === 'test') {
    console.log('\n=== Test mode (single recipient) ===');
    const tName = (await prompt('Test user name (Enter to skip): ')).trim();
    throwIfSessionLost();
    const tPhoneRaw = await promptRequired('Test phone number (e.g. 0123456789, +60-1162383838, or +1 (415) 555-0101)', '');
    throwIfSessionLost();
    const tPhone = normalizePhoneForWhatsApp(tPhoneRaw, cfg.defaultCountryCode);
    if (!tPhone) {
      console.error('Could not parse the phone number.');
      await client.destroy();
      process.exit(1);
    }
    console.log('');
    await processUser({ name: tName, phone: tPhone }, group, inviteLink, existingIds, null);
    throwIfSessionLost();
    console.log('\nTest complete. Re-run with `npm start` to add everyone from the Excel.');
    await client.destroy();
    process.exit(0);
  }

  // action === 'start' — full Excel run.
  const run = loadUsers();
  const users = run.users;
  for (const item of run.initialStatuses) {
    setExcelStatus(run, item.rowIndex, item.status);
  }
  if (run.initialStatuses.length > 0) saveExcelRun(run);
  console.log(`Loaded ${users.length} unique users from Excel.\n`);
  const log = { added: [], invited: [], skipped: [], notOnWhatsApp: [], failed: [] };
  const total = users.length;
  const startedAt = Date.now();

  for (let i = 0; i < total; i++) {
    if (sessionLoss) break;
    let status;
    try {
      status = await processUser(users[i], group, inviteLink, existingIds, log, { n: i + 1, total });
    } catch (e) {
      if (e !== sessionLoss) throw e;
      break;
    }
    setExcelStatus(run, users[i].rowIndex, status || 'unknown');
    saveExcelRun(run);
    const done = i + 1;
    const elapsedSec = Math.round((Date.now() - startedAt) / 1000);
    console.log(`   progress: added=${log.added.length} invited=${log.invited.length} skipped=${log.skipped.length} not-on-wa=${log.notOnWhatsApp.length} failed=${log.failed.length}  (${done}/${total}, ${elapsedSec}s elapsed)`);
    if (!sessionLoss && i < total - 1) {
      await new Promise((r) => setTimeout(r, cfg.delayMsBetweenAdds || 1500));
    }
  }

  fs.writeFileSync(path.join(__dirname, 'add_log.json'), JSON.stringify(log, null, 2));
  console.log('\n===== Summary =====');
  console.log(`Added direct:    ${log.added.length}`);
  console.log(`Invited via DM:  ${log.invited.length}`);
  console.log(`Already member:  ${log.skipped.length}`);
  console.log(`Not on WhatsApp: ${log.notOnWhatsApp.length}`);
  console.log(`Failed:          ${log.failed.length}`);
  console.log('Detailed log: add_log.json');
  console.log(`Excel updated:   ${cfg.excelFile} (${STATUS_COLUMN})`);

  throwIfSessionLost();
  await client.destroy();
  process.exit(0);
}

// ---------- event wiring ----------

// Two-path strategy:
//   1. Fast path: poll window.Store after authentication (seconds).
//   2. Fallback: wait for the 'ready' event (minutes for large accounts).

let alreadyProceeded = false;
let fastPathStarted = false;
let fullSyncProgressTimer = null;

function startFullSyncProgress() {
  if (fullSyncProgressTimer) return;
  const startedAt = Date.now();
  const estimateMs = 120000;
  fullSyncProgressTimer = setInterval(() => {
    const elapsed = Date.now() - startedAt;
    const pct = Math.min(99, Math.round((elapsed / estimateMs) * 100));
    process.stdout.write(`Full sync progress... ${pct}% (${Math.round(elapsed / 1000)}s elapsed)\r`);
  }, 1000);
}

function stopFullSyncProgress(done = false) {
  if (!fullSyncProgressTimer) return;
  clearInterval(fullSyncProgressTimer);
  fullSyncProgressTimer = null;
  if (done) process.stdout.write('Full sync progress... 100%\n');
}

/*
client.on('authenticated', async () => {
  if (fastPathStarted || alreadyProceeded) return;
  fastPathStarted = true;

  console.log('\nAuthenticated. Fetching groups (fast path)...');

  // Give the page a few seconds to initialise, then start polling Store.
  await new Promise((r) => setTimeout(r, 8000));
  if (alreadyProceeded) return;

  const groups = await waitForGroups(50000); // 50 s max
  if (groups && groups.length > 0 && !alreadyProceeded) {
    alreadyProceeded = true;
    await proceedWithGroups(groups.map(wrapGroup));
    return;
  }

  if (!alreadyProceeded) {
    console.log('\nFast path didn\'t find groups — waiting for full sync...');
  }
});

client.on('ready', async () => {
  if (alreadyProceeded) return;
  alreadyProceeded = true;
  console.log('\nWhatsApp Web fully synced.');
  const chats = await client.getChats();
  await proceedWithGroups(chats.filter((c) => c.isGroup));
});

*/

// ---------- entry point ----------

async function main() {
  printBanner();

  if (SETUP_ONLY) {
    await runSetup();
    console.log('\nNow logging in to WhatsApp to pick the target group...');
    await initClient();
    return;
  }
  if (LIST_ONLY) {
    await initClient();
    return;
  }
  // npm start path — auto-run wizard if essentials are missing, then proceed to add.
  if (!cfg.excelFile || !cfg.phoneColumn || !cfg.nameColumn || !cfg.defaultCountryCode || !cfg.eventName || !cfg.eventDate) {
    console.log('Some config fields are incomplete — running setup first.');
    await runSetup();
    console.log('');
  }
  await initClient();
}

main().catch((err) => {
  console.error(err);
  if (isAuthStartupTimeout(err?.message || String(err))) {
    console.error('\nWhatsApp Web did not finish starting. Your saved login was not reset.'
      + '\nCheck that https://web.whatsapp.com/ opens in your browser. A startup timeout alone is not a reason to reset your saved login.');
  }
  process.exit(1);
});
