// ─── KNNDCmdb Sync Worker ────────────────────────────────────────────────────
// Runs on a separate browser thread. Receives a queue of records from the main
// thread, sends them to Google Sheets via XHR in parallel batches, and reports
// progress back — without blocking the data entry UI at all.
//
// Messages IN  (from main thread):
//   { cmd: 'flush', queue: [...], scriptUrl: '...' }
//   { cmd: 'push',  records: [...], scriptUrl: '...' }
//
// Messages OUT (to main thread):
//   { type: 'progress', uploaded, total }
//   { type: 'done',     uploaded, failed: [...] }
//   { type: 'error',    message }

const BATCH_SIZE = 8;   // larger batches safe on worker — no UI to block
const TIMEOUT_MS = 25000;

// Send one payload to Apps Script via XHR, returns true/false
function sendOne(scriptUrl, payload) {
  return new Promise(resolve => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', scriptUrl, true);
    xhr.timeout   = TIMEOUT_MS;
    xhr.onload    = () => resolve(true);
    xhr.onerror   = () => resolve(false);
    xhr.ontimeout = () => resolve(false);
    xhr.send(JSON.stringify(payload));
  });
}

// Build the correct payload for each queue item type
function buildPayload(item) {
  if (item.type === 'delete') return { action: 'deleteMember', ...item.data };
  if (item.type === 'update') return { action: 'updateMember', ...item.data };
  return { action: 'upsertMember', ...item.data };
}

// Process an array of items in parallel batches, posting progress updates
async function processItems(scriptUrl, items) {
  const failed   = [];
  let   uploaded = 0;
  const total    = items.length;

  for (let i = 0; i < total; i += BATCH_SIZE) {
    const batch    = items.slice(i, i + BATCH_SIZE);
    const payloads = batch.map(buildPayload);
    const results  = await Promise.all(payloads.map(p => sendOne(scriptUrl, p)));

    results.forEach((ok, j) => {
      if (ok) { uploaded++; }
      else    { failed.push(batch[j]); }
    });

    // Report progress after each batch so the main thread can update the button
    self.postMessage({ type: 'progress', uploaded, total });
  }

  return { uploaded, failed };
}

// Message handler
self.onmessage = async (e) => {
  const { cmd, scriptUrl } = e.data;

  if (!scriptUrl) {
    self.postMessage({ type: 'error', message: 'No Script URL provided.' });
    return;
  }

  try {
    if (cmd === 'flush') {
      const queue = e.data.queue || [];
      if (!queue.length) {
        self.postMessage({ type: 'done', uploaded: 0, failed: [] });
        return;
      }
      const { uploaded, failed } = await processItems(scriptUrl, queue);
      self.postMessage({ type: 'done', uploaded, failed });
    }

    else if (cmd === 'push') {
      const records = (e.data.records || []).map(m => ({ type: 'add', data: m }));
      if (!records.length) {
        self.postMessage({ type: 'done', uploaded: 0, failed: [] });
        return;
      }
      const { uploaded, failed } = await processItems(scriptUrl, records);
      self.postMessage({ type: 'done', uploaded, failed });
    }

  } catch (err) {
    self.postMessage({ type: 'error', message: err.message });
  }
};
