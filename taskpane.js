/**
 * CBRE Smart Reply - Outlook Add-in
 * v3.0 - Personal sideload
 *
 * Flow:
 *  1. User clicks "Smart Reply" button on a read email
 *  2. Taskpane opens, reads the email's subject/body/sender
 *  3. Calls the Power Automate HTTP endpoint with email context
 *  4. Flow returns: detected listing + reply body + OM file (base64)
 *  5. User reviews preview, clicks "Open Reply with OM Attached"
 *  6. Outlook opens a reply form pre-populated with body and attached OM
 */

// ============================================================================
// CONFIGURATION — UPDATE BEFORE DEPLOYING
// ============================================================================

// Replace this with your Power Automate HTTP trigger URL after creating the flow
const FLOW_URL = 'PASTE_YOUR_POWER_AUTOMATE_HTTP_TRIGGER_URL_HERE';

// ============================================================================
// STATE
// ============================================================================

let detectedListing = null;
let allListings = [];
let currentReplyBody = '';
let currentOmFilename = '';
let currentOmBase64 = '';

// ============================================================================
// OFFICE READY
// ============================================================================

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();
  }
});

async function initialize() {
  try {
    setStatus('Reading email context…');

    const emailContext = await readEmailContext();

    setStatus('Looking up listing in SharePoint…');
    const flowResponse = await callFlow(emailContext);

    handleFlowResponse(flowResponse);
  } catch (err) {
    showError('Could not initialize: ' + (err.message || err));
  }
}

// ============================================================================
// READ EMAIL CONTEXT
// ============================================================================

function readEmailContext() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error('No email item available.'));
      return;
    }

    const subject = item.subject || '';
    const senderEmail = (item.from && item.from.emailAddress) || '';
    const senderName = (item.from && item.from.displayName) || '';

    // Body comes async
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({
          subject: subject,
          senderEmail: senderEmail,
          senderName: senderName,
          body: result.value || ''
        });
      } else {
        // Even if body fails, we can still try with subject only
        resolve({
          subject: subject,
          senderEmail: senderEmail,
          senderName: senderName,
          body: ''
        });
      }
    });
  });
}

// ============================================================================
// CALL POWER AUTOMATE FLOW
// ============================================================================

async function callFlow(emailContext) {
  if (!FLOW_URL || FLOW_URL.includes('PASTE_YOUR')) {
    throw new Error('Flow URL not configured. Edit taskpane.js and set FLOW_URL.');
  }

  const response = await fetch(FLOW_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      subject: emailContext.subject,
      body: emailContext.body.substring(0, 4000),
      senderEmail: emailContext.senderEmail,
      senderName: emailContext.senderName
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Flow returned ${response.status}: ${errText.substring(0, 200)}`);
  }

  return await response.json();
}

// ============================================================================
// HANDLE RESPONSE
// ============================================================================

function handleFlowResponse(data) {
  // Expected response shape:
  // {
  //   detected: boolean,
  //   confidence: 'high' | 'low' | 'none',
  //   listingName: 'Wiley Canyon',
  //   replyBody: '...',
  //   omFilename: 'Santa Clarita_...pdf',
  //   omBase64: 'base64encoded...',
  //   allListings: [{ name, filename }, ...]   // for dropdown override
  // }

  if (!data) {
    showError('Empty response from flow.');
    return;
  }

  allListings = data.allListings || [];
  detectedListing = data.listingName || '';
  currentReplyBody = data.replyBody || '';
  currentOmFilename = data.omFilename || '';
  currentOmBase64 = data.omBase64 || '';

  populateListingDropdown(allListings, detectedListing);

  if (data.detected && data.listingName) {
    setStatus(`Detected: ${data.listingName}`, 'success');
    showDetection(`Auto-detected from email subject. Change above if wrong.`, 'detected');
    showPreview(currentReplyBody, currentOmFilename);
    showInsertButton();
  } else {
    setStatus('No listing auto-detected.', 'warning');
    showDetection('Pick the listing manually from the dropdown above.', 'not-detected');
  }
}

// ============================================================================
// UI HELPERS
// ============================================================================

function setStatus(message, type) {
  const statusCard = document.getElementById('status-card');
  const statusText = document.getElementById('status-text');
  const statusDot = statusCard.querySelector('.status-dot');

  statusText.textContent = message;

  if (type === 'success') {
    statusCard.style.borderLeftColor = '#17E88F';
    statusDot.style.background = '#17E88F';
    statusDot.style.animation = 'none';
  } else if (type === 'warning') {
    statusCard.style.borderLeftColor = '#f59e0b';
    statusDot.style.background = '#f59e0b';
    statusDot.style.animation = 'none';
  } else if (type === 'error') {
    statusCard.style.borderLeftColor = '#dc2626';
    statusDot.style.background = '#dc2626';
    statusDot.style.animation = 'none';
  }
}

function populateListingDropdown(listings, selectedName) {
  const select = document.getElementById('listing-select');
  select.innerHTML = '';

  if (listings.length === 0) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '(no listings available)';
    select.appendChild(opt);
    return;
  }

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '— pick a listing —';
  select.appendChild(placeholder);

  listings.forEach(l => {
    const opt = document.createElement('option');
    opt.value = l.name;
    opt.textContent = l.name;
    opt.dataset.filename = l.filename || '';
    if (selectedName && selectedName.toLowerCase() === l.name.toLowerCase()) {
      opt.selected = true;
    }
    select.appendChild(opt);
  });

  // When user changes, re-fetch reply body + OM for that listing
  select.addEventListener('change', onListingChange);

  document.getElementById('detection-card').classList.remove('hidden');
}

async function onListingChange() {
  const select = document.getElementById('listing-select');
  const chosenName = select.value;

  if (!chosenName) {
    document.getElementById('preview-card').classList.add('hidden');
    document.getElementById('insert-button').classList.add('hidden');
    return;
  }

  setStatus(`Fetching ${chosenName}…`);

  try {
    const response = await fetch(FLOW_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        listingNameOverride: chosenName,
        subject: '',
        body: '',
        senderEmail: '',
        senderName: ''
      })
    });

    if (!response.ok) throw new Error(`Flow error ${response.status}`);

    const data = await response.json();
    currentReplyBody = data.replyBody || '';
    currentOmFilename = data.omFilename || '';
    currentOmBase64 = data.omBase64 || '';

    showPreview(currentReplyBody, currentOmFilename);
    showInsertButton();
    setStatus(`Loaded: ${chosenName}`, 'success');
  } catch (err) {
    showError('Could not load that listing: ' + err.message);
  }
}

function showDetection(text, cls) {
  const hint = document.getElementById('detection-hint');
  hint.textContent = text;
  hint.className = 'hint ' + (cls || '');
}

function showPreview(body, filename) {
  document.getElementById('preview-card').classList.remove('hidden');
  document.getElementById('preview-body').textContent = body;
  document.getElementById('attachment-name').textContent = filename || 'No file';
}

function showInsertButton() {
  const btn = document.getElementById('insert-button');
  btn.classList.remove('hidden');
  btn.onclick = onInsertReply;
}

function showError(message) {
  const card = document.getElementById('error-card');
  const text = document.getElementById('error-text');
  text.textContent = message;
  card.classList.remove('hidden');
  setStatus('Error', 'error');
}

// ============================================================================
// INSERT REPLY WITH ATTACHMENT
// ============================================================================

function onInsertReply() {
  const btn = document.getElementById('insert-button');
  btn.disabled = true;
  btn.textContent = 'Opening reply…';

  // Convert plain body to HTML so line breaks render correctly
  const htmlBody = currentReplyBody
    .split('\n')
    .map(line => line.trim() === '' ? '<br>' : `<div>${escapeHtml(line)}</div>`)
    .join('');

  const replyOptions = {
    htmlBody: htmlBody
  };

  // Only attach if we have an OM file
  if (currentOmBase64 && currentOmFilename) {
    replyOptions.attachments = [
      {
        type: 'file',
        name: currentOmFilename,
        inLine: false,
        // Office.js expects the URL/content as a string. For base64, use:
        url: 'data:application/pdf;base64,' + currentOmBase64,
        isInline: false
      }
    ];
  }

  // Use displayReplyAllForm if the original was sent to multiple people,
  // otherwise displayReplyForm. We use displayReplyAllForm to be safe.
  Office.context.mailbox.item.displayReplyAllForm(replyOptions);

  btn.disabled = false;
  btn.textContent = 'Open Reply with OM Attached';
  setStatus('Reply opened. Review and send from Outlook.', 'success');
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
