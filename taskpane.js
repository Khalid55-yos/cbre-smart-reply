// Smart Reply v3.2 — Outlook Add-in
// Dropdown now sends SharePoint Item ID (more reliable than Name)

const GET_LISTINGS_ENDPOINT = "https://852cb0da7269e604b176df59b67472.1c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7fc7fb5aa81947908488ccf927345ed1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=bfZCgAT0cnWLXMuels1niYu7edMMOLxrirCEGjGnKrk";

const V31_ENDPOINT = "https://852cb0da7269e604b176df59b67472.1c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e01367d97a65415ca2e4b29a436bfb71/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=UHevbE4iKfL2wqlAxYv-10pcbfYu4WArj6J5Dfqr4gQ";

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("generate-btn").onclick = handleGenerate;
    loadListings();
  }
});

async function loadListings() {
  const select = document.getElementById("listing-select");
  const btn = document.getElementById("generate-btn");

  select.innerHTML = '<option value="">Loading listings...</option>';
  select.disabled = true;
  btn.disabled = true;

  try {
    const response = await fetch(GET_LISTINGS_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({})
    });

    if (!response.ok) {
      throw new Error("HTTP " + response.status);
    }

    const data = await response.json();

    let listings = [];
    if (Array.isArray(data)) {
      listings = data;
    } else if (data && Array.isArray(data.listings)) {
      listings = data.listings;
    } else if (data && Array.isArray(data.value)) {
      listings = data.value;
    } else {
      throw new Error("Unexpected response format");
    }

    // Sort by address (fallback to name) since names may be empty
    listings.sort(function (a, b) {
      const aKey = ((a.crcce_address || a.crcce_newcolumn || "")).toLowerCase();
      const bKey = ((b.crcce_address || b.crcce_newcolumn || "")).toLowerCase();
      return aKey.localeCompare(bKey);
    });

    if (listings.length === 0) {
      select.innerHTML = '<option value="">No listings found</option>';
      return;
    }

    let html = '<option value="">Select a listing...</option>';
    listings.forEach(function (listing) {
      const itemId = listing.crcce_sharepointitemid || "";
      const name = listing.crcce_newcolumn || "";
      const address = listing.crcce_address || "";

      // Skip if no Item ID (can't be matched)
      if (!itemId) return;

      // Display: prefer "Name — Address", but dedupe if they match (case-insensitive)
      let display;
      const nameTrim = name.trim();
      const addrTrim = address.trim();
      if (nameTrim && addrTrim && nameTrim.toLowerCase() !== addrTrim.toLowerCase()) {
        display = nameTrim + " — " + addrTrim;
      } else if (addrTrim) {
        display = addrTrim;
      } else if (nameTrim) {
        display = nameTrim;
      } else {
        display = "Listing " + itemId;
      }

      html += '<option value="' + escapeHtmlAttr(itemId) + '">' + escapeHtmlText(display) + '</option>';
    });

    select.innerHTML = html;
    select.disabled = false;
    btn.disabled = false;
  } catch (err) {
    console.error("Failed to load listings:", err);
    select.innerHTML =
      '<option value="">Could not load \u2014 retry</option>';
    select.disabled = false;
    btn.disabled = false;
    showStatus("\u26a0\ufe0f Couldn't load listings: " + err.message, "error");
  }
}

async function handleGenerate() {
  const btn = document.getElementById("generate-btn");
  const select = document.getElementById("listing-select");
  const listingId = select.value;
  // Display label for the chosen listing (used in the reply body)
  const listingLabel = select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
  const includeVdr = document.getElementById("include-vdr").checked;
  const includeOm = document.getElementById("include-om").checked;

  if (!listingId) {
    showStatus("\u26a0\ufe0f Pick a listing first", "error");
    return;
  }

  btn.disabled = true;
  showStatus("Reading email content...", "loading");

  try {
    const emailData = await readEmail();

    showStatus("Generating AI reply...", "loading");

    const response = await fetch(V31_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        // listingName field now carries the SharePoint Item ID (v3.1 filter uses Item ID)
        listingName: listingId,
        emailFrom: emailData.from,
        emailSubject: emailData.subject,
        emailBody: emailData.body
      })
    });

    if (!response.ok) {
      throw new Error("Flow returned HTTP " + response.status);
    }

    const data = await response.json();

    showStatus("Opening reply...", "loading");

    // Use dropdown label for display since data.listingName is now an Item ID
    const replyHtml = formatReplyBody(data, {
      includeVdr: includeVdr,
      includeOm: includeOm,
      listingLabel: listingLabel
    });

    Office.context.mailbox.item.displayReplyForm({
      htmlBody: replyHtml
    });

    showStatus("\u2705 Reply draft opened. Review before sending!", "success");
  } catch (err) {
    console.error(err);
    showStatus("\u274c " + err.message, "error");
  } finally {
    btn.disabled = false;
  }
}

function readEmail() {
  return new Promise(function (resolve, reject) {
    const item = Office.context.mailbox.item;

    const subject = item.subject || "";
    const from = item.from
      ? (item.from.displayName || "") + " <" + item.from.emailAddress + ">"
      : "";

    item.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({
          subject: subject,
          from: from,
          body: result.value
        });
      } else {
        reject(new Error("Could not read email body: " + (result.error && result.error.message)));
      }
    });
  });
}

function formatReplyBody(data, opts) {
  opts = opts || {};
  const aiBody = (data.body || "").trim();
  const omUrl = data.omUrl || "";
  const vdrAccess = data.vdrAccess || "";
  // Prefer the dropdown label (human readable) over the raw Item ID returned from the flow
  const listingName = opts.listingLabel || data.listingName || "this listing";

  const aiBodyHtml = escapeHtmlText(aiBody).replace(/\n/g, "<br>");

  const showVdr = opts.includeVdr && vdrAccess;
  const showOm = opts.includeOm && omUrl;

  let linksHtml = "";
  if (showVdr || showOm) {
    linksHtml = '<p style="margin-top: 16px;"><strong>Materials for ' + escapeHtmlText(listingName) + ':</strong><br>';
    if (showVdr) {
      linksHtml += '\ud83d\udccb <a href="' + vdrAccess + '">Virtual Deal Room (sign CA to access full materials)</a>';
      if (showOm) linksHtml += '<br>';
    }
    if (showOm) {
      linksHtml += '\ud83d\udcc4 <a href="' + omUrl + '">Offering Memorandum (direct PDF)</a>';
    }
    linksHtml += '</p>';
  }

  return (
    '<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #222;">' +
    '<div style="background: #fff3cd; border-left: 4px solid #ffc107; padding: 10px 14px; margin-bottom: 18px; color: #856404; font-size: 10pt;">' +
    '<strong>\u26a0\ufe0f AI-GENERATED DRAFT</strong><br>' +
    'Review and edit before sending.' +
    '</div>' +
    '<p>' + aiBodyHtml + '</p>' +
    linksHtml +
    '</div>'
  );
}

function escapeHtmlText(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeHtmlAttr(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function showStatus(message, type) {
  const status = document.getElementById("status");
  status.textContent = message;
  status.className = "status " + (type || "");
  status.classList.remove("hidden");
}
