// Smart Reply v3.1 — Outlook Add-in
// Reads current email, calls v3.1 flow, opens reply with AI-generated body + selected URLs

const V31_ENDPOINT = "https://852cb0da7269e604b176df59b67472.1c.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e01367d97a65415ca2e4b29a436bfb71/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=UHevbE4iKfL2wqlAxYv-10pcbfYu4WArj6J5Dfqr4gQ";

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("generate-btn").onclick = handleGenerate;
  }
});

async function handleGenerate() {
  const btn = document.getElementById("generate-btn");
  const listingName = document.getElementById("listing-select").value;
  const includeVdr = document.getElementById("include-vdr").checked;
  const includeOm = document.getElementById("include-om").checked;

  btn.disabled = true;
  showStatus("Reading email content...", "loading");

  try {
    // Read current email content
    const emailData = await readEmail();

    showStatus("Generating AI reply with Dataverse + AI Builder...", "loading");

    // Call the v3.1 flow
    const response = await fetch(V31_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        listingName: listingName,
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

    // Build reply HTML — only include URLs the broker chose
    const replyHtml = formatReplyBody(data, { includeVdr: includeVdr, includeOm: includeOm });

    // Open Outlook reply form
    Office.context.mailbox.item.displayReplyForm({
      htmlBody: replyHtml
    });

    showStatus("✅ Reply draft opened. Review before sending!", "success");
  } catch (err) {
    console.error(err);
    showStatus("❌ " + err.message, "error");
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
  const listingName = data.listingName || "this listing";

  // Convert AI body's line breaks to <br> for HTML
  const aiBodyHtml = escapeHtml(aiBody).replace(/\n/g, "<br>");

  // Build the URL section only if at least one is checked AND we have URLs
  const showVdr = opts.includeVdr && vdrAccess;
  const showOm = opts.includeOm && omUrl;

  let linksHtml = "";
  if (showVdr || showOm) {
    linksHtml = '<p style="margin-top: 16px;"><strong>Materials for ' + escapeHtml(listingName) + ':</strong><br>';
    if (showVdr) {
      linksHtml += '📋 <a href="' + vdrAccess + '">Virtual Deal Room (sign CA to access full materials)</a>';
      if (showOm) linksHtml += '<br>';
    }
    if (showOm) {
      linksHtml += '📄 <a href="' + omUrl + '">Offering Memorandum (direct PDF)</a>';
    }
    linksHtml += '</p>';
  }

  return (
    '<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #222;">' +
    '<div style="background: #fff3cd; border-left: 4px solid #ffc107; padding: 10px 14px; margin-bottom: 18px; color: #856404; font-size: 10pt;">' +
    '<strong>⚠️ AI-GENERATED DRAFT</strong><br>' +
    'Review and edit before sending.' +
    '</div>' +
    '<p>' + aiBodyHtml + '</p>' +
    linksHtml +
    '</div>'
  );
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function showStatus(message, type) {
  const status = document.getElementById("status");
  status.textContent = message;
  status.className = "status " + (type || "");
  status.classList.remove("hidden");
}
